const fs = require("fs/promises");
const path = require("path");
const fetch = require("node-fetch");
const yargs = require('yargs/yargs')
const { hideBin } = require('yargs/helpers');
const mkdirp = require("mkdirp");
const { isString } = require("util");

const argv = yargs(hideBin(process.argv))
    .option('subject', {
        type: 'string',
        description: 'Email subject',
        required: true
    })
    .option('dir', {
        type: 'string',
        description: 'Destination directory'
    })
    .option('inbox', {
        type: "boolean",
        description: "Process messages from Inbox?",
        default: false
    })
    .option('cleanup', {
        type: "boolean",
        description: "Cleanup after download?",
        default: true
    })
    .option('limit', {
        type: "number",
        description: "Process X messages",
        default: 0
    })
    .argv;

async function main() {
    // Load the auth token
    var token = (await fs.readFile("token.txt")).toString();
    const dataDir = argv.dir || path.join(process.cwd(), "data");
    await mkdirp(dataDir);

    // Find the Inbox folder ids
    const inboxFolderId = await getFolderId("Inbox", token);
    const sumoFolderId = await getChildFolderId(inboxFolderId, "SumoLogic", token);
    const downloadedFolderId = await getChildFolderId(sumoFolderId, "Downloaded", token);

    const sourceFolder = argv.inbox ? inboxFolderId : sumoFolderId;   

    // Start pulling emails
    let emailsUrl = `https://graph.microsoft.com/v1.0/me/mailFolders/${sourceFolder}/messages?$filter=subject eq '${argv.subject}'`;

    let stop = false;
    let count = 0;

    do {
        const processedEmailIds = [];
        const emailsResponse = await loadEmails(emailsUrl, token);

        // For each email, download the attachment
        for(email of emailsResponse.value) {
            const attachments = await getAttachments(email.id, token);

            if (attachments.value.length != 1) {
                throw new Error("Wrong number of attachments, " + email.id);
            }

            const attachment = attachments.value[0];
            const data = Buffer.from(attachment.contentBytes, 'base64');
            console.log("Writing file", attachment.name, "...");
            const filePath = path.join(dataDir, attachment.name);
            const fileDate = new Date(email.receivedDateTime);
            await fs.writeFile(filePath, data);
            await fs.utimes(filePath, fileDate, fileDate);

            processedEmailIds.push(email.id);

            count++;
            stop = argv.limit > 0 && count >= argv.limit;
            if (stop) break;
        };

        // Move processed emails to folder.
        if (argv.cleanup) {
            for (processedEmailId of processedEmailIds) {
                await moveEmailToFolder(processedEmailId, downloadedFolderId, token);
            }
        }

        emailsUrl = emailsResponse["@odata.nextLink"];
    } while (!stop && typeof emailsUrl != "undefined");

    console.log("Done!");
}

main();

// ====================================

async function loadEmails(url, token) {

    console.log("Fetching emails...");
    return await callApi(url, token);
}

async function getAttachments(id, token) {
    const url = `https://graph.microsoft.com/v1.0/me/messages/${id}/attachments`;

    console.log("Fetching attachments...");
    return await callApi(url, token);
}

async function callApi(url, options) {
    // Options can either be the token or an options object
    if (typeof options === "string") {
        options = {
            headers: {
                Authorization: "Bearer " + options
            }
        };
    }

    const response = await fetch(url, options);

    if (!response.ok) {
        if (response.status == 429) { // Too many requests 
            const retryAfter = response.headers.get("Retry-After");
            console.log("'Too many requests' received. Waiting", retryAfter);
            console.log(Date.now());
            await new Promise(resolve => setTimeout(resolve, parseInt(retryAfter) * 1000));
            console.log(Date.now());
            return await callApi(url, token);
        } else {
            console.error("ERROR", url, response);
            throw new Error("Failure while calling API");
        }
    } else {
        const json = await response.json();
        return json;
    }
}

async function getFolderId(name, token) {
    console.log(`Fetching folder ${name} ...`);
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders?filter=displayName eq '${name}'`;
    const rootFolder = await callApi(url, token);
    return rootFolder.value[0].id;
}


async function getChildFolderId(parentId, name, token) {
    console.log(`Fetching child folder ${name} ...`);
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${parentId}/childFolders?filter=displayName eq '${name}'`;
    const rootFolder = await callApi(url, token);
    return rootFolder.value[0].id;
}

async function moveEmailToFolder(emailId, folderId, token) {
    console.log("Moving email...");
    const url = `https://graph.microsoft.com/v1.0/me/messages/${emailId}/move`;
    await callApi(url, {
        method: "POST",
        body: JSON.stringify({
            destinationId: folderId
        }),
        headers: {
            Authorization: "Bearer " + token,
            "Content-type": "application/json"
        }
    });
}
