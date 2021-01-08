const fs = require("fs/promises");
const path = require("path");
const fetch = require("node-fetch");
const yargs = require('yargs/yargs')
const { hideBin } = require('yargs/helpers')

const argv = yargs(hideBin(process.argv))
    .option('subject', {
        type: 'string',
        description: 'Email subject',
        required: true
    })
    .argv;

async function main() {
    // Load the auth token
    var token = (await fs.readFile("token.txt")).toString();

    // Start pulling emails
    let emailsUrl = `https://graph.microsoft.com/v1.0/me/messages?$filter=subject eq '${argv.subject}'`;

    do {
        const emailsResponse = await loadEmails(emailsUrl, token);

        // For each email, download the attachment
        for(email of emailsResponse.value) {
            const attachments = await getAttachments(email.id, token);

            if (attachments.value.length != 1) {
                throw new Error("Wrong number of attachments, " + email.id);
            }

            const attachment = attachments.value[0];
            const data = Buffer.from(attachment.contentBytes, 'base64');
            await fs.writeFile(path.join(process.cwd(), "data", attachment.name), data);
        };

        emailsUrl = emailsResponse["@odata.nextLink"];
    } while (typeof emailsUrl != "undefined");
    
}

main();

async function loadEmails(url, token) {

    console.log("Fetching emails...");
    return await callApi(url, token);
}

async function getAttachments(id, token) {
    const url = `https://graph.microsoft.com/v1.0/me/messages/${id}/attachments`;

    console.log("Fetching attachments...");
    return await callApi(url, token);
}

async function callApi(url, token) {
    const response = await fetch(url, {
        headers: {
            Authorization: "Bearer " + token
        }
    });

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