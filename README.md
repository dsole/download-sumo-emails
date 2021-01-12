# Outlook email attachments downloader

Simple script to download the attachments from a set or reoccuring emails.
I get data files (CSV) emailed to be regularly and wanted a way to save all the attachments to a folder. Unforatunately Outlook Rules don't support this.

By default the script will look in the folder Inbox > Sumologic , save the files to a directory called "data", and clean up (move) the emails to Inbox > Sumologic > Downloaded. It processes the emails in batches.

## Usage

Paste an access token in  the token.txt file before running the script.

Then run with :

```
node index.js --subject "email subject here"
```

## Help

(Generated using `node index.js --help` )

```
Options:
  --help     Show help                                                 [boolean]
  --version  Show version number                                       [boolean]
  --subject  Email subject                                   [string] [required]
  --dir      Destination directory                                      [string]
  --inbox    Process messages from Inbox?             [boolean] [default: false]
  --cleanup  Cleanup after download?                   [boolean] [default: true]
  --limit    Process X messages                            [number] [default: 0]
```

