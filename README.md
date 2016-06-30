# EWSPuller
<h2> Microsoft Exchange to ES email puller</h2>

Gets emails off the folder in MS Exchange and pushes them to ES index. Supports attachments.
Can read .pst files to get your archives to ES index as well.
Use [Outlook4All app](https://github.com/dsoin/outlook4all) for searching over indexed emails.

Usage:
* -esport N    : ES port (default: 9300)
* -ews_url VAL : Exchange/365 EWS URL, defaults to public 365 URL (default:
                https://outlook.office365.com/EWS/Exchange.asmx)
* -folder VAL  : Exchange Folder to pull emails from. *Note emails will be
                deleted after syncing (default: )
* -interval N  : Pull interval in minutes (default: 60)
* -pass VAL    : Exchange password. Will be read from console if omitted.
                (default: )
* -pst VAL     : PST file to index (default: )
* -user VAL    : Exchange username/email (default: )
