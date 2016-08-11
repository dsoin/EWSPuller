# EWSPuller
<h2> Microsoft Exchange to ES email puller</h2>

Gets emails off the folder in MS Exchange and pushes them to ES index. Supports attachments.
Can read .pst files to get your archives to ES index as well. 
Runs in background to get the emails pulled automatically (see -interval options)
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

<h1>License</h1>

MIT License

Copyright (c) 2016 Dmitrii Soin

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
