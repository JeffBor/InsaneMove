# InsaneMove

## Description

ShareGate offers “insane mode” which uploads to Azure blob storage for fast cloud migration to Office 365 (<https://en.share-gate.com/sharepoint-migration/insane-mode>).     It’s an excellent program for copying SharePoint sites.   However, I wanted to research ways to run that even faster by leveraging parallel processing and came up with “Insane MOVE.”

[![image](https://raw.githubusercontent.com/spjeff/InsaneMove/master/doc/download.png)](https://github.com/spjeff/InsaneMove/releases/download/InsaneMove/InsaneMove.zip)

## Key Features

* Migrate SharePoint sites to Office 365
* Bulk CSV of source/destination URLs
* Powershell `InsaneMOVE.ps1` runs on any 1 server in the farm
* Auto detects *SharePoint* servers with ShareGate Desktop installed
* *Optional list of specific servers on which migration jobs will run.  Requires ShareGate Desktop installed.*
* Opens remote PowerShell to each server
* Creates secure string for passwords (per server)
* Creates Task Scheduler job local on each server
* Creates “worker*X*.ps1” file on each server to run one copy job
* Automatic queuing to start new jobs as current ones complete
* Status report to CSV with ShareGate session ID#, errors, warnings, and error detail XML  (if applicable)
* LOG for both “InsaneMove” centrally and each remote worker PS1

## Requirements

* ShareGate Desktop installed, licensed on remote servers
* Module Microsoft.Online.SharePoint.PowerShell
* Module SharePointPnPPowerShellOnline
* Module CredentialManager
* PowerShell Remoting from control instance to remote servers with ShareGate

## Quick Start

1. Clone or Download `InsaneMove.zip` and extract
1. Populate `wave.csv` with source/destination URLs
1. Run `InsaneMove.ps1 -v wave.csv` to verify all destination site collection exists (and will create if missing)
1. Run `InsaneMove.ps1 wave.csv` to begin copy jobs across remote servers
1. Sit back and enjoy!

## Parameters

* `[string]$fileCSV` > CSV list of source and destination SharePoint site URLs to copy to Office 365.
* `-v[switch]$verifyCloudSites` > Verify all Office 365 site collections.  Prep step before real migration.
* `-i[switch]$incremental` > Copy incremental changes only. <http://help.share-gate.com/article/443-incremental-copy-copy-sharepoint-content>
* `-m[switch]$measure` > Measure size of site collections in GB.
* `-e[switch]$email` > Send email notifications with summary of migration batch progress.
* `-ro[switch]$readOnly` > Lock sites read-only.
* `-rw[switch]$readWrite` > Unlock sites read-write.
* `-sca[switch]$siteCollectionAdmin` > Grant Site Collection Admin rights to the migration user specified in XML settings file.

## Screenshots

![image](https://raw.githubusercontent.com/spjeff/InsaneMove/master/doc/rocket.png)

![image](https://raw.githubusercontent.com/spjeff/InsaneMove/master/doc/diagram.png)

![image](https://raw.githubusercontent.com/spjeff/InsaneMove/master/doc/1.png)

![image](https://raw.githubusercontent.com/spjeff/InsaneMove/master/doc/2.png)

![image](https://raw.githubusercontent.com/spjeff/InsaneMove/master/doc/3.png)

![image](https://raw.githubusercontent.com/spjeff/InsaneMove/master/doc/4.png)

![image](https://raw.githubusercontent.com/spjeff/InsaneMove/master/doc/5.png)

## Contact

Contact @JeffBor <https://github.com/JeffBor>

## Credits

Forked from [@spjeff](https://twitter.com/spjeff) or [spjeff@spjeff.com](mailto:spjeff@spjeff.com)

## License

The MIT License (MIT)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
