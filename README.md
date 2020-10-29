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
* QA checking and reporting.
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
1. Prepare the ShareGate user mapping file `\insanemove\usermap.sgum`
1. Run `InsaneMove.ps1 -v wave.csv` to verify all destination site collection exists (and will create if missing)
1. Run `InsaneMove.ps1 wave.csv` to begin copy jobs across remote servers
1. Sit back and enjoy!

## Parameters

    -fileCSV <String>
    CSV list of source and destination SharePoint site URLs to copy to Office 365.
        Required?                    false
        Position?                    1
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -verifyCloudSites [<SwitchParameter>]
    Verify all Office 365 site collections.  Prep step before real migration.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -verifyWiki [<SwitchParameter>]
    Verify Wiki Libraries exist on Office 365 sites.  After site collections created OK (-verify).
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -incremental [<SwitchParameter>]
    Copy incremental changes only. http://help.share-gate.com/article/443-incremental-copy-copy-sharepoint-content
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -measure [<SwitchParameter>]
    Measure size of site collections in GB.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -readOnly [<SwitchParameter>]
    Lock sites read-only.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -readWrite [<SwitchParameter>]
    Unlock sites read-write.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -noAccess [<SwitchParameter>]
    Lock sites no access.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -userProfile [<SwitchParameter>]
    Update local User Profile Service with cloud personal URL.  Helps with Hybrid Onedrive audience rules.  Need to recompile audiences after running this.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -dryRun [<SwitchParameter>]
    Dry run replaces core "Copy-Site" with "NoCopy-Site" to execute all queueing but not transfer any data.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -clean [<SwitchParameter>]
    Clean servers to prepare for next migration batch.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -deleteSource [<SwitchParameter>]
    Delete source SharePoint sites on-premise.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -deleteDest [<SwitchParameter>]
    Delete destination SharePoint sites in cloud
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -qualityAssurance [<SwitchParameter>]
    Compare source and destination lists for QA check.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -migrate [<SwitchParameter>]
    Copy sites to Office 365.  This is the default method.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -whatif [<SwitchParameter>]
    Pre-Migration Report.  Runs Copy-Site with -WhatIf.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -mini [<SwitchParameter>]
    Leverage different narrow set of servers. MINI line from XML input file.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -prepSource [<SwitchParameter>]
    Prep source by Allow Multi Response on Survey and Update URL metadata fields with "rootfolder" on the source.  Replace with O365 compatible shorter URL.
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

## Screenshots

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
