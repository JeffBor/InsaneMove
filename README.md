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
    Prep source by Allow Multi Response on Survey and Update URL metadata fields with "rootfolder" on the source.  Replace with M365-compatible shorter URL.
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

## Settings File

A settings file is read for persistent settings:

```xml

<?xml version="1.0" standalone="yes"?>
<settings>
	<tenant>
		<adminURL>https://tenant-admin.sharepoint.com</adminURL>
		<adminUser>spadmin@tenant.onmicrosoft.com</adminUser>
		<adminRole>c:0-.f|rolemanager|s-1-5-21-1260933108-580135205-188289138-1234567</adminRole>
		<adminPass>pass@word1</adminPass>
		<suffix>tenant.onmicrosoft.com</suffix>
		<uploadUsers>spmigrate1@tenant.onmicrosoft.com,spmigrate2@tenant.onmicrosoft.com,spmigrate3@tenant.onmicrosoft.com,spmigrate4@tenant.onmicrosoft.com,spmigrate5@tenant.onmicrosoft.com,spmigrate6@tenant.onmicrosoft.com,spmigrate7@tenant.onmicrosoft.com,spmigrate8@tenant.onmicrosoft.com,spmigrate9@tenant.onmicrosoft.com,spmigrate10@tenant.onmicrosoft.com,spmigrate11@tenant.onmicrosoft.com,spmigrate12@tenant.onmicrosoft.com</uploadUsers>
		<uploadPass>pass@word1</uploadPass>
		<timezoneId>11</timezoneId>
	</tenant>
	<maxWorker>1</maxWorker>
	<notify>
		<smtpServer>mailrelay</smtpServer>
		<from>migration-no-reply@company.com</from>
		<to>sharepoint_team@company.com</to>
	</notify>
	<optionalLimitServers>wfe120,wfe121,wfe122,wfe220,wfe221,wfe222</optionalLimitServers>
	<optionalLimitServersMini>wfe121,wfe220,wfe222</optionalLimitServersMini>
	<optionalSchtaskUser>spservice</optionalSchtaskUser>
	<sourceFarmServer>wfe05</sourceFarmServer>
</settings>
```

### timezoneID Setting

When tenant site collection is created in the destination, a timezone ID must be specified.  Use the following to set the TimeZoneID value in the settings file:

Id|Description|                                        Identifier
--|-----------|                                        ----------
 0|None|                                               No:ne
 2|GREENWICH MEAN TIME DUBLIN EDINBURGH LISBON LONDON ||UTC
 3|BRUSSELS COPENHAGEN MADRID PARIS                   ||UTC+01:00
 4|AMSTERDAM BERLIN BERN ROME STOCKHOLM VIENNA        ||UTC+01:00
 5|ATHENS BUCHAREST ISTANBUL                          ||UTC+02:00
 6|BELGRADE BRATISLAVA BUDAPEST LJUBLJANA PRAGUE      ||UTC+01:00
 7|MINSK                                              ||UTC+02:00
 8|BRASILIA                                           ||UTC-03:00
 9|ATLANTIC TIME CANADA                               ||UTC-04:00
10|EASTERN TIME US AND CANADA                         ||UTC-05:00
11|CENTRAL TIME US AND CANADA                         ||UTC-06:00
12|MOUNTAIN TIME US AND CANADA                        ||UTC-07:00
13|PACIFIC TIME US AND CANADA                         ||UTC-08:00
14|ALASKA                                             ||UTC-09:00
15|HAWAII                                             ||UTC-10:00
16|MIDWAY ISLAND SAMOA                                ||UTC-11:00
17|AUKLAND WELLINGTON                                 ||UTC+12:00
17|AUKLAND WELLINGTON                                 ||UTC+12:00
18|BRISBANE                                           ||UTC+10:00
19|ADELAIDE                                           ||UTC+09:30
20|OSAKA SAPPORO TOKYO                                ||UTC+09:00
21|KUALA LUMPUR SINGAPORE                             ||UTC+08:00
22|BANGKOK HANOI JAKARTA                              ||UTC+07:00
23|CHENNAI KOLKATA MUMBAI NEW DELHI                   ||UTC+05:30
24|ABU DHABI MUSCAT                                   |UTC+04:00
25|TEHRAN                                             |UTC+03:30
26|BAGHDAD                                            |UTC+03:00
27|JERUSALEM                                          |UTC+02:00
28|NEWFOUNDLAND AND LABRADOR                          |UTC-03:30
29|AZORES                                             |UTC-01:00
30|MID ATLANTIC                                       |UTC-02:00
31|MONROVIA                                           |UTC
32|CAYENNE                                            |UTC-03:00
33|GEORGETOWN LA PAZ SAN JUAN                         |UTC-04:00
34|INDIANA EAST                                       |UTC-05:00
35|BOGOTA LIMA QUITO                                  |UTC-05:00
36|SASKATCHEWAN                                       |UTC-06:00
37|GUADALAJARA MEXICO CITY MONTERREY                  |UTC-06:00
38|ARIZONA                                            |UTC-07:00
39|INTERNATIONAL DATE LINE WEST                       |UTC-12:00
40|FIJI ISLANDS MARSHALL ISLANDS                      |UTC+12:00
41|MADAGAN SOLOMON ISLANDS NEW CALENDONIA             |UTC+11:00
42|HOBART                                             |UTC+10:00
43|GUAM PORT MORESBY                                  |UTC+10:00
44|DARWIN                                             |UTC+09:30
45|BEIJING CHONGQING HONG KONG SAR URUMQI             |UTC+08:00
46|NOVOSIBIRSK                                        |UTC+06:00
47|TASHKENT                                           |UTC+05:00
48|KABUL                                              |UTC+04:30
49|CAIRO                                              |UTC+02:00
50|HARARE PRETORIA                                    |UTC+02:00
51|MOSCOW STPETERSBURG VOLGOGRAD                      |UTC+03:00
53|CAPE VERDE ISLANDS                                 |UTC-01:00
54|BAKU                                               |UTC+04:00
55|CENTRAL AMERICA                                    |UTC-06:00
56|NAIROBI                                            |UTC+03:00
57|SARAJEVO SKOPJE WARSAW ZAGREB                      |UTC+01:00
58|EKATERINBURG                                       |UTC+05:00
59|HELSINKI KYIV RIGA SOFIA TALLINN VILNIUS           |UTC+02:00
60|GREENLAND                                          |UTC-03:00
61|YANGON RANGOON                                     |UTC+06:30
62|KATHMANDU                                          |UTC+05:45
63|IRKUTSK                                            |UTC+08:00
64|KRASNOYARSK                                        |UTC+07:00
65|SANTIAGO                                           |UTC-04:00
66|SRI JAYAWARDENEPURA                                |UTC+05:30
67|NUKU ALOFA                                         |UTC+13:00
68|VLADIVOSTOK                                        |UTC+10:00
69|WEST CENTRAL AFRICA                                |UTC+01:00
70|YAKUTSK                                            |UTC+09:00
71|ASTANA DHAKA                                       |UTC+06:00
72|SEOUL                                              |UTC+09:00
73|PERTH                                              |UTC+08:00
74|KUWAIT RIYADH                                      |UTC+03:00
75|TAIPEI                                             |UTC+08:00
76|CANBERRA MELBOURNE SYDNEY                          |UTC+10:00
77|CHIHUAHUA LA PAZ MAZATLAN                          |UTC-07:00
78|TIJUANA BAJA CALFORNIA                             |UTC-08:00
79|AMMAN                                              |UTC+02:00
80|BEIRUT                                             |UTC+02:00
81|MANAUS                                             |UTC-04:00
82|TBILISI                                            |UTC+04:00
83|WINDHOEK                                           |UTC+02:00
84|YEREVAN                                            |UTC+04:00
85|BUENOS AIRES                                       |UTC-03:00
86|CASABLANCA                                         |UTC
87|ISLAMABAD KARACHI                                  |UTC+05:00
88|CARACAS                                            |UTC-04:30
89|PORT LOUIS                                         |UTC+04:00
90|MONTEVIDEO                                         |UTC-03:00
91|ASUNCION                                           |UTC-04:00
92|PETROPAVLOVSK KACHATSKY                            |UTC+12:00
93|COORDINATED UNIVERSAL TIME                         |UTC
94|ULAANBAATAR                                        |UTC-08:00

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
