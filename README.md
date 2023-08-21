Name: GPOBackup.ps1

This script will check for GPO's modified in the last day and then only export the data if changes have been made.  This will keep the number of backups and files down to the minimun needed.

This script will create GPO Reports to track changes and backup the GPO's so you can easily locate changes that have been made and recover.
Below is a list of the files created from this script:

- '<Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>.zip'                      - This contains a backup of all your GPO's the folder is create with the name for each GPO for easy Recovery or viewing
- <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOChanges.csv           - This file contains the Changed GPO Information like Name, ID, Owner, Domain, Creation Time, & Modified Time.
- <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOList.csv              - This file contains GPO Information like Name, ID, Owner, Domain, Creation Time, & Modified Time.
- <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOReport.csv            - This file Contains GPO Information like Name, Links, Revision Number, & Security Filtering
- <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOReport.xml            - This a backup of all the GPO's in a single file incase you need to look for a setting and do not know which GPO it is on.
- <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOReport.html           - This a backup of all the GPO's in a single file incase you need to look for a setting and do not know which GPO it is on.
- <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-UnlinkedGPOReport.csv    - This file Contains GPO Information like Name, ID, Status, Creation Time, & Modified Time
- <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-WMIFiltersExport.csv     - This file Contains WMI Filters Information configured in GPMC
- <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-OrphanedGPOs.txt         - This file Contains Orphaned GPO Report
- <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-OrphanedGPOsSYSVOL.csv   - This file Contains list of Orphaned GPOs in SYSVOL
- <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-OrphanedGPOsAD.csv       - This file Contains list of Orphaned GPOs in AD


This script was based off of one from Microsoft to backup GPO's by name, I have added more as the need and to make things simpler when backup up GPO's

Michael Patterson
scripts@mwpatterson.com

Revision History

- 2017-08-18 - Initial Release
- 2017-08-18 - Added and change/notes for drive mapping incase user was mapping to the full path of the folder
- 2017-08-18 - Added Changed GPO Report and eMail Notification for GPO's that changed
- 2017-08-25 - Added Unlinked GPO Report
- 2017-08-31 - Changed Location for Variables to start of Script, Code Cleanup & Formatting
- 2017-12-08 - Added moving to sub folder for yearto keep clutter down, could take it down to mmonth as well by changing $year from "yyyy" to "yyyy-MM"
- 2017-12-27 - Cleanup
- 2018-01-31 - Added check to not send emails
- 2019-01-02 - Changed Text color and added message abount which GPO it was backing up incase it gives an error on backup
- 2019-01-10 - Added PolicyDefinition Folder Backup
- 2019-01-10 - Cleanup
- 2019-08-22 - Added ability to copy to SharePoint
- 2019-08-22 - Added Deleting files older than X Days
- 2019-12-17 - Cleanup
- 2020-01-22 - Added Comment for GPO being backed up & Added All GPO's into One folder Options
- 2020-04-23 - Added WMI Filter Export
- 2020-06-23 - Added HTML Report
- 2020-07-08 - Added Way to Turn of HTML Report if not needed
- 2020-08-25 - Cleanup
- 2020-08-31 - Added Server Setting to specify Domain Controller
- 2020-11-27 - Cleanup
- 2021-04-14 - Added GPO Change Count messages
- 2021-05-13 - Added HTML Reporting for Individual GPO's
- 2022-09-01 - Remove GUID from the Folder path to all long GPO Names
- 2023-03-16 - Script Cleanup
- 2023-08-16 - Adding ability to use 7-Zip from compression
- 2023-08-21 - Added Orphaned GPO Report, Add 14 Char from GUID for GPO Backups, Cleanup


Thanks for others on here that I have pulled parts from to make a more comprehensive script
    WMI Filter Export
    http://www.jhouseconsulting.com/2014/06/09/script-to-create-import-and-export-group-policy-wmi-filters-1354
    ManageWMIFilters.ps1
    Other Parts taken from other scripts on the web

This script is for backups.  To restore you can do the following steps
    Extract the zip file to a location for use
    Open Admin PowerShell
    import-gpo -BackupGpoName <Origional GPO Name> -TargetName <Destination GPO Name> -path <Full Path to GPO Backup>
    EX: import-gpo -BackupGpoName "DC - PDC as Authoritative Time Server" -TargetName "DC - PDC as Authoritative Time Server" -path "C:\GPOBackupByName\2021-05-26-16-03-home.local\DC - PDC as Authoritative Time Server_{38bc3df6-b1f1-4a81-93b2-b9412c0f059d}"
    Open GPMC
    Verify GPO is restored
