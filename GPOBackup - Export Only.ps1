<#
Name: GPOBackup - Export Only.ps1

This script will check for GPO's modified in the last day and then export the data.  This will keep the number of backups and files down to the minimun needed.

This script will create GPO Reports to track changes and backup the GPO's so you can easily locate changes that have been made and recover.
Below is a list of the files created from this script:
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>.zip                      - This contains a backup of all your GPO's the folder is create with the name for each GPO for easy Recovery or viewing
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOChanges.csv           - This file contains the Changed GPO Information like Name, ID, Owner, Domain, Creation Time, & Modified Time.
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOList.csv              - This file contains GPO Information like Name, ID, Owner, Domain, Creation Time, & Modified Time.
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOReport.csv            - This file Contains GPO Information like Name, Links, Revision Number, & Security Filtering
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOReport.xml            - This a backup of all the GPO's in a single file incase you need to look for a setting and do not know which GPO it is on.
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOReport.html           - This a backup of all the GPO's in a single file incase you need to look for a setting and do not know which GPO it is on.
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-UnlinkedGPOReport.csv    - This file Contains GPO Information like Name, ID, Status, Creation Time, & Modified Time
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-WMIFiltersExport.csv     - This file Contains WMI Filters Information configured in GPMC


This script was based off of one from Microsoft to backup GPO's by name, I have added more as the need and to make things simplier when backup up GPO's

Michael Patterson
scripts@mwpatterson.com

Revision History
    2017-08-18 - Initial Release
    2017-08-18 - Added and change/notes for drive mapping incase user was mapping to the full path of the folder
    2017-08-18 - Added Changed GPO Report and eMail Notification for GPO's that changed
    2017-08-25 - Added Unlinked GPO Report
    2017-08-31 - Changed Location for Variables to start of Script, Code Cleanup & Formatting
    2017-12-08 - Added moving to sub folder for yearto keep clutter down, could take it down to mmonth as well by changing $year from "yyyy" to "yyyy-MM"
    2017-12-27 - Cleanup
    2018-01-31 - Added check to not send emails
    2019-01-02 - Changed Text color and added message abount which GPO it was backing up incase it gives an error on backup
    2019-01-10 - Added PolicyDefinition Folder Backup
    2019-01-10 - Cleanup
    2019-08-22 - Added ability to copy to SharePoint
    2019-08-22 - Added Deleting files older than X Days
    2019-12-17 - Cleanup
    2020-01-22 - Added Comment for GPO being backed up & Added All GPO's into One folder Options
    2020-04-23 - Added WMI Filter Export
    2020-06-23 - Added HTML Report
    2020-07-08 - Added Way to Turn of HTML Report if not needed
    2020-08-25 - Cleanup
    2020-08-31 - Added Server Setting to specify Domain Controller
    2020-11-27 - Cleanup
    2021-04-14 - Added GPO Change Count messages
    2021-05-13 - Added HTML Reporting for Individual GPO's
    2022-02-18 - Removed all but stuff needed just for Exporting the GPO's
    2022-09-01 - Remove GUID from the Folder path to all long GPO Names
    2023-03-16 - Script Cleanup

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

#>
# Clear Screen
Clear-Host


# Set Variables
# HTML Report 
#$HTMLReport = "Yes"
$HTMLReport = "No"

# Individual Backup 
$individualBackup = "Yes"
#$individualBackup = "No"

# Single Backup 
#$singleBackup = "Yes"
$singleBackup = "No"

# WMI Filters Backup
$WMIFilters = "Yes"
#$WMIFilters = "No"

# Set Domain
#$domain = $env:USERDNSDOMAIN #Full Domain Name
#$domain = $env:USERDOMAIN #Short Domain Name

# Set Domain Controller
$setServer = "No"
#$setServer = "Yes"
#$server = "<FQDN>"

# Get Date & Backup Locations
$date = get-date -Format "yyyy-MM-dd-HH-mm"
$backupRoot = "C:\" #Can use another drive if available
$backupFolder = "GPOBackupByName\"
$backupFolderPath = $backupRoot + $backupFolder
#$backupFileName = $date + "-" + $domain 
$backupFileName = $date + "-" + $env:USERDNSDOMAIN #Full Domain Name 
#$backupFileName = $date + "-" + $env:USERDOMAIN #Short Domain Name
#$backupPath = $backupRoot + $backupFolder + $date + "-" + $domain
$backupPath = $backupFolderPath + $backupFileName


# Begin Processing GPO's
# Verify GPO BackupFolder
if ((Test-Path $backupFolderPath) -eq $false) {
    New-Item -Path $backupFolderPath -ItemType directory
}


# Export GPO List
Write-Host "`tPlease Wait - Creating GPO List" -Fore Yellow
If ($setServer -eq "Yes") {
    Get-GPO -All -Server $server| Export-csv $backupPath-GPOList.csv -NoTypeInformation
}
Else {
    Get-GPO -All | Export-csv $backupPath-GPOList.csv -NoTypeInformation
}
Write-Host "`t`tCreated GPO List" -Fore Yellow


# Export GPO Report - XML
Write-Host "`tPlease Wait - Creating GPO Report - XML" -Fore Yellow
If ($setServer -eq "Yes") {
    Get-GPOReport -All -Server $server -ReportType xml -Path $backupPath-GPOReport.xml
}
Else {
    Get-GPOReport -All -ReportType xml -Path $backupPath-GPOReport.xml
}
Write-Host "`t`tCreated GPO Report - XML" -Fore Yellow


# Export GPO Report - HTML
If ($HTMLReport -eq "Yes") {
    Write-Host "`tPlease Wait - Creating GPO Report - HTML" -Fore Yellow
    If ($setServer -eq "Yes") {
        Get-GPOReport -All -Server $server -ReportType Html -Path $backupPath-GPOReport.html
    }
    Else {
        Get-GPOReport -All -ReportType Html -Path $backupPath-GPOReport.html
    }
    Write-Host "`t`tCreated GPO Report - HTML" -Fore Yellow
}


# Export GPO Properties Report
Write-Host "`tPlease Wait - Creating GPO Properties Report" -Fore Yellow
If ($setServer -eq "Yes") {
    $GPOList = (Get-Gpo -All -Server $server).DisplayName
}
Else {
    $GPOList = (Get-Gpo -All).DisplayName
}
$colGPOLinks = @()
foreach ($GPOItem in $GPOList) {
    If ($setServer -eq "Yes") {
        [xml]$gpocontent = Get-GPOReport $GPOItem -ReportType xml -Server $server
        $LinksPaths = $gpocontent.GPO.LinksTo | Where-Object { $_.Enabled -eq $True } | ForEach-Object { $_.SOMPath }
        $Wmi = Get-GPO $GPOItem -Server $server | Select-Object WmiFilter
    }
    Else {
        [xml]$gpocontent = Get-GPOReport $GPOItem -ReportType xml
        $LinksPaths = $gpocontent.GPO.LinksTo | Where-Object { $_.Enabled -eq $True } | ForEach-Object { $_.SOMPath }
        $Wmi = Get-GPO $GPOItem | Select-Object WmiFilter
    }
    $CreatedTime = $gpocontent.GPO.CreatedTime
    $ModifiedTime = $gpocontent.GPO.ModifiedTime
    $CompVerDir = $gpocontent.GPO.Computer.VersionDirectory
    $CompVerSys = $gpocontent.GPO.Computer.VersionSysvol
    $CompEnabled = $gpocontent.GPO.Computer.Enabled
    $UserVerDir = $gpocontent.GPO.User.VersionDirectory
    $UserVerSys = $gpocontent.GPO.User.VersionSysvol
    $UserEnabled = $gpocontent.GPO.User.Enabled
    If ($setServer -eq "Yes") {
        $SecurityFilter = ((Get-GPPermissions -Name $GPOItem -All -Server $server | Where-Object { $_.Permission -eq "GpoApply" }).Trustee | Where-Object { $_.SidType -ne "Unknown" }).name -Join ','
    }
    Else {
        $SecurityFilter = ((Get-GPPermissions -Name $GPOItem -All | Where-Object { $_.Permission -eq "GpoApply" }).Trustee | Where-Object { $_.SidType -ne "Unknown" }).name -Join ','
    }
    foreach ($LinksPath in $LinksPaths) {
        $objGPOLinks = New-Object System.Object
        $objGPOLinks | Add-Member -type noteproperty -name GPOName -value $GPOItem
        $objGPOLinks | Add-Member -type noteproperty -name LinksPath -value $LinksPath
        $objGPOLinks | Add-Member -type noteproperty -name WmiFilter -value ($wmi.WmiFilter).Name
        $objGPOLinks | Add-Member -type noteproperty -name CreatedTime -value $CreatedTime
        $objGPOLinks | Add-Member -type noteproperty -name ModifiedTime -value $ModifiedTime
        $objGPOLinks | Add-Member -type noteproperty -name ComputerRevisionsAD -value $CompVerDir
        $objGPOLinks | Add-Member -type noteproperty -name ComputerRevisionsSYSVOL -value $CompVerSys
        $objGPOLinks | Add-Member -type noteproperty -name UserRevisionsAD -value $UserVerDir
        $objGPOLinks | Add-Member -type noteproperty -name UserRevisionsSYSVOL -value $UserVerSys
        $objGPOLinks | Add-Member -type noteproperty -name ComputerSettingsEnabled -value $CompEnabled
        $objGPOLinks | Add-Member -type noteproperty -name UserSettingsEnabled -value $UserEnabled
        $objGPOLinks | Add-Member -type noteproperty -name SecurityFilter -value $SecurityFilter
        $colGPOLinks += $objGPOLinks
    }
}
$colGPOLinks | sort-object GPOName, LinksPath | Export-Csv -Delimiter ',' -Path $backupPath-GPOReport.csv -NoTypeInformation
Write-Host "`t`tCreated GPO Properties Report" -Fore Yellow


# Export Unlinked GPO Report
Write-Host "`tPlease Wait - Creating Unlinked GPO Properties Report" -Fore Yellow
function IsNotLinked($xmldata) {
    If ($null -eq $xmldata.GPO.LinksTo) {
        Return $true
    }
    Return $false
}
$unlinkedGPOs = @()
If ($setServer -eq "Yes") {
    Get-GPO -All -Server $server | ForEach-Object { $gpo = $_ ; $_ | Get-GPOReport -Server $server -ReportType xml | ForEach-Object { If (IsNotLinked([xml]$_)) { $unlinkedGPOs += $gpo } } }
}
Else {
    Get-GPO -All | ForEach-Object { $gpo = $_ ; $_ | Get-GPOReport -ReportType xml | ForEach-Object { If (IsNotLinked([xml]$_)) { $unlinkedGPOs += $gpo } } }
}
If ($unlinkedGPOs.Count -eq 0) {
    Write-Host "`t`tNo Unlinked GPO's Found" -Fore Green
}
Else {
    $unlinkedGPOs | Sort-Object GpoStatus, DisplayName | Select-Object DisplayName, ID, GpoStatus, CreationTime, ModificationTime | Export-Csv -Delimiter ',' -Path $backupPath-UnlinkedGPOReport.csv -NoTypeInformation
}
Write-Host "`t`tCreated Unlinked GPO Properties Report" -Fore Yellow


# Backup WMI Filters
Write-Host "`tPlease Wait - Backing up WMI Filters" -Fore Yellow
#$WMIFilters = @()
#$WmiFilters = Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' -Properties * | Select DistinguishedName, whenCreated, whenChanged, msWMI-Author, msWMI-ID, msWMI-Name, msWMI-Parm1, msWMI-Parm2
$WmiFilters = Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' -Properties * | Select-Object * 
$RowCount = $WMIFilters | Measure-Object | Select-Object -expand count
if ($RowCount -ne 0) {
    write-host -ForeGroundColor Green "`tExporting $RowCount WMI Filters"
    $WMIFilters | export-csv $backupPath-WMIFiltersExport.csv -NoTypeInformation
    } 
else {
    write-host -ForeGroundColor Green "There are no WMI Filters to export`n"
    } 
Write-Host "`t`tBacked up WMI Filters" -Fore Yellow


# Verify GPO BackupPath
Write-Host "`tPlease Wait - Creating Backup Directory" -Fore Yellow
if ((Test-Path $backupPath) -eq $false) {
    New-Item -Path $backupPath -ItemType directory
}


# Backup GPOs into named folders
if ($individualBackup -eq 'Yes'){
    Write-Host "`tPlease Wait - Backing up GPO's" -Fore Yellow
    If ($setServer -eq "Yes") {
        $allGPOs = get-gpo -all -Server $server
        foreach ($gpo in $allGPOs) {
            Write-Host "`t`tProcessing GPO" $gpo.displayname -Fore Yellow
            #$foldername = join-path $backupPath ($gpo.displayname.Replace(" ", "_") + "_{" + $gpo.Id + "}") # Replace " " with "_"
            #$foldername = join-path $backupPath ($gpo.displayname + "_{" + $gpo.Id + "}") # Keep " "
            $foldername = join-path $backupPath ($gpo.displayname) # Keep " "
            if ((Test-Path $foldername) -eq $false) {
                New-Item -Path $foldername -ItemType directory
            }
            Backup-GPO -Server $server -Name $gpo.displayname -Path $foldername -Comment $date
            #$filename = join-path $backupPath ($gpo.displayname.Replace(" ", "_") + ".html") # Replace " " with "_"
            $filename = join-path $backupPath ($gpo.displayname + ".html") # Keep " "
            Get-GPOReport -Name $gpo.displayname -ReportType 'HTML'-Path $filename
        }
    }
    Else {
        $allGPOs = get-gpo -all
        foreach ($gpo in $allGPOs) {
            Write-Host "`t`tProcessing GPO" $gpo.displayname -Fore Yellow
            #$foldername = join-path $backupPath ($gpo.displayname.Replace(" ", "_") + "_{" + $gpo.Id + "}") # Replace " " with "_"
            #$foldername = join-path $backupPath ($gpo.displayname + "_{" + $gpo.Id + "}") # Keep " "
            $foldername = join-path $backupPath ($gpo.displayname) # Keep " "
            if ((Test-Path $foldername) -eq $false) {
                New-Item -Path $foldername -ItemType directory
            }
            Backup-GPO -Name $gpo.displayname -Path $foldername -Comment $date
            #$filename = join-path $backupPath ($gpo.displayname.Replace(" ", "_") + ".html") # Replace " " with "_"
            $filename = join-path $backupPath ($gpo.displayname + ".html") # Keep " "
            Get-GPOReport -Name $gpo.displayname -ReportType 'HTML'-Path $filename
        }
    }
    Write-Host "`t`tBacked up GPO's" -Fore Yellow
}


# Backup All GPOs into one folder
if ($singleBackup -eq 'Yes'){
    Write-Host "`tPlease Wait - Backing up GPO's" -Fore Yellow
    If ($setServer -eq "Yes") {
        $foldername = join-path $backupPath + "_All"
        if ((Test-Path $foldername) -eq $false) {
            New-Item -Path $foldername -ItemType directory
        }
        Backup-GPO -All -Server $server -Path $foldername -Comment $date
        Write-Host "`t`tBacked up GPO's" -Fore Yellow
    }
    Else {
        $foldername = join-path $backupPath + "_All"
        if ((Test-Path $foldername) -eq $false) {
            New-Item -Path $foldername -ItemType directory
        }
        Backup-GPO -All -Path $foldername -Comment $date
        Write-Host "`t`tBacked up GPO's" -Fore Yellow
    }
}


# Backup PolicyDefinition Folder
Write-Host "`tPlease Wait - Backing up PolicyDefinition Folder" -Fore Yellow
$policydefinitionSource = "\\" + $env:USERDOMAIN + "\SYSVOL\" + $env:USERDNSDOMAIN + "\Policies\PolicyDefinitions"
Copy-Item -Path $policydefinitionSource -Recurse -Destination $backupPath -Container
Write-Host "`t`tBacked up PolicyDefinition Folder" -Fore Yellow


# Compress Folders
Write-Host "`tPlease Wait - Creating ZIP File" -Fore Yellow
#PowerShell 5.0
#Compress-Archive -Path $backupPath -DestinationPath $backupPath+".zip"
#PowerShell 2.0-4.x
$source = $backupPath
$destination = $backupPath + ".zip"
If (Test-path $destination) {
    Remove-item $destination
}
Add-Type -assembly "system.io.compression.filesystem"
[io.compression.zipfile]::CreateFromDirectory($Source, $destination)
Write-Host "`t`tCreated ZIP File" -Fore Yellow


# Delete GPO Backup Folder
#Write-Output "`tPlease Wait - Deleting GPO Backup Folder"
Write-Host "`tPlease Wait - Deleting GPO Backup Folder" -Fore Yellow
Remove-item -Path $backupPath -Recurse -Force -ErrorAction SilentlyContinue


# Completed Script
Write-Output "`tGPO Backup - Complete"
