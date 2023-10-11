<#
Name: GPOBackup - Export Only.ps1

This script will export the data if changes have been made.  This will keep the number of backups and files down to the minimum needed.

This script will create GPO Reports to track changes and backup the GPO's so you can easily locate changes that have been made and recover.
Below is a list of the files created from this script:
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>.zip                      - This contains a backup of all your GPO's the folder is create with the name for each GPO for easy Recovery or viewing
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>.7z                       - This contains a backup of all your GPO's the folder is create with the name for each GPO for easy Recovery or viewing
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOChanges.csv           - This file contains the Changed GPO Information like Name, ID, Owner, Domain, Creation Time, & Modified Time.
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOList.csv              - This file contains GPO Information like Name, ID, Owner, Domain, Creation Time, & Modified Time.
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOReport.csv            - This file Contains GPO Information like Name, Links, Revision Number, & Security Filtering
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOReport.xml            - This a backup of all the GPO's in a single file incase you need to look for a setting and do not know which GPO it is on.
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-GPOReport.html           - This a backup of all the GPO's in a single file incase you need to look for a setting and do not know which GPO it is on.
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-UnlinkedGPOReport.csv    - This file Contains GPO Information like Name, ID, Status, Creation Time, & Modified Time
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-WMIFiltersExport.csv     - This file Contains WMI Filters Information configured in GPMC
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-OrphanedGPOs.txt         - This file Contains Orphaned GPO Report
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-OrphanedGPOsSYSVOL.txt   - This file Contains list of Orphaned GPOs in SYSVOL
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-OrphanedGPOsAD.txt       - This file Contains list of Orphaned GPOs in AD
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-EmptyPOReport.csv        - This file Contains list of Empty GPOs in AD


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
    2023-08-16 - Adding ability to use 7-Zip from compression
    2023-08-21 - Added Orphaned GPO Report, Add 14 Char from GUID for GPO Backups, Cleanup
    2023-10-08 - Moved order to longer processing at the end
    2023-10-09 - Script Optimization
    2023-10-10 - Added EmptyPOReport.csv
    2023-10-11 - Cleanup

Thanks for others on here that I have pulled parts from to make a more comprehensive script

WMI Filter Export:  
    - http://www.jhouseconsulting.com/2014/06/09/script-to-create-import-and-export-group-policy-wmi-filters-1354  
    - ManageWMIFilters.ps1  
Other Parts taken from other scripts on the web

This script is for backups.  To restore you can do the following steps:

1. Extract the zip file to a location for use
2. Open Admin PowerShell
3. import-gpo -BackupGpoName "Original GPO Name" -TargetName "Destination GPO Name" -path "Full Path to GPO Backup":  
    - EX: import-gpo -BackupGpoName "DC - PDC as Authoritative Time Server" -TargetName "DC - PDC as Authoritative Time Server" -path "C:\GPOBackupByName\2021-05-26-16-03-home.local\DC - PDC as Authoritative Time Server_{38bc3df6-b1f1-4a81-93b2-b9412c0f059d}"
4. Open GPMC
5. Verify GPO is restored

#>
# Clear Screen
Clear-Host

# Funtions
# Clear Varables
function Get-UserVariable ($Name = '*') {
    # these variables may exist in certain environments (like ISE, or after use of foreach)
    $special = 'ps', 'psise', 'psunsupportedconsoleapplications', 'foreach', 'profile'

    $ps = [PowerShell]::Create()
    $null = $ps.AddScript('$null=$host;Get-Variable') 
    $reserved = $ps.Invoke() | 
    Select-Object -ExpandProperty Name
    $ps.Runspace.Close()
    $ps.Dispose()
    Get-Variable -Scope Global | 
    Where-Object Name -like $Name |
    Where-Object { $reserved -notcontains $_.Name } |
    Where-Object { $special -notcontains $_.Name } |
    Where-Object Name 
}

# End Function(s)

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
Write-Host "`tPlease Wait - Creating GPO List" -ForeGroundColor Yellow
If ($setServer -eq "Yes") {
    $Script:GPOs = Get-GPO -All -Server $server
    $Script:GPOs | Export-Csv $backupPath-GPOList.csv -NoTypeInformation
}
Else {
    $Script:GPOs = Get-GPO -All
    $Script:GPOs | Export-Csv $backupPath-GPOList.csv -NoTypeInformation
}
Write-Host "`t`tCreated GPO List" -ForeGroundColor Yellow


# Orphaned GPOs
Write-Host "`tPlease Wait - Creating Orphaned GPO Properties Reports" -ForeGroundColor Yellow
$Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
# Get AD Domain Name
$DomainDNS = $Domain.Name
# Get AD Distinguished Name
$DomainDistinguishedName = $Domain.GetDirectoryEntry() | Select-Object -ExpandProperty DistinguishedName  
# Build out GPO Policy Locations
$GPOPoliciesDN = "CN=Policies,CN=System,$DomainDistinguishedName"
$GPOPoliciesSYSVOLUNC = "\\$DomainDNS\SYSVOL\$DomainDNS\Policies"
"Reading GPO information from Active Directory ($GPOPoliciesDN)..." | Out-File -FilePath $backupPath-OrphanedGPOs.txt
$GPOPoliciesADSI = [ADSI]"LDAP://$GPOPoliciesDN"
[array]$GPOPolicies = $GPOPoliciesADSI.psbase.children
ForEach ($GPO in $GPOPolicies) { 
    [array]$DomainGPOList += $GPO.Name
}
#$DomainGPOList = $DomainGPOList -replace("{","") ; $DomainGPOList = $DomainGPOList -replace("}","")
$DomainGPOList = $DomainGPOList | sort-object 
[int]$DomainGPOListCount = $DomainGPOList.Count
"Discovered $DomainGPOListCount GPCs (Group Policy Containers) in Active Directory ($GPOPoliciesDN)`n" | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append
"Reading GPO information from SYSVOL ($GPOPoliciesSYSVOLUNC)..." | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append
[array]$GPOPoliciesSYSVOL = Get-ChildItem $GPOPoliciesSYSVOLUNC
ForEach ($GPO in $GPOPoliciesSYSVOL) {
    If ($GPO.Name -ne "PolicyDefinitions") { 
        [array]$SYSVOLGPOList += $GPO.Name 
    }
}
#$SYSVOLGPOList = $SYSVOLGPOList -replace("{","") ; $SYSVOLGPOList = $SYSVOLGPOList -replace("}","")
$SYSVOLGPOList = $SYSVOLGPOList | sort-object 
[int]$SYSVOLGPOListCount = $SYSVOLGPOList.Count
"Discovered $SYSVOLGPOListCount GPTs (Group Policy Templates) in SYSVOL ($GPOPoliciesSYSVOLUNC)`n" | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append

# Check for GPTs in SYSVOL that don't exist in AD
[array]$MissingADGPOs = Compare-Object $SYSVOLGPOList $DomainGPOList -passThru | Where-Object { $_.SideIndicator -eq '<=' }
[int]$MissingADGPOsCount = $MissingADGPOs.Count
$MissingADGPOsPCTofTotal = $MissingADGPOsCount / $DomainGPOListCount
$MissingADGPOsPCTofTotal = "{0:p2}" -f $MissingADGPOsPCTofTotal  
"There are $MissingADGPOsCount GPTs in SYSVOL that don't exist in Active Directory ($MissingADGPOsPCTofTotal of the total)" | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append

If ($MissingADGPOsCount -gt 0 ) {
    "These are:" | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append
    $MissingADGPOs | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append
}
"`n" | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append
# Write Missing GPOs in AD to CSV File
if ($MissingADGPOs.Count -gt 0) {
    $MissingADGPOs | Out-File -FilePath $backupPath-OrphanedGPOsAD.txt
}

# Check for GPCs in AD that don't exist in SYSVOL
[array]$MissingSYSVOLGPOs = Compare-Object $DomainGPOList $SYSVOLGPOList -passThru | Where-Object { $_.SideIndicator -eq '<=' }
[int]$MissingSYSVOLGPOsCount = $MissingSYSVOLGPOs.Count
$MissingSYSVOLGPOsPCTofTotal = $MissingSYSVOLGPOsCount / $DomainGPOListCount
$MissingSYSVOLGPOsPCTofTotal = "{0:p2}" -f $MissingSYSVOLGPOsPCTofTotal  
"There are $MissingSYSVOLGPOsCount GPCs in Active Directory that don't exist in SYSVOL ($MissingSYSVOLGPOsPCTofTotal of the total)" | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append

If ($MissingSYSVOLGPOsCount -gt 0 ) {
    "These are:" | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append
    $MissingSYSVOLGPOs | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append
}
"`n" | Out-File -FilePath $backupPath-OrphanedGPOs.txt -Append
# Write Missing GPOs in SYSVOL to CSV File
if ($MissingSYSVOLGPOs.Count -gt 0) {
    $MissingSYSVOLGPOs | Out-File -FilePath $backupPath-OrphanedGPOsSYSVOL.txt
}


# Export Unlinked GPO Report
Write-Host "`tPlease Wait - Creating Unlinked GPO Properties Report" -ForeGroundColor Yellow
function IsNotLinked($xmldata) {
    If ($null -eq $xmldata.GPO.LinksTo) {
        Return $true
    }
    Return $false
}
$unlinkedGPOs = @()
If ($setServer -eq "Yes") {
    foreach ($gpo in $Script:GPOs) {
        Get-GPOReport -Guid $gpo.Id -Server $server -ReportType xml | ForEach-Object { If (IsNotLinked([xml]$_)) { $unlinkedGPOs += $gpo } }
    }
}
Else {
    foreach ($gpo in $Script:GPOs) {
        Get-GPOReport -Guid $gpo.Id -ReportType xml | ForEach-Object { If (IsNotLinked([xml]$_)) { $unlinkedGPOs += $gpo } }
    }

}
If ($unlinkedGPOs.Count -eq 0) {
    Write-Host "`t`tNo Unlinked GPO's Found" -ForeGroundColor Green
}
Else {
    $unlinkedGPOs | Sort-Object GpoStatus, DisplayName | Select-Object DisplayName, ID, GpoStatus, CreationTime, ModificationTime | Export-Csv -Delimiter ',' -Path $backupPath-UnlinkedGPOReport.csv -NoTypeInformation
}
Write-Host "`t`tCreated Unlinked GPO Properties Report" -ForeGroundColor Yellow


# Empty GPO's & GPO Properties Report
Write-Host "`tPlease Wait - Checking for Empty GPO's" -ForeGroundColor Yellow
Write-Host "`tPlease Wait - Creating GPO Properties Report" -ForeGroundColor Yellow
$emptyGPOs = @()
$colGPOLinks = @()
foreach ($gpo in $Script:GPOs) {
    If ($setServer -eq "Yes") {
        [xml]$gpocontent = Get-GPOReport -Guid $gpo.Id -ReportType xml -Server $server
        If ($NULL -eq $gpocontent.GPO.Computer.ExtensionData -and $NULL -eq $gpocontent.GPO.User.ExtensionData) {
            $emptyGPOs += $gpo
        }
        $LinksPaths = $gpocontent.GPO.LinksTo | Where-Object { $_.Enabled -eq $True } | ForEach-Object { $_.SOMPath }
    }
    Else {
        [xml]$gpocontent = Get-GPOReport -Guid $gpo.Id -ReportType xml
        If ($NULL -eq $gpocontent.GPO.Computer.ExtensionData -and $NULL -eq $gpocontent.GPO.User.ExtensionData) {
            $emptyGPOs += $gpo
        }
        $LinksPaths = $gpocontent.GPO.LinksTo | Where-Object { $_.Enabled -eq $True } | ForEach-Object { $_.SOMPath }
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
        $SecurityFilter = ((Get-GPPermissions -Guid $gpo.Id -All -Server $server | Where-Object { $_.Permission -eq "GpoApply" }).Trustee | Where-Object { $_.SidType -ne "Unknown" }).name -Join ','
    }
    Else {
        $SecurityFilter = ((Get-GPPermissions -Guid $gpo.Id -All | Where-Object { $_.Permission -eq "GpoApply" }).Trustee | Where-Object { $_.SidType -ne "Unknown" }).name -Join ','
    }
    foreach ($LinksPath in $LinksPaths) {
        $objGPOLinks = New-Object System.Object
        $objGPOLinks | Add-Member -type noteproperty -name GPOName -value $gpo.DisplayName
        $objGPOLinks | Add-Member -type noteproperty -name ID -value $gpo.Id
        $objGPOLinks | Add-Member -type noteproperty -name LinksPath -value $LinksPath
        $objGPOLinks | Add-Member -type noteproperty -name WmiFilter -value ($gpo.WmiFilter).Name
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
If ($emptyGPOs.Count -eq 0) {
    Write-Host "`t`tNo Empty GPO's Found" -ForeGroundColor Green
}
Else {
    $emptyGPOs | Sort-Object GpoStatus, DisplayName | Select-Object DisplayName, ID, GpoStatus, CreationTime, ModificationTime | Export-Csv -Delimiter ',' -Path $backupPath-EmptyPOReport.csv -NoTypeInformation
    Write-Host "`t`tCreated Empty GPO Report" -ForeGroundColor Yellow
}
$colGPOLinks | sort-object GPOName, LinksPath | Export-Csv -Delimiter ',' -Path $backupPath-GPOReport.csv -NoTypeInformation
Write-Host "`t`tCreated GPO Properties Report" -ForeGroundColor Yellow



# Backup WMI Filters
Write-Host "`tPlease Wait - Backing up WMI Filters" -ForeGroundColor Yellow
$WmiFilters = Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' -Properties * | Select-Object * 
$RowCount = $WMIFilters | Measure-Object | Select-Object -expand count
if ($RowCount -ne 0) {
    write-host "`t`tExporting $RowCount WMI Filters" -ForeGroundColor Green
    $WMIFilters | Export-Csv $backupPath-WMIFiltersExport.csv -NoTypeInformation
} 
else {
    Write-Host "`t`tThere are no WMI Filters to export" -ForeGroundColor Green 
} 
Write-Host "`t`tBacked up WMI Filters" -ForeGroundColor Yellow


# Export GPO Report - XML
Write-Host "`tPlease Wait - Creating GPO Report - XML" -ForeGroundColor Yellow
If ($setServer -eq "Yes") {
    Get-GPOReport -All -Server $server -ReportType xml -Path $backupPath-GPOReport.xml
}
Else {
    Get-GPOReport -All -ReportType xml -Path $backupPath-GPOReport.xml
}
Write-Host "`t`tCreated GPO Report - XML" -ForeGroundColor Yellow


# Export GPO Report - HTML
If ($HTMLReport -eq "Yes") {
    Write-Host "`tPlease Wait - Creating GPO Report - HTML" -ForeGroundColor Yellow
    If ($setServer -eq "Yes") {
        Get-GPOReport -All -Server $server -ReportType Html -Path $backupPath-GPOReport.html
    }
    Else {
        Get-GPOReport -All -ReportType Html -Path $backupPath-GPOReport.html
    }
    Write-Host "`t`tCreated GPO Report - HTML" -ForeGroundColor Yellow
}


# Verify GPO BackupPath
Write-Host "`tPlease Wait - Creating Backup Directory" -ForeGroundColor Yellow
if ((Test-Path $backupPath) -eq $false) {
    New-Item -Path $backupPath -ItemType directory
}


# Backup GPOs into named folders
if ($individualBackup -eq 'Yes') {
    Write-Host "`tPlease Wait - Backing up GPO's" -ForeGroundColor Yellow
    If ($setServer -eq "Yes") {
        foreach ($gpo in $Script:GPOs) {
            Write-Host "`t`tProcessing GPO" $gpo.displayname -ForeGroundColor Yellow
            #$foldername = join-path $backupPath ($gpo.displayname.Replace(" ", "_") + "_{" + $gpo.Id + "}") # Replace " " with "_"
            #$foldername = join-path $backupPath ($gpo.displayname + "_{" + $gpo.Id + "}") # Keep " "
            #$foldername = join-path $backupPath ($gpo.displayname) # Raw Name # Keep " "
            $foldername = join-path $backupPath ($gpo.displayname + "_{" + $($gpo.Id).ToString().Substring(0,14) + "}") # Keep " "
            #Write-Host $foldername
            if ((Test-Path $foldername) -eq $false) {
                New-Item -Path $foldername -ItemType directory
            }
            Backup-GPO -Server $server -Name $gpo.displayname -Path $foldername -Comment $date
            #$filename = join-path $backupPath ($gpo.displayname.Replace(" ", "_") + ".html") # Replace " " with "_"
            #$filename = join-path $backupPath ($gpo.displayname + ".html") # Raw Name # Keep " "
            $filename = join-path $backupPath ($gpo.displayname + "_{" + $($gpo.Id).ToString().Substring(0,14) + "}" + ".html") # Raw Name # Keep " "
            Get-GPOReport -Name $gpo.displayname -ReportType 'HTML'-Path $filename
        }
    }
    Else {
        foreach ($gpo in $Script:GPOs) {
            Write-Host "`t`tProcessing GPO" $gpo.displayname -ForeGroundColor Yellow
            #$foldername = join-path $backupPath ($gpo.displayname.Replace(" ", "_") + "_{" + $gpo.Id + "}") # Replace " " with "_"
            #$foldername = join-path $backupPath ($gpo.displayname + "_{" + $gpo.Id + "}") # Keep " "
            #$foldername = join-path $backupPath ($gpo.displayname) # Raw Name # Keep " "
            $foldername = join-path $backupPath ($gpo.displayname + "_{" + $($gpo.Id).ToString().Substring(0,14) + "}") # Keep " "
            if ((Test-Path $foldername) -eq $false) {
                New-Item -Path $foldername -ItemType directory
            }
            Backup-GPO -Name $gpo.displayname -Path $foldername -Comment $date
            #$filename = join-path $backupPath ($gpo.displayname.Replace(" ", "_") + ".html") # Replace " " with "_"
            #$filename = join-path $backupPath ($gpo.displayname + ".html") # Raw Name # Keep " "
            $filename = join-path $backupPath ($gpo.displayname + "_{" + $($gpo.Id).ToString().Substring(0,14) + "}" + ".html") # Raw Name # Keep " "
            Get-GPOReport -Name $gpo.displayname -ReportType 'HTML'-Path $filename
        }
    }
    Write-Host "`t`tBacked up GPO's" -ForeGroundColor Yellow
}


# Backup All GPOs into one folder
if ($singleBackup -eq 'Yes') {
    Write-Host "`tPlease Wait - Backing up GPO's" -ForeGroundColor Yellow
    If ($setServer -eq "Yes") {
        $foldername = join-path $backupPath + "_All"
        if ((Test-Path $foldername) -eq $false) {
            New-Item -Path $foldername -ItemType directory
        }
        Backup-GPO -All -Server $server -Path $foldername -Comment $date
        Write-Host "`t`tBacked up GPO's" -ForeGroundColor Yellow
    }
    Else {
        $foldername = join-path $backupPath + "_All"
        if ((Test-Path $foldername) -eq $false) {
            New-Item -Path $foldername -ItemType directory
        }
        Backup-GPO -All -Path $foldername -Comment $date
        Write-Host "`t`tBacked up GPO's" -ForeGroundColor Yellow
    }
}


# Backup PolicyDefinition Folder
Write-Host "`tPlease Wait - Backing up PolicyDefinition Folder" -ForeGroundColor Yellow
$policydefinitionSource = "\\" + $env:USERDOMAIN + "\SYSVOL\" + $env:USERDNSDOMAIN + "\Policies\PolicyDefinitions"
Copy-Item -Path $policydefinitionSource -Recurse -Destination $backupPath -Container
Write-Host "`t`tBacked up PolicyDefinition Folder" -ForeGroundColor Yellow


# Compress Folders
Write-Host "`tPlease Wait - Checking for 7-Zip" -ForeGroundColor Yellow
# Path to 7-Zip
$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"
# Compress Folders to 7-Zip File
if ((Test-Path $7zipPath) -eq $true) {
    Write-Host "`tPlease Wait - Creating 7-ZIP File" -ForeGroundColor Yellow
    # Create Alias
    Set-Alias Compress-7Zip $7ZipPath
    # Set Source & Destination
    $source = $backupPath
    $destination = $backupPath + ".7z"
    # Compress Files/Folder
    Compress-7zip a -mx9 -r -t7z $destination $source
}

# Compress Folders to Zip File
if ((Test-Path $7zipPath) -eq $false) {
    Write-Host "`tPlease Wait - Creating ZIP File" -ForeGroundColor Yellow
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
    Write-Host "`t`tCreated ZIP File" -ForeGroundColor Yellow
}


# Delete GPO Backup Folder
Write-Host "`tPlease Wait - Deleting GPO Backup Folder" -ForeGroundColor Yellow
Remove-item -Path $backupPath -Recurse -Force -ErrorAction SilentlyContinue


# Completed Script
Write-Host "`tGPO Backup - Complete" -ForeGroundColor Yellow


# Clear Variables
Write-Host "`tScript Cleanup" -ForeGroundColor Yellow
Get-UserVariable | Remove-Variable -ErrorAction SilentlyContinue


# End
Exit
