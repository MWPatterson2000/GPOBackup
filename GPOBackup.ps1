<#
Name: GPOBackup.ps1

This script will check for GPO's modified in the last day and then only export the data if changes have been made.  This will keep the number of backups and files down to the minimun needed.

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
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-OrphanedGPOsSYSVOL.csv   - This file Contains list of Orphaned GPOs in SYSVOL
    <Year>-<Month>-<Date>-<Hour>-<Minuite>-<Domain>-OrphanedGPOsAD.csv       - This file Contains list of Orphaned GPOs in AD


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
    2022-09-01 - Remove GUID from the Folder path to all long GPO Names
    2023-03-16 - Script Cleanup
    2023-08-16 - Adding ability to use 7-Zip from compression
    2023-08-21 - Added Orphaned GPO Report

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

# Funtions
# SharePoint Upload Function
Function UploadFileInSlice ($ctx, $libraryName, $fileName, $fileChunkSizeInMB) {
    $fileChunkSizeInMB = 9
    # Each sliced upload requires a unique ID.
    $UploadId = [GUID]::NewGuid()
    # Get the name of the file.
    $UniqueFileName = [System.IO.Path]::GetFileName($fileName)
    # Get the folder to upload into.
    $Docs = $ctx.Web.Lists.GetByTitle($libraryName)
    $ctx.Load($Docs)
    $ctx.Load($Docs.RootFolder)
    $ctx.ExecuteQuery()
    # Get the information about the folder that will hold the file.
    $ServerRelativeUrlOfRootFolder = $Docs.RootFolder.ServerRelativeUrl
    # File object.
    [Microsoft.SharePoint.Client.File] $upload
    # Calculate block size in bytes.
    $BlockSize = $fileChunkSizeInMB * 1024 * 1024
    # Get the size of the file.
    $FileSize = (Get-Item $fileName).length
    if ($FileSize -le $BlockSize) {
        # Use regular approach.
        $FileStream = New-Object IO.FileStream($fileName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
        $FileCreationInfo.Overwrite = $true
        $FileCreationInfo.ContentStream = $FileStream
        $FileCreationInfo.URL = $List.RootFolder.ServerRelativeUrl + "/" + $UniqueFileName
        $Upload = $Docs.RootFolder.Files.Add($FileCreationInfo)
        $ctx.Load($Upload)
        $ctx.ExecuteQuery()
        return $Upload
    }
    else {
        # Use large file upload approach.
        $BytesUploaded = $null
        $Fs = $null
        Try {
            $Fs = [System.IO.File]::Open($fileName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
            $br = New-Object System.IO.BinaryReader($Fs)
            $buffer = New-Object System.Byte[]($BlockSize)
            $lastBuffer = $null
            $fileoffset = 0
            $totalBytesRead = 0
            $bytesRead
            $first = $true
            $last = $false
            # Read data from file system in blocks.
            while (($bytesRead = $br.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $totalBytesRead = $totalBytesRead + $bytesRead
                # You’ve reached the end of the file.
                if ($totalBytesRead -eq $FileSize) {
                    $last = $true
                    # Copy to a new buffer that has the correct size.
                    $lastBuffer = New-Object System.Byte[]($bytesRead)
                    [array]::Copy($buffer, 0, $lastBuffer, 0, $bytesRead)
                }
                If ($first) {
                    $ContentStream = New-Object System.IO.MemoryStream
                    # Add an empty file.
                    $fileInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                    $fileInfo.ContentStream = $ContentStream
                    $fileInfo.Url = $List.RootFolder.ServerRelativeUrl + "/" + $UniqueFileName
                    $fileInfo.Overwrite = $true
                    $Upload = $Docs.RootFolder.Files.Add($fileInfo)
                    $ctx.Load($Upload)
                    # Start upload by uploading the first slice.
                    $s = [System.IO.MemoryStream]::new($buffer)
                    # Call the start upload method on the first slice.
                    $BytesUploaded = $Upload.StartUpload($UploadId, $s)
                    $ctx.ExecuteQuery()
                    # fileoffset is the pointer where the next slice will be added.
                    $fileoffset = $BytesUploaded.Value
                    # You can only start the upload once.
                    $first = $false
                }
                Else {
                    # Get a reference to your file.
                    $Upload = $ctx.Web.GetFileByServerRelativeUrl($Docs.RootFolder.ServerRelativeUrl + [System.IO.Path]::AltDirectorySeparatorChar + $UniqueFileName);
                    If ($last) {
                        # Is this the last slice of data?
                        $s = [System.IO.MemoryStream]::new($lastBuffer)
                        # End sliced upload by calling FinishUpload.
                        $Upload = $Upload.FinishUpload($UploadId, $fileoffset, $s)
                        $ctx.ExecuteQuery()
                        #Write-Host “File upload complete”
                        # Return the file object for the uploaded file.
                        return $Upload
                    }
                    else {
                        $s = [System.IO.MemoryStream]::new($buffer)
                        # Continue sliced upload.
                        $BytesUploaded = $Upload.ContinueUpload($UploadId, $fileoffset, $s)
                        $ctx.ExecuteQuery()
                        # Update fileoffset for the next slice.
                        $fileoffset = $BytesUploaded.Value
                    }
                }
            } #// while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
        }
        Catch {
            Write-Host “`t`tError occurred - $fileName”  -Fore Red
        }
        Finally {
            if ($null -ne $Fs) {
                $Fs.Dispose()
            }
        }
    }
    return $null
}

# Email Function
#function send_email ($exportPath, $email) 
function send_email () {
    $today = Get-Date
    $today = $today.ToString("dddd MMMM-dd-yyyy hh:mm tt")
    $SmtpClient = new-object system.net.mail.smtpClient 
    $mailmessage = New-Object system.net.mail.mailmessage 
    #$SmtpClient.Host = "outlook.office365.com"
    #$SmtpClient.Port = "587"
    #$SmtpClient.EnableSsl = "True"
    #$SmtpClient.Credentials = New-Object System.Net.NetworkCredential("User", "Password"); 
    #$SmtpClient.Host = "<SMTP Relay Server>"
    $SmtpClient.Host = $smtpserver
    #$SmtpClient.Port = "25"
    $SmtpClient.Port = $smtpport
    #$mailmessage.from = "Sender <sender@test.local>" 
    $mailmessage.from = $emailfrom 
    $mailmessage.To.add($email1)
    #$mailmessage.To.add($email2)
    $mailmessage.Subject = "PLEASE READ: GPO's have been changed."
    $mailmessage.IsBodyHtml = $true
    #$mailmessage.Attachments.Add($emailFile)
    $mailmessage.Attachments.Add($backupFolderPath + $backupFileName + "-GPOChanges.csv")
    $mailmessage.Body = @"
<!--<strong>GPO's have been Changed in $domain</strong><br />-->
GPO's have been changed in <span style="background-color:yellow;color:black;"><strong>$domain</strong></span>.<br /> <br />

Generated on : $today<br /><br />
<br /></font></h5>
"@

    $smtpclient.Send($mailmessage) 
}

# Set Variables
#Configure Email notification recipient
$smtpserver = "<SMTP Relay Server>"
$smtpport = "25"
$emailfrom = "Sender <sender@test.local>"
$email1 = "user1@test.local"
#$email2 = "user2@test.local"

# Send Email
#$sendEmail = "Yes"
$sendEmail = "No"

# Copy to SharePoint Library
#$useSharePoint = "Yes"
$useSharePoint = "No"
# Specify site URL
$SiteURL = "https://<Company>.sharepoint.com/sites/<Share>"
# Set SharePoint Folder
$DocLibName = "Group Policy Backup" # Document Libraty on SharePoint Site

# Delete Older Backups
#$deleteOlder = 'Yes' # Yes
$deleteOlder = 'No' # No

# Set min age of files
$max_days = '-7'

# Get the current date
$curr_date = Get-Date

# Determine how far back we go based on current date
$del_date = $curr_date.AddDays($max_days)

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

# Move GPOBackup off of System
#$moveBackups = "Yes"
$moveBackups = "No"

# Set Share Location - Used for Direct Share Access
$useShare = "Yes"
#$useShare = "No"
$year = get-date -Format "yyyy"
$shareLocation = "\\Server.local\Share\Folder\GPO"
$shareLocation = $shareLocation + "\" + $year

# Set Network Location - Used when Mapping a Drive
#$useMapShare = "Yes"
$useMapShare = "No"
$year = get-date -Format "yyyy"
#$networkDrive = "\\Server.local\Share\Folder"
$networkDrive = "\\Server.local\Share\Folder\GPO"
$networkDrive = $networkDrive + "\" + $year

# Set Drive Letter - Used when Mapping a Drive
$driveLetter = "Z"

# Set Folder Location - Used when Mapping a Drive - Needed if not directly mapping directly to full path
$folderLocation = "GPO"

# Combine Network Drive and Folder Location - Used when Mapping a Drive
if ($useMapShare -eq "Yes") {
    $shareLocation = $driveLetter + ":\" + $folderLocation #Used if not directly mapping directly to full path
    #$shareLocation = $driveLetter +":\" #Used if directly mapping directly to path
}

# Get Account to copy with - Used when Mapping a Drive
if ($useShare -eq "No") {
    $user = Read-Host "Enter User Name"# -AsString
    $password = Read-Host "Enter Password" -AsSecureString
    $mycreds = New-Object System.Management.Automation.PSCredential ($user, $password)
}

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
# Check if GPO Changes in last Day, Exit if no changes made in last day
Write-Host "`n`tPlease Wait - Checking for GPO Changes in the last 24 hours" -Fore Yellow
$modifiedGPOs = @(Get-GPO -All | Where-Object { $_.ModificationTime -ge $(Get-Date).AddDays(-1) }).count
If ($modifiedGPOs -eq "0") {
    Write-Host "`tNo Changes in last Day" -Fore Green
    Exit   #Exit if no changes made in last day
}
Write-Host "`n`tPlease Wait - GPO Changes in the last 24 hours" -Fore Yellow
Write-Host "`t`tGPO(s) Changes: $modifiedGPOs" -Fore Yellow


# Verify GPO BackupFolder
if ((Test-Path $backupFolderPath) -eq $false) {
    New-Item -Path $backupFolderPath -ItemType directory
}


# Generate List of changes
Write-Host "`n`tPlease Wait - Creating GPO Email Report" -Fore Yellow
Get-GPO -All | Where-Object { $_.ModificationTime -ge $(Get-Date).AddDays(-1) } | Export-csv $backupPath-GPOChanges.csv -NoTypeInformation
Write-Host "`t`tCreated GPO Email Report" -Fore Yellow


# Send email Notification
if ($sendEmail -eq "Yes") {
    Write-Host "`tPlease Wait - Sending Email Report" -Fore Yellow
    send_email
    Write-Host "`t`tSent Email Report" -Fore Yellow
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


# Orphaned GPOs


# Backup WMI Filters
Write-Host "`tPlease Wait - Backing up WMI Filters" -Fore Yellow
#$WMIFilters = @()
#$WmiFilters = Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' -Properties * | Select DistinguishedName, whenCreated, whenChanged, msWMI-Author, msWMI-ID, msWMI-Name, msWMI-Parm1, msWMI-Parm2
$WmiFilters = Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' -Properties * | Select-Object * 
$RowCount = $WMIFilters | Measure-Object | Select-Object -expand count
if ($RowCount -ne 0) {
    write-host -ForeGroundColor Green "`tExporting $RowCount WMI Filters"
    $WMIFilters | export-csv $backupPath-WMIFiltersExport.csv -NoTypeInformation
    #write-host -ForeGroundColor Green "An export of the WMI Filters has been stored at $backupPath-WMIFiltersExport.csv`n"
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
            #$foldername = join-path $backupPath ($gpo.displayname + "_{" + $gpo.Id + "}") # Keep " "
            $foldername = join-path $backupPath ($gpo.displayname) # Keep " "
            Get-GPOReport -Name $gpo.displayname -ReportType 'HTML'-Path $filename
        }
    }
    Else {
        $allGPOs = get-gpo -all
        foreach ($gpo in $allGPOs) {
            Write-Host "`t`tProcessing GPO" $gpo.displayname -Fore Yellow
            #$foldername = join-path $backupPath ($gpo.displayname.Replace(" ", "_") + "_{" + $gpo.Id + "}") # Replace " " with "_"
            $foldername = join-path $backupPath ($gpo.displayname + "_{" + $gpo.Id + "}") # Keep " "
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


#<#
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
#>


# Backup PolicyDefinition Folder
Write-Host "`tPlease Wait - Backing up PolicyDefinition Folder" -Fore Yellow
$policydefinitionSource = "\\" + $env:USERDOMAIN + "\SYSVOL\" + $env:USERDNSDOMAIN + "\Policies\PolicyDefinitions"
Copy-Item -Path $policydefinitionSource -Recurse -Destination $backupPath -Container
Write-Host "`t`tBacked up PolicyDefinition Folder" -Fore Yellow


# Compress Folders
Write-Host "`tPlease Wait - Checking for 7-Zip" -Fore Yellow
# Path to 7-Zip
$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"
# Compress Folders to 7-Zip File
if ((Test-Path $7zipPath) -eq $true) {
    Write-Host "`tPlease Wait - Creating 7-ZIP File" -Fore Yellow
    # Create Alias
    Set-Alias Compress-7Zip $7ZipPath
    # Set Source & Destination
    $source = $backupPath
    $destination = $backupPath + ".7z"
    # Compress Files/Folder
    Compress-7zip a -mx9 -r -tzip $destination $source
}

# Compress Folders to Zip File
if ((Test-Path $7zipPath) -eq $false) {
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
}


# Delete GPO Backup Folder
#Write-Output "`tPlease Wait - Deleting GPO Backup Folder"
Write-Host "`tPlease Wait - Deleting GPO Backup Folder" -Fore Yellow
Remove-item -Path $backupPath -Recurse -Force -ErrorAction SilentlyContinue


# Calculate Time for Files to Copy to
$updload_date = $curr_date.AddHours(-1)


# Copy to SharePoint
if ($useSharePoint -eq "Yes") {
    # Upload file(s)
    # Add references to SharePoint client assemblies and authenticate to Office 365 site – required for CSOM
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

    # Specify username site URL
    $User = $env:username +"@<Domain>"

    # Get Password
    $Credentials = Get-Credential "$User" -ErrorAction Stop

    # Split username and password
    $Username = $credentials.username
    $Password = $credentials.GetNetworkCredential().password

    # Convert Password to Secure String
    $Pass = ConvertTo-SecureString $Password -AsPlainText -Force

    # Bind to site collection
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User, $Pass)
    $Context.Credentials = $Creds

    # Retrieve list
    $List = $Context.Web.Lists.GetByTitle($DocLibName)
    $Context.Load($List.RootFolder)
    $Context.ExecuteQuery()

    # Upload file(s)
    Write-Host "`n`tCopying the files - Please Wait" -Fore Yellow
    #Foreach ($File in (Get-ChildItem -Path $backupFolderPath -File -Recurse)) {
    #Foreach ($File in (Get-ChildItem -Path $backupFolderPath -File)) {
    Foreach ($File in (Get-ChildItem -Path $backupFolderPath -File | Where-Object { $_.LastWriteTime -gt $updload_date })) {
        Write-Host "`t`tCopying the file: "$File.FullName -Fore Yellow
        #$UpFile = UploadFileInSlice -ctx $Context -libraryName $DocLibName -fileName $File.FullName
        $fileName = $File.FullName
        $UpFile = UploadFileInSlice -ctx $Context -libraryName $DocLibName -fileName $fileName
        $Context.Dispose();
    }
}


# Delete Old Backup Files
If ($deleteOlder -eq 'Yes') {
    Write-Output '`tDeleting older GPO Backup files' -Fore Yellow
    Get-ChildItem $backupFolderPath -Recurse | Where-Object { $_.LastWriteTime -lt $del_date } | Remove-Item
}


# Complete if not moving off of System
if ($moveBackups -eq "No") {
    #Write-Output "`tGPO Backup - Complete"
    Write-Host "`tGPO Backup - Complete" -Fore Yellow
    Exit
}


# Copy/Move to File Share
if ($useShare -eq "Yes") {
    Write-Output "`tPlease Wait - Moving GPO Backup Files to Network Backup Folder" -Fore Yellow
    #Get-ChildItem $backupFolderPath -Recurse | Copy-Item -Destination $shareLocation # Copy Backups
    Get-ChildItem $backupFolderPath -Recurse | Where-Object { $_.LastWriteTime -gt $updload_date } | Copy-Item -Destination $shareLocation # Copy Backups
    Get-ChildItem $backupFolderPath -Recurse | Move-Item -Destination $shareLocation # Move Backups
    Write-Output "`t`tCompleted Moving GPO Backup Files to Network Backup Folder" -Fore Yellow
}


# Copy to Mapped Network Drive
if ($useMapShare -eq "Yes") {
    #Net Use $driveLetter $networkDrive /User:$user $pwd
    New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkDrive -Credential $mycreds
    #Copy/Move GPO`t Backups
    Write-Output "`tPlease Wait - Moving GPO Backup Files to Network Backup Folder" -Fore Yellow
    #Get-ChildItem $backupFolderPath -Recurse | Copy-Item -Destination $shareLocation # Copy Backups
    Get-ChildItem $backupFolderPath -Recurse | Where-Object { $_.LastWriteTime -gt $updload_date } | Copy-Item -Destination $shareLocation # Copy Backups
    Get-ChildItem $backupFolderPath -Recurse | Move-Item -Destination $shareLocation # Move Backups
    Write-Output "`t`tCompleted Moving GPO Backup Files to Network Backup Folder" -Fore Yellow
    #Disconnect Network Drive
    #Net Use $driveLetter /D
}


# Completed Script
Write-Output "`tGPO Backup - Complete" -Fore Yellow
