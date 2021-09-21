##################################
# update email addresses!
#
$toEmail = ""
$ccEmail = ""
#
$TodayDate = (Get-Date -UFormat %m-%d-%Y)
$TodayShortDate = (Get-Date -UFormat %m%d%Y)
$TodayTime = (Get-Date -UFormat %H:%M:%S)
$TodayShortTime = (Get-Date -UFormat %H%M%S)
#
$File1 = "ser7.dat"
$File2 = "finals.daily"
#
# Ordered dictionary:
$URLs = [Ordered]@{$File1="https://datacenter.iers.org/data/latestVersion/bulletinA.txt"; $File2="https://datacenter.iers.org/data/latestVersion/finals.daily.iau1980.txt"}
#
$OutPath = "$Env:homeshare\IERS"
$ArchivePath = "$OutPath\archive"
$LogFile = "$OutPath\_iers.log"
#
# Email Body Creation:
$BodyNote = ""
#
###################################

if ($toEmail -eq "" -Or $ccEmail -eq "" ){
	clear
	echo "`n`n`tUpdate the email addresses before using`n`n"
	exit
}

# Test if the directory(ies) exists; if not, create them.
if(!(Test-Path $ArchivePath)) { 
	New-Item -ItemType Directory -Force -Path $ArchivePath
}

function iersArchive {
	param ( 
		$FileName 
	)
	if (Test-Path $OutPath\$FileName".bak") {
		$NewFileName = "$FileName.$TodayShortDate.$TodayShortTime"
		Move-Item $OutPath\$FileName".bak" $ArchivePath\$NewFileName
	}
}

function iersBackup {
	param ( 
		$FileName
	)
	if (Test-Path $OutPath\$FileName) {
		Rename-Item -Path $OutPath\$FileName -Force -NewName $FileName".bak"
	}
}

function iersChecksum {
	param (
		$FileName
	)
	$before = (certutil -hashfile $OutPath\$FileName".bak" MD5 | find /v /i '"md5"' | find /v /i '"certutil"')
	$after = (certutil -hashfile $OutPath\$FileName MD5 | find /v /i '"md5"' | find /v /i '"certutil"')
	if ( $before -eq "" -Or $after -eq "" ) {
		Write-Output "  One or more checksums failed"
	}
	if ( $before -eq $after ) {
		$TempNote += "    $FileName has not changed since the last pull`n"
		Write-Output "$TodayDate - $TodayTime : $FileName has not changed since the last pull" | Out-File -Append -FilePath $LogFile
	}
	else {
		$TempNote += "    $FileName appears to have been updated`n" 
		Write-Output "$TodayDate - $TodayTime : $FileName appears to have been updated" | Out-File -Append -FilePath $LogFile
	}
    return $TempNote
}

function iersEmail {
	# Create new email message, attach the files 
	# and open it for the user to send:
	$Outlook = New-Object -ComObject Outlook.Application
	$Mail = $Outlook.CreateItem(0)
	$Mail.To = "$toEmail"
	$Mail.Cc = "$ccEmail"
	$Mail.Subject = "IERS For $TodayDate"
	$Mail.Body = $BodyNote
	$Mail.attachments.add("$OutPath\$File1")
	$Mail.attachments.add("$OutPath\$File2")

	$Mail.save()
	$Inspector = $Mail.GetInspector
	$Inspector.Display()
}



###########################
### Do Stuff and Things ###


foreach ( $FileName in $URLs.keys ) {
	# Archive Old Backup Files
	iersArchive -FileName $FileName
	
	# Create Backups of Yesterday's Files
	iersBackup -FileName $FileName 
	
	# IERS Download:
	Invoke-WebRequest -uri $URLs[$FileName] -Outfile $OutPath\$FileName
	
	# Retreives the date from the ser7.dat file for use in the message body:
	if ( $FileName -eq $File1 ) {
		$FileDate = (Get-Content $OutPath\$File1 | Select-Object -skip 7 -first 1).substring(6,18)
		$BodyNote += "IERS file date is : $FileDate`n`nNotes:`n"
	}
	
	# Message Body Note(s):
	$BodyNote += iersChecksum -FileName $FileName
}


# Padding the EOF to leave room for classification
$BodyNote += "`n`n`n`n"

# Create and send email:
iersEmail


### Do Stuff and Things ###
###########################
