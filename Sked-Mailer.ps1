<#
.Synopsis
Gathers file names from variable defined folder and emails them to owner.  

.DESCRIPTION
This script will enumerate the directory tree specified in the hoemdrive variable, discover all files contained in the folder defined in the $outbound variable, and who their owner is.  Then it will look up the file owner's email address in Active Directory and mail the files to them one at a time.  Finally it deletes the discovered files.

.NOTES   
Name: Sked-Report-mailer
Author: Lee Buskey
Version: 1.0
DateCreated: 2015-11-17
DateUpdated: 


.EXAMPLE
Sked-Report-mailer

Still to do:  Error checking, Logging, email signing, picking up email.. 


#>

# Import required modules,. 

Import-Module activedirectory
Import-Module pslogging


# specify needed parameters 

# This should be the Windows domain's netbios name, and requires the trailing backslash in the string. This is a case sensitive strin  

$Reports_url = '\\spablskdbc8r01.sked.ablu.navy.mil\Sked 3.2\Feedback Reports'
$domain =  'SKED\'  			
# This should be the path or URL to the root of the users home directories.   
$homedrive = '\\spablskxa8r01\homedrives' 			
# This is the name of the folder in the users homedrive that you want this script to look in for files to send.  
$outbound = 'outbound' 
# The address of the SMTP server you are allowed to relay through.
$PSEmailServer = "10.2.1.9"								
# The utterly fictitious address that the mailed messages should appear to be from.  
$from = "noreply@sked.abl.navy.mil"									 
#Subject of the emails
$subject = "[SKED] Your files are attached - DEMO DEMO DEMO - This is not official data. - DEMO DEMO DEMO"	
#Body of the email 	
$body = "DEMO DEMO DEMO - This is not official data.  `r`n`n This message was sent from an unmonitored account.  Messages sent to this address including return replies to this message will not be received. `r`n`n For questions or support on this system, please refer to your SKED help desk."
# specify where all the logs will go
$logpath = "d:\Logs"
# specify the text suffix to append to log files. 
$logsuffix = "SKED_Mailer.log"
#Generate a timestap for runlog file. 
$timeStamp = Get-Date -format yyyyMd-hhmmss
#Generaste a runlog filename with the time spamp and the suffix. 
$logname = "$timeStamp"+"$logsuffix"
#Build a full path to the logfile
$logfile = "$LogPath"+"\"+"$logname"
#Get a date stamp for the daily log. 
$DailyLog = Get-Date -format yyyyMd
#Buid a log file for the daily log
$dailyLogFile = "$logpath"+"\"+"$DailyLog"+"_"+"$logsuffix"

$ScriptVersion = "1.2"

# Test for log path, create it if not present. 

If (Test-Path $logpath){
  
}Else{
 New-Item -path D:\logs -ItemType Directory
}

# Test for daily log, create it if not present. 

If (Test-Path $dailyLogFile) {
  
}Else{
 New-Item -path $dailyLogFile -ItemType File
}


#Start the Log file for this run
Start-Log -LogPath $logpath -LogName $Logname -ScriptVersion $ScriptVersion


#Find the folders named after the $outbound variable, then find the files in any of those folders that arent *.xls, *.xlsx, *.pdf, and create a variable named $Badfiles containing the path to the files discovered
$Badfiles = Get-childitem -r $homedrive -directory | where name -eq $outbound | get-childitem -r -exclude *.xls, *.xlsx, *.pdf, *.txt, *.mdb, *.dat, *.tfd, *.sug, *.sklcs, *.sked, *.skedb, *.sxdb, *.skrpt, *.xml  |Get-Acl | select @{n="Path";e={Convert-Path ($_.Path)}}, @{n="Owner";e={$_.Owner.replace($domain,"")}},PsChildName

#test to see if any files were discoved, Log messages... 
if ($badfiles -ne $Null) { Write-LogWarning -ToScreen -LogPath $logfile -Message "$(Get-Date -format yyyyMd-hhmmss) : The following disallowed files were found in the outbound folder, they will be deleted"

#Convert list to string
$logmessage = $Badfiles | Format-List | Out-String 

#Log the bad files
Write-LogWarning -ToScreen -LogPath  $logfile -Message "$(Get-Date -format yyyyMd-hhmmss) : $logmessage" 

#clean-up the bad files and log activity 
$Badfiles | foreach ($_ = (Remove-item $badfiles.Path -force))
Write-LogWarning -toscreen -LogPath  $logfile -Message "$(Get-Date -format yyyyMd-hhmmss) : Delteted disallowed file:  $($Badfiles.PsChildName)"
}
else { Write-host ""
       Write-host "No disallowed files were found in the outbound folder"}

#Now look for file to mail.  
#Find the folders named after the $outbound variable, and find the files that match the filter and create a variable $DiscoveredFiles containing the path to the files including file name, and the owner of the files.   
$DiscoveredFiles = Get-childitem -r $homedrive -directory | where name -eq $outbound | get-childitem -r -file -include *.xls, *.xlsx, *.pdf, *.txt, *.mdb, *.dat, *.tfd, *.sug, *.sklcs, *.sked, *.skedb, *.sxdb, *.skrpt, *.xml  |Get-Acl | select @{n="Path";e={Convert-Path ($_.Path)}}, @{n="Owner";e={$_.Owner.replace($domain,"")}}, PSchildName

#test to see if any files were discovered, Log messages... 
if ($DiscoveredFiles -ne $Null) {$DiscoveredFiles | add-member -membertype noteproperty -name Email_Address  -value NotSet | foreach {$_.Email_Address = (get-aduser $_.Owner -Properties mail | select -ExpandProperty mail )} | foreach {$_.Email_Address = (get-aduser $_.Owner -Properties mail | select -ExpandProperty mail )} 

#for each file identified in $DiscoveredFiles, call the Get-Aduser cmdlet to retrieve the users e-mail address, and populate the Email_Address property for that item.     
$DiscoveredFiles | foreach {$_.Email_Address = (get-aduser $_.Owner -Properties mail | select -ExpandProperty mail )} 

# Log list header 
Write-LogInfo -toscreen -logpath $logfile -Message "$(Get-Date -format yyyyMd-hhmmss) : The following files were found to email"

#Convert list to string
$logmessage = $Discoveredfiles | Format-List | Out-String 

#Log what we found
Write-LogInfo -toscreen -logpath $logfile -Message "$(Get-Date -format yyyyMd-hhmmss) : $logmessage"


#Send the files to the users via email.


foreach ($D in $DiscoveredFiles) { 
 
    Try {

              Send-MailMessage -From $from -To $D.Email_Address -attachments $D.Path -Subject $subject -Body $body -ErrorAction Stop
              
        }

        Catch {

           Write-LogError -LogPath $logfile -ToScreen  -Message "$(Get-Date -format yyyyMd-hhmmss) : SMTP error sending: $($D.PsChildName) owned by $($D.owner) to $($D.Email_Address).  The error message is: $_.  Skipping file for next run"
            continue
            

    }  
     Write-LogInfo -LogPath $logfile -ToScreen -Message "$(Get-Date -format yyyyMd-hhmmss) : Mail sent.  The file: $($D.PsChildName |out-string -stream) was sent to $($D.Email_Address)"
    Remove-Item $D.Path -force

}}

#log activity
else { Write-Host "No data files were found in the outbound folder"}


#Find the folders named after the $outbound variable, then find the files in any of those folders that arent *.xls, *.xlsx, *.pdf, and create a variable named $Badfiles containing the path to the files discovered
$Badfiles = Get-childitem -r $homedrive -directory | where name -eq $outbound | get-childitem -r -exclude *.xls, *.xlsx, *.pdf  |Get-Acl | select @{n="Path";e={Convert-Path ($_.Path)}}, @{n="Owner";e={$_.Owner.replace($domain,"")}},PsChildName

#test to see if any files were discoved, Log messages... 
if ($badfiles -ne $Null) { Write-LogWarning -ToScreen -LogPath $logfile -Message "$(Get-Date -format yyyyMd-hhmmss) : The following disallowed files were found in the outbound folder, they will be deleted"

#Convert list to string
$logmessage = $Badfiles | Format-List | Out-String 

#Log the bad files
Write-LogWarning -ToScreen -LogPath  $logfile -Message "$(Get-Date -format yyyyMd-hhmmss) : $logmessage" 

#clean-up the bad files and log activity 
$Badfiles | foreach ($_ = (Remove-item $badfiles.Path -force))
Write-LogWarning -toscreen -LogPath  $logfile -Message "$(Get-Date -format yyyyMd-hhmmss) : Delteted disallowed file:  $($Badfiles.PsChildName)"
}
else { Write-Host "No disallowed files were found in the Reports folder"}

#Now look for Reports to mail.  
#Find the files in the Database Share.. Create a variable $DiscoveredReports containing the path to the files including file name, and the owner of the files.   
$DiscoveredReports = Get-childitem -r $Reports_url -file -include *.xls, *.xlsx, *.pdf, *.txt, *.mdb, *.dat, *.tfd, *.sug, *.sklcs, *.sked, *.skedb, *.sxdb, *.skrpt, *.xml  |Get-Acl  | select @{n="Path";e={Convert-Path ($_.Path)}}, @{n="Owner";e={$_.Owner.replace($domain,"")}}, PSchildName

#test to see if any files were discovered, Log messages... 
if ($DiscoveredReports -ne $Null) {$DiscoveredReports | add-member -membertype noteproperty -name Email_Address  -value NotSet | foreach {$_.Email_Address = (get-aduser $_.Owner -Properties mail | select -ExpandProperty mail )} | foreach {$_.Email_Address = (get-aduser $_.Owner -Properties mail | select -ExpandProperty mail )} 


#for each file identified in $DiscoveredReports, call the Get-Aduser cmdlet to retrieve the users e-mail address, and populate the Email_Address property for that item.     
$DiscoveredReports | foreach {$_.Email_Address = (get-aduser $_.Owner -Properties mail | select -ExpandProperty mail )} 

# Log list header 
Write-LogInfo -toscreen -logpath $logfile -Message "$(Get-Date -format yyyyMd-hhmmss) : The following Reports were found to email"

#Convert list to string
$logmessage = $DiscoveredReports | Format-List | Out-String 

#Log what we found
Write-LogInfo -toscreen -logpath $logfile -Message "$(Get-Date -format yyyyMd-hhmmss) : $logmessage"
}

else { Write-Host 'No Reports were found in the Reports folder'}

#Send the files to the users via email.


foreach ($D in $DiscoveredReports) { 
 
    Try {

              Send-MailMessage -From $from -To $D.Email_Address -attachments $D.Path -Subject $subject -Body $body -ErrorAction Stop
              
        }

        Catch {

           Write-LogError -LogPath $logfile -ToScreen  -Message "$(Get-Date -format yyyyMd-hhmmss) : SMTP error sending: $($D.PsChildName) owned by $($D.owner) to $($D.Email_Address).  The error message is: $_.  Skipping file for next run"
            continue
            

    }  
     Write-LogInfo -LogPath $logfile -ToScreen -Message "$(Get-Date -format yyyyMd-hhmmss) : Mail sent.  The file: $($D.PsChildName |out-string -stream) was sent to $($D.Email_Address)"
    Remove-Item $D.Path -force



#log activity
else { Write-Host "No Reports were found in the outbound folder"}
}


#stop log
Stop-Log -toscreen -logpath $logfile -Noexit
#Append RunLog to DailyLog and delete RunLog.. 
Get-content $logfile >> $dailyLogFile
Remove-Item $logfile
