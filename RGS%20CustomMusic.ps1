########################################################################
# Author: Paul B
# Blog: lyncsorted.blogspot.com
# Purpose: Script to Change Response Group Workflow Hold Audio
# Version: 1.0
# Date: 05/09/2013
#
# Note
#   This script was created to allow a customer to quickly change the audio file
#   (MOH) used by a particular RGS workflow. The idea was initially for an IT Service Desk
#   that would announce known outages etc by means of their hold queue music. 
#     
#  Of course the script can be modified to cycle through marketing or sales promotions.
#  Perhaps team announcements or reminders etc.    
#
#  Select the required Audio Announcement file in the Menu list, or selecting a 
#  custom audio file to set as the Workflows "CustomMusicOnHold".
#
# Instructions
# Edit the below variables to match your requirements as follows:-
#  1. RGSWorkflow MUST match the name of the Lync RGS Workflow to be changed
#  2. lyncpool MUST be in the format Service:ApplicationServer:FEPool.domain.local
#       Where FEPool.domain.local is substituted for your FE Pools FQDN
#  3. logfile is the location of the logs for this script
#  4. AudioDirectory is the location of the pre-recorded Audio file
#
# Please feel free to edit the Menu items to match your own requirements
#
#Version History
# V1.0 first creation
#
########################################################################


Import-Module Lync

#Variables 
$RGSWorkflow = "TestWorkflow" #Lync RGS Workflow Name
$lyncpool = "Service:ApplicationServer:lyncFE1.lyncsorted.local" #<Front End Pool FQDN>
$logfile = "e:\LXLSupport\RGS\Audio_change_log.txt"
$AudioDirectory = "e:\LXLSupport\RGS\"

# Menu
cls
Write-Host "* Change" $RGSWorkflow"Audio Announcement *" -ForegroundColor Cyan
Write-Host " "
Write-Host "Change the" $RGSWorkflow "Audio Announcement to:"
Write-Host " "
Write-Host "1. Network Issues Audio Announcement" -foregroundcolor Yellow
Write-Host "2. Email Issues Audio Announcement" -foregroundcolor Yellow
Write-Host "3. Internet Issues Audio Announcement" -foregroundcolor Yellow
Write-Host "4. Login Issues Audio Announcement" -foregroundcolor Yellow
Write-Host "5. Printer Issues Audio Announcement" -foregroundcolor Yellow
Write-Host "6. New Custom Audio Announcement" -foregroundcolor Yellow
Write-Host "7. Reset to Default Audio Announcement" -foregroundcolor Yellow
Write-Host " "
$a = Read-Host "Select 1-7: "
Write-Host " "

# Logic to set Audio File based on menu selection
switch ($a)
{  1 { $SelectedAudioFileName = "NetworkIssues.wav"
      }
   2 { $SelectedAudioFileName = "EmailIssues.wav"
      }
   3 { $SelectedAudioFileName = "InternetIssues.wav"
      }
   4 { $SelectedAudioFileName = "LoginIssues.wav"
      }
   5 { $SelectedAudioFileName = "PrinterIssues.wav"
      }
   6 {
      # Custom Audio
        Write-Host "Type the name of the custom Audio file - make sure the file exisits in " $AudioDirectory  -ForegroundColor Magenta -NoNewLine
        Write-Host " "        
        $SelectedAudioFileName = Read-Host "Audio File"
       }
   7 { $SelectedAudioFileName = "ServiceDesk.wav"
      }
}

#Load the Audio File - change the details as required
$Audiopath = $AudioDirectory+$SelectedAudioFileName
$Audiopath = Get-Content -ReadCount 0 -Encoding Byte $AudioPath | Import-CsRgsAudioFile -Identity $lyncpool -FileName $SelectedAudioFileName

#Set the RGS Workflow to use the new Audio file
$b = Get-CsRgsWorkflow -Identity $lyncpool -Name $RGSworkflow
$b.CustomMusicOnHoldFile = $Audiopath
Set-CsRgsWorkflow -Instance $b

# Write to log file
$date = Get-Date
Write-output "$date, $RGSWorkflow , $SelectedAudioFileName" | Out-File -append $logfile

Write-Host "  "
Write-Host " The hold music for the RGS Workflow " $RGSWorkflow -ForegroundColor Green " has been changed to: " -NoNewline
Write-Host $SelectedAudioFileName -ForegroundColor Green
Write-Host "  "