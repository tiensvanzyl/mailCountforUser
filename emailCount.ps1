###########################################################################################
#   Get the amount of emails sent and received by a specific user                         #
#   Author: Tiens van Zyl                                                                 #
#   Email: tiens.vanzyl@gmail.com                                                         #
#   Date:   30 April 2018                                                                 #
#   Updated by:                                                                           #
#   Updates:                                                                              #
#   Version: 0.1                                                                          #
#   External link: (credit) http://techgenix.com/numberofe-mailssentandreceivedbyoneuser/ #
#   Description:                                                                          #        
#       This script displays the total mail count for mail sent and received for a        #
#        specific user specified by start and end date.                                   #
#        Adapted the code from the above link                                             #
###########################################################################################

#Prompts the user for the Hub Transport Server names. Use wild card if more than on HT server exists i.e. HTServer0*
$Server = Read-Host -Prompt 'Input your server name'

#Type in the start date i.e. 02/04/2018 (month/day/year)
$StartDate = Read-Host -Prompt 'Enter the start date (month/day/year)'

#Type in the end date i.e. 02/04/2018 (month/day/year)
$EndDate = Read-Host -Prompt 'Enter the end date(month/day/year)'  

#Type the email address of the user you need to check the mail count for
$User = Read-Host -Prompt 'Enter the email address for the user'

#Array used to add the total count of mail sent and received for the user. Do not change this.
[Int] $mailSent = $mailReceived = 0

#Runs the message track for mail sent and adds the total to the mailSent variable
Get-TransportServer $Server | Get-MessageTrackingLog -ResultSize Unlimited -Start $StartDate -End $EndDate -Sender $User -EventID RECEIVE | ? {$_.Source -eq "STOREDRIVER"} | ForEach { $mailSent++ }

#Runs the message track for mail received and adds the total to the mailReceived variable
Get-TransportServer $Server | Get-MessageTrackingLog -ResultSize Unlimited -Start $StartDate -End $EndDate -Recipients $User -EventID DELIVER | ForEach { $mailReceived++ }

#Outputs the total mail sent and received for the user to the console.
Write-Host "E-mails sent:    ", $mailSent

Write-Host "E-mails received:", $mailReceived