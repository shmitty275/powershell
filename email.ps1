# Function to create report email
function SendNotification{
 $Msg = New-Object Net.Mail.MailMessage
 $Smtp = New-Object Net.Mail.SmtpClient($ExchangeServer)
 $Msg.From = $FromAddress
 $Msg.To.Add($ToAddress)
 $Msg.Subject = "Announcement: One Drive Phishing Email"
 $Msg.Body = $EmailBody
 $Msg.IsBodyHTML = $true
 $Smtp.Send($Msg)
}
 
# Define local Exchange server info for message relay. Ensure that any servers running this script have permission to relay.
$ExchangeServer = "192.168.1.30"
$FromAddress = "rlp@apbcpa.com"
 
# Import user list and information from .CSV file
$Users = Import-Csv RLPContacts.csv
 
# Send notification to each user in the list
Foreach ($User in $Users) {
 $ToAddress = $User.Email
 $Name = $User.FirstName
 $EmailBody = @"
 <html>
 <head>
 </head>
 <body>
 <p>Dear $Name,</p>
 
 <p>You may have received an email from me regarding "Action Required One Drive" and attached document. This is not a legitimate email authorized by myself.</p>
 
 <p>Please delete this email. If you have opened the link or have entered any credentials on the site, please change any password immediately.</p>
  
 <p>Regards,</p>
 
 <p>Richard Payne</p>
 </body>
 </html>
"@
 Write-Host "Sending notification to $Name ($ToAddress)" -ForegroundColor Yellow
 SendNotification
}