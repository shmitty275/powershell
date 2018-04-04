#Input variables, largely self-explanatory in this example 
$DomainName = "contoso"
$DNSDomainName = "contoso.com"
$ReplSourceDC = "dc1.contoso.com"
$SafeModePassword = "MyS&cure@ssw0rd!"

#Rename the computer and reboot
Rename-Computer -NewName dc2
Restart-Computer -Force 

#Add the computer to the domain (if necessary) and reboot.
#Mine is joined as part of Azure provisioning process, so I skip this step
Add-Computer -DomainName $DomainName -Credential (Get-Credential)
Restart-Computer -Force

#Install the Active Directory feature 
Install-WindowsFeature -Name AD-Domain-Services

#Convert password to a secure string 
$Password = ConvertTo-SecureString -AsPlainText -String $SafeModePassword -Force 

#Install as a domain controller in specified domain, then reboot
Install-ADDSDomainController -DomainName $DNSDomainName -DatabasePath "%SYSTEMROOT%\NTDS" -LogPath "%SYSTEMROOT%\NTDS" -SysvolPath "%SYSTEMROOT%\SYSVOL" -InstallDns -ReplicationSourceDC $ReplSourceDC -SafeModeAdministratorPassword $Password -NoRebootOnCompletion
Restart-Computer -Force

#Retrieve members of "Domain Controllers" to ensure new DC is present
Get-ADGroupMember "Domain Controllers"