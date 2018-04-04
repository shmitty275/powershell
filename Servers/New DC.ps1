
########################################
# PowerShell Script to Install Domain Controllers #
########################################

Import-Module ADDSDeployment
Install-ADDSDomainController `
-NoGlobalCatalog:$false `
-InstallDns:$false `
-CreateDnsDelegation:$false `
-CriticalReplicationOnly:$false `
-DatabasePath "C:\Windows\NTDS" `
-LogPath "C:\Windows\NTDS" `
-SysvolPath "C:\Windows\SYSVOL" `
-DomainName "contoso.local" `
-NoRebootOnCompletion:$false `
-SiteName "SiteName" `
-Force:$true
