  Import-Module ActiveDirectory
  
  $CurrentDate = (Get-Date).Month
  $DateNow = (Get-Date).AddDays(-90)

  $user = Get-ADUser -Filter * -Properties * | 
        Select @{Name="Name";Expression={$_.Name}},
        lastlogondate, Enabled | 
        Where {($_.LastLogonDate) -le $DateNow -and ($_.Enabled -eq 'true')} | Sort LastLogonDate |
        Export-CsV C:\Windows\LTSvc\"$env:USERDOMAIN$CurrentDate".csv