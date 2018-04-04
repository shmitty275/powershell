<#   
.SYNOPSIS   
Active Directory Script that queries for user accounts that have unchanged passwords for the past 90 days
    
.DESCRIPTION 
This script will return the samaccountname, pwdlastset and if an account is currently enabled or disabled. This script is part of the Active Directory Friday section of my blog.

.NOTES   
Name: Get-UnchangedPwdLastSet.ps1
Author: Jaap Brasser
DateCreated: 2013-07-26
DateUpdated: 2013-07-26
Site: http://www.jaapbrasser.com

.LINK
http://www.jaapbrasser.com/active-directory-friday-find-user-accounts-that-have-not-changed-password-in-90-days/

.PARAMETER PwdAge
The number of days since the password has been changed. This value defaults to 90.
#>
param (
    $PwdAge = 180
)
$PwdDate = (Get-Date).AddDays($PwdAge).ToFileTime()
$Searcher = New-Object DirectoryServices.DirectorySearcher -Property @{
    Filter = "(&(objectclass=user)(objectcategory=person)(pwdlastset<=$PwdDate))"
    PageSize = 500
    SizeLimit = 200
}
$Searcher.FindAll() | ForEach-Object {
    New-Object -TypeName PSCustomObject -Property @{
        samaccountname = $_.Properties.samaccountname -join ''
        pwdlastset = [datetime]::FromFileTime([int64]($_.Properties.pwdlastset -join ''))
        enabled = -not [boolean]([int64]($_.properties.useraccountcontrol -join '') -band 2)
    }
}