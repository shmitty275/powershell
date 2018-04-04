################################################################################################################################################################ 
# Script accepts 2 parameters from the command line 
# 
# Office365Username - Optional - Administrator login ID for the tenant we are querying 
# Office365AdminPassword - Optional - Administrator login password for the tenant we are querying 
# 
# 
# To run the script 
# 
# .\Get-ClutterDetails.ps1 -Office365Username admin@xxxxxx.onmicrosoft.com -Office365AdminPassword Password123 
# 
# NOTE: If you do not pass an input file to the script, it will show if clutter has been enabled or disabled on each of your mailboxes in the tenant.  Not advisable for tenants with large 
# user count (< 3,000)  
# 
# Author:                 Alan Byrne 
# Version:                 1.0 
# Last Modified Date:     19/05/2015
# Last Modified By:     Alan Byrne 
################################################################################################################################################################ 
 
#Accept input parameters 
Param( 
    [Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Office365Username, 
    [Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)] 
    [string] $Office365AdminPassword
) 
 
#Constant Variables 
$OutputFile = "ClutterDetails.csv"   #The CSV Output file that is created, change for your purposes 
 

#Remove all existing Powershell sessions 
Get-PSSession | Remove-PSSession 
  
#Did they provide creds?  If not, ask them for it.
if (([string]::IsNullOrEmpty($Office365Username) -eq $false) -and ([string]::IsNullOrEmpty($Office365AdminPassword) -eq $false))
{
    $SecureOffice365AdminPassword = ConvertTo-SecureString -AsPlainText $Office365AdminPassword -Force     
     
    #Build credentials object 
    $Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365Username, $SecureOffice365AdminPassword 
}
else
{
    #Build credentials object 
    $Office365Credentials  = Get-Credential
}

#Create remote Powershell session 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Office365credentials -Authentication Basic –AllowRedirection         

#Import the session 
Import-PSSession $Session -AllowClobber | Out-Null                     
 
#Prepare Output file with headers 
Out-File -FilePath $OutputFile -InputObject "UserPrincipalName,SamAccountName,ClutterEnabled" -Encoding UTF8 
 
#Get all users 
#No input file found, gather all mailboxes from Office 365 
write-host "Retrieving Users"
$objUsers = get-mailbox -ResultSize Unlimited | select UserPrincipalName, SamAccountName 
 
#Iterate through all users     
Foreach ($objUser in $objUsers) 
{     
    #Prepare UserPrincipalName variable 
    $strUserPrincipalName = $objUser.UserPrincipalName 
    $strSamAccountName = $objUser.SamAccountName 
    
    write-host "Processing $strUserPrincipalName"
    #Get Clutter info to the users mailbox 
    $strClutterInfo = $(get-clutter -Identity $($objUser.UserPrincipalName)).isenabled  
    
    #Prepare the user details in CSV format for writing to file 
    $strUserDetails = "$strUserPrincipalName,$strSamAccountName,$strClutterInfo"
     
    #Append the data to file 
    Out-File -FilePath $OutputFile -InputObject $strUserDetails -Encoding UTF8 -append 
} 

write-host "Completed - data saved to $OutputFile"
 
#Clean up session 
Get-PSSession | Remove-PSSession 