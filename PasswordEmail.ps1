<#
.SYNOPSIS
    Scan Active Directory for passwords that are expiring in the designated time
    frame and send that user an Email warning them of it.
.DESCRIPTION
    Designate the timeframe you want emails to be sent to the users, as well as
    individual days out.  IE send an email at 15 and 10 days that their password is
    expiring, then every day from 5 down.
    
    All configuration changes are made in the pwdcfg.ini file.  Script will look for
    the file in the same path as where the script is located.  If it does not find it
    it will create a baseline configuration file and launch Notepad for you to edit it.
    
    A log file of each run and the emails sent out will also be kept as pwdExpired.log.
    
    A demo mode is available so all emails will be sent to the designated person instead
    of all users so you can test before deploying in production.
    
       
    Settings that must be configured in the pwdcfg.ini file:
    
    SMTPRelay                   The IP address or hostname of your SMTP relay
    SMTPFrom                    Who the emails will be sent out as
    SMTPPort                    Port your SMTP relay uses, default is 25
    UseSSL                      Set to Yes if your SMTP relay requires SSL
    UseCredentials              Set to Yes or No.  Tells the script to use credentials for 
                                sending email.  Script will prompt you for those credentials
                                the first time you run the script, saving the data in an 
                                encrypted file in the same locatoin as the script.  Make sure
                                to edit the $Key variable in the Get-SavedCredential function
                                get get a unique encryption key.  After the first run the script
                                will no longer prompt you for credentials.
    SMTPCredentials             If you set UseCredentials to Yes you can set the user name of
                                the user you want to authenticate with on your SMTP relay. If
                                you leave this blank the script will default to the currently
                                logged in user.
    BodyAsHTML                  Set to Yes if your email body will be formatted with HTML.
    
    SearchBase                  The FQDN of the OU you want the start your search in.  Leave it
                                blank to search the entire Active Directory tree.
    SingleDayNotifications      Designate the individual days you want notifications to be sent
                                out.  If your DaysToExpire is set to 5, these numbers should always be
                                greater than 5.  IE: 14,8 would be 14 days out and 8 days out to
                                send notifications.
    DaysToExpire                The value where notifications will be sent out EVERY day, if a 
                                users password will be expiring in less than this value.
                                
    EMAIL BODY                  A special setting, all lines following this line will be evaluated
                                as the body of the email you want to send. Use the variables 
                                below to insert personalized information into the body of the email.
                                Supports HTML if you set BodyAsHTML above to Yes.
                                
    The body of the email is defined in the pwdcfg.ini file, you can set specialized 
    variables in the text that the script will replace with the actual value:
    
    %DAYS%                      The number of days until the password expires
    %FIRSTNAME%                 The user's first name
    %LASTNAME%                  The user's last name
    %EMAILADDR%                 The user's full email address
    
    Example pwdcfg.ini file:
    ** This file is designed to relay mail against Gmail **
    [SMTP SETUP]
    SMTPRelay=smtp.gmail.com
    SMTPFrom=no-reply@mycompany.com
    SMTPPort=587
    UseSSL=Yes
    UseCredentials=yes
    SMTPCredentials=authorizeduser@mycompany.com
    BodyAsHTML=No

    [SETTINGS]
    SearchBase=
    SingleDayNotifications=14,8
    DaystoExpire=5

    [EMAIL BODY]
    Hi %firstname%, your password is going to die in %Days%.  Change it now, no more questions.
    
    
    Another example:
    ** This one relays off of an Exchange server with no credential requirements **
    [SMTP SETUP]
    SMTPRelay=MyExchange2010
    SMTPFrom=no-reply@mycompany.com
    SMTPPort=25
    UseSSL=No
    UseCredentials=No
    SMTPCredentials=
    BodyAsHTML=No

    [SETTINGS]
    SearchBase=
    SingleDayNotifications=14,8
    DaystoExpire=5

    [EMAIL BODY]
    Hi %firstname%, your password is going to die in %Days%.  Change it now, no more questions.
    
    
    Both examples will send an email to the user at 14 days, 8 days and then every day from 5 down.
    It will also search all of Active Directory, so some service accounts and other user objects may
    get email.
    
.PARAMETER Path
    Designate the path where you want the configuration file, the log file and, if used, the credential
    file.  Default is the same path where you saved the script.
.PARAMETER Demo
    Turn this parameter on to have all emails sent to one email address (To below).  This switch should
    be used when testing.
.PARAMETER To
    When Demo is turned on you must designate a To email address.  All emails will be sent to this email
    address instead of the user's.
.EXAMPLE
    .\Send-PasswordExpirationNotifications.ps1 
    
    All configuration and log files will be saved in default location.  User's will receive emails if
    their password is going to expire in the designated timeframe.
.EXAMPLE
    .\Send-PasswordExpirationNotifications.ps1 -Demo -To administrator@mycompany.com
    
    All configuration and log files will be saved in default location.  User's will NOT receive emails if
    their password is going to expire in the designated timeframe, but instead administrator@mycompany.com
    will receive them all.
.NOTES
    Author:            Martin Pugh
    Twitter:           @thesurlyadm1n
    Spiceworks:        Martin9700
    Blog:              www.thesurlyadmin.com
       
    Changelog:
       1.1             Added ability to use HTML for the email.  This was actually part of the original
                       specification and I forgot to put it in!!
       1.01            Found a bug when sending email in production, it was sending to demo address
                       anyway!  Thanks Kent for telling me about this.
       1.0             Initial Release
.LINK
    http://community.spiceworks.com/scripts/show/2244-email-users-that-their-password-is-expiring
#>
#requires -Version 3.0
[CmdletBinding()]
Param (
    [string]$Path,
    
    [switch]$Demo,
    [string]$To = "you@yourdomain.com"
)

Function Get-SavedCredential {
	<#
	.SYNOPSIS
		Simple function to get and save domain credentials.
	.LINK
		http://community.spiceworks.com/scripts/show/1629-get-secure-credentials-function
	#>
    Param (
	    [String]$AuthUser = $env:USERNAME,
        [string]$PathToCred
    )
    $Key = [byte]29,36,18,74,72,75,85,52,73,44,0,21,98,76,99,28

    #Build the path to the credential file
    $CredFile = $AuthUser.Replace("\","~")
    $File = $PathToCred + "\Credentials-$CredFile.crd"
	#And find out if it's there, if not create it
    If (-not (Test-Path $File))
	{	(Get-Credential $AuthUser).Password | ConvertFrom-SecureString -Key $Key | Set-Content $File
    }
	#Load the credential file 
    $Password = Get-Content $File | ConvertTo-SecureString -Key $Key
    $AuthUser = (Split-Path $File -Leaf).Substring(12).Replace("~","\")
    $AuthUser = $AuthUser.Substring(0,$AuthUser.Length - 4)
	$Credential = New-Object System.Management.Automation.PsCredential($AuthUser,$Password)
    Return $Credential
}

Write-Verbose "$(Get-Date): Loading ActiveDirectory module..."
Try { Import-Module ActiveDirectory -ErrorAction Stop }
Catch { Write-Host "Unable to load Active Directory module, is RSAT installed?" -ForegroundColor Red; Exit }

Write-Verbose "$(Get-Date): Validate To parameter"
If ($Demo -and $To -eq $null)
{   Write-Host "If using Demo mode you must specify the -To parameter" -ForegroundColor Red
    Exit
}

Write-Verbose "$(Get-Date): Validate configuration file exists, otherwise create it, exit the script and open Notepad with the config file in there"
If (-not $Path)
{   $Path = $PSScriptRoot
}

If (-not (Test-Path $Path\pwdcfg.ini -PathType Leaf))
{   $pwdcfg = @"
[SMTP SETUP]
SMTPRelay=<ip address or host name>
SMTPFrom=no-reply@mycompany.com
SMTPPort=25
UseSSL=No
UseCredentials=No
SMTPCredentials=
BodyAsHTML=No

[SETTINGS]
SearchBase=
SingleDayNotifications=14,8
DaystoExpire=5


#
# %DAYS% = replaces with the number of days until the password expires
# %FIRSTNAME% = user's first name
# %EMAILADDR% = user's email address
# %LASTNAME% = user's last name
#


[EMAIL BODY]
*** This is an automatically generated email, please do not reply. ***


Good morning %FIRSTNAME%,

We have detected that your password is going to expire in %DAYS% days.

We strongly suggest you change it immediately.  Once your password expires you will not be able to log into Outlook Web Access or the VPN.  Passwords can only be changed from the office or if you have a VPN connection.  Remote employees may not be able to get into the system at all without assistance from the help desk.



If you need help changing your password, please try this first:
<URL to your change password procedure>

If you are a remote user, we have specific instructions here: 
<URL to your remote user change password procedure>




If you're still unsure how to change your password simply email Helpdesk@mycompany.com for assistance.
"@
    $pwdcfg | Out-File $Path\pwdcfg.ini
    Notepad.exe $Path\pwdcfg.ini
    Write-Host "Configuration file wasn't present so have created one for you.  Please insert the proper information and rerun the script." -ForegroundColor Green
    Exit
}

Write-Verbose "$(Get-Date): Parse the pwdcfg.ini file"
$pwdcfg = Get-Content $Path\pwdcfg.ini
$MailSplat = @{
    From = ($pwdcfg | Select-String "SMTPFrom").Line.SubString(9)
    SMTPServer = ($pwdcfg | Select-String "SMTPRelay").Line.SubString(10)
    Port = [int]($pwdcfg | Select-String "SMTPPort").Line.SubString(9)
    ErrorAction = "Stop"
}
$SMTPAuth = ($pwdcfg | Select-String "SMTPCredentials").Line.SubString(16)
If ($SMTPAuth -eq "")
{   $SMTPAuth = $env:USERNAME
}
If (($pwdcfg | Select-String "UseCredentials").Line.SubString(15).ToUpper() -eq "YES")
{   Write-Verbose "Retrieving credentials..."
    $MailSplat.Add("Credential",(Get-SavedCredential -AuthUser $SMTPAuth -PathToCred $Path))
}
If (($pwdcfg | Select-String "UseSSL").Line.SubString(7).ToUpper() -eq "YES")
{   $MailSplat.Add("UseSSL",$true)
}

If (($pwdcfg | Select-String "BodyAsHTML").Line.SubString(11).ToUpper() -eq "YES")
{   $MailSplat.Add("BodyAsHTML",$true)
}

$SingleDays = ($pwdcfg | Select-String "SingleDayNotifications").Line.SubString(23).Split(",")
$DaysToExpire = [int]($pwdcfg | Select-String "DaystoExpire").Line.SubString(13)
$SearchBase = ($pwdcfg | Select-String "SearchBase").Line.SubString(11)

$Body = ForEach ($Line in $pwdcfg)
{   If ($Found)
    {   $Line
    }
    Else
    {   If ($Line -like "*EMAIL BODY*")
        {   $Found = $true
        }
    }
}

Write-Verbose "$(Get-Date): Start the log file"
$Log = @"
#
#  Password Expiration Log
#  Run on $(Get-Date)
#
"@
Add-Content -Path $Path\pwdExpired.log -Value $Log
If ($Demo)
{   Add-Content -Path $Path\pwdExpired.log -Value "#Demo Mode Detected, all emails to be sent to $To"
}

Write-Verbose "$(Get-Date): Determine the maximum password age for the domain"
$maxPasswordAgeTimeSpan = $null
$dfl = (Get-ADDomain).DomainMode.Value__
$maxPasswordAgeTimeSpan = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
If ($maxPasswordAgeTimeSpan -eq $null -or $maxPasswordAgeTimeSpan.TotalMilliseconds -eq 0) 
{	Write-Host "MaxPasswordAge is not set for the domain or is set to zero!"
	Write-Host "So no password expiration's possible.  Exiting script."
	Exit
}

$SearchSplat = @{
    Properties = "PasswordExpired","PasswordLastSet","PasswordNeverExpires","Mail"
}
If ($SearchBase -ne "")
{   $SearchSplat.Add("SearchBase",$SearchBase)
    $SearchSplat.Add("SearchScope","Subtree")
}

ForEach ($User in (Get-ADUser -Filter * @SearchSplat))
{	If ($User.PasswordNeverExpires -or $User.PasswordLastSet -eq $null -or $User.PasswordExpired -or $User.Enabled -eq $false -or $User.Mail -eq $null)
	{	Continue
	}
    Write-Verbose "$(Get-Date): Working on $($User.SamAccountName)..."
	If ($dfl -ge 3) 
	{	Write-Verbose "$(Get-Date): Greater than Windows2008 domain functional level, determining FGPP"
		$accountFGPP = $null
		$accountFGPP = Get-ADUserResultantPasswordPolicy $User
    	If ($accountFGPP) 
		{	$ResultPasswordAgeTimeSpan = $accountFGPP.MaxPasswordAge
    	} 
		Else 
		{	$ResultPasswordAgeTimeSpan = $maxPasswordAgeTimeSpan
    	}
	}
	Else
	{	$ResultPasswordAgeTimeSpan = $maxPasswordAgeTimeSpan
	}
    $TS = (New-TimeSpan -Start (Get-Date) -End ($User.PasswordLastSet + $ResultPasswordAgeTimeSpan)).Days

    If ($SingleDays -contains $TS -or $TS -le $DaysToExpire)
    {   Write-Verbose "$(Get-Date): $($User.SamAccountName) is expiring, sending email"
        If ($Demo)
        {   $SendTo = $To
        }
        Else
        {   $SendTo = $User.Mail
        }
    
        Add-Content $Path\pwdExpired.log -Value "$($User.SamAccountName) set to expire in $TS days.  Email sent to $SendTo"
        $SendBody = $Body.Replace("%DAYS%",$TS)
        $SendBody = $SendBody.Replace("%FIRSTNAME%",$User.GivenName)
        $SendBody = $SendBody.Replace("%LASTNAME%",$User.Surname)
        $SendBody = $SendBody.Replace("EMAILADDR",$User.Mail)
        $Subject = "Your password is about to expire in $TS days!"
        
        Try {
            Send-MailMessage -To $SendTo -Body ($SendBody | Out-String) -Subject $Subject @MailSplat
        }
        Catch {
            Add-Content $Path\pwdExpired.log -Value "Error sending email: $($Error[0])"
            Write-Warning $Error[0]
        }
    }
}
Add-Content $Path\pwdExpired.log -Value "#####  Script completed: $(Get-Date)  #####"
Write-Verbose "$(Get-Date): Script completed!"