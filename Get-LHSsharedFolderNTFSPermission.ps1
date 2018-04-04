Function Get-LHSsharedFolderNTFSPermission
{
<#
.SYNOPSIS
    Lists NTFS permissions of all Shared Folder on local or remote Computer.

.DESCRIPTION
    Lists NTFS permissions of all Shared Folder on local or remote Computer.
    Lists all Share Permision on local or remote Computer.
    Using CIM cmdlets

    This Script can be used to audit/report about shared folder permissions on
    remote Computers 

    ToDo: to check recursive on subfolders

.PARAMETER ComputerName
    The computer name(s) to retrieve the info from. 
    Default to local Computer

.PARAMETER SharePermission
    Shwitch to list Share permision instead of shared Folder NTFS Permissions

.PARAMETER Credential
    Credential to use to connect to the remote Computer.
    Default to current user.

.EXAMPLE
    PS C:\> Get-LHSsharedFolderNTFSPermission -ComputerName Server1

    ComputerName       : Server1
    ConnectionStatus   : Success
    ShareName          : tmp
    SharedFolderPath   : C:\tmp
    SecurityPrincipal  : NT-AUTORITÄT\Authentifizierte Benutzer
    FileSystemRights   : Modify, Synchronize
    AccessControlType  : AccessAllowed
    AccessControlFalgs : Inherited
    Description        : Temp folder
    ..

    To list all Shared Folder NTFS Permissions. 
    If connection to the remote Computer fail, it will be displayed in the 
    'ConnectionStatus' property as 'Fail'

.EXAMPLE
    Get-LHSsharedFolderNTFSPermission -ComputerName Server1 -SharePermission

    ComputerName      : Server1
    ConnectionStatus  : Success
    ShareName         : print$
    SharedFolderPath  : C:\Windows\system32\spool\drivers
    SecurityPrincipal : VORDEFINIERT\Administratoren
    ShareRights       : FullControl
    Description       : Druckertreiber

    ComputerName      : Server1
    ConnectionStatus  : Success
    ShareName         : Temp
    SharedFolderPath  : C:\Temp
    SecurityPrincipal : Jeder
    ShareRights       : Modify, Synchronize
    Description       : 
    ..

    To list all Share and their Permission only
    Note: Administrative shares cannot be displayed (Admin$,C$,IPC$)

.EXAMPLE
    Get-LHSsharedFolderNTFSPermission -ComputerName Server1 |
    Export-Csv -Path "C:\Temp\Permission.csv" -NoTypeInformation -UseCulture

    Invoke-Item "C:\Temp\Permission.csv"

    Outputs Shareed Folder NTFS Permissions to a CSV file and opens it in Excel


.EXAMPLE
    Get-LHSsharedFolderNTFSPermission -ComputerName Server1 -SharePermission |
    Export-Csv -Path "C:\Temp\shares.csv" -NoTypeInformation -UseCulture

    Invoke-Item "C:\Temp\shares.csv"

    Outputs Share Permissions to a CSV file and opens it in Excel
    Note: Administrative shares cannot be displayed (Admin$,C$,IPC$)
    
.INPUTS
    System.String, you can pipe ComputerNames to this Function

.OUTPUTS
    Custom PSObjects 

.NOTES
    Administrative Shares have not the same properties as normal shares
    (Admin$,C$,IPC$) and cannot be displayed


    AUTHOR: Pasquale Lantella 
    LASTEDIT: 22. Januar 2014
    KEYWORDS: Share,Permission

.LINK
    Microsoft All-In-One Script Framework
    Http://www.ScriptingGuys.com 

#Requires -Version 3.0
#>
   
[cmdletbinding()]  

[OutputType('PSObject')] 

Param(

    [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$True,
        HelpMessage='An array of computer names. The default is the local computer.')]
	[alias("CN")]
	[string[]]$ComputerName = $Env:COMPUTERNAME,

    [parameter()]
    [Switch]$SharePermission,

    [parameter()]
    [Alias('RunAs')]
    [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty

   )

BEGIN {

    Set-StrictMode -Version Latest
    ${CmdletName} = $Pscmdlet.MyInvocation.MyCommand.Name


    Function Test-IsWsman3 {
    # Test if WSMan is greater or eqaul Version 3.0
    # Tested against Powershell 4.0
        [cmdletbinding()]
        Param(
        [Parameter(Position=0,ValueFromPipeline)]
        [string]$Computername=$env:computername
        )
 
        Begin {
            #a regular expression pattern to match the ending
            [regex]$rx="\d\.\d$"
        }
        Process {
	    $result = $Null
            Try {
                $result = Test-WSMan -ComputerName $Computername -ErrorAction Stop
            }
            Catch {
                #Write-Error $_
                $False
            }
            if ($result) {
                $m = $rx.match($result.productversion).value
                if ($m -ge '3.0') {
                    $True
                }
                else {
                    $False
                }
            }
        } #process
        End {}
    } #end Test-IsWSMan



#    $RecordErrorAction = $ErrorActionPreference
    #change the error action temporarily
#    $ErrorActionPreference = "SilentlyContinue"

} # end BEGIN

PROCESS {
    
    ForEach ($Computer in $ComputerName) 
    {
        IF (Test-Connection -ComputerName $Computer -count 2 -quiet) 
        { 
#region CimSession
            # Create a Cim Session using WSMAN or DCOM protocol
            $SessionParams = @{
                ComputerName  = $Computer
                ErrorAction = 'Stop'
            } 
            if ($PSBoundParameters['Credential'])  
            {
                $SessionParams.Credential = $Credential
            }
            If (Test-IsWsman3 –ComputerName $Computer)
            {
	            $Option = New-CimSessionOption -Protocol WSMan
	            $SessionParams.SessionOption = $Option      
            }
            Else
            {
	            $Option = New-CimSessionOption -Protocol DCOM
	            $SessionParams.SessionOption = $Option 
            }
            $CimSession = New-CimSession @SessionParams 
#endregion CimSession 
            
            $ShareList = Get-CimInstance -ClassName Win32_Share -CimSession $CimSession
            
                
		    foreach($Share in $ShareList)
		    {
			    $outputObject = @()

                if ($PSBoundParameters['SharePermission'])  
                {
                    Write-verbose "List Share Permission"
                    $ShareSecs = $ShareSecDescriptor = $Null
                    $ShareSecs = Get-CimInstance -ClassName Win32_LogicalShareSecuritySetting `
                    -Filter "Name='$($Share.Name)'" -CimSession $CimSession

			        $ShareSecDescriptor = $ShareSecs | Invoke-CimMethod -MethodName GetSecurityDescriptor -CimSession $CimSession
			        
                    Try
                    {
                        foreach($DACL in $ShareSecDescriptor.Descriptor.DACL)
			            {  
				            $DACLDomain = $DACL.Trustee.Domain
				            $DACLName = $DACL.Trustee.Name
				            if($DACLDomain -ne $null)
				            {
	           		            $UserName = "$DACLDomain\$DACLName"
				            }
				            else
				            {
					            $UserName = "$DACLName"
				            }

				            $outputObject = New-Object PSObject -Property @{
				                ComputerName = $Computer;
					            ConnectionStatus = "Success";
						        ShareName = $Share.Name;
                                SharedFolderPath = $Share.Path;
						        SecurityPrincipal = $UserName;
						        ShareRights = [Security.AccessControl.FileSystemRights]`
							        $($DACL.AccessMask -as [Security.AccessControl.FileSystemRights]);
						        AccessControlType = [Security.AccessControl.AceType]$DACL.AceType;
						        AccessControlFalgs = [Security.AccessControl.AceFlags]$DACL.AceFlags;
						        Description = $Share.Description;		
    	                    } 
                            $outputObject | Select-Object ComputerName,ConnectionStatus,ShareName,SharedFolderPath,SecurityPrincipal, `
                                ShareRights,Description -Unique

                        }
                    }
                    Catch { Continue}
                }
                Else
                {
                    Write-verbose "List Shared Folder NTFS Permission"
                    $SharedNTFSSecs = $SecDescriptor = $Null

                    # to Escape all '\'
			        $SharedFolderPath = [regex]::Escape($Share.Path)
                    If ($SharedFolderPath)
                    {
				        $SharedNTFSSecs = Get-CimInstance -ClassName Win32_LogicalFileSecuritySetting `
				        -Filter "Path='$SharedFolderPath'" -CimSession $CimSession
                    
			            $SecDescriptor = $SharedNTFSSecs | Invoke-CimMethod -MethodName GetSecurityDescriptor -CimSession $CimSession

                        $DACLs = $SecDescriptor.Descriptor.DACL #| Where-Object {$_ -match "\w+"} #remove empty Objects
                        Write-Verbose "`$DACLs : $DACLs "

		                foreach($DACL in $DACLs)
		                {
                            Try
                            {  
			                    $DACLDomain = $DACL.Trustee.Domain
			                    $DACLName = $DACL.Trustee.Name
			                    if($DACLDomain -ne $null)
			                    {
           		                    $UserName = "$DACLDomain\$DACLName"
			                    }
			                    else
			                    {
				                    $UserName = "$DACLName"
			                    }
			
			                    $outputObject = New-Object PSObject -Property @{
			                        ComputerName = $Computer;
				                    ConnectionStatus = "Success";
					                ShareName = $Share.Name;
                                    SharedFolderPath = $Share.Path;
					                SecurityPrincipal = $UserName;
					                FileSystemRights = [Security.AccessControl.FileSystemRights] `
                                           $($DACL.AccessMask -as [Security.AccessControl.FileSystemRights]);
					                AccessControlType = [Security.AccessControl.AceType]$DACL.AceType;
					                AccessControlFalgs = [Security.AccessControl.AceFlags]$DACL.AceFlags;
					                Description = $Share.Description;		
   	                            } 
                                $outputObject | Select-Object ComputerName,ConnectionStatus,ShareName,SharedFolderPath,SecurityPrincipal, `
                                    FileSystemRights,AccessControlType,AccessControlFalgs,Description -Unique
                            }
                            Catch { Continue }
                        }
                    }
                    Else
                    {
	                    $outputObject = New-Object PSObject -Property @{
	                        ComputerName = $Computer;
		                    ConnectionStatus = "Success";
			                ShareName = $Share.Name;
                                  SharedFolderPath = $Null;
			                SecurityPrincipal = "Not Available";
			                FileSystemRights = "Not Available";
			                AccessControlType = "Not Available";
			                AccessControlFalgs = "Not Available";
			                Description = $Share.Description;		
 	                     } 
                         $outputObject | Select-Object ComputerName,ConnectionStatus,ShareName,SharedFolderPath,SecurityPrincipal, `
                             FileSystemRights,AccessControlType,AccessControlFalgs,Description -Unique
                    }

                } # end if ($PSBoundParameters['SharePermission'])
		    } # end foreach($Share in $ShareList)
            
            Remove-CimSession -CimSession $CimSession    
        } 
        Else 
        {
            Write-Warning "\\$Computer DO NOT reply to ping"
        
            if ($PSBoundParameters['SharePermission'])  
            {
                $outputObject = New-Object PSObject -Property @{
                    ComputerName = $Computer;
                    ConnectionStatus = "Fail";
				    ShareName = "Not Available";
                    SharedFolderPath = "Not Available";
				    SecurityPrincipal = "Not Available";
				    ShareRights = "Not Available";
				    AccessControlType = "Not Available";
				    AccessControlFalgs = "Not Available";
                    Description = "Not Available";
                } 
                $outputObject | Select-Object ComputerName,ConnectionStatus,ShareName,SharedFolderPath,SecurityPrincipal, `
                    ShareRights,Description -Unique
                        
            }
            Else
            {
                $outputObject = New-Object PSObject -Property @{
                    ComputerName = $Computer;
                    ConnectionStatus = "Fail";
				    ShareName = "Not Available";
                    SharedFolderPath = "Not Available";
				    SecurityPrincipal = "Not Available";
				    FileSystemRights = "Not Available";
				    AccessControlType = "Not Available";
				    AccessControlFalgs = "Not Available";
                    Description = "Not Available";
                }
                $outputObject | Select-Object ComputerName,ConnectionStatus,ShareName,SharedFolderPath,SecurityPrincipal, `
                    FileSystemRights,AccessControlType,AccessControlFalgs,Description -Unique         
            }
        } # end IF (Test-Connection -ComputerName $Computer -count 2 -quiet)
      
	   
    } # end ForEach ($Computer in $ComputerName)

} # end PROCESS

END {
    #restore the error action preference
#    $ErrorActionPreference = $RecordErrorAction 
    Write-Verbose "Function ${CmdletName} finished." 
}

} # end Function Get-LHSsharedFolderNTFSPermission                
             
