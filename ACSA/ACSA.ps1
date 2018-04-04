######################################################################
# Released: 7/30/2011 3:09 PM
# Author: Rich Prescott
# Blog: blog.richprescott.com
# Twitter: twitter.com/Arposh
# Requires: Powershell v2
# Optional: Quest ActiveDirectory Tools (Searching AD for users/computers)
# Optional: PSExec.exe (Remotely executing commands on computers)
# Optional: Trace32.exe (Tailing log files)
######################################################################

##########################
#### PC SEARCH FORM 2 ####
##########################
Function FormPCSearch{

# Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
$InitialFormWindowState2 = New-Object System.Windows.Forms.FormWindowState

# -----------------------------------------------

function SelectItem
{
$PCString = '$list2.SelectedItems | foreach-object {$_.text} #| foreach-object {$_.dnsname}'
$PCname = invoke-expression $PCString
$txt1.text = $PCName
$form2.Close()
}

$OnLoadForm_StateCorrection2=
{
	$form2.WindowState = $InitialFormWindowState2
}

$form2 = New-Object System.Windows.Forms.Form
$form2.Text = "Loading..."
$form2.Name = "form2"
$form2.DataBindings.DefaultDataSourceUpdateMode = 0
$form2.StartPosition = "CenterScreen"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 270
$System_Drawing_Size.Height = 300
$form2.ClientSize = $System_Drawing_Size
$Form2.KeyPreview = $True
$Form2.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$Form2.Close()}})

# Label PC Search #
$lblPC = New-Object System.Windows.Forms.Label
$lblPC.TabIndex = 8
$lblPC.TextAlign = 256
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 255
$System_Drawing_Size.Height = 15
$lblPC.Size = $System_Drawing_Size
$lblPC.Text = "Double-click a computer or hit enter to select it."
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 10
$lblPC.Location = $System_Drawing_Point
$lblPC.DataBindings.DefaultDataSourceUpdateMode = 0
$lblPC.Name = "lblCompname"
$lblPC.Visible = $false
$form2.Controls.Add($lblPC)

# Listview PC Search #
$list2 = New-Object System.Windows.Forms.ListView
$list2.UseCompatibleStateImageBehavior = $False
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 255
$System_Drawing_Size.Height = 250
$list2.Size = $System_Drawing_Size
$list2.DataBindings.DefaultDataSourceUpdateMode = 0
$list2.Name = "list2"
$list2.TabIndex = 2
$list2.anchor = "right, top, bottom, left"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 40
$list2.View = [System.Windows.Forms.View]"Details"
$list2.FullRowSelect = $true
$list2.GridLines = $true
$columnnames = "Computer","User"
$list2.Columns.Add("Computer", 125) | out-null
$list2.Columns.Add("User", 125) | out-null
$list2.Location = $System_Drawing_Point
$list2.add_DoubleClick({SelectItem})
$list2.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {SelectItem}})
$form2.Controls.Add($list2)

$progress2 = New-Object System.Windows.Forms.ProgressBar
$progress2.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 255
$System_Drawing_Size.Height = 23
$progress2.Size = $System_Drawing_Size
$progress2.Step = 1
$progress2.TabIndex = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10 #120
$System_Drawing_Point.Y = 10 #13
$progress2.Location = $System_Drawing_Point
$progress2.Name = "p1"
$progress2.text = "Loading..."
$form2.Controls.Add($progress2)

################################
####POPULATE PC SEARCH LIST ####
################################

function updatepclist
{
$PCs = get-qadcomputer $computername | sort-object -property name

$progress2.value = 20
if ($PCs.count){$progress2.step = (80/$PCs.count-1)}
else{$progress2.step = 80}

foreach($PC in $PCs){
    $progress2.value += $progress2.step
    $pingPCname = $PC.Name
    $item2 = new-object System.Windows.Forms.ListViewItem($PC.name)
    if (test-connection $pingPCname -quiet -count 1){
        $PCuser = gwmi win32_computersystem -computername $PC.name -ev pcsearcherror
        if ($pcsearcherror){$item2.subitems.add("Unavailable")}  #PC can be pinged, but is not accessible
        if ($PCuser.username -ne $null){$item2.subitems.add($PCuser.Username)} #PC pinged successfully and user is logged in
        } #End test-connection
    else{$item2.subitems.add("Offline")} # PC cannot be pinged
    $item2.Tag = $PC
    $list2.Items.Add($item2) > $null
    } #End foreach
$progress2.visible = $false
$lblpc.visible = $true
$form2.Text = "Select Computer"
} #End function updatepclist


#Save the initial state of the form
$InitialFormWindowState2 = $form2.WindowState
#Init the OnLoad event to correct the initial state of the form
$form2.add_Load($OnLoadForm_StateCorrection2)
$form2.add_Load({updatepclist})
#Show the Form
$form2.ShowDialog()| Out-Null

}
################
#### FORM 2 ####
################

#Generated Form Function
function GenerateForm {

# Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles();
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$vbmsg = new-object -comobject wscript.shell

###################
#### PC SEARCH ####
###################
$btn0_OnClick= 
{
$computername = $txt1.text
$logPC = $txt1.text
$btn10.visible = $false
$btn11.Visible = $false
if($computername.length -lt 4){$vbpcsearch = $vbmsg.popup("Search queries must include at least four characters.",0,"Error",0)}
else{FormPCSearch}
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $logpc + "," +  "Search for PC," + $txt1.text | out-file -filepath $lfile -append}
}


#####################
#### SYSTEM INFO ####
#####################

$btn1_OnClick= 
{
if ($txt1.text -eq "." -OR $txt1.text -eq "localhost"){$txt1.text = hostname}
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
$lbl2.text = ""
HideUnusedItems

if (test-connection $computername -quiet -count 1){
$stBar1.text = "System Info for " + $computername.ToUpper() + " (Loading...)"
$list1.visible = $false
$lbl2.visible = $true
$systeminfoerror = $null

# Begin query #
$rComp = gwmi win32_computersystem -computername $computername -ev systeminfoerror
if ($systeminfoerror){$stBar1.text = "Error retrieving info from " + $computername.ToUpper()}
else {
$rOS = gwmi win32_operatingsystem -computername $computername
$rComp2 = gwmi win32_computersystemproduct -computername $computername
$rCPU = gwmi win32_processor -computername $computername
$rBIOS = gwmi win32_bios -computername $computername
$rRam = gwmi win32_physicalmemory -computername $computername
$rIP = gwmi win32_networkadapterconfiguration -computername $computername | ?{$_.DNSDomain -ne $null}
$rMon = gwmi win32_desktopmonitor -computername $computername -filter "Availability='3'"
$rVid = gwmi win32_videocontroller -computername $computername
$rDVD = gwmi win32_cdromdrive -computername $computername
$rHD = gwmi win32_logicaldisk -computername $computername -filter "Drivetype='3'"
$rProc = gwmi win32_process -ComputerName $computername
$rOU = Get-QADComputer $computername

# McAfee Info #
$ProductVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$rComp.name).OpenSubKey('SOFTWARE\McAfee\DesktopProtection').GetValue('szProductVer')
$EngineVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$rComp.name).OpenSubKey('SOFTWARE\McAfee\AVEngine').GetValue('EngineVersionMajor')
$DatVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$rComp.name).OpenSubKey('SOFTWARE\McAfee\AVEngine').GetValue('AVDatVersion')
$DatDate = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$rComp.name).OpenSubKey('SOFTWARE\McAfee\AVEngine').GetValue('AVDatDate')

# Separate sticks of memory #
$RAM = $rComp.totalphysicalmemory / 1GB
$mem = "{0:N2}" -f $RAM + " GB Usable -- "
$memcount = 0
foreach ($stick in $rRam){
$mem += "(" + "$($rRam[$memcount].capacity / 1GB) GB" + ") "
$memcount += 1
}
$mem += "Physical Stick(s)"

# Enumerate Monitors #
$monitor = ""
foreach ($mon in $rmon) {
$monitor += "(" + $mon.screenwidth + " x " + $mon.screenHeight + ") "
}

# List IP/MAC Address #
$IP = $rIP.IPAddress
$MAC = $rIP.MACAddress
if ($rIP.count)
    {
    $IP = $rIP[0].IPAddress
    $MAC = $rIP[0].MACAddress
    }

# Convert Date fields #
$imagedate = [System.Management.ManagementDateTimeconverter]::ToDateTime($rOS.InstallDate)
$localdate = [System.Management.ManagementDateTimeconverter]::ToDateTime($rOS.LocalDateTime)

# Format Hard Disk sizes #
$HDfree = $rHD.Freespace / 1GB
$HDSize = $rHD.Size / 1GB

# Screensaver Activity #
$TimeSS = $rProc | ?{$_.Name -match ".scr"}
if (!$TimeSS)
    {
    $Screensaver = "Not Active"
    }
else{
    $Screensaver = "{0:N2}" -f (Compare-DateTime -TimeOfObject $TimeSS -Property "CreationDate")
    }
    
# User Logon Duration #
$explorer = $rProc | ?{$_.name -match "explorer.exe"}
if (!$explorer)
    {
    $userlogonduration = $null
    }
elseif ($explorer.count)
    {
    $explorer = $explorer | sort creationdate
    $UserLogonDuration = $explorer[0]
    }
else
    {
    $UserLogonDuration = $explorer
    }
if ($UserLogonDuration){$ULD = Compare-DateTime $UserLogonDuration "CreationDate"}
else{$ULD = ""}


<#
# Desktop/My Documents folder sizes #
if ($rComp.Username -eq $null){}
else {
    if ($rOS.Caption -match "Windows 7" -OR "Vista"){$userpath = "users\"; $mydocs = "\Documents"}
    if ($rOS.Caption -match "XP"){$userpath = "documents and settings\"; $mydocs = "\My Documents"}
    $path = "\\$computername\c$\$userpath"
    $username = $rComp.Username
    if ($username.indexof("\") -ne -1){$username = $username.remove(0,$username.lastindexof("\")+1)}
        
    # Desktop Folder Size
    $startFolder1 = $path + $username + "\Desktop"
    $colItems1 = (Get-ChildItem $startFolder1 -recurse| Measure-Object -property length -sum)
    $rDesk = "{0:N2}" -f ($colItems1.sum / 1MB)

    # My Documents Folder Size
    $startFolder2 = $path + $username + $mydocs
    $colItems2 = (Get-ChildItem $startFolder2 -recurse| Measure-Object -property length -sum)
    $rMyDoc = "{0:N2}" -f ($colItems2.sum / 1MB)
    }
#>

# Write query results #
$lbl2.text += "Computer Name:`t" + $rComp.name + "`n"
$lbl2.text += "Domain Location:`t" + $rOU.ParentContainer + "`n"
$lbl2.text += "Current User:`t" + $rComp.username + "`n"
$lbl2.text += "User logged on for:`t" + $ULD + "`n"
$lbl2.text += "Screensaver Time:`t" + $Screensaver + "`n"
$lbl2.text += "Last Restart:`t" + (Compare-DateTime -TimeOfObject $rOS -Property "Lastbootuptime") + "`n`n"
$lbl2.text += "Manufacturer:`t" + $rComp.Manufacturer + "`n"
$lbl2.text += "Model:`t`t" + $rComp.Model + "`n"
$lbl2.text += "Chassis:`t`t" + $rComp2.Version + "`n"
$lbl2.text += "Serial:`t`t" + $rBIOS.SerialNumber + "`n`n"
$lbl2.text += "CPU:`t`t" + $rCPU.Name.Trim() + "`n"
$lbl2.text += "RAM:`t`t" + $mem + "`n"
$lbl2.text += "Hard Drive: `t{0:N1} GB Free / {1:N1} GB Total `n" -f $HDfree, $HDsize
$lbl2.text += "Optical Drive:`t" + "(" + $rDVD.Drive + ") " + $rDVD.Caption + "`n"
$lbl2.text += "Video Card:`t" + $rVid.Name + "`n"
$lbl2.text += "Monitor(s):`t" + $monitor + "`n`n"
$lbl2.text += "Local Date/Time:`t" + $localdate + "`n"
$lbl2.text += "Operating System:`t" + $rOS.Caption + "`n"
$lbl2.text += "Service Pack:`t" + $rOS.CSDVersion + "`n"
$lbl2.text += "OS Architecture:`t" + $rComp.SystemType + "`n"
$lbl2.text += "PC imaged on:`t" + $imagedate + "`n`n"
$lbl2.text += "IP Address:`t" + $IP + "`n"
$lbl2.text += "MAC Address:`t" + $MAC + "`n`n"
$lbl2.text += "McAfee Version:`t" + $ProductVer + "`n"
$lbl2.text += "McAfee Engine:`t" + $EngineVer + "`n"
$lbl2.text += "DAT Version:`t" + $DatVer + "`n"
$lbl2.text += "Last Update:`t" + $DatDate + "`n`n"

<#
# Desktop/My Docs labels #
if ($rComp.Username -eq $null){}
else {
    $lbl2.text += "Desktop Folder:`t" + $rDesk + " MB" + "`n"
    $lbl2.text += "My Docs Folder:`t" + $rMyDoc + " MB"
    }
#>

$stBar1.text = "System Info for " + $computername.ToUpper()
}
  }
  else{
  $stBar1.text = "Could not contact " + $computername.ToUpper()
}
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "System Info" | out-file -filepath $lfile -append}
}

######################
#### LOCAL ADMINS ####
######################
$btn2_OnClick= 
{
if ($txt1.text -eq "." -OR $txt1.text -eq "localhost"){$txt1.text = hostname}
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()

HideUnusedItems

if (test-connection $computername -quiet -count 1){
$stBar1.text = "Local Admins on " + $computername.ToUpper() + " (Loading...)"
$list1.visible = $true
$lbl2.visible = $false

$list1.Columns[0].text = "Domain"
$list1.Columns[0].width = 129
$list1.Columns[1].text = "User"
$list1.Columns[1].width = ($list1.width - $list1.columns[0].width - 25)
$List1.items.Clear()

$localgroupName = "Administrators" 
 if ($computerName -eq "") {$computerName = "$env:computername"} 
  
 if([ADSI]::Exists("WinNT://$computerName/$localGroupName,group")) { 
  
     $group = [ADSI]("WinNT://$computerName/$localGroupName,group") 
  
     $members = @() 
     $Group.Members() | 
     % { 
         $AdsPath = $_.GetType().InvokeMember("Adspath", 'GetProperty', $null, $_, $null) 
         # Domain members will have an ADSPath like WinNT://DomainName/UserName. 
         # Local accounts will have a value like WinNT://DomainName/ComputerName/UserName. 
         $a = $AdsPath.split('/',[StringSplitOptions]::RemoveEmptyEntries) 
         $name = $a[-1] 
         $domain = $a[-2] 
         $class = $_.GetType().InvokeMember("Class", 'GetProperty', $null, $_, $null) 
  
         $member = New-Object PSObject 
         $member | Add-Member -MemberType NoteProperty -Name "Name" -Value $name 
         $member | Add-Member -MemberType NoteProperty -Name "Domain" -Value $domain 
         $member | Add-Member -MemberType NoteProperty -Name "Class" -Value $class 
  
         $members += $member 
        }
    foreach ($admin in $members){
        $item = new-object System.Windows.Forms.ListViewItem($admin.domain)
        if ($admin.Name -ne $null){$item.SubItems.Add($admin.Name)}
        $item.Tag = $admin
        $list1.Items.Add($item) > $null
        } #End foreach $admin
    
    $btn12.visible = $true 
    
    }
    
    $stBar1.text = "Local Admins on " + $computername.ToUpper()
    
    } #End test-connection

else{$stBar1.text = "Could not contact " + $computername.ToUpper()}
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "Local Admins" | out-file -filepath $lfile -append}
} #End Local Admins

######################
#### APPLICATIONS ####
######################
$btn3_OnClick= 
{
if ($txt1.text -eq "." -OR $txt1.text -eq "localhost"){$txt1.text = hostname}
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
$lbl2.text = ""
HideUnusedItems


if (test-connection $computername -quiet -count 1){
$stBar1.text = "Applications on " + $computername.ToUpper() + " (Loading...)"

$list1.visible = $true
$lbl2.visible = $false
$columnnames = "Name","Path"
$list1.Columns[1].text = "Install Date"
$list1.Columns[1].width = 129
$list1.Columns[0].text = "Name"
$list1.Columns[0].width = ($list1.width - $list1.columns[1].width - 25)

$List1.items.Clear()
$systeminfoerror = $null
$software = gwmi win32_product -computername $computername -ev systeminfoerror | sort-object -property Name
if ($systeminfoerror){$stBar1.text = "Error retrieving info from " + $computername.ToUpper()}
else {
$columnproperties = "Name","InstallDate"
foreach ($app in $software) {
    $item = new-object System.Windows.Forms.ListViewItem($app.name)
    if ($app.InstallDate -ne $null){
    $item.SubItems.Add($app.InstallDate)
    }
    $item.Tag = $app
    $list1.Items.Add($item) > $null
  }

$btn11.Visible = $true

$stBar1.text = "Applications installed on " + $computername.ToUpper() + " (" + $software.count + ")"
  }
  } #End wmi error check
  else{
  $stBar1.text = "Could not contact " + $computername.ToUpper()
}
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "Applications" | out-file -filepath $lfile -append}
} #End Applications

#######################
#### STARTUP ITEMS ####
#######################
$btn8_OnClick=
{
if ($txt1.text -eq "." -OR $txt1.text -eq "localhost"){$txt1.text = hostname}
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
$lbl2.text = ""
HideUnusedItems

if (test-connection $computername -quiet -count 1){
$stBar1.text = "Startup items on " + $computername.ToUpper() + " (Loading...)"
$list1.visible = $true
$lbl2.visible = $false

$list1.Columns[0].text = "Name"
$list1.Columns[0].width = 175
$list1.Columns[1].text = "Path"
$list1.Columns[1].width = ($list1.width - $list1.columns[0].width - 25)

$List1.items.Clear()
$startup = gwmi win32_startupcommand -computername $computername -filter "User='All Users'" -ev systeminfoerror #"User='Public'" -Windows 7
if ($systeminfoerror){$stBar1.text = "Error retrieving info from " + $computername.ToUpper()}
else {

foreach ($start in $startup){
    $item = new-object System.Windows.Forms.ListViewItem($start.Caption)
    if ($start.Command -ne $null){
    $item.SubItems.Add($start.Command)
    }
    $item.Tag = $start
    $list1.Items.Add($item) > $null
  }


$stBar1.text = "Startup items on " + $computername.ToUpper() + " (" + $startup.count + ")"
$btn13.Visible = $true
  }
  } #End wmi error check
  else{
  $stBar1.text = "Could not contact " + $computername.ToUpper()
}
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "Startup Items" | out-file -filepath $lfile -append}
} #End startup item list

# REMOTE DESKTOP #
$btn4_OnClick= 
{
$computername = $txt1.text
#HideUnusedItems
$remote =  "mstsc.exe /v:" + $computername
iex $remote
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "Remote Desktop" | out-file -filepath $lfile -append}
}

# REMOTE ASSISTANCE #
$btn5_OnClick= 
{
$computername = $txt1.text
HideUnusedItems
$adminOS = gwmi win32_operatingsystem
if ($adminOS.Caption -match "Windows 7"){msra /offerRA $computername}
if ($adminOS.Caption -match "Windows Vista"){msra /offerRA $computername}
if ($adminOS.Caption -match "Windows XP"){
$ie = New-Object -ComObject "InternetExplorer.Application"
$ie.visible = $true
$ie.navigate("hcp://CN=Microsoft%20Corporation,L=Redmond,S=Washington,C=US/Remote%20Assistance/Escalation/Unsolicited/Unsolicitedrcui.htm")
<#
$wshell = Get-Process | Where-Object {$_.Name -eq "HelpCtr"}
    start-sleep -milliseconds 1500
[void] 
[System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
Start-Sleep -Milliseconds 250
[System.Windows.Forms.SendKeys]::SendWait("$computername")
Start-Sleep -Milliseconds 250
[System.Windows.Forms.SendKeys]::SendWait("%C")
Start-Sleep -Milliseconds 250
[System.Windows.Forms.SendKeys]::SendWait("%S")
#>
}
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "Remote Assistance" | out-file -filepath $lfile -append}
}

# FILE STRUCTURE #
$btn6_OnClick= 
{
if ($txt1.text -eq "." -OR $txt1.text -eq "localhost"){$txt1.text = hostname}
$computername = $txt1.text
#HideUnusedItems
$files = "\\" + $computername + "\c$"
explorer $files
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "File Structure" | out-file -filepath $lfile -append}
}

# RESTART COMPUTER #
$btn7_OnClick= 
{
if ($txt1.text -eq "." -OR $txt1.text -eq "localhost"){$txt1.text = hostname}
$computername = $txt1.text
HideUnusedItems
$vbrestart = $vbmsg.popup("Are you sure you want to restart " + $computername.ToUpper() + "?",0,"Restart " + $computername.ToUpper() + "?",4)
switch ($vbrestart)
{
6 {
restart-computer -force -computername $computername
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "Restart Computer" | out-file -filepath $lfile -append}
}
7 {}
}
}

######################
#### SHOW-PROCESS ####
######################
$btn9_OnClick= 
{
if ($txt1.text -eq "." -OR $txt1.text -eq "localhost"){$txt1.text = hostname}
$computername = $txt1.text
HideUnusedItems
$stBar1.text = "Pinging " + $computername.ToUpper()

if (test-connection $computername -quiet -count 1){
$stBar1.text = "Processes on " + $computername.ToUpper() + " (Loading...)"

$list1.visible = $true
$lbl2.Visible = $false

$List1.items.Clear()

$list1.Columns[0].text = "Name"
$list1.Columns[0].width = 150
$list1.Columns[1].text = "Path"
$list1.Columns[1].width = ($list1.width - $list1.columns[0].width - 25)

$systeminfoerror = $null
$procs = gwmi win32_process -computername $computername -ev systeminfoerror | sort-object -property name
if ($systeminfoerror){$stBar1.text = "Error retrieving info from " + $computername.ToUpper()}
else{
$columnproperties = "Name","ExecutablePath"
foreach ($d in $procs) {
    $item = new-object System.Windows.Forms.ListViewItem($d.name)
    if ($d.executablepath -ne $null){
    $item.SubItems.Add($d.executablepath)
    }
    $item.Tag = $d
    $list1.Items.Add($item) > $null
  }
$stBar1.text = "Processes on " + $computername.ToUpper()
$btn10.visible = $true
  }
  } #End wmi error check
  else{
  $stBar1.text = "Could not contact " + $computername.ToUpper() 
    }
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "Processes" | out-file -filepath $lfile -append}
} #End show processs

# END PROCESS #
$btn10_OnClick= 
{
if ($list1.selecteditems.count -gt 1){$vbmsg1 = $vbmsg.popup("You may only select one process to end at a time.",0,"Error",0)}
elseif ($list1.selecteditems.count -lt 1){$vbmsg1 = $vbmsg.popup("Please select a process to end.",0,"Error",0)}
else{
$exprString = '$list1.SelectedItems | foreach-object {$_.tag} | foreach-object {$_.processid}'
$endproc = invoke-expression $exprString
$process = Get-WmiObject -ComputerName $computername -Query "select * from win32_process where processID='$endproc'"
$process.terminate()
if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "End Process," + $process.name | out-file -filepath $lfile -append}
start-sleep 1
$List1.items.Clear()
start-sleep 2
$procs = gwmi win32_process -computername $computername | sort-object -property name
$columnproperties = "Name","ExecutablePath"
foreach ($d in $procs) {
    $item = new-object System.Windows.Forms.ListViewItem($d.name)
    if ($d.executablepath -ne $null){
    $item.SubItems.Add($d.executablepath)
    }
    $item.Tag = $d
    $list1.Items.Add($item) > $null
  }
  
}
}


# UNINSTALL APP #
$btn11_OnClick= 
{
if ($list1.selecteditems.count -gt 1){$vbmsg1 = $vbmsg.popup("You may only select one application to uninstall at a time.",0,"Error",0)}
elseif ($list1.selecteditems.count -lt 1){$vbmsg1 = $vbmsg.popup("Please select an application to uninstall.",0,"Error",0)}
else{

$exprString = '$list1.SelectedItems | foreach-object {$_.tag} | foreach-object {$_.name}'
$endapp = invoke-expression $exprString
$stBar1.text = "Applications on " + $computername.ToUpper() + " (Uninstalling $($Endapp))"

$uninapp = Get-WmiObject -ComputerName $computername -Query "select * from win32_product where name='$endapp'"
$uninapp.uninstall()

if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "Uninstall Application," + $uninapp.name | out-file -filepath $lfile -append}
start-sleep 1

$List1.items.Clear()
$stBar1.text = "Applications on " + $computername.ToUpper() + " (Refreshing...)"

$software = gwmi win32_product -computername $computername | sort-object -property Name
$columnproperties = "Name","InstallDate"
foreach ($app in $software) {
    $item = new-object System.Windows.Forms.ListViewItem($app.name)
    if ($app.InstallDate -ne $null){
    $item.SubItems.Add($app.InstallDate)
    }
    $item.Tag = $app
    $list1.Items.Add($item) > $null
  }
  $stBar1.text = "Applications installed on " + $computername.ToUpper() + " (" + $software.count + ")"
}
}

# REMOVE ADMIN #
$btn12_OnClick= 
{
if ($list1.selecteditems.count -gt 1){$vbmsg1 = $vbmsg.popup("You may only select one account to remove at a time.",0,"Error",0)}
elseif ($list1.selecteditems.count -lt 1){$vbmsg1 = $vbmsg.popup("Please select an account to remove.",0,"Error",0)}
else{
$stBar1.text = "Admins on " + $computername.ToUpper() + " (Removing...)"
$expUser = '$list1.SelectedItems | foreach-object {$_.tag} | foreach-object {$_.name}'
$username = invoke-expression $expUser
$expDomain = '$list1.SelectedItems | foreach-object {$_.tag} | foreach-object {$_.domain}'
$domain = invoke-expression $expDomain

$computer = [ADSI]("WinNT://" + $computername + ",computer")
$Group = $computer.psbase.children.find("administrators")
$Group.Remove("WinNT://" + $domain + "/" + $username)

if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + "," +  "Remove Admin," + $domain + "\" + $username | out-file -filepath $lfile -append}
start-sleep 1

$List1.items.Clear()
$stBar1.text = "Local Admins on " + $computername.ToUpper() + " (Refreshing...)"

$localgroupName = "Administrators" 
 if ($computerName -eq "") {$computerName = "$env:computername"} 
  
 if([ADSI]::Exists("WinNT://$computerName/$localGroupName,group")) { 
  
     $group = [ADSI]("WinNT://$computerName/$localGroupName,group") 
  
     $members = @() 
     $Group.Members() | 
     % { 
         $AdsPath = $_.GetType().InvokeMember("Adspath", 'GetProperty', $null, $_, $null) 
         # Domain members will have an ADSPath like WinNT://DomainName/UserName. 
         # Local accounts will have a value like WinNT://DomainName/ComputerName/UserName. 
         $a = $AdsPath.split('/',[StringSplitOptions]::RemoveEmptyEntries) 
         $name = $a[-1] 
         $domain = $a[-2] 
         $class = $_.GetType().InvokeMember("Class", 'GetProperty', $null, $_, $null) 
  
         $member = New-Object PSObject 
         $member | Add-Member -MemberType NoteProperty -Name "Name" -Value $name 
         $member | Add-Member -MemberType NoteProperty -Name "Domain" -Value $domain 
         $member | Add-Member -MemberType NoteProperty -Name "Class" -Value $class 
  
         $members += $member 
    } 
    }
    
    foreach ($admin in $members){
    $item = new-object System.Windows.Forms.ListViewItem($admin.domain)
    if ($admin.Name -ne $null){
    $item.SubItems.Add($admin.Name)
    }
    $item.Tag = $admin
    $list1.Items.Add($item) > $null
    }
    
$stBar1.text = "Local Admins on " + $computername.ToUpper()

  }
}

# Remove Startup Items Button #
$btn13_OnClick=
{ 
if ($list1.selecteditems.count -gt 1){$vbmsg1 = $vbmsg.popup("You may only select one account to remove at a time.",0,"Error",0)}
elseif ($list1.selecteditems.count -lt 1){$vbmsg1 = $vbmsg.popup("Please select an account to remove.",0,"Error",0)}
else{
$stBar1.text = "Startup Items on " + $computername.ToUpper() + " (Removing...)"
$expStartUp = '$list1.SelectedItems | foreach-object {$_.tag} | foreach-object {$_.Name}'
$RemStartUp = invoke-expression $expStartUp
$RemoveStartUpItem = Get-WmiObject -ComputerName $computername -Query "select * from win32_startupcommand where name='$RemStartUp'"

foreach($remstitem in $removestartupitem)
    {
    $path = $remstitem.location #$removestartupitem.location

    if($path -match "HKLM")
        {
        $path = $path.Replace('HKLM\','')
        $BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computername).OpenSubKey($Path, $true)
        foreach ($key in $basekey.GetValueNames())
            {   
            if ($key -eq $remstitem.Name)
                {
                $basekey.DeleteValue($key)
                }
        }
    }

    if($path -match "HKU")
        {
        $vbmsg1 = $vbmsg.popup($remstitem.Name + " is located in the user registry hive and may not be removed fully.",0,"Error",0)
        }

    if($path -match "Common Startup")
        {
        $commonxp = "\\" + $computername + "\c$\documents and settings\all users\start menu\programs\startup\" + $remstitem.Command # + ".lnk"
        if(test-path $commonxp){remove-item $commonxp}
        }
    }# End foreach startupitem removal

$List1.items.Clear()
$startup = gwmi win32_startupcommand -computername $computername -filter "User='All Users'"

foreach ($start in $startup){
    $item = new-object System.Windows.Forms.ListViewItem($start.Caption)
    if ($start.Command -ne $null){
    $item.SubItems.Add($start.Command)
    }
    $item.Tag = $start
    $list1.Items.Add($item) > $null
  }

$stBar1.text = "Startup items on " + $computername.ToUpper() + " (" + $startup.count + ")"
}
}

Function HideUnusedItems{
$btn10.visible = $false
$btn11.Visible = $false
$btn12.Visible = $false
$btn13.Visible = $false
} #End Function HideUnusedItems

function McAfeeLogs
{
$computername = $txt1.text
if (test-connection $computername -quiet -count 1){

    $userOS = gwmi win32_operatingsystem -computername $computername

    if ($userOS.Caption -match "Windows 7"){$McAfeepath = "notepad.exe \\$computername\c$\ProgramData\McAfee\DesktopProtection"}
    if ($userOS.Caption -match "2008"){$McAfeepath = "notepad.exe \\$computername\c$\ProgramData\McAfee\DesktopProtection"}
    if ($userOS.Caption -match "Windows XP"){$McAfeepath = "notepad.exe \\$computername\c$\Documents and Settings\All Users\Application Data\McAfee\DesktopProtection"}

    if ($McAfeeFile -eq "AP"){$McAfeelog = "$mcafeepath\AccessProtectionLog.txt"}
    if ($McAfeeFile -eq "OAS"){$McAfeelog = "$mcafeepath\OnAccessScanLog.txt"}
    if ($McAfeeFile -eq "ODS"){$McAfeelog = "$mcafeepath\OnDemandScanLog.txt"}
    if ($McAfeeFile -eq "UD"){$McAfeelog = "$mcafeepath\UpdateLog.txt"}
    if ($McAfeeFile -eq "Agent" -AND ($userOS.Caption -match "Windows 7" -OR $userOS.Caption -match "2008")){$McAfeeLog = "notepad.exe \\$computername\c$\ProgramData\McAfee\Common Framework\DB\Agent_$computername.log"}
    if ($McAfeeFile -eq "Agent" -AND $userOS.Caption -match "Windows XP"){$McAfeeLog = "notepad.exe \\$computername\c$\Documents and Settings\All Users\Application Data\McAfee\Common Framework\DB\Agent_$computername.log"}

    iex $McAfeelog
    } #End Test-Connection
else{$stBar1.text = "Could not contact " + $computername.ToUpper()}
} #End Function McAfeeLogs

function WSUSLogs
{
$computername = $txt1.text

if (test-connection $computername -quiet -count 1)
    {
    $WSUSPath = "\\$computername\c$\Windows"
    if ($WSUSfile -eq "Updates"){$WSUSlog = "$WSUSpath\WindowsUpdate.log"}
    if ($WSUSfile -eq "Report"){$WSUSlog = "$WSUSpath\SoftwareDistribution\ReportingEvents.log"}
    iex $WSUSlog
    } #End Test-Connection
else{$stBar1.text = "Could not contact " + $computername.ToUpper()}
} #End Function WSUSLogs

$FindUser=
{
$findusername = $txt1.text
if ($findusername -eq ""){$vbmsg1 = $vbmsg.popup("Please enter a full or partial username into the textbox.",0,"Error",0)}
else{
    $userlist = get-qaduser $findusername | sort name | select Name, samaccountname, Department, Company, Description, telephoneNumber, email
    $userlist | out-gridview -title "Find Users"
    if (!$userlist){$vbmsg1 = $vbmsg.popup("No users were found matching your query.",0,"Error",0)}
    } #End Else
} #End $FindUser

$FindPCUser=
{
$computername = $txt1.text
if($computername.length -lt 4){$vbpcsearch = $vbmsg.popup("Search queries must include at least four characters.",0,"Error",0)}
else{
    $pc1 = $computername + "$"
    $findPCusername = get-qadcomputer $pc1
    
    if ($findPCusername){
        if (test-connection $findpcusername.name -quiet -count 1){
            $wmifindpcuser = gwmi win32_computersystem -computername $findpcusername.name
            $wmipcusername = $wmifindpcuser.username
            if ($wmipcusername.indexof("\") -ne -1){$wmipcusername = $wmipcusername.remove(0,$wmipcusername.lastindexof("\")+1)}
            $userlist = get-qaduser $wmipcusername | sort name | select Name, samaccountname, Department, Company, Description, telephoneNumber, email
            $userlist | out-gridview -title "Find PC User"
            if (!$userlist){$vbmsg1 = $vbmsg.popup("Noone is logged into " + $findpcusername.name.ToUpper(),0,"Error",0)}
        } #End Ping
        else{$vbmsg1 = $vbmsg.popup("Could not contact " + $findpcusername.name.ToUpper(),0,"Error",0)}
    } #End If $FindPCusername
    
    if (!$findPCusername){
        $findPCusername = get-qadcomputer $computername
        if (!$findPCusername){$vbmsg1 = $vbmsg.popup("No computer was found matching your query. Please try again.",0,"Error",0)}
        elseif ($findPCusername.count){$vbmsg1 = $vbmsg.popup("Multiple machines were found matching your query.  Please narrow your search.",0,"Error",0)}
        else {
            if (test-connection $findpcusername.name -quiet -count 1){
            $wmifindpcuser = gwmi win32_computersystem -computername $findpcusername.name
            $wmipcusername = $wmifindpcuser.username
            if ($wmipcusername.indexof("\") -ne -1){$wmipcusername = $wmipcusername.remove(0,$wmipcusername.lastindexof("\")+1)}
            $userlist = get-qaduser $wmipcusername | sort name | select Name, samaccountname, Department, Company, Description, telephoneNumber, email
            $userlist | out-gridview -title "Find PC User"
            if (!$userlist){$vbmsg1 = $vbmsg.popup("Noone is logged into " + $findpcusername.name.ToUpper(),0,"Error",0)}
            } #End Ping $Findpcusername.name
            else{$vbmsg1 = $vbmsg.popup("Could not contact " + $findpcusername.name.ToUpper(),0,"Error",0)}
            } #End Else
        } #End If !$FindPCUserName
    } #End $Computername.Length -gt 4
} #End $FindPCUser


function EventViewer
{
$computername = $txt1.text
$eventvwr = "eventvwr.exe $computername"
iex $eventvwr
}

function UsersGroups
{
$computername = $txt1.text
$UserGrps = "lusrmgr.msc -a /computer=$computername"
iex $UserGrps
}

function Services
{
$computername = $txt1.text
$Services = "services.msc /computer=$computername"
iex $Services
}

function Compare-DateTime($TimeOfObject,$Property)
{
$TimeOfObject = $TimeOfObject.converttodatetime($TimeOfObject.$Property)
$TimeOfObject = (get-date) - $TimeOfObject
$days = " Day "
if ($TimeOfObject.days -ne 1){$days = $days.replace('Day ','Days ')}
$hours = " Hour "
if ($TimeOfObject.hours -ne 1){$hours = $hours.replace('Hour ','Hours ')}
$minutes = " Minute "
if ($TimeOfObject.minutes -ne 1){$minutes = $minutes.replace('Minute ','Minutes ')}
$TimeComparison = $TimeOfObject.days.tostring() + $days + $TimeOfObject.hours.tostring() + $hours + $TimeOfObject.minutes.tostring() + $minutes
if ($TimeOfObject.days -eq 0){$TimeComparison = $TimeComparison.Replace('0 Days ','')}
if ($TimeOfObject.days -eq 0 -AND $TimeOfObject.hours -eq 0){$TimeComparison = $TimeComparison.Replace('0 Hours ','')}
return $TimeComparison
}


function ReaderIE
{
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
if (test-connection $computername -quiet -count 1)
    {
    $stBar1.text = "Updating Adobe Reader executable path on " + $computername.ToUpper()
    6..14 | %{cmd /c if exist "\\$Computername\c$\program files\adobe\reader $_.0\Reader\AcroRd32.exe" psexec \\$Computername -d reg add HKCR\SOFTWARE\Adobe\Acrobat\Exe /ve /d "`\`"C:\Program Files\Adobe\Reader $_.0\Reader\AcroRd32.exe`\`"" /f}
    6..14 | %{cmd /c if exist "\\$Computername\c$\program files (x86)\adobe\reader $_.0\Reader\AcroRd32.exe" psexec \\$Computername -d reg add HKCR\SOFTWARE\Adobe\Acrobat\Exe /ve /d "`\`"C:\Program Files (x86)\Adobe\Reader $_.0\Reader\AcroRd32.exe`\`"" /f}
    
    if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + ",QFix-ReaderIEPlugin" | out-file -filepath $lfile -append}
    $stBar1.text = "Adobe Reader executable path updated on " + $computername.ToUpper()
    }
else{$stBar1.text = "Could not contact " + $computername.ToUpper()}
    
} #End Function ReaderIE


function Update-FormTitle
{
$form1.Text = "Arposh Admin Tool $Version - Connected to $((Get-QADRootDSE).dnshostname)"
}


function Reset-SUSClientID
{
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
if (Test-Connection $Computername -quiet -count 1)
    {
    $stBar1.text = "Running Reset WSUS Client ID on " + $computername.ToUpper()
    $service = "wuauserv"
    $pcservice = gwmi win32_service -computername $computername -filter "name='$service'"
    $stopsvc = $pcservice.stopservice()

    do {
        start-sleep -m 500
        $pcservice = gwmi win32_service -computername $computername -filter "name='$service'"
        }
    while ($pcservice.state -eq "Running")

    if ($stopsvc.returnvalue -eq 0){"$service stopped on $computername"}
    else{"Error stopping $service on $computername"}

    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computerName)
    $regKey = $reg.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\WindowsUpdate", $true) 
    if ($regkey.getvalue("SUSclientid")){$regkey.deletevalue("SUSclientid")}
    if ($regkey.getvalue("SusClientIdValidation")){$regkey.deletevalue("SusClientIdValidation")}
    if ($regkey.getvalue("PingID")){$regkey.deletevalue("PingID")}
    if ($regkey.getvalue("AccountDomainSid")){$regkey.deletevalue("AccountDomainSid")}

    <#
    $SDpath = "\\$computername\c$\Windows\SoftwareDistribution.old"
    if (test-path $SDpath){remove-item $SDpath -recurse}
    rename-item -path \\$computername\c$\Windows\SoftwareDistribution -newname $SDpath
    #>
    
    start-sleep -m 500

    $startsvc = $pcservice.startservice()
    start-sleep -m 500
    if ($startsvc.returnvalue -eq 0){"$service started on $computername"}
    else{"Error starting $service on $computername"}

    $cmd = "cmd.exe /c psexec.exe \\$computername -d C:\windows\system32\wuauclt.exe /resetauthorization /detectnow"
    $CheckForUpdates = Invoke-Expression $cmd
    $stBar1.text = "Reset WSUS Client ID completed on " + $computername.ToUpper()
    if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + ",QFix-ResetWSUSClientID" | out-file -filepath $lfile -append}    
    }
    else{$stBar1.text = "Could not contact " + $computername.ToUpper()}
} #End function Reset-SUSClientID

function Update-GroupPolicy
{
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
if (Test-Connection $Computername -quiet -count 1)
    {
    $stBar1.text = "Updating Group Policy on " + $computername.ToUpper()
    $cmd = "cmd.exe /c psexec.exe \\$computername -d gpupdate /force"
    $GroupPolicy = Invoke-Expression $cmd
    $stBar1.text = "Group Policy updated on " + $computername.ToUpper()
    if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + ",QFix-GPUpdate" | out-file -filepath $lfile -append}    
    }
else{$stBar1.text = "Could not contact " + $computername.ToUpper()}
} #End function Update-GroupPolicy


function Invoke-WSUSReport
{
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
if (Test-Connection $Computername -quiet -count 1)
    {
    $stBar1.text = "Reporting into WSUS on " + $computername.ToUpper()
    $cmd = "cmd.exe /c psexec.exe \\$computername -d wuauclt.exe /reportnow"
    $WSUSUpdate = Invoke-Expression $cmd
    $stBar1.text = "WSUS reporting started on " + $computername.ToUpper()
    if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + ",QFix-WSUSReport" | out-file -filepath $lfile -append}    
    }
else{$stBar1.text = "Could not contact " + $computername.ToUpper()}
} # End function Invoke-WSUSReport


function Invoke-WSUSDetect
{
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
if (Test-Connection $Computername -quiet -count 1)
    {
    $stBar1.text = "Checking into WSUS on " + $computername.ToUpper()
    $cmd = "cmd.exe /c psexec.exe \\$computername -d wuauclt.exe /detectnow"
    $WSUSUpdate = Invoke-Expression $cmd
    $stBar1.text = "WSUS detect started on " + $computername.ToUpper()
    if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + ",QFix-WSUSDetect" | out-file -filepath $lfile -append}    
    }
else{$stBar1.text = "Could not contact " + $computername.ToUpper()}
} # End function Invoke-WSUSReport


function Rename-Computer
{
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
if (Test-Connection $Computername -quiet -count 1)
    {
    $stBar1.text = "Renaming " + $computername.ToUpper()
    $newPCname = Read-Host "Warning!  This will reboot $computername.  Enter a new name to continue."

    $Cred = Get-credential
    $User = ($cred.GetNetworkCredential()).UserName
    $pwd = ($cred.GetNetworkCredential()).Password

    $cmd = "netdom renamecomputer $computername /newname:$newPCname /userd:$user /passwordd:$pwd /reboot:5 /force"

    iex $cmd
    $cmd = ""
    $pwd = ""
    
    $stBar1.text = "$Computername has been renamed to $NewPCName"
    if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + ",QFix-RenamePC" | out-file -filepath $lfile -append}    
    }
else{$stBar1.text = "Could not contact " + $computername.ToUpper()}
} #End function Rename-Computer


function Lock-Computer
{
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
if (Test-Connection $Computername -quiet -count 1)
    {
    $stBar1.text = "Locking Workstation - " + $computername.ToUpper()
    $cmd = "cmd.exe /c psexec.exe \\$computername c:\Windows\System32\rundll32.exe user32.dll,LockWorkStation"
    $WSUSUpdate = Invoke-Expression $cmd
    $stBar1.text = "Workstation locked - " + $computername.ToUpper()
    if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + ",QFix-LockPC" | out-file -filepath $lfile -append}    
    }
else{$stBar1.text = "Could not contact " + $computername.ToUpper()}
}


function Update-McAfeeDAT
{
$computername = $txt1.text
$stBar1.text = "Pinging " + $computername.ToUpper()
if (Test-Connection $Computername -quiet -count 1)
    {
    $stBar1.text = "Contacting McAfee ePO - " + $computername.ToUpper()
    $Path64 = "\\$Computername\C$\Program Files (x86)"
    If (Test-Path $Path64){$cmd = "cmd.exe /c psexec.exe \\$Computername -d `"C:\Program Files (x86)\McAfee\VirusScan Enterprise\MCUpdate.exe`" /update /quiet"}
    Else {$cmd = "cmd.exe /c psexec.exe \\$Computername -d `"C:\Program Files\McAfee\VirusScan Enterprise\MCUpdate.exe`" /update /quiet"}
    $McAfeeUpdate = Invoke-Expression $cmd
    $stBar1.text = "McAfee DAT update started on " + $computername.ToUpper()
    if(test-path $lfile){(get-date -uformat "%Y-%m-%d-%H:%M") + "," + $user + "," + $computername + ",QFix-UpdateMcAfeeDAT" | out-file -filepath $lfile -append}    
    }
else{$stBar1.text = "Could not contact " + $computername.ToUpper()}

}

$SetDomain=
{
$CTD = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a domain to connect to", "Connect to domain", "")
if ($CTD){
    Connect-QADService $CTD
    Update-FormTitle
    }
}


$OnLoadForm_StateCorrection=
{
	$form1.WindowState = $InitialFormWindowState #Correct the initial state of the form to prevent the .Net maximized form issue
}

#----------------------------------------------
#region Generated Form Code
$form1 = New-Object System.Windows.Forms.Form
Update-FormTitle
$form1.Name = "form1"
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 750
$System_Drawing_Size.Height = 621
$form1.ClientSize = $System_Drawing_Size
$form1.StartPosition = "CenterScreen"
$Form1.KeyPreview = $True
$Form1.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$Form1.Close()}})

# Menu Strip #
$MenuStrip = new-object System.Windows.Forms.MenuStrip
$MenuStrip.backcolor = "ControlLight"
$FileMenu = new-object System.Windows.Forms.ToolStripMenuItem("&File")
$ViewMenu = new-object System.Windows.Forms.ToolStripMenuItem("&View")
$QFixMenu = new-object System.Windows.Forms.ToolStripMenuItem("&Quick Fix")


$FileDomain = new-object System.Windows.Forms.ToolStripMenuItem("Connect to domain...")
$FileDomain.add_Click($SetDomain)
$FileMenu.DropDownItems.Add($FileDomain) > $null

$FileUser = new-object System.Windows.Forms.ToolStripMenuItem("Find &User in AD")
$FileUser.add_Click($FindUser)
$FileMenu.DropDownItems.Add($FileUser) > $null

$FilePCUser = new-object System.Windows.Forms.ToolStripMenuItem("Find User on &PC")
$FilePCUser.add_Click($FindPCUser)
$FileMenu.DropDownItems.Add($FilePCUser) > $null

$FileExit = new-object System.Windows.Forms.ToolStripMenuItem("E&xit")
$FileExit.add_Click({$form1.close()})
$FileMenu.DropDownItems.Add($FileExit) > $null

$McAfeeMenu = new-object System.Windows.Forms.ToolStripMenuItem("&McAfee Logs")
$ViewMenu.DropdownItems.Add($McAfeeMenu) > $null

    $McAfeeAP = new-object System.Windows.Forms.ToolStripMenuItem("&Access Protection")
    $McAfeeAP.add_Click({$McAfeeFile = "AP"; McAfeeLogs})
    $McAfeeMenu.DropDownItems.Add($McAfeeAP) > $null
    
    $McAfeeAgent = new-object System.Windows.Forms.ToolStripMenuItem("A&gent")
    $McAfeeAgent.add_Click({$McAfeeFile = "Agent"; McAfeeLogs})
    $McAfeeMenu.DropDownItems.Add($McAfeeAgent) > $null

    $McAfeeOAS = new-object System.Windows.Forms.ToolStripMenuItem("&On Access Scan")
    $McAfeeOAS.add_Click({$McAfeeFile = "OAS"; McAfeeLogs})
    $McAfeeMenu.DropDownItems.Add($McAfeeOAS) > $null

    $McAfeeODS = new-object System.Windows.Forms.ToolStripMenuItem("On &Demand Scan")
    $McAfeeODS.add_Click({$McAfeeFile = "ODS"; McAfeeLogs})
    $McAfeeMenu.DropDownItems.Add($McAfeeODS) > $null

    $McAfeeUD = new-object System.Windows.Forms.ToolStripMenuItem("&Updates")
    $McAfeeUD.add_Click({$McAfeeFile = "UD"; McAfeeLogs})
    $McAfeeMenu.DropDownItems.Add($McAfeeUD) > $null

$WSUSMenu = new-object System.Windows.Forms.ToolStripMenuItem("&WSUS Logs")
$ViewMenu.DropdownItems.Add($WSUSMenu) > $null
    
    $WSUSReport = new-object System.Windows.Forms.ToolStripMenuItem("Reporting")
    $WSUSReport.add_Click({$WSUSFile = "Report"; WSUSLogs})
    $WSUSMenu.DropDownItems.Add($WSUSReport) > $null
    
    $WSUSUpdates = new-object System.Windows.Forms.ToolStripMenuItem("Updates")
    $WSUSUpdates.add_Click({$WSUSFile = "Updates"; WSUSLogs})
    $WSUSMenu.DropDownItems.Add($WSUSUpdates) > $null

$ViewEventVwr = new-object System.Windows.Forms.ToolStripMenuItem("Event Viewer")
$ViewEventVwr.add_Click({EventViewer})
$ViewMenu.DropdownItems.Add($ViewEventVwr) > $null

$ViewServices = new-object System.Windows.Forms.ToolStripMenuItem("Services")
$ViewServices.add_Click({Services})
$ViewMenu.DropdownItems.Add($ViewServices) > $null

$ViewUsersGroups = new-object System.Windows.Forms.ToolStripMenuItem("Users/Groups")
$ViewUsersGroups.add_Click({UsersGroups})
$ViewMenu.DropdownItems.Add($ViewUsersGroups) > $null

$QFixGPUpdate = new-object System.Windows.Forms.ToolStripMenuItem("Group Policy - Update")
$QFixGPUpdate.add_Click({Update-GroupPolicy})
$QFixMenu.DropdownItems.Add($QFixGPUpdate) > $null

$QFixLockPC = new-object System.Windows.Forms.ToolStripMenuItem("Lock Computer")
$QFixLockPC.add_Click({Lock-Computer})
$QFixMenu.DropdownItems.Add($QFixLockPC) > $null

$QFixMcAfeeDAT = new-object System.Windows.Forms.ToolStripMenuItem("McAfee - Update DAT")
$QFixMcAfeeDAT.add_Click({Update-McAfeeDAT})
$QFixMenu.DropdownItems.Add($QFixMcAfeeDAT) > $null

$QFixReaderIE = new-object System.Windows.Forms.ToolStripMenuItem("Reader - Fix IE Plugin")
$QFixReaderIE.add_Click({ReaderIE})
$QFixMenu.DropdownItems.Add($QFixReaderIE) > $null

$QFixRenamePC = new-object System.Windows.Forms.ToolStripMenuItem("Rename Computer")
$QFixRenamePC.add_Click({Rename-Computer})
$QFixMenu.DropdownItems.Add($QFixRenamePC) > $null

$QFixSUSDetect = new-object System.Windows.Forms.ToolStripMenuItem("WSUS - Detect")
$QFixSUSDetect.add_Click({Invoke-WSUSDetect})
$QFixMenu.DropdownItems.Add($QFixSUSDetect) > $null

$QFixSUSReport = new-object System.Windows.Forms.ToolStripMenuItem("WSUS - Report")
$QFixSUSReport.add_Click({Invoke-WSUSReport})
$QFixMenu.DropdownItems.Add($QFixSUSReport) > $null

$QFixSUSClientID = new-object System.Windows.Forms.ToolStripMenuItem("WSUS - Reset Client ID")
$QFixSUSClientID.add_Click({Reset-SUSClientID})
$QFixMenu.DropdownItems.Add($QFixSUSClientID) > $null


$MenuStrip.Items.Add($FileMenu) > $null
$MenuStrip.Items.Add($ViewMenu) > $null
$MenuStrip.Items.Add($QFixMenu) > $null
$form1.Controls.Add($MenuStrip)

# Textbox 1 - Computer Name #
$txt1 = New-Object System.Windows.Forms.TextBox
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 125
$System_Drawing_Size.Height = 20
$txt1.Size = $System_Drawing_Size
$txt1.DataBindings.DefaultDataSourceUpdateMode = 0
$txt1.Name = "txt1"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 30
$txt1.Location = $System_Drawing_Point
$txt1.TabIndex = 0
$form1.Controls.Add($txt1)


# Label 2 - Results #
$lbl2 = New-Object System.Windows.Forms.Richtextbox
$lbl2.TabIndex = 7
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 600
$System_Drawing_Size.Height = ($form1.height - 96)
$lbl2.Size = $System_Drawing_Size
$lbl2.BorderStyle = 2
$lbl2.anchor = "bottom, left, top, right"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 145
$System_Drawing_Point.Y = 30
$lbl2.Location = $System_Drawing_Point
$lbl2.DataBindings.DefaultDataSourceUpdateMode = 0
$lbl2.Name = "lbl2"
$lbl2.Visible = $true
$form1.Controls.Add($lbl2)


# Group 1 - Information #
$grp1 = New-Object System.Windows.Forms.GroupBox
$grp1.Name = "grp1"
$grp1.Text = "Information"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 125
$System_Drawing_Size.Height = 184
$grp1.Size = $System_Drawing_Size
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 86
$grp1.Location = $System_Drawing_Point
$grp1.TabStop = $False
$grp1.TabIndex = 4
$grp1.DataBindings.DefaultDataSourceUpdateMode = 0
$form1.Controls.Add($grp1)


# Group 2 - Tools #
$grp2 = New-Object System.Windows.Forms.GroupBox
$grp2.Name = "grp2"
$grp2.Text = "Tools"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 125
$System_Drawing_Size.Height = 154
$grp2.Size = $System_Drawing_Size
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 277
$grp2.Location = $System_Drawing_Point
$grp2.TabStop = $False
$grp2.TabIndex = 5
$grp2.DataBindings.DefaultDataSourceUpdateMode = 0
$form1.Controls.Add($grp2)

# Button 0 - PC Search #
$btn0 = New-Object System.Windows.Forms.Button
$btn0.TabIndex = 0
$btn0.Name = "btn0"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn0.Size = $System_Drawing_Size
$btn0.UseVisualStyleBackColor = $True
$btn0.Text = "&Search for PC"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 19
$System_Drawing_Point.Y = 56
$btn0.Location = $System_Drawing_Point
$btn0.DataBindings.DefaultDataSourceUpdateMode = 0
$btn0.add_Click($btn0_OnClick)
$form1.Controls.Add($btn0)

# Button 1 - System Info #
$btn1 = New-Object System.Windows.Forms.Button
$btn1.TabIndex = 1
$btn1.Name = "btn1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn1.Size = $System_Drawing_Size
$btn1.UseVisualStyleBackColor = $True
$btn1.Text = "System &Info"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 21
$btn1.Location = $System_Drawing_Point
$btn1.DataBindings.DefaultDataSourceUpdateMode = 0
$btn1.add_Click($btn1_OnClick)
$grp1.Controls.Add($btn1)


# Button 2 - Local Admins #
$btn2 = New-Object System.Windows.Forms.Button
$btn2.TabIndex = 2
$btn2.Name = "btn2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn2.Size = $System_Drawing_Size
$btn2.UseVisualStyleBackColor = $True
$btn2.Text = "&Local Admins"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 52
$btn2.Location = $System_Drawing_Point
$btn2.DataBindings.DefaultDataSourceUpdateMode = 0
$btn2.add_Click($btn2_OnClick)
$grp1.Controls.Add($btn2)


# Button 3 - Applications #
$btn3 = New-Object System.Windows.Forms.Button
$btn3.TabIndex = 3
$btn3.Name = "btn3"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn3.Size = $System_Drawing_Size
$btn3.UseVisualStyleBackColor = $True
$btn3.Text = "&Applications"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 83
$btn3.Location = $System_Drawing_Point
$btn3.DataBindings.DefaultDataSourceUpdateMode = 0
$btn3.add_Click($btn3_OnClick)
$grp1.Controls.Add($btn3)


# Button 4 - Remote Desktop #
$btn4 = New-Object System.Windows.Forms.Button
$btn4.TabIndex = 4
$btn4.Name = "btn4"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn4.Size = $System_Drawing_Size
$btn4.UseVisualStyleBackColor = $True
$btn4.Text = "Remote &Desktop"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 20
$btn4.Location = $System_Drawing_Point
$btn4.DataBindings.DefaultDataSourceUpdateMode = 0
$btn4.add_Click($btn4_OnClick)
$grp2.Controls.Add($btn4)


# Button 5 - Remote Assistance #
$btn5 = New-Object System.Windows.Forms.Button
$btn5.TabIndex = 5
$btn5.Name = "btn5"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn5.Size = $System_Drawing_Size
$btn5.UseVisualStyleBackColor = $True
$btn5.Text = "R&emote Assistance"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 51
$btn5.Location = $System_Drawing_Point
$btn5.DataBindings.DefaultDataSourceUpdateMode = 0
$btn5.add_Click($btn5_OnClick)
$grp2.Controls.Add($btn5)


# Button 6 - File Structure #
$btn6 = New-Object System.Windows.Forms.Button
$btn6.TabIndex = 6
$btn6.Name = "btn6"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn6.Size = $System_Drawing_Size
$btn6.UseVisualStyleBackColor = $True
$btn6.Text = "View &C Drive"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 83
$btn6.Location = $System_Drawing_Point
$btn6.DataBindings.DefaultDataSourceUpdateMode = 0
$btn6.add_Click($btn6_OnClick)
$grp2.Controls.Add($btn6)


# Button 7 - Restart Computer #
$btn7 = New-Object System.Windows.Forms.Button
$btn7.TabIndex = 7
$btn7.Name = "btn7"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn7.Size = $System_Drawing_Size
$btn7.UseVisualStyleBackColor = $True
$btn7.Text = "&Restart Computer"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 115
$btn7.Location = $System_Drawing_Point
$btn7.DataBindings.DefaultDataSourceUpdateMode = 0
$btn7.add_Click($btn7_OnClick)
$grp2.Controls.Add($btn7)


# Button 8 - Startup Items #
$btn8 = New-Object System.Windows.Forms.Button
$btn8.TabIndex = 9
$btn8.Name = "btn8"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn8.Size = $System_Drawing_Size
$btn8.UseVisualStyleBackColor = $True
$btn8.Text = "Startup I&tems"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 114
$btn8.Location = $System_Drawing_Point
$btn8.DataBindings.DefaultDataSourceUpdateMode = 0
$btn8.add_Click($btn8_OnClick)
$grp1.Controls.Add($btn8)

# Button 9 - Show Processes #
$btn9 = New-Object System.Windows.Forms.Button
$btn9.TabIndex = 10
$btn9.Name = "btn9"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn9.Size = $System_Drawing_Size
$btn9.UseVisualStyleBackColor = $True
$btn9.Text = "&Processes"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 147
$btn9.Location = $System_Drawing_Point
$btn9.DataBindings.DefaultDataSourceUpdateMode = 0
$btn9.add_Click($btn9_OnClick)
$grp1.Controls.Add($btn9)

# Button 10 - End Process #
$btn10 = New-Object System.Windows.Forms.Button
$btn10.TabIndex = 11
$btn10.Name = "btn10"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn10.Size = $System_Drawing_Size
$btn10.anchor = "bottom, left"
$btn10.UseVisualStyleBackColor = $True
$btn10.Text = "End Process"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 19
$System_Drawing_Point.Y = ($form1.height - 90)
$btn10.Location = $System_Drawing_Point
$btn10.DataBindings.DefaultDataSourceUpdateMode = 0
$btn10.add_Click($btn10_OnClick)
$btn10.Visible = $false
$form1.Controls.Add($btn10)

# Button 11 - Uninstall App #
$btn11 = New-Object System.Windows.Forms.Button
$btn11.TabIndex = 12
$btn11.Name = "btn11"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn11.Size = $System_Drawing_Size
$btn11.anchor = "bottom, left"
$btn11.UseVisualStyleBackColor = $True
$btn11.Text = "Uninstall App"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 19
$System_Drawing_Point.Y = ($form1.height - 90)
$btn11.Location = $System_Drawing_Point
$btn11.DataBindings.DefaultDataSourceUpdateMode = 0
$btn11.add_Click($btn11_OnClick)
$btn11.Visible = $false
$form1.Controls.Add($btn11)

# Button 12 - Remove Admin #
$btn12 = New-Object System.Windows.Forms.Button
$btn12.TabIndex = 13
$btn12.Name = "btn12"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn12.Size = $System_Drawing_Size
$btn12.anchor = "bottom, left"
$btn12.UseVisualStyleBackColor = $True
$btn12.Text = "Remove Admin"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 19
$System_Drawing_Point.Y = ($form1.height - 90)
$btn12.Location = $System_Drawing_Point
$btn12.DataBindings.DefaultDataSourceUpdateMode = 0
$btn12.add_Click($btn12_OnClick)
$btn12.Visible = $false
$form1.Controls.Add($btn12)

# Button 13 - Remove Startup #
$btn13 = New-Object System.Windows.Forms.Button
$btn13.TabIndex = 14
$btn13.Name = "btn13"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 110
$System_Drawing_Size.Height = 25
$btn13.Size = $System_Drawing_Size
$btn13.anchor = "bottom, left"
$btn13.UseVisualStyleBackColor = $True
$btn13.Text = "Remove Item"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 19
$System_Drawing_Point.Y = ($form1.height - 90)
$btn13.Location = $System_Drawing_Point
$btn13.DataBindings.DefaultDataSourceUpdateMode = 0
$btn13.add_Click($btn13_OnClick)
$btn13.Visible = $false
$form1.Controls.Add($btn13)

## Listview 1 ##
$list1 = New-Object System.Windows.Forms.ListView
$list1.DataBindings.DefaultDataSourceUpdateMode = 0
$list1.Name = "list1"
$list1.anchor = "bottom, left, top, right"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 145
$System_Drawing_Point.Y = 30
$list1.Location = $System_Drawing_Point
$list1.TabIndex = 3
$list1.View = [System.Windows.Forms.View]"Details"
$list1.Size = new-object System.Drawing.Size(600, ($form1.height - 96))
$list1.FullRowSelect = $true
$list1.GridLines = $true
$columnnames = "Name","Path"
$list1.Columns.Add("Name", 150) | out-null
$list1.Columns.Add("Path", 450) | out-null
$list1.visible = $false
$form1.Controls.Add($list1)

## Status Bar ##
$stBar1 = New-Object System.Windows.Forms.StatusBar
$stBar1.Name = "stBar1"
$stBar1.Text = "Ready"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 750
$System_Drawing_Size.Height = 22
$stBar1.Size = $System_Drawing_Size
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 0
$System_Drawing_Point.Y = 380
$stBar1.Location = $System_Drawing_Point
$stBar1.DataBindings.DefaultDataSourceUpdateMode = 0
$stBar1.TabIndex = 0
$form1.Controls.Add($stBar1)


#endregion Generated Form Code

$InitialFormWindowState = $form1.WindowState #Save the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection) #Init the OnLoad event to correct the initial state of the form
$form1.ShowDialog()| Out-Null #Show the Form

} #End Function GenerateForm

# Load Quest ActiveRoles Snapin
$Quest = Get-PSSnapin Quest.ActiveRoles.ADManagement -ea silentlycontinue
if (!$Quest) {
   "Loading Quest.ActiveRoles.ADManagement Snapin"
   Add-PSSnapin Quest.ActiveRoles.ADManagement
   if (!$?) {"Need to install AD Snapin from http://www.quest.com/powershell";exit}
}

# Enable VB messageboxes
$vbmsg = new-object -comobject wscript.shell

# Get local user/computer info
$user = $env:username
$userPC = $env:computername
$userdomain = $env:userdomain
$lfile = "C:\ACSA\logs.log"

$Trace32 = "C:\ACSA\Trace32.exe"
if (!(Test-Path (Join-Path $Env:windir "system32\trace32.exe"))) {Copy-Item $Trace32 (Join-Path $Env:windir "system32")}

$PSExec = "C:\ACSA\psexec.exe"
if (!(Test-Path (Join-Path $Env:windir "system32\psexec.exe"))) {Copy-Item $PSExec (Join-Path $Env:windir "system32")}

$Version = "v1.0.2"

#Call the Function
GenerateForm