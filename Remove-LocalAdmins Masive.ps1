#Remove-LocalAdmins Masive
#Author: Alvaro Saenz
#Creation Date: 24/September/2016

#This script will do the following actions:
#1.- Get a computer list from a TXT file
#2.- Get a list of users from a TXT to be removed from the local admin group
#3.- Do a ping to every computer on the list, if the computer is offline it will skip it and pass to the next one
#4.- If the computer answers the ping it will search into the local admins group and if a user matches with a user from the user list it will be removed
#5.- Creates a log with all the transactions

# Log Time Variables
$Date = Get-Date -UFormat %b-%m-%Y
$Hour = (Get-Date).Hour
$Minuntes = (Get-Date).Minute
$Log = "C:\Scripts\Remove-LocalAdmins Masive-" + $Date + "-" + $Hour + "-" + $Minuntes + ".log"

#Creates a log file for this process
Start-Transcript -Path $Log  -Force 

#List of computers to be check
$ComputerNames = Get-Content ".\TestLocal.txt"

#Ping the computers on the list
foreach ($ComputerName in $ComputerNames) {

#If theres no ping answer pass to the next one
if ( -not(Test-Connection $ComputerName -Quiet -Count 1 -ErrorAction Continue )) {
Write-Output "Computer $ComputerName not reachable (PING) - Skipping this computer..." }

#If computer does answer the ping
Else { Write-Output "Computer $computerName is online"

#Search into the local Administrators group
$LocalGroupName = "Administrators"
$Group = [ADSI]("WinNT://$computerName/$localGroupName,group")
$Group.Members() |
foreach {
$AdsPath = $_.GetType().InvokeMember('Adspath', 'GetProperty', $null, $_, $null)
$A = $AdsPath.split('/',[StringSplitOptions]::RemoveEmptyEntries)
$Names = $a[-1] 
$Domain = $a[-2]

#Gets the list of users to be removed from a TXT that you specify and checks if theres a match in the local group
foreach ($name in $names) {
Write-Output "Verifying the local admin users on computer $computerName" 
$Admins = Get-Content ".\TestUsers.txt"
foreach ($Admin in $Admins) {
if ($name -eq $Admin) {

#If it finds a match it will notify you and then remove the user from the local administrators group
Write-Output "User $Admin found on computer $computerName ... "
$Group.Remove("WinNT://$computerName/$domain/$name")
Write-Output "Removed" }}}}}

#Passes all the information of the operations made into the log file
}Stop-Transcript