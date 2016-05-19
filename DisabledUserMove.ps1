#requires -Modules ActiveDirectory

<#
.SYNOPSIS
    This script will scan AD and move disabled users.

.NOTES
	Author		: Benjamin Hassen
	File Name	: DisabledUserMove.ps1
    Created		: 02/08/2016
    Version		: 1.0

.CHANGELOG
    1.0 - 02/18/2016
        Created
    1.1 - 02/26/2016
        Added Hide from Address Book function

.SYNTAX
    DisabledUserMove.ps1

.EXAMPLE
    .\DisabledUserMove

#>

[CmdletBinding(SupportsShouldProcess)]
Param()

####################################################
#region: Config
###########################
 [RegEx] $Ignore 		= 'Service|Exchange|Termed'					# Used to ignore accounts that you do not want moved.
[String] $Domain 		= 'Contoso.com'								# Domain to clean up.
[String] $SearchBase 	= 'OU=Users,DC=Contoso,DC=com'				# Root OU that you would like to move users from.
[String] $DisabledOU	= 'OU=Termed,OU=Disabled,DC=Contoso,DC=com'	# Where to put disabled users.
[String] $Exchange		= 'Exchange.Contoso.com'					# FQDN of Exchange Server
[String] $LogDir		= 'C:\Temp'									# Where to write the log file
###########################
#endregion: Config
####################################################


####################################################
#region: Functions
###########################

Function Get-ExchSession 
{
	Param (
        [String]$ExchServer
    )

	$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
                            -ConnectionUri http://$ExchServer/PowerShell/ `
                            -Authentication Kerberos 

	Import-PSSession $Session -AllowClobber | Write-Verbose
}

###########################
#endregion: Functions
####################################################


####################################################
#region: Main
###########################

Try 
{
  Import-Module ActiveDirectory -ErrorAction Stop
} 
Catch 
{
  Write-Output "[ERROR]`t ActiveDirectory Module couldn't be loaded. Script will stop!"
  Exit 1
}

If(!(Test-Path $LogDir)){New-Item -Path $LogDir -ItemType Directory -Force}

[String]$Date = Get-Date -Format "MM-dd-yy hh:mm"
[String]$Log = "$LogDir\DisabledUserMove.Log"

# Get a list of DSB disabled users that need moved
[Array]$DisabledUsers = Get-ADUser -Filter {Enabled -eq $false} `
                                -SearchBase $SearchBase `
                                -SearchScope Subtree `
                                -Server $Domain `
                                -Properties Mail | 
                                Where-Object {$_.DistinguishedName -notmatch $Ignore}


If($DisabledUsers.Count -gt '0') 
{
    Get-ExchSession $Exchange 

    ForEach($Usr in $DisabledUsers) 
    {
        Try 
        {
            Set-Mailbox -Identity $Usr.Mail -HiddenFromAddressListsEnabled $true
            Set-ADUser -Identity $Usr -Description "Moved to Disabled OU $Date" -Server $Domain
            Move-ADObject -Identity $Usr -TargetPath $DisabledOU -Server $Domain			
            "[Info]$Date - $($Usr.Name) has been moved.<br>" | Out-File $Log -Append
        } 
        Catch 
        {
            "[Error]$Date - $($Usr.Name) couldn't be moved.<br>" | Out-File $Log -Append
        }
    }
    Remove-PSSession -ComputerName $Exchange
} 
Else 
{
    "[Info]$Date - There were no disabled users to move.<br>" | Out-File $Log -Append
}

$Results = Get-Content $Log
$ArchLog = "DisabledUserMove-$(Get-Date -Format 'MMddyy').log"
Rename-Item -Path $Log -NewName $ArchLog
