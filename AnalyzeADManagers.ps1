<#
.SYNOPSIS
    A script to analyze the "manager" attribute on users in Active Directory
.DESCRIPTION
    Script to analyze the "manager" attribute on users in Active Directory. Checks if the value of the "manager"
    attribute in a certain user points to a disabled user or is empty. The scripts generates a CSV file at the end.
.EXAMPLE
    AnalyzeADManagers.ps1 -OutputPath C:\Reports\InvalidManagers.csv
#>

[CmdLetBinding()]
Param
(
    [Parameter(Mandatory=$false)]
    $OrganizationalUnit
)

#Region Messages
$LoadSuccess = "Successfully loaded "
$LoadError = "Error loading "
$LoadData = "data for "
$LoadGroupMembers = "group members of "
$AddSuccess = "Successfully added "
$AddError = "Error adding "
$ModifySuccess = "Successfully modified "
$ModifyError = "Error modifying "
$RemoveSuccess = "Successfully removed "
$RemoveError = "Error removing "
$EnableSuccess = "Successfully enabled "
$EnableError = "Error enabling "
$DisableSuccess = "Successfully disabled "
$DisableError = "Error disabling "
$InvalidOUSelection = "An invalid OU was selected. Falling back to previous value..."
$InvalidDomainSelection = "An invalid Domain was selected. Falling back to previous value..."
$LoadModuleSuccess = " Module successfully loaded."
$LoadModuleError = "Error loading module "
$ModuleNotInstalled = " Module not installed"
$LoadSnapinSuccess = " Snapin successfully loaded."
$LoadSnapinError = "Error loading Snapin "
$EX_RecipientMGMTSnapin_NotInstalled = "Exchange Recipient Management Shell not installed. Please install the Exchange 2019 CU12 and above Management Tools-Only install. See: https://docs.microsoft.com/en-us/Exchange/manage-hybrid-exchange-recipients-with-management-tools"
$RegInfo = "configuration information from registry."
$Exists  = "exists."
$MessageConfirmationEnabled = "Confirmation for all destructive actions is enabled." 
$MessageConfirmationDisabled = "Confirmation for all destructive actions is disabled."
$MessageConfirmNoConfirmation = "Are you sure you want to disable confirmation dialogs for all destructive actions?"
$MessageLoggingEnabled = "Logging successfully enabled!"
$MessageLoggingDisabled = "Logging successfully disabled!"
$MessageNoDomainFound = "No Domainname could be determined. Please logon with a domain account!"
$MessageNoAdSite = "AD Site could not be determined."
$MessageNoDomainInfo = "No domain information could be retrieved."
$MessageDomainFQDN = "Current Domain FQDN: "
$MessageDomainDN = "Current Domain DN: "
$MessageDomainCN = "Domain canonical name: "
$MessageDCsFouned = "Domain Controllers found: "
$MessageTargetDC = "Target Domain Controller: "
$MessageDefaultOU = "Default AD container is "
$MessageUPNSuffixes = "UPN Suffixes found:"
$MessageNoObjectSelected = "No object was selected!"
$MessageConfirmRemove = "Do you really want to remove "
$MessageConfirmDisable = "Do you really want to disable "
$MessageNOUGSelected = "Selected Group is not a Universal Group! Please select another group."
$MessageAlreadyMailEnabled = "The selected object is already mail-enabled. Please select another object!"
$MessageLastEmailAddress = "Cannot remove last email address!"
$MessageMandatoryAttributesMissing = "Onject not created. Mandatory attributes missing..."
$MessageProxyAddressNotRemoved = "Address not be removed"
$MessageMemberNotRemoved = "Member not be removed"
$MessageRecipientNotRemoved = "Recipient not removed"
#EndRegion Messages
#Region Logging
[string]$Script:LogFileNamePrefix = "AnalyzeADManagers"
[string]$Script:LogfileName = ($LogFileNamePrefix + "_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now)
[string]$Script:LogPath = $ScriptParentPath
[string]$script:LogFilePath = Join-Path -Path $Script:LogPath -ChildPath $Script:LogfileName    
[string]$Script:LogFileStart = "Logging started"
[string]$Script:LogFileStop = "Logging stopped"
$Script:NoLogging = $true
#EndRegion Logging
function Write-LogFile
{
    # Logging function, used for progress and error logging...
    # Uses the globally (script scoped) configured variables 'LogFilePath' to identify the logfile and 'NoLogging' to disable it.
    #
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$LogPrefix,
        [System.Management.Automation.ErrorRecord]$ErrorInfo = $null
    )

    # Prefix the string to write with the current Date and Time, add error message if present...
    if ($ErrorInfo)
    {
        $logLine = "{0:d.M.y H:mm:ss} : {1}: {2} Error: {3}" -f [DateTime]::Now, $LogPrefix, $Message, $ErrorInfo.Exception.Message
    }

    Else
    {
        $logLine = "{0:d.M.y H:mm:ss} : {1}: {2}" -f [DateTime]::Now, $LogPrefix, $Message
    }

    if (-not $NoLogging)
    {
        # Create the Script:Logfile and folder structure if it doesn't exist
        if (-not (Test-Path $Script:LogFilePath -PathType Leaf))
        {
            New-Item -ItemType File -Path $Script:LogFilePath -Force -Confirm:$false -WhatIf:$false | Out-Null
            #Add-Content -Value "Logging started." -Path $Script:LogFilePath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        }

        # Write to the Script:Logfile
        Add-Content -Value $logLine -Path $Script:LogFilePath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        Write-Verbose $logLine
    }
    
    Else
    {
        Write-Host $logLine
    }
}
Function ManageModuleAndSnapInLoading
{
    # Function to check for and import modules

    Param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [switch]$Snapin,
        [switch]$IsFile
    )

    if ($Snapin)
    {
        if ((Get-PSSnapin $Name -Registered -ErrorAction SilentlyContinue))
        {
            try
            {
                Add-PSSnapin -Name $Name -ErrorAction Stop
                Write-LogFile -Message "$($Name)$($LoadSnapinSuccess)"
            }
            
            catch
            {
                Write-LogFile -Message "$($LoadSnapinError)$($Name)." -ErrorInfo $_
            }
        }

        Else
        {
            Write-LogFile -Message "$($EX_RecipientMGMTSnapin_NotInstalled)"
        }
    }

    Else
    {
        $IsModuleInstalled = (Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1)
    
        if ($IsFile)
        {
            [System.IO.FileInfo]$Filename = $Name
            $MatchedName = $Filename.Name.Replace(".psm1", "")
        }

        Else
        {
            $MatchedName = $Name
        }

        if ($IsModuleInstalled.Name -match "$($MatchedName)")
        {   
            try
            {
                Import-Module -Name $Name -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
                Write-LogFile -Message "$($Name)$($LoadModuleSuccess)"
            }
            
            catch
            {
                Write-LogFile -Message "$($LoadModuleError)$($Name)." -ErrorInfo $_
            }
        }
    
        Else
        {
            Write-LogFile = "$($Name)$($ModuleNotInstalled)."
            Break
        }    
    }
}
function Show-MessageBox
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $BoxTitle,
        [Parameter(Mandatory=$true)]
        [string]
        $BoxMessageText,
        [Parameter(Mandatory=$true)]
        [string]
        $BoxIcon,
        [Parameter(Mandatory=$true)]
        [int]
        $Buttons
    )

    [System.Windows.Forms.MessageBox]::Show($BoxMessageText, $BoxTitle, $Buttons, $BoxIcon)
}
function Get-ManagerUser
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $Username
    )

    
}
function Get-ADUserManagerData
{
    Param
    (
        [Parameter(Mandatory=$False)]
        $OrganizationalUnit
    )

    try
    {
        $UserManagerData = Get-ADuser -
    }

    catch
    {
        
    }
}

# Do not load AD PSDrive
$Env:ADPS_LoadDefaultDrive = 0
# Import ActiveDirectory Module
ManageModuleAndSnapInLoading -Name ActiveDirectory | Out-Null

# Enumerate Domain Information
if (! ($env:USERDNSDOMAIN -like "*"))
{
    Write-LogFile -Message $MessageNoDomainFound
    Show-MessageBox -BoxTitle $MSGBoxTitleNoDomainFound -BoxMessageText $MessageNoDomainFound -BoxIcon $MSGBoxStop -Buttons $MSGBoxOK
    Break
}

Else
{
    $Script:Domain = Get-ADDomain -Identity $env:USERDNSDOMAIN
    $Script:DomainFQDN = $Script:Domain.DNSRoot
    Write-LogFile -Message "$($MessageDomainFQDN)$($Script:DomainFQDN)."
    $Script:DomainDN = $Script:Domain.DistinguishedName
    Write-LogFile -Message "$($MessageDomainDN)$($Script:DomainDN)."
    $Script:DomainControllers = Get-DomainControllersforDomain -DomainName $Script:DomainFQDN
    Write-LogFile -Message "$($MessageDCsFouned)$($Script:DomainControllers.Hostname)."
    $Script:CurrentDomainController = $Script:DomainControllers[0].Hostname.Value
    Write-LogFile -Message "$($MessageTargetDC)$($Script:CurrentDomainController)."
    
    if (!($Script:DomainFQDN -like "*" -and $Script:CurrentDomainController -Like "*"))
    {
        Write-LogFile -Message $MessageNoDomainInfo
    }
}

# Query users
#
# If an OU was specified
if ($OrganizationalUnit)
{
    $UserManagerData = Get-ADUserManagerData -OrganizationalUnit $OrganizationalUnit
}

else
{
    $UserManagerData = Get-ADUserManagerData -OrganizationalUnit $Script:DomainDN    
}
