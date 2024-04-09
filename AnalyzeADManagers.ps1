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
    [Parameter(Mandatory=$true)]
    $OutputPath,
    [Parameter(Mandatory=$false)]
    $OrganizationalUnit
)

# Region Output
$OutputFileDisabledManagers = Join-Path -Path $OutputPath -ChildPath "DisabledManagerUsers.txt"
$OutputFileEmptyManagers = Join-Path -Path $OutputPath -ChildPath "NoManagerSet.txt"
# End Region Output
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
$MessageTargetDC = "Target Domain Controller selected: "
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
function Get-DomainControllersforDomain
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $DomainName
    )

    Try
    {
        [array]$DCs = Get-ADDomainController -DomainName $DomainName -Discover -ErrorAction Stop| Select-Object $ColumnsDCsDatagrid
        Write-LogFile -Message "$($LoadSuccess) Domain Controllers."
    }

    catch
    {
        Write-LogFile "$($LoadError)$($ADDCs)." -ErrorInfo $_
    }

    Return $DCs
}
function Get-ManagerUser
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        $Username
    )

    $Filter = "Enabled -eq ""False"" -and DistinguishedName -eq ""$UserName"""
    $Enabled = Get-ADUser -Filter $Filter -Properties DisplayName | Select-Object DisplayName,SamAccountName
    Return $Enabled
}
function Get-ADUserManagerData
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$False)]
        $OrganizationalUnit
    )

    try
    {
        $UserManagerData = Get-ADUser -LDAPFilter '(manager=*)' -SearchBase $OrganizationalUnit -Properties Manager | Select-Object  SamAccountName,Manager
        Write-LogFile -Message "Successfully retrieved $($UserManagerData.Count) Users from $($OrganizationalUnit)"
    }

    catch
    {
        Write-LogFile -Message "Failure loading users from $($OrganizationalUnit)" -ErrorInfo $_
    }

    Return $UserManagerData
}
function Get-EmptyManagerAttribute
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $OrganizationalUnit
    )

    try
    {
        $EmptyManagerData = Get-ADUser -LDAPFilter '(!(manager=*))' -SearchBase $OrganizationalUnit -Properties Displayname | Select-Object  DisplayName,SamAccountName
        Write-LogFile -Message "Successfully retrieved $($EmptyManagerData.Count) Users from $($OrganizationalUnit)"
    }

    catch
    {
        Write-LogFile -Message "Failure loading users from $($OrganizationalUnit)" -ErrorInfo $_
    }

    Return $EmptyManagerData
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
    $Script:CurrentDomainController = ($Script:DomainControllers | Select-Object -First 1).Hostname.Value
    Write-LogFile -Message "$($MessageTargetDC)$($Script:CurrentDomainController)."
    
    if (!($Script:DomainFQDN -like "*" -and $Script:CurrentDomainController -Like "*"))
    {
        Write-LogFile -Message $MessageNoDomainInfo
    }
}


# If an OU was specified
if ($OrganizationalUnit)
{
    # the container equals $OrganizationalUnit
    $Container = $OrganizationalUnit
}

# If no OU was specified (e.g. the whole domain should be searched)
else
{
    # the container is set to the domain DN
    $Container = $Script:DomainDN
}

# Query users with manager set
$UserManagerData = Get-ADUserManagerData -OrganizationalUnit $Container
# Query users without a manager
$EmptyManagerData = Get-EmptyManagerAttribute -OrganizationalUnit $Container

# Write CSV file with user where no manager is set
$EmptyManagerData | Export-Csv -Path $OutputFileEmptyManagers -NoTypeInformation

# Create a custom object to store disabled manager users
$DisabledManagers = New-Object -TypeName pscustomobject

# iterate through all users with manager set
foreach ($User in $UserManagerData)
{
    # Retrieve user account for manager of current loop object
    $Enabled = Get-ManagerUser -Username $user.Manager

    # if the manager user is disabled
    if ($Enabled -match "\S")
    {
        $DisabledManagers | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $Enabled.SamAccountName
        $DisabledManagers | Add-Member -MemberType NoteProperty -Name DisplayName -Value $Enabled.DisplayName
    }
}

# Write CSV file with user where manager is disabled
$DisabledManagers | Export-Csv -Path $OutputFileDisabledManagers -NoTypeInformation
