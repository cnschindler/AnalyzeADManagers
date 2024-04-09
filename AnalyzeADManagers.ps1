<#
.SYNOPSIS
    A script to analyze the "manager" attribute on users in Active Directory
.DESCRIPTION
    Script to analyze the "manager" attribute on users in Active Directory. Checks if the value of the "manager"
    attribute in a certain user points to a disabled user or is empty. The scripts generates a CSV file at the end.
.EXAMPLE
    AnalyzeADManagers.ps1 -OutputFolder C:\Reports
.EXAMPLE
    AnalyzeADManagers.ps1 -OutputFolder C:\Reports -OrganizationalUnit "Ou=Accounts,dc=domain,dc=com"
.PARAMETER OutputFolder
    This parameter is mandatory. The Full Path to the folder where CSV output files are stored
.PARAMETER OrganizationalUnit
    This parameter is optional. The distinguishedname of an organizational unit in Active Directory to start the search in
#>

[CmdLetBinding()]
Param
(
    [Parameter(Mandatory = $true)]
    $OutputFolder,
    [Parameter(Mandatory = $false)]
    $OrganizationalUnit
)

# User filter. Modify if you want to include/exclude users from the "empty manger" list
[scriptblock]$Userfilter = {$_.SamAccountName -ne "krbtgt" -and $_.SamAccountName -ne "Guest" -and $_.Displayname -ne "Discovery Search Mailbox" -and $_.Displayname -ne "E4E Encryption Store - Active" -and $_.Displayname -notlike "SystemMailbox*" -and $_.SamAccountName -notlike "Healthmailbox*" -and $_.Displayname -notlike "Microsoft Exchange*"}

# Where to log? $true = Console, $false = File (by default in the folder where the script was started)
$Script:NoLogging = $true

#Region Logging
$ScriptParentPath = Split-Path -Path $MyInvocation.MyCommand.Path -Parent # modify if you want to place the logfile in a different folder
[string]$Script:LogFileNamePrefix = 'AnalyzeADManagers' # modify if you want to specify another logfile prefix
[string]$Script:LogfileName = ($LogFileNamePrefix + '_{0:yyyyMMdd-HHmmss}.log' -f [DateTime]::Now)
[string]$Script:LogPath = $ScriptParentPath
[string]$script:LogFilePath = Join-Path -Path $Script:LogPath -ChildPath $Script:LogfileName    
[string]$Script:LogFileStart = 'Logging started'
[string]$Script:LogFileStop = 'Logging stopped'
#EndRegion Logging
# Region Output
$DisabledManagersFilename = 'DisabledManagerUsers.txt' # modify if you want another file name for the list of disabled manager users
$EmptyManagersFileName = 'NoManagerSet.txt' # modify if you want another file name for the list of empty manager attributes
$OutputFileDisabledManagers = Join-Path -Path $OutputFolder -ChildPath $DisabledManagersFilename
$OutputFileEmptyManagers = Join-Path -Path $OutputFolder -ChildPath $EmptyManagersFileName
# End Region Output
#Region Messages
$LoadSuccess = 'Successfully loaded '
$LoadError = 'Error loading '
$LoadModuleSuccess = ' Module successfully loaded.'
$LoadModuleError = 'Error loading module '
$ModuleNotInstalled = ' Module not installed'
$LoadSnapinSuccess = ' Snapin successfully loaded.'
$LoadSnapinError = 'Error loading Snapin '
$EX_RecipientMGMTSnapin_NotInstalled = 'Exchange Recipient Management Shell not installed. Please install the Exchange 2019 CU12 and above Management Tools-Only install. See: https://docs.microsoft.com/en-us/Exchange/manage-hybrid-exchange-recipients-with-management-tools'
$MessageNoDomainFound = 'No Domainname could be determined. Please logon with a domain account!'
$MessageNoDomainInfo = 'No domain information could be retrieved.'
$MessageDomainFQDN = 'Current Domain FQDN: '
$MessageDomainDN = 'Current Domain DN: '
$MessageDCsFouned = 'Domain Controllers found: '
$MessageTargetDC = 'Target Domain Controller selected: '
#EndRegion Messages

function Write-LogFile {
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
    if ($ErrorInfo) {
        $logLine = '{0:d.M.y H:mm:ss} : {1}: {2} Error: {3}' -f [DateTime]::Now, $LogPrefix, $Message, $ErrorInfo.Exception.Message
    }

    Else {
        $logLine = '{0:d.M.y H:mm:ss} : {1}: {2}' -f [DateTime]::Now, $LogPrefix, $Message
    }

    if (-not $NoLogging) {
        # Create the Script:Logfile and folder structure if it doesn't exist
        if (-not (Test-Path $Script:LogFilePath -PathType Leaf)) {
            New-Item -ItemType File -Path $Script:LogFilePath -Force -Confirm:$false -WhatIf:$false | Out-Null
            #Add-Content -Value "Logging started." -Path $Script:LogFilePath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        }

        # Write to the Script:Logfile
        Add-Content -Value $logLine -Path $Script:LogFilePath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        Write-Verbose $logLine
    }
    
    Else {
        Write-Host $logLine
    }
}
Function ManageModuleAndSnapInLoading {
    # Function to check for and import modules

    Param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [switch]$Snapin,
        [switch]$IsFile
    )

    if ($Snapin) {
        if ((Get-PSSnapin $Name -Registered -ErrorAction SilentlyContinue)) {
            try {
                Add-PSSnapin -Name $Name -ErrorAction Stop
                Write-LogFile -Message "$($Name)$($LoadSnapinSuccess)"
            }
            
            catch {
                Write-LogFile -Message "$($LoadSnapinError)$($Name)." -ErrorInfo $_
            }
        }

        Else {
            Write-LogFile -Message "$($EX_RecipientMGMTSnapin_NotInstalled)"
        }
    }

    Else {
        $IsModuleInstalled = (Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1)
    
        if ($IsFile) {
            [System.IO.FileInfo]$Filename = $Name
            $MatchedName = $Filename.Name.Replace('.psm1', '')
        }

        Else {
            $MatchedName = $Name
        }

        if ($IsModuleInstalled.Name -match "$($MatchedName)") {   
            try {
                Import-Module -Name $Name -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
                Write-LogFile -Message "$($Name)$($LoadModuleSuccess)"
            }
            
            catch {
                Write-LogFile -Message "$($LoadModuleError)$($Name)." -ErrorInfo $_
            }
        }
    
        Else {
            Write-LogFile = "$($Name)$($ModuleNotInstalled)."
            Break
        }    
    }
}
function Show-MessageBox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $BoxTitle,
        [Parameter(Mandatory = $true)]
        [string]
        $BoxMessageText,
        [Parameter(Mandatory = $true)]
        [string]
        $BoxIcon,
        [Parameter(Mandatory = $true)]
        [int]
        $Buttons
    )

    [System.Windows.Forms.MessageBox]::Show($BoxMessageText, $BoxTitle, $Buttons, $BoxIcon)
}
function Get-DomainControllersforDomain {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $DomainName
    )

    Try {
        [array]$DCs = Get-ADDomainController -DomainName $DomainName -Discover -ErrorAction Stop | Select-Object $ColumnsDCsDatagrid
        Write-LogFile -Message "$($LoadSuccess) Domain Controllers."
    }

    catch {
        Write-LogFile "$($LoadError)$($ADDCs)." -ErrorInfo $_
    }

    Return $DCs
}
function Get-ManagerUser {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        $Username
    )

    $Filter = "Enabled -eq ""False"" -and DistinguishedName -eq ""$UserName"""
    $Enabled = Get-ADUser -Filter $Filter -Properties DisplayName | Select-Object DisplayName, SamAccountName
    Return $Enabled
}
function Get-ADUserManagerData {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $False)]
        $OrganizationalUnit
    )

    try {
        $UserManagerData = Get-ADUser -LDAPFilter '(manager=*)' -SearchBase $OrganizationalUnit -Properties Manager | Select-Object SamAccountName, Manager
        Write-LogFile -Message "Successfully retrieved $($UserManagerData.Count) Users with manager attribute set from $($OrganizationalUnit)"
    }

    catch {
        Write-LogFile -Message "Failure loading users from $($OrganizationalUnit)" -ErrorInfo $_
    }

    Return $UserManagerData
}
function Get-EmptyManagerAttribute {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        $OrganizationalUnit
    )

    try {
        $EmptyManagerData = Get-ADUser -LDAPFilter '(!(manager=*))' -SearchBase $OrganizationalUnit -Properties Displayname | Select-Object DisplayName, SamAccountName
        Write-LogFile -Message "Successfully retrieved $($EmptyManagerData.Count) Users without manager set from $($OrganizationalUnit)"
    }

    catch {
        Write-LogFile -Message "Failure loading users from $($OrganizationalUnit)" -ErrorInfo $_
    }

    $EmptyManagerData = $EmptyManagerData | Where-Object -FilterScript $Userfilter
    Return $EmptyManagerData
}

# Do not load AD PSDrive
$Env:ADPS_LoadDefaultDrive = 0
# Import ActiveDirectory Module
ManageModuleAndSnapInLoading -Name ActiveDirectory | Out-Null

# Enumerate Domain Information
if (! ($env:USERDNSDOMAIN -like '*')) {
    Write-LogFile -Message $MessageNoDomainFound
    Show-MessageBox -BoxTitle $MSGBoxTitleNoDomainFound -BoxMessageText $MessageNoDomainFound -BoxIcon $MSGBoxStop -Buttons $MSGBoxOK
    Break
}

Else {
    $Script:Domain = Get-ADDomain -Identity $env:USERDNSDOMAIN
    $Script:DomainFQDN = $Script:Domain.DNSRoot
    Write-LogFile -Message "$($MessageDomainFQDN)$($Script:DomainFQDN)."
    $Script:DomainDN = $Script:Domain.DistinguishedName
    Write-LogFile -Message "$($MessageDomainDN)$($Script:DomainDN)."
    $Script:DomainControllers = Get-DomainControllersforDomain -DomainName $Script:DomainFQDN
    Write-LogFile -Message "$($MessageDCsFouned)$($Script:DomainControllers.Hostname)."
    $Script:CurrentDomainController = ($Script:DomainControllers | Select-Object -First 1).Hostname.Value
    Write-LogFile -Message "$($MessageTargetDC)$($Script:CurrentDomainController)."
    
    if (!($Script:DomainFQDN -like '*' -and $Script:CurrentDomainController -Like '*')) {
        Write-LogFile -Message $MessageNoDomainInfo
    }
}


# If an OU was specified
if ($OrganizationalUnit) {
    # the container equals $OrganizationalUnit
    $Container = $OrganizationalUnit
}

# If no OU was specified (e.g. the whole domain should be searched)
else {
    # the container is set to the domain DN
    $Container = $Script:DomainDN
}

# Query users without a manager
$EmptyManagerData = Get-EmptyManagerAttribute -OrganizationalUnit $Container

# Query users with manager set
$UserManagerData = Get-ADUserManagerData -OrganizationalUnit $Container

# Create a hashtable to store disabled manager users
$DisabledManagers = @{}

# iterate through all users with manager set
foreach ($User in $UserManagerData) {
    # Retrieve user account for manager of current loop object
    $Enabled = Get-ManagerUser -Username $user.Manager

    # if the manager user is disabled
    if ($Enabled -match '\S') {
        # if the user was not already added to the hashtable
        if (-Not $DisabledManagers.ContainsKey($Enabled.SamAccountName)) {
            $DisabledManagers.Add($Enabled.DisplayName, $enabled.SamAccountName)
        }
    }

    Clear-Variable Enabled
}

Write-LogFile -Message "Disabled manager users found: $($DisabledManagers.Count)"

# Write CSV file with user where no manager is set
Write-LogFile -Message "Writing outputfile $($OutputFileEmptyManagers)"
$EmptyManagerData | Export-Csv -Path $OutputFileEmptyManagers -NoTypeInformation

# Write CSV file with disabled manager users
Write-LogFile -Message "Writing outputfile $($OutputFileDisabledManagers)"
$DisabledManagers.GetEnumerator() | Select-Object @{l = 'Displayname'; e = { $_.Key } }, @{l = 'SamAccountName'; e = { $_.Value } } | Export-Csv -Path $OutputFileDisabledManagers -NoTypeInformation
