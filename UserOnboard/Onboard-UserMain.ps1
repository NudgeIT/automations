param (
    #If a user is specified here, all group membership will be inherited to the new user
    [Parameter(Mandatory = $false, HelpMessage = '{
        "XType": {
            "Class": "User",
            "IdProperty": "Id",
            "TitleProperty": "FullName",
            "DescriptionProperty": "FullName"
        },
        "Title": "TemplateUser",
        "Description": "Template User",
        "Placeholder": "i.e. John Snow"
    }')]
    [object]$TemplateUser,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "UserName",
        "Description": "UserName",
        "Placeholder": "i.e. jsmith",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value."
        }
    }')]
    [String]$SamAccountName,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "GivenName",
        "Description": "GivenName",
        "Placeholder": "i.e. John",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value."
        }
    }')]
    [String]$GivenName,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "Surname",
        "Description": "Surname",
        "Placeholder": "i.e. Smith",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value."
        }
    }')]
    [String]$Surname,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "E-Mail",
        "Description": "E-Mail",
        "Placeholder": "i.e. $john.smith@yourdomain.com",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please select a value.",
            "ValidatePattern": "This is not a valid e-mail address. i.e. $john.smith@yourdomain.com"
        }
    }')]
    [ValidatePattern('([\w\.\-_]+)?\w+@[\w\-_]+(\.\w+){1,}')]
    [String]$Email,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "ProjectX",
        "Description": "ProjectX",
        "Placeholder": "i.e. Yes/No",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please select a value."
        }
    }')]
    [Boolean]$ProjectX,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "Department",
        "Description": "Department",
        "Placeholder": "i.e. Sales",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please select a value.",
            "ValidateSet": "This is not an accepted value."
        }
    }')]
    [ValidateSet(
        "IT",
        "Marketing",
        "Sales"
    )]
    [String]$Department,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "Title",
        "Description": "Title",
        "Placeholder": "i.e. IT Support",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please select a value."
        }
    }')]
    [String]$Title,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "OfficePhone",
        "Description": "OfficePhone",
        "Placeholder": "i.e. +42 41 139 75 71",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please select a value.",
            "ValidatePattern": "The format is not correct. i.e. i.e. +42 41 139 75 71"
        }
    }')]
    [ValidatePattern('^(\+41 \d{2} \d{3} \d{2} \d{2})$')]
    [String]$OfficePhone
    
)

#Script Configuration
$TerminalServicesProfilePath = ""  #Ex: "\\filesrv\profiles$\$UserName"
$HomeDirectoryPath = "" #Ex: "\\filesrv\users$\$UserName"
$HomeDriveLetter = "" #Ex: "M"
$Domain = "" #Ex: "domain.local"
$domainAdminsSID = "" #Fill in your domain admins group SID
$ExchangeServerUri = "" #Ex: "https://exchange.domain.local:443/PowerShell/"
$WsldPath       = "" #Ex: "D:\automation\pbx\pbx.wsdl"
$ApiUrl         = "" #Ex: "https://10.0.0.100"
###

#Logging configuration
$Script:VerbosePreference = 'SilentlyContinue'
$Script:CustomVerbosePreference = $true

#Catch unhandled errors
$ErrorActionPreference = 'Stop'
trap {
    Throw "There is an error on line '$($_.InvocationInfo.ScriptLineNumber)', character '$($_.InvocationInfo.OffsetInLine)'. $_"
}

#Create AD User
Write-Verbose "Creating AD User." -Verbose:$CustomVerbosePreference
$accountPassword = New-NiboRandomPassword -passLength 10 -lowerCase -upperCase -numbers
$newUserProperties = @{
    DisplayName                 = "$Surname $GivenName"
    UserPrincipalName           = "$GivenName.$Surname@$Domain"
    Name                        = "$Surname $GivenName"
    Title                       = $Title
    Description                 = $Department
    AccountPassword             = ($accountPassword | ConvertTo-SecureString -AsPlainText -Force)
    TerminalServicesProfilePath = $TerminalServicesProfilePath
    HomeDirectory               = $HomeDirectoryPath
    HomeDrive                   = $HomeDriveLetter
}
$boundProperties = $PSBoundParameters
$removeProperties = @(
    'SubstituteEmployee'
    'Department'
    'Framework'
)
foreach ($property in $removeProperties) {
    $boundProperties.Remove($property)
}
$phoneWithoutSpaces = $OfficePhone.Replace(' ', '')
$extension = $phoneWithoutSpaces.Substring($phoneWithoutSpaces.Length - 4, 4)
$customAttributes = @{
    customAttr1 = 'value1'
    customAttr2 = 'value2'
}
$newADUser = .\Create-ADUser.ps1 @newUserProperties @boundProperties -CustomAttributes $customAttributes
##

#Add user to ProjectX groups
$ProjectXGroups = @(
    'ProjectX_Security'
    'ProjectX_EMail'
)
if ($ProjectX) {
    Write-Verbose "Adding user to ProjectX groups." -Verbose:$CustomVerbosePreference
    foreach ($group in $ProjectXGroups) {
        Write-Verbose "Adding user to group '$group'" -Verbose:$CustomVerbosePreference
        .\Add-UserToADGroup.ps1 -SamAccountName $SamAccountName -GroupName $group
    }
}
##

#All new users are added to these groups
$allGroups = @(
    'security_all',
    'vpn_access',
    'distribution_all'
)
Write-Verbose "Adding user to groups." -Verbose:$CustomVerbosePreference
foreach ($group in $allGroups) {
    Write-Verbose "Adding user to group '$group'" -Verbose:$CustomVerbosePreference
    .\Add-UserToADGroup.ps1 -SamAccountName $SamAccountName -GroupName $group
}


#Copy template employee groups
if ($TemplateUser) {
    Write-Verbose "Copy template user groups." -Verbose:$CustomVerbosePreference
    $ldapFilter = "(UserPrincipalName=$($TemplateUser.UPN))"
    $substituteAdUser = Get-ADUser -LDAPFilter $ldapFilter -Properties MemberOf
    if (-not $substituteAdUser) {
        Throw "Unable to retrieve substitute user AD Account using filter '$ldapFilter'"
    }
    foreach ($group in $substituteAdUser.MemberOf) {
        .\Add-UserToADGroup.ps1 -SamAccountName $SamAccountName -GroupName $group
    }    
}
##

#Configure Shares
Write-Verbose "Assigning permissions." -Verbose:$CustomVerbosePreference
if (-not (Test-Path $newUserProperties.TerminalServicesProfilePath)) {
    New-Item -ItemType Directory -Path $newUserProperties.TerminalServicesProfilePath -Force
}
if (-not (Test-Path $newUserProperties.HomeDirectory)) {
    New-Item -ItemType Directory -Path $newUserProperties.HomeDirectory -Force
}
$systemUserSID = 'S-1-5-18'
Add-NTFSAccess -Path $newUserProperties.TerminalServicesProfilePath -Account $systemUserSID -AccessRights FullControl
Add-NTFSAccess -Path $newUserProperties.TerminalServicesProfilePath -Account $domainAdminsSID -AccessRights FullControl
Add-NTFSAccess -Path $newUserProperties.TerminalServicesProfilePath -Account $newADUser.SID -AccessRights Read, ReadAndExecute, Modify
Disable-NTFSAccessInheritance $newUserProperties.TerminalServicesProfilePath -RemoveInheritedAccessRules
Add-NTFSAccess -Path $newUserProperties.HomeDirectory -Account $systemUserSID -AccessRights FullControl
Add-NTFSAccess -Path $newUserProperties.HomeDirectory -Account $domainAdminsSID -AccessRights FullControl
Add-NTFSAccess -Path $newUserProperties.HomeDirectory -Account $newADUser.SID -AccessRights Read, ReadAndExecute, Modify
Disable-NTFSAccessInheritance $newUserProperties.HomeDirectory -RemoveInheritedAccessRules

#Configure Exchange
Write-Verbose "Configure Exchange Account" -Verbose:$CustomVerbosePreference

$ExchangeCredential = Get-AutomationPSCredential -Name 'exchange-admin'
$ImportedCommands = @(
    'Get-Mailbox'
    'Set-Mailbox'
    'Set-CasMailbox'
    'Enable-Mailbox'
    'Add-MailboxFolderPermission'
)
$ExchangeSession = .\Create-ExchangeSession.ps1 -ExchangeCredential $ExchangeCredential -ExchangeServerUri $ExchangeServerUri -ImportedCommands $ImportedCommands
try {
    #Enable Mailbox
    Write-Verbose "Enable Mailbox." -Verbose:$CustomVerbosePreference
    Enable-Mailbox -Identity $SamAccountName -ActiveSyncMailboxPolicy "ActiveSync01"
    Start-Sleep -Seconds 20

    #Add SMTP address
    Write-Verbose "Add SMTP address." -Verbose:$CustomVerbosePreference
    Set-Mailbox -Identity $SamAccountName -EmailAddressPolicyEnabled $false
    Set-Mailbox -Identity $SamAccountName -PrimarySmtpAddress $Email

    #Add calendar access rights
    Write-Verbose "Add calendar permissions." -Verbose:$CustomVerbosePreference
    Add-MailboxFolderPermission "$SamAccountName`:\calendar" -user '' -AccessRights reviewer #Add the user you want to assign permissions to
}
catch {
    Remove-PSSession -Session $ExchangeSession
    Throw "Error configuring Exchange. $_"
}

#Add user to InovaPhone System
Write-Verbose "Add user to Innovaphone system." -Verbose:$CustomVerbosePreference
$newPbxUserProperties = @{
    WsldPath       = $WsldPath
    ApiUrl         = $ApiUrl
    ApiCredential  = Get-AutomationPSCredential -Name 'pbx-admin'
    ApiUserName    = '' #The user that will make the API calls
    DisplayName    = $newUserProperties.DisplayName
    Extension      = $extension
    Email          = "$GivenName.$Surname"
    SamAccountName = $SamAccountName

}
.\Create-PbxUser.ps1 @newPbxUserProperties

Write-Output "New user created. Username: '$SamAccountName'  Password: $accountPassword"