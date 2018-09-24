[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "FirstName",
        "Description": "First Name",
        "Placeholder": "i.e. John",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value. i.e. John"
        }
    }')]
    [String]$GivenName,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "LastName",
        "Description": "Last Name",
        "Placeholder": "i.e. Smith",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value. i.e. Smith"
        }
    }')]
    [String]$Surname,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "Name",
        "Description": "Name",
        "Placeholder": "i.e. John Smith",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value. i.e. John Smith"
        }
    }')]
    [String]$Name,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "Title",
        "Description": "Title",
        "Placeholder": "i.e. Assistant",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value. i.e. Assistant"
        }
    }')]
    [String]$Title,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "HomeDirectory",
        "Description": "HomeDirectory",
        "Placeholder": "i.e. \\server\samaccountname",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value. i.e. \\server\samaccountname"
        }
    }')]
    [String]$HomeDirectory,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "HomeDrive",
        "Description": "HomeDrive",
        "Placeholder": "i.e. M",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value. i.e. M",
            "ValidatePattern": "Please enter input should be one letter from A to Z. i.e. M"
        }
    }')]
    [ValidatePattern('^[A-Za-z]{1}$')]
    [String]$HomeDrive,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "DisplayName",
        "Description": "Display Name",
        "Placeholder": "i.e. John Smith",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value. i.e. John Smith"
        }
    }')]
    [String]$DisplayName,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "SamAccountName",
        "Description": "SamAccountName",
        "Placeholder": "i.e. John Smith",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value. i.e. jsmith"
        }
    }')]
    [String]$SamAccountName,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "UserPrincipalName",
        "Description": "UserPrincipalName",
        "Placeholder": "i.e. john.smith@contoso.com",
        "ErrorMessage": {
            "Mandatory": "The input is not a valid UserPrincipalName. i.e. john.smith@contoso.com"
        }
    }')]
    [ValidatePattern('([\w\.\-_]+)?\w+@[\w-_]+(\.\w+){1,}')]
    [String]$UserPrincipalName,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "E-Mail",
        "Description": "E-Mail",
        "Placeholder": "i.e. john.smith@contoso.com",
        "ErrorMessage": {
            "validatePattern": "The input is not a valid e-mail address. i.e. john.smith@contoso.com"
        }
    }')]
    [ValidatePattern('([\w\.\-_]+)?\w+@[\w-_]+(\.\w+){1,}')]
    [String]$EMail,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "Description",
        "Description": "Description",
        "Placeholder": "i.e. Accounting User"
    }')]
    [String]$Description,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "AccountPassword",
        "Description": "AccountPassword",
        "Placeholder": "i.e. password",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value. i.e. password"
        }
    }')]
    [String]$AccountPassword,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "PasswordNeverExpires",
        "Description": "PasswordNeverExpires",
        "Placeholder": "i.e. Yes/No",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please provide a value. i.e. Yes/No"
        }
    }')]
    [Boolean]$PasswordNeverExpires,
    
    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "OfficePhone",
        "Description": "OfficePhone",
        "Placeholder": "i.e. +40726457812"
    }')]
    [String]$OfficePhone,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "TerminalServicesProfilePath",
        "Description": "TerminalServicesProfilePath",
        "Placeholder": "i.e. \\srv.domain.local\username"
    }')]
    [String]$TerminalServicesProfilePath,

    [Parameter(Mandatory = $false, HelpMessage = '{
        "Title": "TerminalServicesProfilePath",
        "Description": "TerminalServicesProfilePath",
        "Placeholder": "i.e. KeyValue pairs"
    }')]
    [HashTable]$CustomAttributes

)

#Logging configuration
$Script:VerbosePreference = 'SilentlyContinue'
$Script:CustomVerbosePreference = $true

#Functions declaration
Function Set-ADUserRDSProfilePath {
    [CmdletBinding()]
    Param (
        [parameter(ValueFromPipeline)]
        [Microsoft.ActiveDirectory.Management.ADUser]$Identity,
        [string]$RDSProfilePath
    )
    Process {
        $ADUser = Get-ADUser $Identity -Properties *  | Select-Object -ExpandProperty disting*
        $ADUser = [ADSI]"LDAP://$ADUser"
        $ADUser.psbase.Invokeset("terminalservicesprofilepath", $RDSProfilePath)
        $ADUser.setinfo()
    }
}

#Check for property availability
$UniqueProperties = @('DisplayName', 'UserPrincipalName', 'SamAccountName')
foreach ($key in $UniqueProperties) {
    $findExistingUser = Get-ADUser -LDAPFilter "($key=$($PSBoundParameters.$key))"
    if ($findExistingUser) {
        [Boolean]$stopScript = $true
        Write-Warning -Message "Property '$key' with value '$($PSBoundParameters.$key)' is already assigned to account '$($findExistingUser.DistinguishedName)'"
    }
}
if ($stopScript) {
    Throw [Microsoft.ActiveDirectory.Management.ADIdentityAlreadyExistsException] "Some of the properties are already present in the directory. Please check the logs."
}

#Create Active Directory user object
[hashtable]$newUserProperties = $PSBoundParameters
$newUserProperties.Remove('TerminalServicesProfilePath') | Out-Null
$newUserProperties.Remove('CustomAttributes') | Out-Null
$newUserProperties.Remove('AccountPassword') | Out-Null
try {
    $newUser = New-ADUser @newUserProperties -AccountPassword ($AccountPassword | ConvertTo-SecureString -AsPlainText -Force) -PassThru -ErrorAction Stop
}
catch {
    Throw "Unable to create ActiveDirectory User using the supplied parameters. $_"
}

#Set Remote Desktop Services profile path if specified in bound parameters
if ($TerminalServicesProfilePath) {
    try {
        Set-ADUserRDSProfilePath -Identity $SamAccountName -RDSProfilePath $TerminalServicesProfilePath -ErrorAction Stop
    }
    catch {
        Throw "Unable to set RDS Profile Path for user '$SamAccountName'. $_"
    }
}

#Set custom attributes if any are passed in bound parameters
foreach ($key in $CustomAttributes.Keys) {
    try {
        Set-ADUser -Identity $SamAccountName -replace @{$key = "$($CustomAttributes.$key)"}
    }
    catch {
        Throw "Unable to set attribute '$key' with value '$($CustomAttributes.$key)' for user '$SamAccountName'. $_"
    }
}

Return $newUser

#Active Directory User Properties
# GivenName
# Surname
# DisplayName
# SamAccountName
# UserPrincipalName
# Title
# AccountPassword
# ChangePasswordAtLogon
# CannotChangePassword
# PasswordNeverExpires
# City
# Company
# EmployeeID
# EmployeeNumber
# Enabled
# HomeDirectory
# HomeDrive
# Instance
# Manager
# MobilePhone
# OfficePhone
# Office
# StreetAddress
# Organization
# OtherAttributes