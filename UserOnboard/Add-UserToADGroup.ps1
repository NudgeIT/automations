param (
    
    [Parameter(Mandatory = $false, HelpMessage = '{
    "Title": "SamAccountName",
    "Description": "SamAccountName",
    "Placeholder": "i.e. jsmith",
    "ErrorMessage": {
        "Mandatory": "This field is mandatory, please provide a value. i.e. jsmith"
    }
    }')]
    [String]$SamAccountName,

    [Parameter(Mandatory = $false, HelpMessage = '{
    "Title": "GroupName",
    "Description": "Group Name",
    "Placeholder": "i.e. Sales",
    "ErrorMessage": {
        "Mandatory": "This field is mandatory, please provide a value. i.e. Sales"
    }
    }')]
    [String]$GroupName

)

#Logging configuration
$Script:VerbosePreference = 'SilentlyContinue'
$Script:CustomVerbosePreference = $true

$adUser = Get-ADUser -Identity $SamAccountName -ErrorAction Stop
$group = Get-ADGroup -Identity $GroupName -ErrorAction Stop

Write-Verbose "Adding user '$SamAccountName' to group '$GroupName'" -Verbose:$CustomVerbosePreference
Add-ADGroupMember -Identity $group -Members $adUser -ErrorAction Stop