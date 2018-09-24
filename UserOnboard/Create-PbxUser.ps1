param (
    [String]$WsldPath,
    [String]$ApiUrl,
    [PSCredential]$ApiCredential,
    [string]$ApiUserName,
    [String]$DisplayName, 
    [String]$Extension, 
    [String]$Email, 
    [String]$SamAccountName 
)

#Logging configuration
$Script:VerbosePreference = 'SilentlyContinue'
$Script:CustomVerbosePreference = $true

if (-not (Test-Path $WsldPath)) {
    Throw "File not found. Please check path '$WsldPath'"
}
$webServiceProxy = New-WebServiceProxy -Uri $WsldPath
$webServiceProxy.Url = "$ApiUrl/PBX0/user.soap"
$ApiUser = $ApiCredential.UserName
$ApiPassword = $ApiCredential.GetNetworkCredential().Password
$webServiceProxy.Credentials = new-object System.Net.NetworkCredential("$ApiUser", "$ApiPassword", "")

[int]$key = $null
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
try {
    $result = $webServiceProxy.Initialize($ApiUserName, 'PowerShell', $true, $true, $true, $true, $true, ([ref]$key))
}
catch {
    Throw "Unable to initialize Web Service Proxy. $_"
}
if ($result = 0) {
    Throw "Unable to authenticate to Innovaphone."
}

$newUserProperties = @"
<add>
    <user 
        cn="$DisplayName" 
        dn="$DisplayName"
        text="$DisplayName"
        e164="$Extension" 
        h323="$SamAccountName"
        email="$Email"
        cfnr="" 
        busy-out="" 
        loc="" 
        node="root" 
        filter="normal" 
        cd-filter="normal" 
        type="">
        <cd type="">
            <ep e164=""/>
        </cd>
        <cd type="">
            <ep e164=""/>
        </cd>
        <grp name=""/>
        <grp name=""/>
    </user>
</add>
"@

try {
    $result = $webServiceProxy.Admin($newUserProperties)
}
catch {
    Throw "Error creating Innovaphone user. $_"
}
Write-Verbose "Create Innovaphone user result: $result" -Verbose:$CustomVerbosePreference
 


