param (
    [PSCredential]$ExchangeCredential,
    [String]$ExchangeServerUri,
    [array]$ImportedCommands
)

try {
    if ($exchangeServerUri -notlike '*outlook.office365.com/powershell-liveid*') {
        Write-Verbose -message "Create exchange on-premise remote session." -Verbose:$CustomVerbosePreference
        $sessionOptions = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeServerUri -Credential $ExchangeCredential -Authentication Basic -AllowRedirection -SessionOption $sessionOptions -WarningAction 'SilentlyContinue' -ErrorAction 'Stop'
    }
    else {
        Write-Verbose -message "Create exchange online remote session." -Verbose:$CustomVerbosePreference
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeServerUri -Credential $ExchangeCredential -Authentication Basic -AllowRedirection -WarningAction 'SilentlyContinue' -ErrorAction 'Stop'
    }
    Write-Verbose -message "Import exchange commands." -Verbose:$CustomVerbosePreference
    Import-PSSession -Session $exchangeSession -CommandName $ImportedCommands -AllowClobber | Out-Null
}
catch {
    Remove-PSSession -Session $exchangeSession -ErrorAction SilentlyContinue
    Throw "Unable to create exchange session. $_"
}

Return $exchangeSession