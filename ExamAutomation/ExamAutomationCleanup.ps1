#Script configuration
[string]$AutomationFolder = '' #path that will be used by the script - can be network path as well

#Go through each active exam
foreach ($file in (Get-ChildItem -Path "$AutomationFolder\ExamsPendingDeprovision")) {
    if ($file.Extension.ToLower() -ne '.json') {
        Continue
    }
    
    #Get active exam details
    $examDetails = ConvertFrom-Json -InputObject (Get-Content -Path $file.FullName -Raw)
    
    #Check for deprovision date
    if ((Get-Date) -lt $examDetails.Date) {
        $timeSpan = New-TimeSpan -Start (Get-Date) -End $examDetails.Date
        Write-Verbose -Message "Exam '$($file.Name.ToUpper().Replace('.JSON',''))' will be deprovisioned in $($timeSpan.Days) days, $($timeSpan.Hours) hours, $($timeSpan.Minutes) minutes." -Verbose
        Continue
    }

    #Delete user accounts
    foreach ($user in $examDetails.Users) {
        Write-Verbose -Message "Removing user $($User.Username)." -Verbose
        Remove-ADUser -Identity $user.Username `
        -Confirm:$false 
    }

    #Remove exam deprovision file
    Remove-Item -Path $file.FullName -Force

}
