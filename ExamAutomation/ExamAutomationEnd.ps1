param (
    [Parameter(Mandatory = $true, HelpMessage = '{
        "XType": {
            "Class": "User",
            "IdProperty": "Id",
            "TitleProperty": "FullName",
            "DescriptionProperty": "FullName"
        },
        "Title": "Teacher",
        "Description": "Teacher",
        "Placeholder": "i.e. John Snow",
    }')]
    [Object]$Teacher,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "Exam Name",
        "Description": "Exam Name",
        "Placeholder": "i.e. MATH 01",
        "ErrorMessage": {
            "Pattern": "Exam name does not match the correct format. i.e. MATH 01."
        }
    }')]
    [ValidatePattern('[A-Za-z]{4}\ [0-9]{2} - [0-9]{5}')]
    [string]$ExamNameAndNumber
)

#Script configuration
$Global:ErrorActionPreference = 'Stop'
$Global:VerbosePreference = 'SilentlyContinue'
[boolean]$Global:CustomVerbosePreference = $true

[string]$ExamsNetworkPath = '' #path that will be used by the script - can be network path as well
[string]$AutomationFolder = "$ExamsNetworkPath\Configuration"

#Check for active exam path
if (-not "$AutomationFolder\ActiveExamsDetails\$ExamNameAndNumber.JSON") {
    Throw "Cannot find '$AutomationFolder\ActiveExamsDetails\$ExamNameAndNumber.JSON'."
}

#Get active exam details
$examDetails = ConvertFrom-Json -InputObject (Get-Content -Path "$AutomationFolder\ActiveExamsDetails\$ExamNameAndNumber.JSON" -Raw)

#Check if teacher is authorized to close the exam
if ($examDetails.Teacher -ne $Teacher.UPN) {
    Throw "You are not the teacher assigned for this exam. Only the assigned teacher can end an exam."
}

#Disable user accounts
foreach ($user in $examDetails.Users) {
    Write-Verbose "Disabling account '$user'." -Verbose:$CustomVerbosePreference
    Disable-ADAccount -Identity $user.Username
}

#Move exam to deprovision folder
ConvertTo-Json -InputObject @{
    Date = (Get-Date).AddDays(4)
    Users = $examDetails.Users
    Group = $examDetails.Group
} | Out-File "$AutomationFolder\ExamsPendingDeprovision\$ExamNameAndNumber.json"

#Archive exam folder
$examFolder = Get-Item -Path "$ExamsNetworkPath\$ExamNameAndNumber" -ErrorAction Stop
$destination = "$ExamsNetworkPath\$ExamNameAndNumber.zip"
Add-Type -assembly "system.io.compression.filesystem"
if (Test-Path $destination) {
    Remove-Item -Path $destination -Force
}
[System.IO.Compression.ZipFile]::CreateFromDirectory($examFolder, $destination)
Copy-Item -Path $destination -Destination "$AutomationFolder\ExamsArchive"

#Send send exam end confirmation to teacher
Write-Verbose "Sending e-mail to '$($Teacher.UPN)'." -Verbose:$CustomVerbosePreference
[array]$to = @($Teacher.UPN)
Send-MailMessage -From '' ` #e-mail address that will appear as sender
    -to $to `
    -Subject "Exam ended - $ExamNameAndNumber" `
    -Body "Please see attached the archive containing exam folder." `
    -Attachments $destination `
    -SmtpServer '' #fill in server that will be used for email relay

#Cleanup exam folder
Remove-Item -Path "$AutomationFolder\ActiveExamsDetails\$ExamNameAndNumber.JSON"
Remove-Item -Path "$ExamsNetworkPath\$ExamNameAndNumber" -Recurse -Force
Remove-Item -Path $destination -Force