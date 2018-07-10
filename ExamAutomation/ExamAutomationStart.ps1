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
            "Pattern": "Exam name does not match the correct format. i.e. i.e. GY GYM 1604 A 07."
        }
    }')]
    [ValidatePattern('[A-Za-z]{4}\ [0-9]{2}')]
    [string]$ExamName,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "Number of students",
        "Description": "Number of students attending the exam",
        "Placeholder": "i.e. 25",
        "ErrorMessage": {
            "Min": "Please enter a number between 1 and 90.",
            "Max": "Please enter a number between 1 and 90."
        }
    }')]
    [ValidateRange(1, 90)]
    [int]$NumberOfStudents,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "Exam Date",
        "Description": "Exam Date",
        "Placeholder": "Please select date",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please select a Date."
        }
    }')]
    [DateTimeOffset]$ExamDate,

    [Parameter(Mandatory = $true, HelpMessage = '{
        "Title": "Internet Access",
        "Description": "Internet Access",
        "Placeholder": "i.e. Yes/No",
        "ErrorMessage": {
            "Mandatory": "This field is mandatory, please select a value."
        }
    }')]
    [bool]$InternetAccess
)

#Script configuration
$Global:ErrorActionPreference = 'Stop'
$Global:VerbosePreference = 'SilentlyContinue'
[boolean]$Global:CustomVerbosePreference = $true

[string]$ExamUsersOu = '' #fill in OU that will hold the exam users
[string]$ExamsNetworkPath = '' #path that will be used by the script - can be network path as well
[string]$AutomationFolder = "$ExamsNetworkPath\Configuration"
[string]$ConfigFilePath = "$AutomationFolder\config.json"

#Functions
function Close-WordApplication {
    param (
        [object]$Application,
        [object]$WordDocument,
        [boolean]$Save
    )
    if ($Save) {
        $WordDocument.Save()
        $WordDocument.Saved = $true
    }
    $WordDocument.Close()
    $Application.Quit()
}

#Test for config file path
if (-not (Test-Path $ConfigFilePath)) {
    Throw [System.IO.FileNotFoundException] "Config file path not found."
}

#Test if exam folder already exists
if (Test-Path -Path "$ExamsNetworkPath\$ExamName - $examNumber") {
    Throw "Folder '$ExamName' already created. Stopping automation."
}

#Get exam number from config
$config = ConvertFrom-Json -InputObject (Get-Content -Path $ConfigFilePath -Raw) 
$examNumber = ($config.ExamNumber + 1).ToString("00000")

#Create exam shared folder
$examFolder = New-Item -ItemType Directory -Path "$ExamsNetworkPath\$ExamName - $examNumber" 

#Find teacher's account
$userUpn = $Teacher.UPN
$teacherADAccount = Get-ADUser -LDAPFilter "(UserPrincipalName=$userUpn)"
if (-not $teacherADAccount) {
    Throw "Active directory account not found using filter: UserPrincipalName=$userUpn."
}

#Assign Write permissions for teacher
Add-NTFSAccess -Path $examFolder -Account $teacherADAccount.SID -AccessRights ReadAndExecute, Write, Read

#Create child subjects folder - this will inherit teacher's permissions
$subjectsFolder = New-Item -ItemType Directory -Path "$($examFolder.FullName)\Aufgabe"

#Open word application
$wordApplication = New-Object -ComObject Word.Application 
#Open word document
try {
    $document = $wordApplication.documents.open("$AutomationFolder\Exam-Template.docx")
    $selection = $wordApplication.Selection
}
catch {
    Close-WordApplication -Save $false -WordDocument $document -Application $wordApplication
    Throw "Unable to open template document. $_"
}

#Create ActiveDirectory Group
$adGroup = New-ADGroup -Name $ExamName `
-SamAccountName $ExamName `
-GroupCategory Security `
-GroupScope DomainLocal `
-DisplayName $ExamName `
-Path $ExamUsersOu `
-Description "Members of this group attend exam $ExamName"

#Update user details in onboard document
[System.Collections.ArrayList]$allUsers = @()
for ($i = 1; $i -le $NumberOfStudents; $i++) {
    $userName = 'ex-{0}-{1}' -f $examNumber, $i.ToString("00")
    $password = NiboUtils\New-NiboRandomPassword -passLength 8 -numbers -lowerCase -upperCase
    [void]$allUsers.Add(
        [PSCustomObject]@{
            Username = $userName
            Password = $password
        }
    )
    Write-Verbose  "Creating user: $userName." -Verbose:$CustomVerbosePreference
    [void]$selection.GoTo([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToPage,
        [Microsoft.Office.Interop.Word.WdGoToDirection]::wdGoToNext, 1
    )
    try {
        $newADUser = New-ADUser -DisplayName $userName `
            -Name $userName `
            -SamAccountName $userName `
            -Description $ExamName `
            -AccountPassword ($password | ConvertTo-SecureString -AsPlainText -Force) `
            -ChangePasswordAtLogon $false `
            -PasswordNeverExpires $true `
            -Enabled $true `
            -Path $ExamUsersOu `
            -PassThru
        
        
    }
    catch {
        Close-WordApplication -Save $false -WordDocument $document -Application $wordApplication
        Throw "Unable to create AD user '$userName'. $_"
    }
    #Create student folder
    Write-Verbose "Creating user folder '$examFolder\$username'." -Verbose:$CustomVerbosePreference
    $userFolder = New-Item -ItemType Directory -Path "$examFolder\$username" 
    
    #Assign permissions for the student
    Write-Verbose "Assigning permissions to folder '$examFolder\$username'." -Verbose:$CustomVerbosePreference
    Add-NTFSAccess -Path $subjectsFolder -Account $newADUser.SID -AccessRights Read, ReadAndExecute
    Add-NTFSAccess -Path $userFolder -Account $newADUser.SID -AccessRights Read, ReadAndExecute, Write
    
    #Add student details to exam document
    Write-Verbose "Add student details to exam document." -Verbose:$CustomVerbosePreference
    try {
        $selection.Font.Size = 18; $selection.Font.Bold = $true
        [void]$selection.TypeText("Username: $userName")
        [void]$selection.TypeParagraph() 
        [void]$selection.TypeText("Password: $password")
        [void]$selection.TypeParagraph()
        [void]$selection.TypeText("Student folder: ")
        $selection.Font.Size = 14; $selection.Font.Bold = $false
        [void]$selection.TypeText("$($userFolder.FullName)")
        [void]$selection.InsertNewPage()
    }
    catch {
        Close-WordApplication -Save $false -WordDocument $document -Application $wordApplication
        Throw "Error editing template document. $_"
    }
}

#Update group membership
Add-ADGroupMember -Identity $ExamName -Members $allUsers.Username

#Edit first page with exam details
Write-Verbose "Edit first page in the exam document." -Verbose:$CustomVerbosePreference
try {
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    $document.Tables.Item(1).Columns(2).Cells(1).Range.Text = ($ExamDate).ToString()
    $document.Tables.Item(1).Columns(2).Cells(2).Range.Text = $ExamName
    $document.Tables.Item(1).Columns(2).Cells(3).Range.Text = $examFolder.FullName
    $j = 1; $i = 0
    foreach ($user in $allUsers) {
        $i++
        $document.Tables.Item(2).Columns($j).Cells($i).Range.Text = $user.Username
        if ($i % ($document.Tables(2).Rows.Count) -eq 0) {
            $j++; $i = 0
        }
    }
    #Save new document
    $document.SaveAs("$AutomationFolder\$ExamName.docx")
    #Save new document as PDF
    $wdFormatPDF = 17
    $document.SaveAs("$AutomationFolder\$ExamName.pdf", [ref]$wdFormatPDF)
    Close-WordApplication -Save $false -WordDocument $document -Application $wordApplication
}
catch {
    Close-WordApplication -Save $false -WordDocument $document -Application $wordApplication
    Throw "Error editing template document. $_"
}

#Save exam details to disk
Write-Verbose "Export exam details to JSON." -Verbose:$CustomVerbosePreference
ConvertTo-Json -InputObject @{
    Group = $ExamName
    Teacher = $userUpn
    Users   = $allUsers
} | Out-File "$AutomationFolder\ActiveExamsDetails\$ExamName - $examNumber.json"

#Increment exam number in config file
Write-Verbose "Incrementing exam number in config file." -Verbose:$CustomVerbosePreference
$config.ExamNumber = $config.ExamNumber + 1
ConvertTo-Json -InputObject $config | Out-File $ConfigFilePath 

#Send email to teacher containing exam details and student list
Write-Verbose "Sending e-mail to '$userUpn'." -Verbose:$CustomVerbosePreference
[array]$to = @($userUpn)
Send-MailMessage -From '' ` #e-mail address that will appear as sender
    -to $to `
    -Subject "New Exam ready - $ExamName" `
    -Body "Please see attached the list with the user accounts." `
    -Attachments "$AutomationFolder\$ExamName.pdf" `
    -SmtpServer "" #fill in server that will be used for email relay

#Automation end
Write-Output "Exam created successfully, please check e-mail for exam document."





