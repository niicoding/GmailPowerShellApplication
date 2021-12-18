Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$MyCredentials
$global:filePaths

$LoginForm = New-Object system.Windows.Forms.Form
    $LoginForm.ClientSize = '400,200'
    $LoginForm.text = "PowerShell Gmail GUI"
    $LoginForm.BackColor = "#ffffff"

$EmailForm = New-Object System.Windows.Forms.Form
    $EmailForm.ClientSize = '500,500'
    $EmailForm.text = "PowerShell Gmail GUI"
    $EmailForm.BackColor = "#ffffff"

$FromLabel = New-Object system.Windows.Forms.Label
    $FromLabel.text = "From:"
    $FromLabel.AutoSize = $true
    $FromLabel.width = 25
    $FromLabel.height = 20
    $FromLabel.location = New-Object System.Drawing.Point(20, 60)
    $FromLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$From = New-Object system.Windows.Forms.TextBox
    $From.multiline = $false
    $From.width = 250
    $From.height = 20
    $From.location = New-Object System.Drawing.Point(120, 60)
    $From.Font = 'Microsoft Sans Serif,10'


$ToLabel = New-Object system.Windows.Forms.Label
    $ToLabel.text = "To:"
    $ToLabel.AutoSize = $true
    $ToLabel.width = 25
    $ToLabel.height = 20
    $ToLabel.location = New-Object System.Drawing.Point(20, 100)
    $ToLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$To = New-Object system.Windows.Forms.TextBox
    $To.multiline = $false
    $To.width = 250
    $To.height = 20
    $To.location = New-Object System.Drawing.Point(120, 100)
    $To.Font = 'Microsoft Sans Serif,10'

$SubjectLabel = New-Object system.Windows.Forms.Label
    $SubjectLabel.text = "Subject:"
    $SubjectLabel.AutoSize = $true
    $SubjectLabel.width = 25
    $SubjectLabel.height = 20
    $SubjectLabel.location = New-Object System.Drawing.Point(20, 140)
    $SubjectLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$Subject = New-Object system.Windows.Forms.TextBox
    $Subject.multiline = $false
    $Subject.width = 250
    $Subject.height = 20
    $Subject.location = New-Object System.Drawing.Point(120, 140)
    $Subject.Font = 'Microsoft Sans Serif,10'

$MessageLabel = New-Object system.Windows.Forms.Label
    $MessageLabel.text = "Message:"
    $MessageLabel.AutoSize = $true
    $MessageLabel.width = 25
    $MessageLabel.height = 20
    $MessageLabel.location = New-Object System.Drawing.Point(20, 180)
    $MessageLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$Message = New-Object System.Windows.Forms.TextBox
    $Message.Multiline = $true
    $Message.width = 250
    $Message.height = 200
    $Message.location = New-Object System.Drawing.Point(120, 180)
    $Message.Font = 'Microsoft Sans Serif,10'

$EmailFormCancelButton = New-Object System.Windows.Forms.Button
    $EmailFormCancelButton.Location = New-Object System.Drawing.Point(20, 450)
    $EmailFormCancelButton.Text = 'Cancel'
    $EmailFormCancelButton.width = 90
    $EmailFormCancelButton.height = 30
    $EmailFormCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $EmailFormCancelButton.Font = 'Microsoft Sans Serif, 10'
    $EmailForm.CancelButton = $EmailFormCancelButton

$SendEmailButton = New-Object system.Windows.Forms.Button
    $SendEmailButton.location = New-Object System.Drawing.Point(400, 450)
    $SendEmailButton.text = 'Send Email'
    $SendEmailButton.width = 90
    $SendEmailButton.height = 30
    $SendEmailButton.location = New-Object System.Drawing.Point(400, 450)
    $SendEmailButton.Font = 'Microsoft Sans Serif,10'
    $SendEmailButton.Add_Click({ SendEmail })

$EmailFormTitle = New-Object system.Windows.Forms.Label
    $EmailFormTitle.text = "PowerShell Gmail"
    $EmailFormTitle.AutoSize = $true
    $EmailFormTitle.width = 25
    $EmailFormTitle.height = 10
    $EmailFormTitle.location = New-Object System.Drawing.Point(175, 20)
    $EmailFormTitle.Font = 'Microsoft Sans Serif,13'

$LoginFormTitle = New-Object system.Windows.Forms.Label
    $LoginFormTitle.text = "Gmail Login"
    $LoginFormTitle.AutoSize = $true
    $LoginFormTitle.width = 25
    $LoginFormTitle.height = 10
    $LoginFormTitle.location = New-Object System.Drawing.Point(160, 20)
    $LoginFormTitle.Font = 'Microsoft Sans Serif,13'

$UsernameLabel = New-Object system.Windows.Forms.Label
    $UsernameLabel.text = "Username:"
    $UsernameLabel.AutoSize = $true
    $UsernameLabel.width = 25
    $UsernameLabel.height = 20
    $UsernameLabel.location = New-Object System.Drawing.Point(20, 60)
    $UsernameLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$Username = New-Object system.Windows.Forms.TextBox
    $Username.multiline = $false
    $Username.width = 250
    $Username.height = 20
    $Username.location = New-Object System.Drawing.Point(120, 60)
    $Username.Font = 'Microsoft Sans Serif,10'

$PasswordLabel = New-Object system.Windows.Forms.Label
    $PasswordLabel.text = "Password:"
    $PasswordLabel.AutoSize = $true
    $PasswordLabel.width = 25
    $PasswordLabel.height = 20
    $PasswordLabel.location = New-Object System.Drawing.Point(20, 100)
    $PasswordLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$Password = New-Object system.Windows.Forms.TextBox
    $Password.multiline = $false
    $Password.UseSystemPasswordChar = $true
    $Password.width = 250
    $Password.height = 20
    $Password.location = New-Object System.Drawing.Point(120, 100)
    $Password.Font = 'Microsoft Sans Serif, 10'

$LoginFormOKButton = New-Object System.Windows.Forms.Button
    $LoginFormOKButton.Location = New-Object System.Drawing.Point(300, 150)
    $LoginFormOKButton.Text = 'Login'
    $LoginFormOKButton.Add_Click({ VerifyData })
    $LoginFormOKButton.Font = 'Microsoft Sans Serif, 10'
    $LoginForm.AcceptButton = $LoginFormOKButton

$LoginFormCancelButton = New-Object System.Windows.Forms.Button
    $LoginFormCancelButton.Location = New-Object System.Drawing.Point(215, 150)
    $LoginFormCancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $LoginFormCancelButton.Text = 'Cancel'
    $LoginFormCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $LoginFormCancelButton.Font = 'Microsoft Sans Serif, 10'
    $LoginForm.CancelButton = $LoginFormCancelButton



$buttonFileBrowser = New-Object System.Windows.Forms.Button
    $buttonFileBrowser.text = 'Browse'
    $buttonFileBrowser.width = 90
    $buttonFileBrowser.height = 30
    $buttonFileBrowser.location = New-Object System.Drawing.Point(300, 450)
    $buttonFileBrowser.Font = 'Microsoft Sans Serif,10'
    $buttonFileBrowser.Add_Click({ BrowseFiles })

function SendEmail {
    Send-MailMessage -SmtpServer smtp.gmail.com -Port 587 -UseSsl -DeliveryNotificationOption OnSuccess -From $From.Text -To $To.Text -Subject $Subject.Text -Body $Message.Text -Attachments $global:filePaths -Credential $MyCredentials
    Write-Host -Object "Email Sent."
}
function BrowseFiles {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $OpenDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenDialog.Multiselect = $true
    $OpenDialog.initialDirectory = "C:\"
    $OpenDialog.ShowDialog() | Out-Null
    $global:filePaths = $OpenDialog.FileNames
}
function VerifyData {
    if ([string]::IsNullOrEmpty($Username.Text) -or [string]::IsNullOrEmpty($Password.Text)) {
        $LoginForm = [System.Windows.Forms.DialogResult]::None
    }
    else {
        $LoginForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
}

$LoginForm.controls.addRange(@($LoginFormTitle, $UsernameLabel, $Username, $PasswordLabel, $Password, 
        $LoginFormOKButton, $LoginFormCancelButton))

$EmailForm.controls.addRange(@($EmailFormTitle, $FromLabel, $From, $ToLabel, $To, $SubjectLabel, 
        $Subject, $MessageLabel, $Message, $EmailFormCancelButton, $SendEmailButton, $buttonFileBrowser))

$result = $LoginForm.ShowDialog()

if (($result -eq [System.Windows.Forms.DialogResult]::OK)) {
    $From.Text = $Username.Text
    $From.Select(0, 0)
    $MyCredentials = New-Object System.Management.Automation.PSCredential ($Username.Text, (ConvertTo-SecureString -String $Password.text -AsPlainText -Force))
    $EmailForm.showDialog()
} 
