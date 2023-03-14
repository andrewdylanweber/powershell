#Install-RDGCert
#Written by drdweb Oct 2022
#Only designed for RDGs running all 4 RD Roles

Function Get-FilePath {
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter           = 'Personal Information Exchange (*.pfx)|*.pfx'
    }
    $FileBrowser.ShowDialog() | Out-Null
    $FileBrowser.FileName
}

function Get-StringInput {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Encryption Password'
    $form.Size = New-Object System.Drawing.Size(300, 200)
    $form.StartPosition = 'CenterScreen'
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75, 120)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150, 120)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(280, 20)
    $label.Text = 'What is the encryption password?'
    $form.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 40)
    $textBox.Size = New-Object System.Drawing.Size(260, 20)
    $form.Controls.Add($textBox)
    
    $form.Topmost = $true
    $form.Add_Shown({ $textBox.Select() })
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $x = $textBox.Text
        $x
    }        
}

Function Install-RDGCert() {
    #User select .pfx to import
    $Path = Get-FilePath
    #User define pfx decrypt password
    $Password = Get-StringInput
    #Convert plaintext to a usable securestring
    $Pass = ConvertTo-SecureString $Password -AsPlainText -Force
    #Set the $ConnectionBroker to the current machine's FQDN
    $ConnectionBroker = ([System.Net.Dns]::GetHostByName($env:computerName)).HostName

    Set-RDCertificate -Role RDRedirector -Password $Pass -ConnectionBroker $ConnectionBroker -ImportPath $Path -Force
    Set-RDCertificate -Role RDGateway -Password $Pass -ConnectionBroker $ConnectionBroker -ImportPath $Path -Force
    Set-RDCertificate -Role RDWebAccess -Password $Pass -ConnectionBroker $ConnectionBroker -ImportPath $Path -Force
    Set-RDCertificate -Role RDPublishing -Password $Pass -ConnectionBroker $ConnectionBroker -ImportPath $Path -Force
}

Install-RDGCert

