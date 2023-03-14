#Exchange Certificate Renewal Tool
#drdweb 2022

<# Environment and GUI functions #>
#Load Exchange Powershell SnapIn
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

###Windows forms to select the name and filepath of the request through a GUI
function Save-Req([string] $initialDirectory ) {

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "REQ Files (*.req)| *.req"
    $OpenFileDialog.ShowDialog() |  Out-Null

    return $OpenFileDialog.filename
}
function Save-Cert([string] $initialDirectory ) {

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "PFx Files (*.pfx)| *.pfx"
    $OpenFileDialog.ShowDialog() |  Out-Null

    return $OpenFileDialog.filename
}

#File selection menu for new .crt
Function Find-FilePath($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Certificate (*.crt)| *.crt"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

function Get-FileNameInput {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Friendly Name'
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
    $label.Text = 'Please enter the friendly name for the new certificate:'
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

function Get-StringInput {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Security Key Entry'
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
    $label.Text = 'Enter the security key for the .pfx'
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

function New-WPFControl { 
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $True)]
        [string]$WindowTitle,
        [Parameter(Mandatory = $True)]
        [string]$ButtonText,
        [string]$Buttonaction
    )

    #Import Assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
     
    $Form1 = New-Object System.Windows.Forms.Form 
    $ActionButton = New-Object System.Windows.Forms.Button 
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState 
     
    # Check for ENTER and ESC presses
    $Form1.KeyPreview = $True
    $Form1.Add_KeyDown({ if ($_.KeyCode -eq "Enter") {
                # if enter, perform click
                $ActionButton.PerformClick()
            }
        })
    $Form1.Add_KeyDown({ if ($_.KeyCode -eq "Escape") {
                # if escape, exit
                $Form1.Close()
            }
        })
     
    # The action on the button
    $handler_Button_Click = 
    {
        <# Action to perform on Ok button click#>
        Invoke-Expression $Buttonaction
        $Form1.Close()
    }
     
    $OnLoadForm_StateCorrection = 
    {
        $Form1.WindowState = $InitialFormWindowState 
    }   
     
    # Form Code 
    $Form1.Name = "Data_Form"
    $Form1.Text = $WindowTitle
    $Form1.MaximizeBox = $false #lock form
    $Form1.FormBorderStyle = 'FixedDialog'
    # None,FixedDialog,FixedSingle,FixedToolWindow,Sizable,SizableToolWindow
     
    # Icon
    $Form1.Icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)
    # $NotifyIcon.Icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)
     
    $Form1.DataBindings.DefaultDataSourceUpdateMode = 0 
    $Form1.StartPosition = "CenterScreen"# moves form to center of screen
    $System_Drawing_Size = New-Object System.Drawing.Size 
    $System_Drawing_Size.Width = 700 # sets X
    $System_Drawing_Size.Height = 180 # sets Y
    $Form1.ClientSize = $System_Drawing_Size
     
    $ActionButton.Name = "OK_Button" 
    $System_Drawing_Size = New-Object System.Drawing.Size 
    $System_Drawing_Size.Width = 200
    $System_Drawing_Size.Height = 55
     
    $ActionButton.Size = $System_Drawing_Size 
    $ActionButton.UseVisualStyleBackColor = $True
    $ActionButton.Text = $ButtonText
    $System_Drawing_Point = New-Object System.Drawing.Point 
    $System_Drawing_Point.X = 250 
    $System_Drawing_Point.Y = 70
     
    $ActionButton.Location = $System_Drawing_Point 
    $ActionButton.DataBindings.DefaultDataSourceUpdateMode = 0 
    $ActionButton.add_Click($handler_Button_Click)
    $Form1.Controls.Add($ActionButton)
     
    $InitialFormWindowState = $Form1.WindowState 
    $Form1.add_Load($OnLoadForm_StateCorrection) 
      
    # Show Form 
    $Form1.ShowDialog()
}


# Import Section #
function Import-Cert {
    $newcrt = Find-FilePath

    $friendname = Get-FileNameInput
    
    Import-ExchangeCertificate -FileData ([System.IO.File]::ReadAllBytes($newcrt)) -FriendlyName $friendname  -PrivateKeyExportable $true
    
    $newthmbprntsel = (Get-ExchangeCertificate | Select-Object -Property * | Out-GridView -Title "Available Exchange Certificates" -PassThru).Thumbprint
    
    Enable-exchangecertificate -thumbprint $newthmbprntsel  -services iis, smtp, pop, imap    

}

# Export Section #
function Export-CertwPK {
    #Get's user input for .pfx password and converts it to a secure string
    $pwdinput = Get-StringInput

    $pwdsecure = ConvertTo-SecureString -String $pwdinput -Force -AsPlainText

    #Select your path and name for the cert export
    $filepath = Save-Cert

    #A verifacation that we are exporting the correct cert using the thumbprint from the import stage
    $certsel = Get-ChildItem -Path cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $newthmbprntsel }

    #exporting the cert to the chosen path using our password
    Export-PfxCertificate -Cert $certsel -FilePath $filepath -Password $pwdsecure 

}

#Nested functions for finishing the cert install
Function Start-CertRenewal {
    Import-Cert

    Export-CertwPK

}


<# New request generation #>
#Sets the certificate thumbprint selection with GUI
$thmbprntsel = (Get-ExchangeCertificate | Where-Object { $_.Services -like "IMAP, POP, IIS, SMTP" } | 
    Select-Object -Property * | Out-GridView -Title "IMAP, POP, IIS, & SMTP Certificates" -PassThru).Thumbprint

$txtrequest = Get-ExchangeCertificate -Thumbprint $thmbprntsel | New-ExchangeCertificate -GenerateRequest -KeySize 2048 -PrivateKeyExportable $true

#Sets the file and folder path to the .req
$reqfilepth = Save-Req

#executes the request generation
[System.IO.File]::WriteAllBytes($reqfilepth, [System.Text.Encoding]::Unicode.GetBytes($txtrequest))

#Pause script to retrieve new cert.
New-WPFControl -WindowTitle "Click Ok to continue" -ButtonText "Ok" -Buttonaction "Start-CertRenewal"








