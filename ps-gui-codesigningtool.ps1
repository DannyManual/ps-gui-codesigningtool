# This script creates a GUI for generating and managing code signing certificates using PowerShell.
# It allows users to create self-signed certificates, store them in a specified directory, and sign PowerShell scripts with those certificates.
# Author: Daniel Butz
# Date: 2025-09-05
# Version: 1.0
# This script is intended for educational purposes and should be used with caution in production environments.
# This script is provided "as-is" without any warranties or guarantees.
# This application does not collect or transmit any personal data.

# Füge die erforderlichen Assemblys hinzu
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Sicherstellen, dass die Datei mit UTF-8-Encoding gespeichert wird
[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Class with Fullpath and Shortpath properties
class PathEntity {
    [string]$Fullpath
    [string]$Shortpath

    PathEntity([string]$full, [string]$short) {
        $this.Fullpath = $full
        $this.Shortpath = $short
    }   
}

# BindingList
$blstCurrentCertList = New-Object System.ComponentModel.BindingList[PathEntity]

$noteText = "Dieses Tool ermöglich auf einfache Weise selbst erstellte Skripte für PowerShell zu signieren.
Ist diese Signatur auf einem Computer als vertrauenswürdig eingestuft, kann das Skript ohne weitere Bestätigungen ausgeführt werden.


Für alle Vorgänge muss zunächst ein Arbeitsverzeichnis ausgewählt werden. 
Der Punkt rechts zeigt an, ob das Verzeichnis beschreibbar ist oder nicht.
Auf der linken Seite lässt sich ein CodeSigning-Zertifikat erzeugen. 
Auf der rechten Seite kann ein Skript mit dem gewählten Zertifikat signiert werden."


# Erstelle das Hauptformular
$frmStart = New-Object System.Windows.Forms.Form
$frmStart.Text = "CodeSigning Tool"
$frmStart.MinimumSize = New-Object System.Drawing.Size(800, 640)
$frmStart.StartPosition = "CenterScreen"
$frmStart.Add_Shown({ loadFormData })
$frmStart.Add_FormClosing({
    saveFormData
    $frmStart.Dispose()
})

$frmStart.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($MyInvocation.MyCommand.Path)


# Erstelle ein Panel für die Hauptaufteilung
$tlpMain = New-Object System.Windows.Forms.TableLayoutPanel
$tlpMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpMain.RowCount = 3
$tlpMain.ColumnCount = 2
$tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpMain.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tlpMain.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$tlpMain.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$frmStart.Controls.Add($tlpMain)

# Datei für die Speicherung des Arbeitsverzeichnisses
$configFilePath = "config.txt"

$gbTitleNotes = New-Object System.Windows.Forms.GroupBox
$gbTitleNotes.Text = "Info"
$gbTitleNotes.Dock = "Fill"
$gbTitleNotes.AutoSize = $true
$gbTitleNotes.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$tlpMain.Controls.Add($gbTitleNotes, 0, 0)
$tlpMain.SetColumnSpan($gbTitleNotes, 2)

$lblTitleNotes = New-Object System.Windows.Forms.Label
$lblTitleNotes.AutoSize = $true
$lblTitleNotes.Dock = [System.Windows.Forms.DockStyle]::Fill
$gbTitleNotes.Controls.Add($lblTitleNotes)
$lblTitleNotes.Text = $noteText

# Bereich oben: Arbeitsverzeichnis auswählen
$gbWorkspace = New-Object System.Windows.Forms.GroupBox
$gbWorkspace.Text = "Working directory"
$gbWorkspace.Dock = "Fill"
$gbWorkspace.AutoSize = $true
$gbWorkspace.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$tlpMain.Controls.Add($gbWorkspace, 0, 1)
$tlpMain.SetColumnSpan($gbWorkspace, 2)

$tlpWorkspace = New-Object System.Windows.Forms.TableLayoutPanel
$tlpWorkspace.Dock = "Fill"
$tlpWorkspace.AutoSize = $true
$tlpWorkspace.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$tlpWorkspace.RowCount = 1
$tlpWorkspace.ColumnCount = 4
$tlpWorkspace.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpWorkspace.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tlpWorkspace.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpWorkspace.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpWorkspace.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$gbWorkspace.Controls.Add($tlpWorkspace)

$lblWorkspace = New-Object System.Windows.Forms.Label
$lblWorkspace.Text = "Working directory:"
$lblWorkspace.AutoSize = $true
$lblWorkspace.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblWorkspace.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$tlpWorkspace.Controls.Add($lblWorkspace, 0, 0)

$txtWorkspace = New-Object System.Windows.Forms.TextBox
$txtWorkspace.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtWorkspace.Add_TextChanged({ refresh })
$tlpWorkspace.Controls.Add($txtWorkspace, 1, 0)

$btnWorkspaceSelect = New-Object System.Windows.Forms.Button
$btnWorkspaceSelect.Text = "Select..."
$btnWorkspaceSelect.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnWorkspaceSelect.AutoSize = $true
$btnWorkspaceSelect.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$btnWorkspaceSelect.Add_Click({ selectWorkspaceFolder })
$tlpWorkspace.Controls.Add($btnWorkspaceSelect, 2, 0)

$rdbWorkspaceCheck = New-Object System.Windows.Forms.RadioButton
$rdbWorkspaceCheck.Text = ""
$rdbWorkspaceCheck.Dock = [System.Windows.Forms.DockStyle]::Fill
$rdbWorkspaceCheck.Checked = $true
$rdbWorkspaceCheck.AutoSize = $true
$rdbWorkspaceCheck.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$tlpWorkspace.Controls.Add($rdbWorkspaceCheck, 3, 0)

# Bereich links: Zertifikatserzeugung
$certGroupBox = New-Object System.Windows.Forms.GroupBox
$certGroupBox.Text = "Certificate creation"
$certGroupBox.Dock = "Fill"
$tlpMain.Controls.Add($certGroupBox, 0, 2)

$tlpCertDetail = New-Object System.Windows.Forms.TableLayoutPanel
$tlpCertDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.RowCount = 10
$tlpCertDetail.ColumnCount = 2
$tlpCertDetail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpCertDetail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpCertDetail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpCertDetail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpCertDetail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpCertDetail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpCertDetail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpCertDetail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpCertDetail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpCertDetail.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tlpCertDetail.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpCertDetail.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$certGroupBox.Controls.Add($tlpCertDetail)

$lblName = New-Object System.Windows.Forms.Label
$lblName.Text = "Name:"
$lblName.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblName.AutoSize = $true
$lblName.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($lblName, 0, 0)

$txtName = New-Object System.Windows.Forms.TextBox
$txtName.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($txtName, 1, 0)

$lblOrg = New-Object System.Windows.Forms.Label
$lblOrg.Text = "Organisation:"
$lblOrg.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblOrg.AutoSize = $true
$lblOrg.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($lblOrg, 0, 1)

$txtOrg = New-Object System.Windows.Forms.TextBox
$txtOrg.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($txtOrg, 1, 1)

$lblPassword = New-Object System.Windows.Forms.Label
$lblPassword.Text = "Password:"
$lblPassword.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblPassword.AutoSize = $true
$lblPassword.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($lblPassword, 0, 2)

$txtPassword = New-Object System.Windows.Forms.TextBox
$txtPassword.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtPassword.UseSystemPasswordChar = $true
$tlpCertDetail.Controls.Add($txtPassword, 1, 2)

$lblCN = New-Object System.Windows.Forms.Label
$lblCN.Text = "Common Name:"
$lblCN.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblCN.AutoSize = $true
$lblCN.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($lblCN, 0, 3)

$txtCN = New-Object System.Windows.Forms.TextBox
$txtCN.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($txtCN, 1, 3)

$lblEmail = New-Object System.Windows.Forms.Label
$lblEmail.Text = "E-Mail:"
$lblEmail.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblEmail.AutoSize = $true
$lblEmail.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($lblEmail, 0, 4)

$txtEmail = New-Object System.Windows.Forms.TextBox
$txtEmail.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($txtEmail, 1, 4)

$lblCountry = New-Object System.Windows.Forms.Label
$lblCountry.Text = "Country:"
$lblCountry.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblCountry.AutoSize = $true
$lblCountry.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($lblCountry, 0, 5)

$txtCountry = New-Object System.Windows.Forms.TextBox
$txtCountry.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($txtCountry, 1, 5)

$lblState = New-Object System.Windows.Forms.Label
$lblState.Text = "State:"
$lblState.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblState.AutoSize = $true
$lblState.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($lblState, 0, 6)

$txtState = New-Object System.Windows.Forms.TextBox
$txtState.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($txtState, 1, 6)

$lblCity = New-Object System.Windows.Forms.Label
$lblCity.Text = "City:"
$lblCity.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblCity.AutoSize = $true
$lblCity.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($lblCity, 0, 7)

$txtCity = New-Object System.Windows.Forms.TextBox
$txtCity.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpCertDetail.Controls.Add($txtCity, 1, 7)

$btnCreateCert = New-Object System.Windows.Forms.Button
$btnCreateCert.Text = "Create certificate"
$btnCreateCert.Anchor = [System.Windows.Forms.AnchorStyles]::Top
$btnCreateCert.AutoSize = $true
$btnCreateCert.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$btnCreateCert.Add_Click({ createCert })
$tlpCertDetail.Controls.Add($btnCreateCert, 0, 8)
$tlpCertDetail.SetColumnSpan($btnCreateCert, 2)

# Bereich rechts: Zertifikatsliste und Signierung
$gpSigningTool = New-Object System.Windows.Forms.GroupBox
$gpSigningTool.Text = "Certificates and Signing"
$gpSigningTool.Dock = "Fill"
$tlpMain.Controls.Add($gpSigningTool, 1, 2)

$tlpSigningTool = New-Object System.Windows.Forms.TableLayoutPanel
$tlpSigningTool.Dock = "Fill"
$tlpSigningTool.AutoSize = $true
$tlpSigningTool.RowCount = 10
$tlpSigningTool.ColumnCount = 1
$tlpSigningTool.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpSigningTool.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
$tlpSigningTool.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpSigningTool.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpSigningTool.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpSigningTool.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpSigningTool.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpSigningTool.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpSigningTool.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tlpSigningTool.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tlpSigningTool.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$gpSigningTool.Controls.Add($tlpSigningTool)

$lblCertsList = New-Object System.Windows.Forms.Label
$lblCertsList.Text = "Available certificates:"
$lblCertsList.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpSigningTool.Controls.Add($lblCertsList, 0, 0)

$lstCertsList = New-Object System.Windows.Forms.ListBox
$lstCertsList.Dock = [System.Windows.Forms.DockStyle]::Fill
$lstCertsList.DataSource = $blstCurrentCertList
$lstCertsList.DisplayMember = "Shortpath"
$tlpSigningTool.Controls.Add($lstCertsList, 0, 1)

$btnRefreshList = New-Object System.Windows.Forms.Button
$btnRefreshList.Text = "Refresh"
$btnRefreshList.AutoSize = $true
$btnRefreshList.Anchor = [System.Windows.Forms.AnchorStyles]::Top
$btnRefreshList.Add_Click({ refresh })
$tlpSigningTool.Controls.Add($btnRefreshList, 0, 2)

$lblFileToSign = New-Object System.Windows.Forms.Label
$lblFileToSign.Text = "PowerShell-Script File:"
$lblFileToSign.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpSigningTool.Controls.Add($lblFileToSign, 0, 3)

$txtFileToSign = New-Object System.Windows.Forms.TextBox
$txtFileToSign.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpSigningTool.Controls.Add($txtFileToSign, 0, 4)

$btnSelectFileToSign = New-Object System.Windows.Forms.Button
$btnSelectFileToSign.Text = "Select..."
$btnSelectFileToSign.AutoSize = $true
$btnSelectFileToSign.Add_Click({ selectFileToSign })
$tlpSigningTool.Controls.Add($btnSelectFileToSign, 0, 5)

$lblSignPasswd = New-Object System.Windows.Forms.Label
$lblSignPasswd.Text = "Password for certificate:"
$lblSignPasswd.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblSignPasswd.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$tlpSigningTool.Controls.Add($lblSignPasswd, 0, 6)

$txtSignPasswd = New-Object System.Windows.Forms.TextBox
$txtSignPasswd.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtSignPasswd.UseSystemPasswordChar = $true
$tlpSigningTool.Controls.Add($txtSignPasswd, 0, 7)

$btnSign = New-Object System.Windows.Forms.Button
$btnSign.Text = "Sign"
$btnSign.AutoSize = $true
$btnSign.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$btnSign.Add_Click({ signFile })
$tlpSigningTool.Controls.Add($btnSign, 0, 8)

function selectWorkspaceFolder {    
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtWorkspace.Text = $folderDialog.SelectedPath
        Set-Content -Path $configFilePath -Value $folderDialog.SelectedPath
    }       
}

function selectFileToSign {
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Choose PowerShell Script File"
    $fileDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1"
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFileToSign.Text = $fileDialog.FileName
    }
}

function signFile {
    $selectedCert = $lstCertsList.SelectedItem.Fullpath
    $scriptPath = $txtFileToSign.Text

    if (-not $selectedCert -or -not $scriptPath) {
        [System.Windows.Forms.MessageBox]::Show("Please select certificate and script file.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        $certPassword = $txtSignPasswd.Text
        $securePassword = ConvertTo-SecureString -String $certPassword -Force -AsPlainText
        $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $certificate.Import($selectedCert, $securePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet)

        # Füge einen Zeitstempel hinzu
        Set-AuthenticodeSignature -FilePath $scriptPath -Certificate $certificate -TimestampServer "http://timestamp.digicert.com"

        [System.Windows.Forms.MessageBox]::Show("Script successfully signed.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Errors occured while signing script file: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}
function createCert {
    $name = $txtName.Text
    $org = $orgTextBox.Text
    $workspace = $txtWorkspace.Text
    $password = $passwordTextBox.Text
    $cn = $cnTextBox.Text
    $email = $txtEmail.Text
    $country = $countryTextBox.Text
    $state = $stateTextBox.Text
    $city = $cityTextBox.Text

    if (-not $name -or -not $org -or -not $workspace -or -not $password -or -not $cn -or -not $email -or -not $country -or -not $state -or -not $city) {
        [System.Windows.Forms.MessageBox]::Show("Please fill all fields.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        $cert = New-SelfSignedCertificate -DnsName $name -CertStoreLocation "Cert:\CurrentUser\My" -KeyUsage DigitalSignature -Type CodeSigningCert -Subject "CN=$cn, E=$email, C=$country, S=$state, L=$city, O=$org"
        $pfxPath = Join-Path -Path $workspace -ChildPath "$name.pfx"
        $securePassword = ConvertTo-SecureString -String $password -Force -AsPlainText
        Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $securePassword
        [System.Windows.Forms.MessageBox]::Show("Certificate was created and stored under: $pfxPath", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Errors while creating the certificate: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function updateRadioWSCheck {
    $pfad = $txtWorkspace.Text
    
    if ($pfad.Length -eq 0) {
        $rdbWorkspaceCheck.ForeColor = [System.Drawing.Color]::Red
        return
    }

    $ordnerExistiert = (Test-Path $pfad -PathType Container)
    function Test-OrdnerBeschreibbar($ordner) {
        try {
            $testfile = [System.IO.Path]::Combine($ordner, [System.IO.Path]::GetRandomFileName())
            $file = [System.IO.File]::Create($testfile)
            $file.Close()
            Remove-Item $testfile -Force
            return $true
        }
        catch {
            return $false
        }
    }

    if ($ordnerExistiert -and (Test-OrdnerBeschreibbar $pfad)) {
        $rdbWorkspaceCheck.ForeColor = [System.Drawing.Color]::Green
    }
    else {
        $rdbWorkspaceCheck.ForeColor = [System.Drawing.Color]::Red
    } 
}

function updateCertList {
    $blstCurrentCertList.Clear()
    if ($txtWorkspace.Text.Length -gt 0) {
        if (Test-Path $txtWorkspace.Text) {
            Get-ChildItem -Path $txtWorkspace.Text -Filter "*.pfx" | ForEach-Object {
                $blstCurrentCertList.Add((New-Object PathEntity -ArgumentList $_.FullName, $_.Name))
            }
        }
    }
}

function refresh {
    updateRadioWSCheck
    updateCertList
}

function loadFormData {
    if (Test-Path $configFilePath) {
        $savedWorkspace = Get-Content $configFilePath
        if (-not [string]::IsNullOrWhiteSpace($savedWorkspace)) {
            $txtWorkspace.Text = $savedWorkspace
        }
    }  
}

function saveFormData {
    $workspacePath = $txtWorkspace.Text
    if (-not [string]::IsNullOrWhiteSpace($workspacePath)) {
        Set-Content -Path $configFilePath -Value $workspacePath
    }
}

# Show Main Form
[void]$frmStart.ShowDialog()
