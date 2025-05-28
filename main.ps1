Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#Import-Module ActiveDirectory -ErrorAction Stop

function Show-ModeChoiceWindow {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Choix du mode"
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ShowIcon = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Choisissez le mode d'exécution :"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(50,20)
    $form.Controls.Add($label)

    $buttonDemo = New-Object System.Windows.Forms.Button
    $buttonDemo.Text = "Mode Démo (simulation)"
    $buttonDemo.Size = New-Object System.Drawing.Size(200,30)
    $buttonDemo.Location = New-Object System.Drawing.Point(50,50)
    $buttonDemo.Add_Click({ $form.Tag = 'Demo'; $form.Close() })
    $form.Controls.Add($buttonDemo)

    $buttonReal = New-Object System.Windows.Forms.Button
    $buttonReal.Text = "Mode Réel (exécution)"
    $buttonReal.Size = New-Object System.Drawing.Size(200,30)
    $buttonReal.Location = New-Object System.Drawing.Point(50,85)
    $buttonReal.Add_Click({ $form.Tag = 'Real'; $form.Close() })
    $form.Controls.Add($buttonReal)



    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Show-ActionChoiceWindow {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Choix de l'action"
    $form.Size = New-Object System.Drawing.Size(300,180)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ShowIcon = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Choisissez une action :"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(80,20)
    $form.Controls.Add($label)

    $buttonCreate = New-Object System.Windows.Forms.Button
    $buttonCreate.Text = "Création d'un utilisateur"
    $buttonCreate.Size = New-Object System.Drawing.Size(200,30)
    $buttonCreate.Location = New-Object System.Drawing.Point(50,50)
    $buttonCreate.Add_Click({ $form.Tag = 'Create'; $form.Close() })
    $form.Controls.Add($buttonCreate)

    $buttonCopy = New-Object System.Windows.Forms.Button
    $buttonCopy.Text = "Copie d'un utilisateur"
    $buttonCopy.Size = New-Object System.Drawing.Size(200,30)
    $buttonCopy.Location = New-Object System.Drawing.Point(50,85)
    $buttonCopy.Add_Click({ $form.Tag = 'Copy'; $form.Close() })
    $form.Controls.Add($buttonCopy)

    # Bouton Retour
    $buttonReturn = New-Object System.Windows.Forms.Button
    $buttonReturn.Text = "Retour"
    $buttonReturn.Size = New-Object System.Drawing.Size(200,30)
    $buttonReturn.Location = New-Object System.Drawing.Point(50,120)
    $buttonReturn.Add_Click({
        $form.Tag = 'Return'
        $form.Close()
    })
    $form.Controls.Add($buttonReturn)

    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Generate-StrongPassword {
    # Exemple simple, à améliorer selon besoins
    Add-Type -AssemblyName System.Web
    $length = 12
    $password = [System.Web.Security.Membership]::GeneratePassword($length,3)
    return $password
}

function Log-UserCreation {
    param(
        [string]$SamAccountName,
        [string]$DisplayName,
        [string]$Mode,       # 'Demo' ou 'Real'
        [string]$Action     # 'Create' ou 'Copy'
    )
    $desktop = [Environment]::GetFolderPath("Desktop")
    $logFile = Join-Path $desktop "ADUserCreationLog.csv"
    if (-not (Test-Path $logFile)) {
        "Timestamp;SamAccountName;DisplayName;Mode;Action" | Out-File -FilePath $logFile -Encoding UTF8
    }
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "$timestamp;$SamAccountName;$DisplayName;$Mode;$Action"
    $line | Out-File -FilePath $logFile -Append -Encoding UTF8
}

function Create-ADUserReal {
    param(
        [string]$SamAccountName,
        [string]$GivenName,
        [string]$Surname,
        [string]$DisplayName,
        [string]$UserPrincipalName,
        [string]$Password
    )
    try {
        New-ADUser -SamAccountName $SamAccountName `
                   -GivenName $GivenName `
                   -Surname $Surname `
                   -Name $DisplayName `
                   -UserPrincipalName $UserPrincipalName `
                   -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
                   -Enabled $true `
                   -ChangePasswordAtLogon $true `
                   -PasswordNeverExpires $false `
                   -Path "OU=Users,DC=domain,DC=local"  # A adapter à ton environnement
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la création AD:`n$_","Erreur", 'OK', 'Error') | Out-Null
        return $false
    }
}

function Show-MainForm {
    param(
        [string]$Mode,   # 'Demo' ou 'Real'
        [string]$Action  # 'Create' ou 'Copy'
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Gestion utilisateur Active Directory"
    $form.Size = New-Object System.Drawing.Size(600, 550)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ShowIcon = $false

    # --- Ajout du logo ---
$logoPath = "C:\Adista\adista.jpg"

if (Test-Path $logoPath) {
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    $pictureBox.Size = New-Object System.Drawing.Size(80, 50)  # taille du logo

    # Positionnement en bas à droite :
    $x = $form.ClientSize.Width - $pictureBox.Width - 20  # 20 px de marge à droite
    $y = $form.ClientSize.Height - $pictureBox.Height - 40 # 40 px de marge en bas (pour tenir compte de la bordure)

    $pictureBox.Location = New-Object System.Drawing.Point($x, $y)

    $pictureBox.Image = [System.Drawing.Image]::FromFile($logoPath)
    $form.Controls.Add($pictureBox)

    $startY = 35  # On peut garder les champs en haut comme avant, car logo en bas
}
else{
    $startY = 35
}


    # Taille du formulaire
    $formWidth = $form.ClientSize.Width

    # Largeur du label
    $labelWidth = 200

    # Calcul de la position X pour centrer le label
    $xPos = [int](($formWidth - $labelWidth) / 2)

    # Création du label avec la position centrée
    $labelMode = New-Object System.Windows.Forms.Label
    $labelMode.Text = "Mode : $Action"
    $labelMode.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Bold)
    $labelMode.Size = New-Object System.Drawing.Size($labelWidth, 25)
    $labelMode.Location = New-Object System.Drawing.Point($xPos, 5)
    $form.Controls.Add($labelMode)

    $labels = @{
        'GivenName'        = "Prénom :"
        'Surname'          = "Nom :"
        'DisplayName'      = "Nom complet (displayName) :"
        'SamAccountName'   = "Nom de connexion (samAccountName) :"
        'UserPrincipalName' = "User Principal Name (UPN) :"
        'Email'            = "Adresse email :"
        'Password'         = "Mot de passe :"
    }

    $controls = @{}
    $spacingY = 40
    $y = $startY

    function Add-LabeledTextbox {
        param($labelText, $name, [ref]$currentY, [bool]$isPassword = $false)

        [int]$yValue = $currentY.Value

        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labelText
        $label.Location = New-Object System.Drawing.Point(20, ($yValue + 5))  # Décalage 5px vers le bas
        $label.Size = New-Object System.Drawing.Size(200, 20)
        $form.Controls.Add($label)

        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Location = New-Object System.Drawing.Point(230, ($yValue - 3))
        $textbox.Size = New-Object System.Drawing.Size(320, 25)

        if ($isPassword) {
            $textbox.UseSystemPasswordChar = $true
            $textbox.ReadOnly = $true
        }

        $form.Controls.Add($textbox)
        $controls[$name] = $textbox
        $currentY.Value += $spacingY
    }

    # Création des champs texte
    foreach ($key in 'GivenName', 'Surname', 'DisplayName', 'SamAccountName', 'UserPrincipalName', 'Email') {
        Add-LabeledTextbox -labelText $labels[$key] -name $key -currentY ([ref]$y)
    }

    # En mode Copy, ajout du champ source utilisateur
    if ($Action -eq 'Copy') {
        $labelSourceUser = New-Object System.Windows.Forms.Label
        $labelSourceUser.Text = "Utilisateur Source :"
        $labelSourceUser.Location = New-Object System.Drawing.Point(20, ($y + 5))
        $labelSourceUser.Size = New-Object System.Drawing.Size(200, 20)
        $form.Controls.Add($labelSourceUser)

        $txtSource = New-Object System.Windows.Forms.TextBox
        $txtSource.Location = New-Object System.Drawing.Point(230, ($y - 3))
        $txtSource.Size = New-Object System.Drawing.Size(320, 25)
        $form.Controls.Add($txtSource)

        $y += $spacingY
    }
    else {
        $txtSource = $null
    }

    # Champ mot de passe
    Add-LabeledTextbox -labelText $labels['Password'] -name 'Password' -currentY ([ref]$y) -isPassword:$true

    $btnGenPass = New-Object System.Windows.Forms.Button
    $btnGenPass.Text = "Générer mot de passe"
    $btnGenPass.Size = New-Object System.Drawing.Size(200, 30)
    $btnGenPass.Location = New-Object System.Drawing.Point(230, $y)
    $form.Controls.Add($btnGenPass)

    $y += 50

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = "Créer"
    $btnCreate.Size = New-Object System.Drawing.Size(100, 30)
    $btnCreate.Location = New-Object System.Drawing.Point(230, $y)
    $form.Controls.Add($btnCreate)

    $btnReturn = New-Object System.Windows.Forms.Button
    $btnReturn.Text = "Retour"
    $btnReturn.Size = New-Object System.Drawing.Size(100, 30)
    $btnReturn.Location = New-Object System.Drawing.Point(350, $y)
    $form.Controls.Add($btnReturn)

    # Ligne horizontale sous bouton Retour
    $borderLine = New-Object System.Windows.Forms.Label
    $borderLine.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $borderLine.AutoSize = $false
    $borderLine.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 40), 2)
    $borderLine.Location = New-Object System.Drawing.Point(20, ($y + 40))
    $form.Controls.Add($borderLine)

    # Évènements
    $btnGenPass.Add_Click({
        $pw = Generate-StrongPassword
        $controls['Password'].Text = $pw
    })

    $btnReturn.Add_Click({
        $form.Tag = 'Return'
        $form.Close()
    })

    # Validation champs
    function Validate-Inputs {
        foreach ($key in 'GivenName','Surname','DisplayName','SamAccountName','UserPrincipalName') {
            if ([string]::IsNullOrWhiteSpace($controls[$key].Text)) {
                [System.Windows.Forms.MessageBox]::Show("Le champ '$($labels[$key])' est obligatoire.","Erreur","OK","Error") | Out-Null
                return $false
            }
        }
        if ([string]::IsNullOrWhiteSpace($controls['Password'].Text)) {
            [System.Windows.Forms.MessageBox]::Show("Générez un mot de passe avant de continuer.","Erreur","OK","Error") | Out-Null
            return $false
        }
        if ($Action -eq 'Copy' -and ([string]::IsNullOrWhiteSpace($txtSource.Text))) {
            [System.Windows.Forms.MessageBox]::Show("Le champ 'SamAccountName source' est obligatoire en mode copie.","Erreur","OK","Error") | Out-Null
            return $false
        }
        return $true
    }

    $btnCreate.Add_Click({
        if (-not (Validate-Inputs)) { return }

        $sam = $controls['SamAccountName'].Text.Trim()
        $given = $controls['GivenName'].Text.Trim()
        $surname = $controls['Surname'].Text.Trim()
        $display = $controls['DisplayName'].Text.Trim()
        $upn = $controls['UserPrincipalName'].Text.Trim()
        $email = $controls['Email'].Text.Trim()
        $pass = $controls['Password'].Text.Trim()

        if ($Action -eq 'Copy') {
            $sourceSam = $txtSource.Text.Trim()
            try {
                $sourceUser = Get-ADUser -Identity $sourceSam -Properties * -ErrorAction Stop
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Utilisateur source introuvable : $sourceSam","Erreur","OK","Error") | Out-Null
                return
            }
        }

        if ($Mode -eq 'Demo') {
            [System.Windows.Forms.MessageBox]::Show("Simulation : création utilisateur :`nSamAccountName = $sam`nDisplayName = $display","Mode Démo") | Out-Null
            Log-UserCreation -SamAccountName $sam -DisplayName $display -Mode $Mode -Action $Action
        }
        else {
            $success = Create-ADUserReal -SamAccountName $sam -GivenName $given -Surname $surname -DisplayName $display -UserPrincipalName $upn -EmailAddress $email -Password $pass
            if ($success) {
                [System.Windows.Forms.MessageBox]::Show("Utilisateur créé avec succès.","Succès") | Out-Null
                Log-UserCreation -SamAccountName $sam -DisplayName $display -Mode $Mode -Action $Action
            }
        }
    })

    $form.ShowDialog() | Out-Null
    return $form.Tag
}


# ------------------------------
# Script principal (boucle navigation)
do {
    $mode = Show-ModeChoiceWindow
    if (-not $mode) { break }
    
    do {
        $action = Show-ActionChoiceWindow
        if (-not $action) { break 2 }  # Sort de la boucle principale si annulation
        if ($action -eq 'Return') { break } # Revenir au choix du mode
        
        $result = Show-MainForm -Mode $mode -Action $action
    } while ($result -eq 'Return')
} while ($true)
