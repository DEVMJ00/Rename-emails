Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Création de la fenêtre
$form = New-Object System.Windows.Forms.Form
$form.Text = "Renommer les mails"
$form.Size = New-Object System.Drawing.Size(600, 300)
$form.FormBorderStyle = 'FixedDialog'
$form.StartPosition = "CenterScreen"

# Texte d'instruction
$labelIntro = New-Object System.Windows.Forms.Label
$labelIntro.Text = "Selectionner le fichier a renommer, puis choisir un nouveau nom :"
$labelIntro.AutoSize = $true
$labelIntro.Location = New-Object System.Drawing.Point(20, 20)
$form.Controls.Add($labelIntro)

# Liste des fichiers .msg
$comboFiles = New-Object System.Windows.Forms.ComboBox
$comboFiles.Location = New-Object System.Drawing.Point(20, 50)
$comboFiles.Size = New-Object System.Drawing.Size(540, 20)
$comboFiles.DropDownStyle = 'DropDownList'
$form.Controls.Add($comboFiles)

# Liste des options de renommage
$comboOptions = New-Object System.Windows.Forms.ComboBox
$comboOptions.Location = New-Object System.Drawing.Point(20, 90)
$comboOptions.Size = New-Object System.Drawing.Size(300, 20)
$comboOptions.DropDownStyle = 'DropDownList'
$comboOptions.Items.AddRange(@(
    "1 - 1ere relance",
    "2 - 2eme relance",
    "3 - 3eme relance",
    "4 - 4eme relance",
    "5 - Information complementaire",
    "6 - Demande de devis",
    "7 - Devis",
    "8 - Demande d intervention",
    "9 - Rapport d'intervention",
    "10 - Facture",
    
    "R - Reponse"
))
$form.Controls.Add($comboOptions)

# Bouton de renommage
$btnRenommer = New-Object System.Windows.Forms.Button
$btnRenommer.Text = "Renommer"
$btnRenommer.Location = New-Object System.Drawing.Point(350, 88)
$btnRenommer.Size = New-Object System.Drawing.Size(100, 25)
$form.Controls.Add($btnRenommer)

# Label d'information
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.AutoSize = $true
$labelStatus.Location = New-Object System.Drawing.Point(20, 130)
$form.Controls.Add($labelStatus)

# Chargement des fichiers .msg à l'ouverture
$script:msgFiles = Get-ChildItem -Path "." -Filter "*.msg" | Select-Object -ExpandProperty Name
if ($msgFiles.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("Aucun fichier .msg existant dans le dossier.","Information",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    $form.Close()
}
else {
    $comboFiles.Items.AddRange($msgFiles)
    $comboFiles.SelectedIndex = 0
    $comboOptions.SelectedIndex = 0
}

# Action de renommage
$btnRenommer.Add_Click({
    $fichierSelectionne = $comboFiles.SelectedItem
    $choix = $comboOptions.SelectedItem
    if (-not $fichierSelectionne -or -not $choix) {
        $labelStatus.Text = "Veuillez choisir un fichier et une action."
        return
    }

    $prefixe = switch -Regex ($choix) {
        "^1" { "1ere_relance" }
        "^2" { "2eme_relance" }
        "^3" { "3eme_relance" }
        "^4" { "4eme_relance" }
	"^5" { "Information_complementaire" }
	"^6" { "Demande_de_devis" }
        "^7" { "Devis" }
	"^8" { "Demande_d_intervention" }
        "^9" { "Rapport_d_intervention" }
        "^10" { "Facture" }
        "^R" { "Reponse" }
        default { "Fichier" }
    }

    # Génération de la date du jour au format yyyyMMdd
    $date = Get-Date -Format "yyyyMMdd"

    # Séparation du nom de fichier en nom sans extension et extension
    $nomSansExtension = [System.IO.Path]::GetFileNameWithoutExtension($fichierSelectionne)
    $extension = [System.IO.Path]::GetExtension($fichierSelectionne)
    $nouveauNom = "${date}_${prefixe}${extension}"

    try {
        Rename-Item -Path $fichierSelectionne -NewName $nouveauNom -ErrorAction Stop
        $labelStatus.ForeColor = 'DarkGreen'
	$labelStatus.MaximumSize = New-Object System.Drawing.Size(550, 0)  # Largeur max, hauteur auto
        $labelStatus.Font = New-Object System.Drawing.Font($labelStatus.Font, [System.Drawing.FontStyle]::Bold)
	$labelStatus.Text = "'$fichierSelectionne' s'appelle maintenant : '$nouveauNom'."
        # Mise à jour de la liste
        $comboFiles.Items.Clear()
        $script:msgFiles = Get-ChildItem -Path "." -Filter "*.msg" | Select-Object -ExpandProperty Name
        $comboFiles.Items.AddRange($msgFiles)
        if ($msgFiles.Count -gt 0) {
            $comboFiles.SelectedIndex = 0
        }
    } catch {
        $labelStatus.ForeColor = 'DarkRed'
        $labelStatus.Text = "Erreur : $_"
    }
})

# Lancement
[void]$form.ShowDialog()
