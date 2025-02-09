# Fonction pour écrire dans le fichier de logs
$logFilePath = "C:\Path\To\Wotomatics\logs\logs.txt" # Chemin du fichier de logs
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$timestamp - $message" | Out-File -FilePath $logFilePath -Append
}

Write-Host ""

# Creer une instance de l'application Word
try {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false # Permettre a Word de s'executer de maniere visible
} catch {
    Write-Host "Erreur: Impossible de demarrer Microsoft Word." -ForegroundColor Red
    Write-Log "Erreur: Impossible de demarrer Microsoft Word."
    exit
}

# Definir le chemin du template utilise pour le publipostage
$templatePath = 'C:\Path\To\Wotomatics\annexes\modelTest.docx'
$template = $null
try {
    $template = $wordApp.Documents.Open($templatePath)
} catch {
    Write-Host "Erreur: Impossible d'ouvrir le fichier Word a partir de '$templatePath'." -ForegroundColor Red
    Write-Log "Erreur: Impossible d'ouvrir le fichier Word a partir de '$templatePath'."
    $wordApp.Quit() | Out-Null
    exit
}

# Creer une instance de l'application Excel
try {
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false # Laisser Excel invisible (s'execute en arriere-plan)
} catch {
    Write-Host "Erreur: Impossible de demarrer Microsoft Excel." -ForegroundColor Red
    Write-Log "Erreur: Impossible de demarrer Microsoft Excel."
    $wordApp.Quit() | Out-Null
    exit
}

# Definir le chemin de la base de donnees qui sera utilisee
$dataPath = 'C:\Path\To\Wotomatics\annexes\dataTest.xlsx'
$workBook = $null
try {
    $workBook = $excelApp.Workbooks.Open($dataPath)
} catch {
    Write-Host "Erreur: Impossible d'ouvrir le fichier Excel a partir de '$dataPath'." -ForegroundColor Red
    Write-Log "Erreur: Impossible d'ouvrir le fichier Excel a partir de '$dataPath'."
    $wordApp.Quit() | Out-Null
    $excelApp.Quit() | Out-Null
    exit
}

$workSheet = $workBook.Sheets.Item(1)
$rowCount = $workSheet.UsedRange.Rows.Count

# Stocker les donnees du fichier Excel dans le tableau data
$data = @()
for ($i = 2; $i -le $rowCount; $i++) {
    $dataRow = @{
        "date" = $workSheet.Cells.Item($i, 1).Text
        "nom" = $workSheet.Cells.Item($i, 2).Text
        "prenom" = $workSheet.Cells.Item($i, 3).Text
        "equipe" = $workSheet.Cells.Item($i, 5).Text
    }
    $data += $dataRow
}

# Fermer l'instance d'Excel
$workBook.Close() | Out-Null
$excelApp.Quit() | Out-Null

# Afficher un message de chargement
Write-Host "Creation des fichiers en cours..." -ForegroundColor White

# Creer autant de documents Word qu'il y a de lignes dans notre tableau
if ($data.Count -gt 0) {
    $progress = 0
    $progressStep = [Math]::Round(100 / $data.Count)

    foreach ($row in $data) {
        try {
            # Dupliquer le modele pour chaque enregistrement
            $newDocument = $wordApp.Documents.Add($templatePath) # Ajouter un nouveau document a partir du modele
            if ($newDocument -eq $null) {
                Write-Host "Erreur: Impossible de creer un nouveau document Word pour '$($row['prenom']) $($row['nom'])." -ForegroundColor Red
                Write-Log "Erreur: Impossible de creer un nouveau document Word pour '$($row['prenom']) $($row['nom'])."
                continue
            }

            # Remplacer les champs de fusion par les donnees du fichier Excel
            $findText = @("{MERGEFIELD nom}", "{MERGEFIELD prenom}", "{MERGEFIELD date}", "{MERGEFIELD equipe}")
            $replaceText = @($row["nom"], $row["prenom"], $row["date"], $row["equipe"])

            for ($j = 0; $j -lt $findText.Length; $j++) {
                # Chercher et remplacer chaque champ
                $result = $newDocument.Content.Find.Execute($findText[$j], $true, $false, $false, $false, $false, $false, 1, $false, $replaceText[$j], 2) 
                if (-not $result) {
                    Write-Host "Erreur: Impossible de remplacer le champ '$($findText[$j])' dans le document pour '$($row['prenom']) $($row['nom'])." -ForegroundColor Red
                    Write-Log "Erreur: Impossible de remplacer le champ '$($findText[$j])' dans le document pour '$($row['prenom']) $($row['nom'])."
                    continue
                }
            }

            # Definir ou seront sauvegardes les fichiers du publipostage
            $outputPath = "C:\Path\To\Wotomatics\publipostage\$($row['prenom'])_$($row['nom']).pdf"
            $newDocument.SaveAs([ref] $outputPath, 17) | Out-Null # 17 est le format PDF
            Write-Log "Le fichier PDF a bien ete cree pour '$($row['prenom']) $($row['nom'])."

            # Fermer le document sans sauvegarder les modifications dans le modele
            $newDocument.Close($false) | Out-Null
        } catch {
            Write-Host "Erreur: Une erreur est survenue pendant la creation du fichier pour '$($row['prenom']) $($row['nom'])." -ForegroundColor Red
            Write-Log "Erreur: Une erreur est survenue pendant la creation du fichier pour '$($row['prenom']) $($row['nom'])."
        }

        # Mettre a jour la barre de progression
        $progress = [Math]::Min(100, $progress + $progressStep) # S'assurer que la progression ne depasse pas 100
        Write-Progress -PercentComplete $progress -Status "Creation des fichiers" -Activity "$($progress)% termine" | Out-Null
    }
    # Afficher un message de confirmation
    Write-Host "L'operation s'est deroulee avec succes !" -ForegroundColor Green
    Write-Host ""

    # Arreter toute barre de progression qui pourrait etre encore en cours
    Write-Progress -PercentComplete 0 -Status " " -Activity " " | Out-Null

    # Demander si l'utilisateur veut envoyer les fichiers par mail
    $sendEmail = Read-Host "Voulez-vous envoyer les fichiers par mail ? [Y/n]"

    # Si l'utilisateur repond "Y", appeler le script d'envoi des e-mails
    if ($sendEmail -eq "Y" -or $sendEmail -eq "y") {
        # Lancer le script d'envoi des emails
        .\sendMails.ps1
    } else {
        Write-Host "Aucun e-mail n'a ete envoye." -ForegroundColor Yellow
    }
} else {
    Write-Host "Aucune donnee a traiter." -ForegroundColor Yellow
}

# Fermer l'instance de Word
$template.Close($false) | Out-Null # Ne pas enregistrer le modele
$wordApp.Quit() | Out-Null

# Liberer les objets COM sans generer de sortie
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null

