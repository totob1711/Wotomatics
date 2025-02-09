# Paramètres SMTP pour Outlook/Office 365
$smtpServer = "smtp.office365.com"
$smtpPort = 587
$logFilePath = "C:\Path\To\Wotomatics\logs\logs.txt" # Chemin du fichier de logs

# Fonction pour écrire dans le fichier de logs
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$timestamp - $message" | Out-File -FilePath $logFilePath -Append
}

# Demander l'adresse e-mail et le mot de passe à l'utilisateur
Write-Host ""
Write-Host "Veuillez entrer vos informations de connexion." -ForegroundColor Yellow
Write-Log "Tentative de connexion au serveur SMTP"

# Cacher toute interaction visuelle avant de demander l'email et le mot de passe
$smtpEmail = Read-Host "Entrez votre adresse e-mail (ex: votre-adresse@outlook.com)"
$smtpPassword = Read-Host "Entrez votre mot de passe" -AsSecureString

# Créer une instance de SmtpClient pour tester la connexion
try {
    # Créer une instance de SmtpClient
    $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
    $smtpClient.EnableSsl = $true

    # Créer un objet NetworkCredential avec les informations de connexion
    $networkCredential = New-Object System.Net.NetworkCredential($smtpEmail, $smtpPassword)

    # Tenter de se connecter avec les informations d'identification
    $smtpClient.Credentials = $networkCredential

    # Création d'un e-mail de test pour vérifier la connexion
    $testMail = New-Object System.Net.Mail.MailMessage
    $testMail.From = $smtpEmail
    $testMail.To.Add($smtpEmail)  # Envoi à l'adresse de l'utilisateur pour vérifier
    $testMail.Subject = "Test de connexion SMTP"
    $testMail.Body = "Ce message est un test pour verifier la connexion SMTP."

    # Essayer d'envoyer le message de test
    $smtpClient.Send($testMail)
    
    # Si l'e-mail est envoyé avec succès
    Write-Host "Connexion SMTP reussie." -ForegroundColor Green
    Write-Log "Connexion SMTP reussie avec l'adresse e-mail: $smtpEmail"
    Write-Host "Patientez pendant l'envoi des mails..." -ForegroundColor Yellow
} catch {
    Write-Host "Erreur: Impossible de se connecter au serveur SMTP. Details: $_" -ForegroundColor Red
    Write-Log "Erreur: Impossible de se connecter au serveur SMTP. Details: $_"
    exit
}


Write-Host ""

# Créer une instance de l'application Excel
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false # Laisser Excel invisible (s'exécute en arrière-plan)

# Définir le chemin de la base de données qui sera utilisée
$dataPath = 'C:\Path\To\Wotomatics\annexes\data.xlsx'
Write-Log "Ouverture du fichier Excel reussie"
$workBook = $excelApp.Workbooks.Open($dataPath)
$workSheet = $workBook.Sheets.Item(1) # Sélectionner la première feuille
$rowCount = $workSheet.UsedRange.Rows.Count # Récupérer le nombre de lignes dans la feuille

# Boucle sur chaque ligne de données pour envoyer un e-mail à chaque utilisateur
for ($i = 2; $i -le $rowCount; $i++) {
    # Récupérer les informations de l'utilisateur depuis Excel
    $prenom = $workSheet.Cells.Item($i, 3).Text
    $nom = $workSheet.Cells.Item($i, 2).Text

    # Construire l'adresse e-mail et le nom du fichier
    $pdfFileName = "$prenom`_$nom.pdf"

    # Convertir prénom et nom en minuscules pour l'adresse e-mail
    $prenom = $prenom.ToLower()
    $nom = $nom.ToLower()
    $emailTo = "$prenom.$nom@mail.fr"
    $pdfFilePath = "C:\Path\To\Wotomatics\publipostage\$pdfFileName"

    # Vérifier si le fichier PDF existe
    if (Test-Path $pdfFilePath) {
        $emailSubject = "Document"
        $emailBody = "Bonjour $prenom, veuillez trouver ci-joint votre document."

        # Créer l'objet PSCredential pour l'authentification
        $credential = New-Object System.Management.Automation.PSCredential($smtpEmail, $smtpPassword)

        # Envoi du mail
        try {
            Send-MailMessage -From $smtpEmail -To $emailTo -Subject $emailSubject -Body $emailBody -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $credential -Attachments $pdfFilePath
            Write-Host "E-mail envoye a $emailTo pour le fichier $($pdfFileName)." -ForegroundColor Green
            Write-Log "E-mail envoye a $emailTo pour le fichier $pdfFileName."
        } catch {
            Write-Host "Erreur lors de l'envoi de l'e-mail à $emailTo pour le fichier $($pdfFileName). Details: $_" -ForegroundColor Red
            Write-Log "Erreur lors de l'envoi de l'e-mail a $emailTo pour le fichier $pdfFileName. Details: $_"
        }
    } else {
        Write-Host "Le fichier $pdfFileName n'existe pas dans le repertoire specifie." -ForegroundColor Red
        Write-Log "Le fichier $pdfFileName n'existe pas dans le repertoire specifie."
    }
}

# Fermer l'instance d'Excel
$workBook.Close()
$excelApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
Write-Log "Fermeture de l'instance Excel et liberation des objets COM."
