# yyyyMMdd - REMESSA ESTRATEGICA - TALENTOS.txt
# Este script move a Remessa Estratégica da Via Varejo para o diretório carga 18 de importação do CRM, replicada automaticamente para a via 55.
# This script moves the Remessa Estratégica from Via Varejo to the Carga 18 import directory of the CRM, which is automatically replicated to Via 55.

function Load-EnvVariables {
    $envFile = "C:\Users\ayla.atilio\dev\aut\config.env"  
    if (Test-Path $envFile) {
        Get-Content $envFile | ForEach-Object {
            if ($_ -match "^\s*([^#].+?)=(.+)$") {
                [System.Environment]::SetEnvironmentVariable($matches[1], $matches[2], "Process")
            }
        }
    } else {
        throw "Arquivo .env não encontrado."
    }
}

Load-EnvVariables

# WinSCP module import
Add-Type -Path "C:\Users\ayla.atilio\dev\libraries\WinSCP\WinSCPnet.dll"

# SFTP extcubo Config
$sessionOptions = New-Object WinSCP.SessionOptions
$sessionOptions.Protocol = [WinSCP.Protocol]::Sftp
$sessionOptions.HostName = [System.Environment]::GetEnvironmentVariable("SFTP_HOST_CUBO")
$sessionOptions.UserName = [System.Environment]::GetEnvironmentVariable("SFTP_USER_CUBO")
$sessionOptions.SshPrivateKeyPath = [System.Environment]::GetEnvironmentVariable("SFTP_KEY_CUBO")
$sessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true

# EnvVariables config verification
if (-not $sessionOptions.HostName) { throw "SFTP_HOST não está configurado." }

# SFTP extcubo Initialize
$session = New-Object WinSCP.Session
$session.SessionLogPath = $null

# Error and Logs variables
$errorLog = ""
$successLog = ""

# Local Paths
$localPath18 = [System.Environment]::GetEnvironmentVariable("LOCAL_PATH_18")
$localPath55 = [System.Environment]::GetEnvironmentVariable("LOCAL_PATH_55")

# Backup Verification Primary Loads
$verificationPath18 = [System.Environment]::GetEnvironmentVariable("VERIFY18")
$verificationPath55 = [System.Environment]::GetEnvironmentVariable("VERIFY55")

# Remote Paths
$remotePath = [System.Environment]::GetEnvironmentVariable("REMOTE_PATH_RE")
$remoteProcessedPath = [System.Environment]::GetEnvironmentVariable("REMOTE_PROCESSED_PATH_RE")

# Get Date yyyyMMdd
$currentDateStr = (Get-Date).ToString("yyyyMMdd")

# File Name
$fileNamePattern = "$currentDateStr - REMESSA ESTRATEGICA - TALENTOS.txt"

# Function to check if there are files with today's modification date
function HasFileWithCurrentDate($folderPath) {
    $currentDate = Get-Date
    $files = Get-ChildItem -Path $folderPath -File | Where-Object {
        $_.LastWriteTime.Date -eq $currentDate.Date
    }
    return $files.Count -gt 0
}

try {
    # check if there are files with today's modification date
    if (-not (HasFileWithCurrentDate $verificationPath18) -or -not (HasFileWithCurrentDate $verificationPath55)) {
        throw "Nenhum arquivo com a data de modificação de hoje encontrado nos diretórios de verificação '18' e '55'. Processamento abortado."
    }

    # Verification for folder containing files (ignore subfolders)
    function ContainsFiles($folderPath) {
        $files = Get-ChildItem -Path $folderPath -File
        return $files.Count -gt 0
    }

    if (ContainsFiles $localPath18 -or ContainsFiles $localPath55) {
        throw "As pastas locais 18 ou 55 contêm arquivos. Processamento abortado."
    }

    # SFTP Server Conection
    $session.Open($sessionOptions) | Out-Null

    # Locate the remote file
    $remoteFiles = $session.EnumerateRemoteFiles($remotePath, $fileNamePattern, [WinSCP.EnumerationOptions]::None)
    $targetFile = $remoteFiles | Where-Object { $_.Name -eq $fileNamePattern -and $_.IsDirectory -eq $false }

    if (-not $targetFile) {
        throw "Arquivo '$fileNamePattern' não encontrado no diretório remoto '$remotePath'."
    }

    $remoteFilePath = [WinSCP.RemotePath]::Combine($remotePath, $targetFile.Name)

    # Download files to the server paths
    $destinationFilePath18 = [System.IO.Path]::Combine($localPath18, $targetFile.Name)

    $session.GetFiles($remoteFilePath, $destinationFilePath18).Check() | Out-Null

    # Move file to PROCESSADOS folder on the remote server
    $remoteProcessedFilePath = [WinSCP.RemotePath]::Combine($remoteProcessedPath, $targetFile.Name)
    $session.MoveFile($remoteFilePath, $remoteProcessedFilePath) | Out-Null

    $successLog += "Arquivo '$($targetFile.Name)' processado com sucesso.`n"

} catch {
    $errorLog += "Erro durante o processamento: $($_.Exception.Message)`n"
} finally {
    # SFTP Session Closed
    $session.Dispose() | Out-Null
}

# E-mail Error Function
if ($errorLog -ne "") {
    $smtpServer = [System.Environment]::GetEnvironmentVariable("SMTP_SERVER")
    $port = [System.Environment]::GetEnvironmentVariable("SMTP_PORT")
    $userName = [System.Environment]::GetEnvironmentVariable("EMAIL_USER")
    $password = [System.Environment]::GetEnvironmentVariable("EMAIL_PASS")
    $to = [System.Environment]::GetEnvironmentVariable("EMAIL_TO")

    # Create SMTP Client
    $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $port)
    $smtpClient.EnableSsl = $true
    $smtpClient.Credentials = New-Object System.Net.NetworkCredential($userName, $password)

    # E-mail Config
    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.From = $userName
    $mailMessage.To.Add($to)
    $mailMessage.Subject = "ERRO NA REMESSA ESTRATÉGICA"
    $mailMessage.Body = $errorLog
    $mailMessage.IsBodyHtml = $false

    # Send E-mail
    $smtpClient.Send($mailMessage)

    # Session closed
    $mailMessage.Dispose()
    $smtpClient.Dispose()
}

# Success E-mail
if ($successLog -ne "" -and $errorLog -eq "") {
    $smtpServer = [System.Environment]::GetEnvironmentVariable("SMTP_SERVER")
    $port = [System.Environment]::GetEnvironmentVariable("SMTP_PORT")
    $userName = [System.Environment]::GetEnvironmentVariable("EMAIL_USER")
    $password = [System.Environment]::GetEnvironmentVariable("EMAIL_PASS")
    $to = [System.Environment]::GetEnvironmentVariable("EMAIL_TO")

	# Create SMTP Client
    $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $port)
    $smtpClient.EnableSsl = $true
    $smtpClient.Credentials = New-Object System.Net.NetworkCredential($userName, $password)

	# E-mail Config
    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.From = $userName
    $mailMessage.To.Add($to)
    $mailMessage.Subject = "REMESSA ESTRETÉGICA PROCESSADA COM SUCESSO!"
    $mailMessage.Body = $successLog
    $mailMessage.IsBodyHtml = $false

    # Send E-mail
    $smtpClient.Send($mailMessage)

    # Session closed
    $mailMessage.Dispose()
    $smtpClient.Dispose()
}