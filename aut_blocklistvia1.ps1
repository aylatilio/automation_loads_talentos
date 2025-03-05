# BADLIST_VVAR.zip
# Este script move o arquivo BADLIST_VVAR.zip do SFTP para um diretório local, onde é extraído em .csv e salvo em .txt. Envia e-mail em caso de erro.
# This script moves the BADLIST_VVAR.zip file from the SFTP to a local directory, where it is extracted as a .csv, saved as a .txt and sends an email in case of an error.

# Sensitive Data .env
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

# Paths
$remotePath = [System.Environment]::GetEnvironmentVariable("REMOTE_PATH_B1")
$remoteCapturedPath = [System.Environment]::GetEnvironmentVariable("REMOTE_CAPTURE_PATH_B1")
$localPath = [System.Environment]::GetEnvironmentVariable("LOCAL_PATH_B1")


# Get today date yyyyMMdd
$currentDateStr = (Get-Date).ToString("yyyyMMdd")

# File Name
$fileNamePattern = "BADLIST_VVAR_$currentDateStr.zip"

# E-mail Error Function
function Send-ErrorEmail($errorMessage) {
    $smtpServer = [System.Environment]::GetEnvironmentVariable("SMTP_SERVER")
    $port = [System.Environment]::GetEnvironmentVariable("SMTP_PORT")
    $userName = [System.Environment]::GetEnvironmentVariable("EMAIL_USER")
    $password = [System.Environment]::GetEnvironmentVariable("EMAIL_PASS")
    $to = [System.Environment]::GetEnvironmentVariable("EMAIL_TO")

    $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $port)
    $smtpClient.EnableSsl = $true
    $smtpClient.Credentials = New-Object System.Net.NetworkCredential($userName, $password)

    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.From = $userName
    $mailMessage.To.Add($to)
    $mailMessage.Subject = "Erro no processamento BADLIST_VVAR"
    $mailMessage.Body = $errorMessage
    $mailMessage.IsBodyHtml = $false

    $smtpClient.Send($mailMessage)

    $mailMessage.Dispose()
    $smtpClient.Dispose()
}

try {
    # Connect to the SFTP server
    $session.Open($sessionOptions)

    # Locate the remote file
    $remoteFiles = $session.EnumerateRemoteFiles($remotePath, $fileNamePattern, [WinSCP.EnumerationOptions]::None)
    if (-not $remoteFiles) {
        throw "Nenhum arquivo remoto encontrado em '$remotePath'."
    }

    $targetFile = $remoteFiles | Where-Object { $_.Name -eq $fileNamePattern -and $_.IsDirectory -eq $false }
    if (-not $targetFile) {
        throw "Arquivo '$fileNamePattern' não encontrado no diretório remoto '$remotePath'."
    }

    $remoteFilePath = [WinSCP.RemotePath]::Combine($remotePath, $targetFile.Name)
    $localFilePath = [System.IO.Path]::Combine($localPath, $targetFile.Name)

    # Download .zip file to the localpath
    $downloadResult = $session.GetFiles($remoteFilePath, $localFilePath)
    if (-not $downloadResult.IsSuccess) {
        throw "Falha ao baixar o arquivo '$remoteFilePath'."
    }

    # .zip verification
    if (-not (Test-Path $localFilePath)) {
        throw "Arquivo ZIP '$localFilePath' não foi baixado corretamente."
    }

    # Extract the .zip file contents
    Expand-Archive -Path $localFilePath -DestinationPath $localPath -Force

    # Process the extracted .csv file
    $csvFile = Get-ChildItem -Path $localPath -Filter "*.csv" | Select-Object -First 1
    if ($csvFile -eq $null) {
        throw "Nenhum arquivo CSV encontrado no ZIP extraído."
    }

    # Convert .csv in .txt
    $txtFilePath = [System.IO.Path]::ChangeExtension($csvFile.FullName, ".txt")
    Rename-Item -Path $csvFile.FullName -NewName $txtFilePath -ErrorAction Stop

    # Remove the local ZIP file after extraction
    Remove-Item -Path $localFilePath -Force

    # Move .zip file to CAPTURADOS
    $remoteCapturedFilePath = [WinSCP.RemotePath]::Combine($remoteCapturedPath, $targetFile.Name)
    if ([string]::IsNullOrWhiteSpace($remoteCapturedFilePath)) {
        throw "Caminho remoto para CAPTURADOS é inválido."
    }

    # Move .zip file to CAPTURADOS with aditional verification
    try {
        $session.MoveFile($remoteFilePath, $remoteCapturedFilePath)
    } catch {
        throw "Falha ao mover o arquivo '$remoteFilePath' para '$remoteCapturedFilePath': $($_.Exception.Message)"
    }

    # Moving verification
    if ($session.FileExists($remoteFilePath)) {
        throw "O arquivo ainda existe em '$remoteFilePath' após tentativa de mover."
    }

} catch {
    # Get error and send e-mail
    $errorMessage = "Erro durante o processamento: $($_.Exception.Message)"
    Write-Output $errorMessage
    Send-ErrorEmail $errorMessage
} finally {
    # SFTP Session Closed
    $session.Dispose()
}
