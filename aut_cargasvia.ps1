# yyyyMMdd(-1)_remessa_consolidada_base_out.txt
# Este script move as cargas primárias da Via Varejo, bases 18 e 55, para os diretórios de importação do CRM.
# This script moves the Via Varejo primary loads, bases 18 and 55, to the CRM import directories.

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
$sessionOptions.HostName = [System.Environment]::GetEnvironmentVariable("SFTP_HOST_VV")
$sessionOptions.UserName = [System.Environment]::GetEnvironmentVariable("SFTP_USER_VV")
$sessionOptions.SshPrivateKeyPath = [System.Environment]::GetEnvironmentVariable("SFTP_KEY_VV")
$sessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true

# EnvVariables config verification
if (-not $sessionOptions.HostName) { throw "SFTP_HOST não está configurado." }

# SFTP vvarejo Initialize
$session = New-Object WinSCP.Session
$session.SessionLogPath = $null

# Error and Logs variables
$errorLog = ""
$successLog = ""

try {
    # SFTP Server Conection
    $session.Open($sessionOptions) | Out-Null

    # Paths
    $remotePath = [System.Environment]::GetEnvironmentVariable("REMOTE_PATH_RC")
	$remoteProcessedPath = [System.Environment]::GetEnvironmentVariable("REMOTE_PROCESSED_PATH_RC")
    $localPath18 = [System.Environment]::GetEnvironmentVariable("LOCAL_PATH_18")
	$localPath55 = [System.Environment]::GetEnvironmentVariable("LOCAL_PATH_55")
    
    # EnumerateRemoteFiles list files in remote path
    $remoteFiles = $session.EnumerateRemoteFiles($remotePath, "*", [WinSCP.EnumerationOptions]::None)

    # Get day before yyyyMMdd(-1)
    $previousDateStr = (Get-Date).AddDays(-1).ToString("yyyyMMdd")

    # Filter files based on the date in filename
    $filteredFiles = $remoteFiles | Where-Object { $_.Name -match "^$previousDateStr" }

    # If no files are found, log the error and terminate the script
    if (-not $filteredFiles) {
        $errorLog += "Nenhum arquivo com a data $previousDateStr foi encontrado no diretório remoto.`n"
        throw "Nenhum arquivo encontrado."
    }

	# Function to check if the folder contains files (ignore subfolders)
    function IsFolderEmptyOrContainsOnlySubfolders($folderPath) {
        $files = Get-ChildItem -Path $folderPath -File
        return $files.Count -eq 0
    }

    # Verification for folder containing files (ignore subfolders)
    if (-not (IsFolderEmptyOrContainsOnlySubfolders $localPath18)) {
        $errorLog += "A pasta local 18 contém arquivos e não pode receber novas cargas.`n"
    }
    if (-not (IsFolderEmptyOrContainsOnlySubfolders $localPath55)) {
        $errorLog += "A pasta local 55 contém arquivos e não pode receber novas cargas.`n"
    }

    # Download files to the paths
    foreach ($fileInfo in $filteredFiles) {
        if ($fileInfo.IsDirectory -eq $false -and $fileInfo.Name -match "\.txt$") {
            $remoteFilePath = [WinSCP.RemotePath]::Combine($remotePath, $fileInfo.Name)
            
            try {
                if ($fileInfo.Name -match "_remessa_consolidada_18_out") {
                    if (IsFolderEmptyOrContainsOnlySubfolders $localPath18) {
                        $destinationFilePath18 = [System.IO.Path]::Combine($localPath18, $fileInfo.Name)
                        $session.GetFiles($remoteFilePath, $destinationFilePath18).Check() | Out-Null

                        # Move file to PROCESSADOS folder on the remote server
                        $remoteProcessedPath = [WinSCP.RemotePath]::Combine($remotePath, "PROCESSADOS")
                        $remoteProcessedFile = [WinSCP.RemotePath]::Combine($remoteProcessedPath, $fileInfo.Name)
                        $session.MoveFile($remoteFilePath, $remoteProcessedFile) | Out-Null

                        $successLog += "Arquivo '$($fileInfo.Name)' movido com sucesso para Via 18.`n"
                    } else {
                        $errorLog += "Arquivo '$($fileInfo.Name)' não foi movido pois a pasta local 18 contém arquivos.`n"
                    }
                } elseif ($fileInfo.Name -match "_remessa_consolidada_55_out") {
                    if (IsFolderEmptyOrContainsOnlySubfolders $localPath55) {
                        $destinationFilePath55 = [System.IO.Path]::Combine($localPath55, $fileInfo.Name)
                        $session.GetFiles($remoteFilePath, $destinationFilePath55).Check() | Out-Null

                        # Move file to PROCESSADOS folder on the remote server
                        $remoteProcessedPath = [WinSCP.RemotePath]::Combine($remotePath, "PROCESSADOS")
                        $remoteProcessedFile = [WinSCP.RemotePath]::Combine($remoteProcessedPath, $fileInfo.Name)
                        $session.MoveFile($remoteFilePath, $remoteProcessedFile) | Out-Null

                        $successLog += "Arquivo '$($fileInfo.Name)' movido com sucesso para Via 55.`n"
                    } else {
                        $errorLog += "Arquivo '$($fileInfo.Name)' não foi movido pois a pasta local 55 contém arquivos.`n"
                    }
                } else {
                    $errorLog += "Arquivo '$($fileInfo.Name)' não corresponde aos padrões específicos. Não movido.`n"
                }
            } catch {
                $errorLog += "Erro ao processar arquivo '$($fileInfo.Name)': $($_.Exception.Message)`n"
            }
        }
    }
} catch {
    $errorLog += "Ocorreu um erro: $($_.Exception.Message)`n"
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
    $mailMessage.Subject = "ERRO EM PROCESSAMENTO DE CARGAS VIA"
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
    $mailMessage.Subject = "PROCESSAMENTO DE CARGAS VIA REALIZADO COM SUCESSO!"
    $mailMessage.Body = $successLog
    $mailMessage.IsBodyHtml = $false

    # Send E-mail
    $smtpClient.Send($mailMessage)

    # Session closed
    $mailMessage.Dispose()
    $smtpClient.Dispose()
}
