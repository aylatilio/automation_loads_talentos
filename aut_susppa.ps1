# suspcob_18_yyyyMMdd.txt and PA_18_yyyyMMdd.txt
# Este script move um único arquivo suspcoob ou PA da Via Varejo 18 ou 55 para os diretórios de importação do CRM, em seguida, move os arquivos para PROCESSADOS no SFTP.
# This script moves a single "suspcoob" or "PA" file from Via Varejo 18 or 55 to the CRM import directories, then moves the files to the "PROCESSADOS" folder on the SFTP.

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

# Error Logs and array variables
$errorLog = ""
$processedFiles = @()  # Array

# E-mail Config
$smtpServer = [System.Environment]::GetEnvironmentVariable("SMTP_SERVER")
$port = [System.Environment]::GetEnvironmentVariable("SMTP_PORT")
$userName = [System.Environment]::GetEnvironmentVariable("EMAIL_USER")
$password = [System.Environment]::GetEnvironmentVariable("EMAIL_PASS")
$to = [System.Environment]::GetEnvironmentVariable("EMAIL_TO")

# E-mail Function
function Send-Email {
    param (
        [string]$Subject,
        [string]$Body
    )

    $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $port)
    $smtpClient.EnableSsl = $true
    $smtpClient.Credentials = New-Object System.Net.NetworkCredential($userName, $password)

    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.From = $userName
    $mailMessage.To.Add($to)
    $mailMessage.Subject = $Subject 
    $mailMessage.Body = $Body

    $smtpClient.Send($mailMessage)
}

	# Function to check if the folder contains files (ignore subfolders)
function HasFiles($path) {
    Get-ChildItem -Path $path -File | Where-Object { $_.Name -notmatch "\.bak$" } | ForEach-Object { return $true }
    return $false
}

try {
    # SFTP Server Conection
    $session.Open($sessionOptions)
    Write-Host "Conexão SFTP estabelecida com sucesso."

    # Paths
    $remotePath = [System.Environment]::GetEnvironmentVariable("REMOTE_PATH_RC")
	$remoteProcessedPath = [System.Environment]::GetEnvironmentVariable("REMOTE_PROCESSED_PATH_RC")
    $localPath18 = [System.Environment]::GetEnvironmentVariable("LOCAL_PATH_18")
	$localPath55 = [System.Environment]::GetEnvironmentVariable("LOCAL_PATH_55")

    # List Files
    $remoteFiles = $session.EnumerateRemoteFiles($remotePath, "*", [WinSCP.EnumerationOptions]::None)
    $processedFilesRemote = $session.EnumerateRemoteFiles($remoteProcessedPath, "*", [WinSCP.EnumerationOptions]::None)

    # Get today and yesterday date
    $currentDateStr = (Get-Date).ToString("yyyyMMdd")
    $previousDateStr = (Get-Date).AddDays(-1).ToString("yyyyMMdd")

    # Verify remessa consolidada presence in the PROCESSADOS directory
    $remessaProcessed18 = $processedFilesRemote | Where-Object { $_.Name -match "^$currentDateStr|$previousDateStr.*_remessa_consolidada_18_out\.txt$" }
    $remessaProcessed55 = $processedFilesRemote | Where-Object { $_.Name -match "^$currentDateStr|$previousDateStr.*_remessa_consolidada_55_out\.txt$" }

    # Counter variables
    $canProcess18 = $remessaProcessed18.Count -gt 0
    $canProcess55 = $remessaProcessed55.Count -gt 0

    $notMovedDetails = @()
    $filesMoved = @()

    # 18 Base
    if ($canProcess18 -and -not (HasFiles $localPath18)) {
        $suspcob18 = $remoteFiles | Where-Object { $_.Name -match "^suspcob_18_$currentDateStr\.txt$" }
        $pa18 = $remoteFiles | Where-Object { $_.Name -match "^PA_18_$currentDateStr\.txt$" }

        if ($suspcob18) {
            Write-Host "Movendo $($suspcob18.Name) para $localPath18"
            $session.GetFiles($remotePath + "/" + $suspcob18.Name, $localPath18 + "\").Check()
            $session.MoveFile($remotePath + "/" + $suspcob18.Name, $remoteProcessedPath + "/" + $suspcob18.Name)
            $filesMoved += $suspcob18.Name
        } elseif ($pa18) {
            Write-Host "Movendo $($pa18.Name) para $localPath18"
            $session.GetFiles($remotePath + "/" + $pa18.Name, $localPath18 + "\").Check()
            $session.MoveFile($remotePath + "/" + $pa18.Name, $remoteProcessedPath + "/" + $pa18.Name)
            $filesMoved += $pa18.Name
        } else {
            $notMovedDetails += "Nenhum arquivo suspcoob ou PA encontrado para Via 18."
        }
    } else {
        $notMovedDetails += "Diretório $localPath18 já contém arquivos ou remessa consolidada ausente."
    }

    # 55 Base
    if ($canProcess55 -and -not (HasFiles $localPath55)) {
        $suspcob55 = $remoteFiles | Where-Object { $_.Name -match "^suspcob_55_$currentDateStr\.txt$" }
        $pa55 = $remoteFiles | Where-Object { $_.Name -match "^PA_55_$currentDateStr\.txt$" }

        if ($suspcob55) {
            Write-Host "Movendo $($suspcob55.Name) para $localPath55"
            $session.GetFiles($remotePath + "/" + $suspcob55.Name, $localPath55 + "\").Check()
            $session.MoveFile($remotePath + "/" + $suspcob55.Name, $remoteProcessedPath + "/" + $suspcob55.Name)
            $filesMoved += $suspcob55.Name
        } elseif ($pa55) {
            Write-Host "Movendo $($pa55.Name) para $localPath55"
            $session.GetFiles($remotePath + "/" + $pa55.Name, $localPath55 + "\").Check()
            $session.MoveFile($remotePath + "/" + $pa55.Name, $remoteProcessedPath + "/" + $pa55.Name)
            $filesMoved += $pa55.Name
        } else {
            $notMovedDetails += "Nenhum arquivo suspcoob ou PA encontrado para Via 55."
        }
    } else {
        $notMovedDetails += "Diretório $localPath55 já contém arquivos ou remessa consolidada ausente."
    }

    # E-mail send
    if ($filesMoved.Count -gt 0) {
        $subject = "PROCESSAMENTO DE BAIXAS VIA REALIZADO COM SUCESSO!"
        $body = "Os seguintes arquivos foram processados com sucesso:`n`n" + ($filesMoved -join "`n")
        Send-Email -Subject $subject -Body $body
    } else {
        $subject = "ERRO EM PROCESSAMENTO DE BAIXAS VIA"
        $body = "O processamento falhou pelos seguintes motivos:`n`n" + ($notMovedDetails -join "`n")
        Send-Email -Subject $subject -Body $body
    }
}
catch {
    Write-Error "Erro encontrado: $_"
}
finally {
    # SFTP session closed
    if ($session.Opened) {
        $session.Dispose()
    }
    Write-Host "Conexão SFTP fechada."
}