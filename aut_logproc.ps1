# TALENTOS_ddMMyyyy.txt (Log de Processamento)
# Este script concatena o conteúdo do arquivo lognectar_VIA_VAREJO_SA_18.txt da Via Varejo com as linhas de layout 1, 6 e 7 do lognectar_VIA_VAREJO_SA_55.txt
# em um único arquivo TALENTOS_ddMMyyyy.txt e o move para o SFTP Cubo.
# This script concatenates the content of the lognectar_VIA_VAREJO_SA_18.txt file from Via Varejo with lines of layout 1, 6, and 7 from lognectar_VIA_VAREJO_SA_55.txt
# into a single file TALENTOS_ddMMyyyy.txt and moves it to the Cube SFTP.

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

# E-mail Config
$smtpServer = [System.Environment]::GetEnvironmentVariable("SMTP_SERVER")
$port = [System.Environment]::GetEnvironmentVariable("SMTP_PORT")
$userName = [System.Environment]::GetEnvironmentVariable("EMAIL_USER")
$password = [System.Environment]::GetEnvironmentVariable("EMAIL_PASS")
$to = [System.Environment]::GetEnvironmentVariable("EMAIL_TO")

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

# Caminhos locais e configuração dos arquivos
$localDir = [System.Environment]::GetEnvironmentVariable("LOCALDIR_LP")
$logFile18 = Join-Path $localDir "lognectar_VIA_VAREJO_SA_18.txt"
$logFile55 = Join-Path $localDir "lognectar_VIA_VAREJO_SA_55.txt"
$outputFileName = "TALENTOS_$(Get-Date -Format 'ddMMyyyy').txt"
$outputFilePath = Join-Path $localDir $outputFileName
$destinationDir = [System.Environment]::GetEnvironmentVariable("DESTINATIONDIR_LP")
$remoteDir = [System.Environment]::GetEnvironmentVariable("REMOTEDIR_LP")

try {
    # Existing Files Verification
    if (!(Test-Path $logFile18) -or !(Test-Path $logFile55)) {
        throw "Os arquivos necessários ($logFile18 ou $logFile55) não foram encontrados."
    }

    # Create output file
    Write-Host "Criando o arquivo $outputFileName..."
    Remove-Item -Path $outputFilePath -Force -ErrorAction SilentlyContinue
    New-Item -Path $outputFilePath -ItemType File -Force

    # lognectar_VIA_VAREJO_SA_18.txt content copy
    Get-Content $logFile18 | Out-File -FilePath $outputFilePath -Append

    # Filter and copy lines from lognectar_VIA_VAREJO_SA_55.txt
    Get-Content $logFile55 | ForEach-Object {
        $columns = $_ -split ";"
        if ($columns[5] -in "1", "6", "7") {
            $_ | Out-File -FilePath $outputFilePath -Append
        }
    }

    Write-Host "Conteúdo copiado para $outputFileName."

    # Delete original files
    Write-Host "Excluindo arquivos originais..."
    Remove-Item -Path $logFile18, $logFile55 -Force

    # SFTP conect and move file
    $session = New-Object WinSCP.Session
    $session.Open($sessionOptions)
    Write-Host "Conexão SFTP estabelecida."

    Write-Host "Movendo o arquivo para o servidor SFTP..."
    $transferOptions = New-Object WinSCP.TransferOptions
    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
    $session.PutFiles($outputFilePath, $remoteDir + "/", $false, $transferOptions).Check()

    # Move file to Local Path
    Write-Host "Movendo o arquivo para o diretório local: $destinationDir..."
    if (!(Test-Path $destinationDir)) {
        New-Item -Path $destinationDir -ItemType Directory -Force
    }
    Move-Item -Path $outputFilePath -Destination $destinationDir -Force

    Write-Host "Arquivo movido para o diretório local e SFTP com sucesso."

    # Success E-mail
    Send-Email -Subject "Processamento concluído com sucesso!" -Body "O arquivo $outputFileName foi gerado, transferido para o servidor SFTP e movido para o diretório local com sucesso."
    Write-Host "E-mail de confirmação enviado."
}
catch {
    Write-Error "Erro durante o processamento: $_"
    Send-Email -Subject "Erro no processamento" -Body "Ocorreu um erro durante o processamento: $_"
}
finally {
    # SFTP Session Closed
    if ($session -and $session.Opened) {
        $session.Dispose()
    }
    Write-Host "Processo finalizado."
}
