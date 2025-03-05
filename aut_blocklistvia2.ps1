# BADLIST_VVAR.txt
# Este script move o arquivo BADLIST_VVAR.txt, obtido pelo script aut_blocklistvia1.ps1, para o diretório de carga do servidor, 
# verifica se já existem arquivos no diretório (que não sejam pastas) e envia um e-mail caso encontre erros.
# This script moves the BADLIST_VVAR.txt file, obtained by the aut_blocklistvia1.ps1 script, to the server's upload directory, 
# checks if there are any files (other than folders) in the directory, and sends an email if any errors are found.

# Sensitive Data .env
$envFile = "C:\Users\ayla.atilio\dev\aut\config.env"
if (Test-Path $envFile) {
    Get-Content $envFile | ForEach-Object {
        if ($_ -match "^\s*([^#].+?)=(.+)$") {
            [System.Environment]::SetEnvironmentVariable($matches[1], $matches[2], "Process")
        }
    }
}

# Paths
$localPath = [System.Environment]::GetEnvironmentVariable("LOCAL_PATH_B2")
$destinationPath = [System.Environment]::GetEnvironmentVariable("DESTINATION_PATH_B2")

# E-mail config
$smtpServer = [System.Environment]::GetEnvironmentVariable("SMTP_SERVER")
$port = [System.Environment]::GetEnvironmentVariable("SMTP_PORT")
$userName = [System.Environment]::GetEnvironmentVariable("EMAIL_USER")
$password = [System.Environment]::GetEnvironmentVariable("EMAIL_PASS")
$to = [System.Environment]::GetEnvironmentVariable("EMAIL_TO")

# E-mail Error Function
function Send-ErrorEmail($subject, $message) {
    $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $port)
    $smtpClient.EnableSsl = $true
    $smtpClient.Credentials = New-Object System.Net.NetworkCredential($userName, $password)

    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.From = $userName
    $mailMessage.To.Add($to)
    $mailMessage.Subject = $subject
    $mailMessage.Body = $message
    $mailMessage.IsBodyHtml = $false

    $smtpClient.Send($mailMessage)

    $mailMessage.Dispose()
    $smtpClient.Dispose()
}

try {
    # Files in destination directory verification
    $existingFiles = Get-ChildItem -Path $destinationPath -File -ErrorAction Stop

    if ($existingFiles.Count -gt 0) {
        # Enviar notificação e abortar operação
        $subject = "Erro: Arquivos existentes no diretório de destino"
        $message = "Há arquivos existentes no diretório '$destinationPath'. Operação de movimentação abortada."
        Send-ErrorEmail $subject $message
        Write-Output $message
        return
    }

    # .txt file integrity verification
    $txtFile = Get-ChildItem -Path $localPath -Filter "*.txt" | Select-Object -First 1
    if ($txtFile -eq $null) {
        throw "Nenhum arquivo .txt encontrado no diretório '$localPath'."
    }

    # .txt file moved
    Move-Item -Path $txtFile.FullName -Destination $destinationPath -ErrorAction Stop
    #Write-Output "Arquivo '${txtFile.Name}' movido com sucesso para '$destinationPath'." 

} catch {
    # E-mail error
    $subject = "Erro no processamento de Blocklist"
    $message = "Erro encontrado: $($_.Exception.Message)"
    Send-ErrorEmail $subject $message
    Write-Output $message
}
