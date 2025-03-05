# E-mail Resumo de Processamento: este script verifica se todos os 7 layouts dos arquivos de carga da Via Varejo 18 e 55 foram processados 
# na base Nectar 1 e se os 8 Arquivos de Retorno com a data de modificação do dia atual, anteriores à 9pm, foram postados no SFTP.
# Email Processing's Summary: This script checks if all 7 layouts of the load files from Via Varejo 18 and 55 have been processed in the 
# Nectar 1 database and if the 8 Return Files, modified on the current day and before 9pm, have been uploaded to the SFTP.

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

# Paths Hash
$localPaths = @{
    "Carga 18" = [System.Environment]::GetEnvironmentVariable("CARGA18")
    "Arquivos PA 18" = [System.Environment]::GetEnvironmentVariable("PA18")
    "Arquivos SUSP 18" = [System.Environment]::GetEnvironmentVariable("SUSP18")
    "Bureau Email 18" = [System.Environment]::GetEnvironmentVariable("BUREAUEMAIL18")
    "Bureau Telefone 18" = [System.Environment]::GetEnvironmentVariable("BUREAUTEL18")
    "Carga 55" = [System.Environment]::GetEnvironmentVariable("CARGA55")
    "Arquivos PA 55" = [System.Environment]::GetEnvironmentVariable("PA55")
    "Arquivos SUSP 55" = [System.Environment]::GetEnvironmentVariable("SUSP55")
    "Bureau Email 55" = [System.Environment]::GetEnvironmentVariable("BUREAUEMAIL55")
    "Bureau Telefone 55" = [System.Environment]::GetEnvironmentVariable("BUREAUTEL55")
    "Blocklist" = [System.Environment]::GetEnvironmentVariable("BADLIST")
    "Remessa Estrategica" = [System.Environment]::GetEnvironmentVariable("REMESSA_ESTRATEGICA")
    "Arquivos Retorno" = [System.Environment]::GetEnvironmentVariable("ARQUIVOS_RETORNO")
}

# Hash table
$folderResults = @()

# Directory Loop
foreach ($key in $localPaths.Keys) {
    $pathOriginal = $localPaths[$key]
    $path = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::Default.GetBytes($pathOriginal))

    if ($key -in @("Remessa Estrategica", "Blocklist", "Carga 18", "Carga 55", "Bureau Email 18", "Bureau Email 55", "Bureau Telefone 18", "Bureau Telefone 55")) {
        # Check if the network drive path is mapped
        if (!(Test-Path $path)) {
            $folderResults += @{
                Path = $path
                Name = $key
                Result = "Erro: Unidade de rede nao encontrada"
            }
            continue
        }

        # Check if there are files in the directory and if there is a file with today's modification date
        $files = Get-ChildItem -Path $path | Where-Object { -not $_.PSIsContainer }
        $currentDateFileExists = $files | Where-Object { $_.LastWriteTime.Date -eq (Get-Date).Date }

        if ($files.Count -eq 0 -or !$currentDateFileExists) {
            $folderResults += @{
                Path = $path
                Name = $key
                Result = "Ausente"
            }
        } else {
            $folderResults += @{
                Path = $path
                Name = $key
                Result = "Sucesso"
            }
        }
    }

    elseif ($key -in @("Arquivos PA 18", "Arquivos SUSP 18", "Arquivos PA 55", "Arquivos SUSP 55")) {
        # Check if there are files in the directory with today's modification date and a specific name
        $files = Get-ChildItem -Path $path | Where-Object { -not $_.PSIsContainer }

        if ($key -eq "Arquivos PA 18") {
            $currentDateFileExists = $files | Where-Object { $_.LastWriteTime.Date -eq (Get-Date).Date -and $_.Name -like "*PA_18*" }
        } elseif ($key -eq "Arquivos SUSP 18") {
            $currentDateFileExists = $files | Where-Object { $_.LastWriteTime.Date -eq (Get-Date).Date -and $_.Name -like "*suspcob_18*" }
        } elseif ($key -eq "Arquivos PA 55") {
            $currentDateFileExists = $files | Where-Object { $_.LastWriteTime.Date -eq (Get-Date).Date -and $_.Name -like "*PA_55*" }
        } elseif ($key -eq "Arquivos SUSP 55") {
            $currentDateFileExists = $files | Where-Object { $_.LastWriteTime.Date -eq (Get-Date).Date -and $_.Name -like "*suspcob_55*" }
        }

        if ($files.Count -eq 0 -or !$currentDateFileExists) {
            $folderResults += @{
                Path = $path
                Name = $key
                Result = "Ausente"
            }
        } else {
            $folderResults += @{
                Path = $path
                Name = $key
                Result = "Sucesso"
            }
        }
    }

		elseif ($key -eq "Arquivos Retorno") {
		# Performs the validation via SFTP for "Return Files"
		$remotePath = $path
		$currentDate = (Get-Date).ToString("yyyyMMdd")

		# Defines the patterns for the expected files
		$expectedFiles = @(
			"${currentDate}_18_acio_in.txt",
			"${currentDate}_55_acio_in.txt",
			"${currentDate}_18_Promessa_Pagamento_in.txt",
			"${currentDate}_55_Promessa_Pagamento_in.txt",
			"${currentDate}_18_phone_in.txt",
			"${currentDate}_55_phone_in.txt",
			"${currentDate}_18_address_in.txt",
			"${currentDate}_55_address_in.txt"
		)

		try {
			if (-not (Test-Path $sessionOptions.SshPrivateKeyPath)) {
				throw "O caminho da chave privada não é válido: $($sessionOptions.SshPrivateKeyPath)"
			}

			# Connect to the SFTP server
			$session.Open($sessionOptions)

			# List the files in the remote directory
			$files = $session.ListDirectory($remotePath).Files | Where-Object { $_.Name -ne "." -and $_.Name -ne ".." }

			# Check if the expected files are present
			$missingFiles = $expectedFiles | Where-Object { -not ($files.Name -contains $_) }

			# Define the status based on the presence of the expected files
			if ($missingFiles.Count -eq 0) {
				$status = "Sucesso"
			} else {
				$status = "Ausente"
			}

			$folderResults += @{
				Path = $remotePath
				Name = "Arquivos Retorno"
				Result = $status
			}
		} catch {
			$folderResults += @{
				Path = $remotePath
				Name = "Arquivos Retorno"
				Result = "Erro: $_"
			}
		} finally {
			# WinSCP Session Closed
			$session.Dispose()
		}
	}

}

# Create HTML E-mail Body
$emailBody = @"
<html>
<head>
    <style>
        body {
            font-family: Arial, sans-serif;
            font-size: 14px;
            background-color: #f4f4f9;
            color: #333;
        }
        h2 {
            text-align: center;
            color: #2350e1;
        }
        h3 {
            color: #333;
        }
        .success {
            color: green;
            font-weight: bold;
        }
        .failure {
            color: red;
            font-weight: bold;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 12px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
            color: #555;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f1f1f1;
        }
        .status-success {
            color: #4CAF50;
            font-weight: bold;
        }
        .status-failure {
            color: #F44336;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <h2>TALENTOS - Resumo de Processamento</h2>
    <h3>Via Varejo</h3>
    <table>
        <tr>
            <th>Processo</th>
            <th>Status</th>
        </tr>
"@

$order = @("Remessa Estrategica", "Blocklist", "Carga 18", "Carga 55", "Arquivos PA 18", "Arquivos PA 55", "Arquivos SUSP 18", "Arquivos SUSP 55", "Bureau Email 18", "Bureau Email 55", "Bureau Telefone 18", "Bureau Telefone 55", "Arquivos Retorno")

# Add the results information to the email body
foreach ($process in $order) {
    $result = $folderResults | Where-Object { $_.Name -eq $process }
    if ($result) {
        $status = $result.Result
        $statusClass = if ($status -eq "Sucesso") { "status-success" } else { "status-failure" }
        $emailBody += "<tr>
            <td>$process</td>
            <td class='$statusClass'>$status</td>
        </tr>"
    }
}

$emailBody += @"
    </table>
</body>
</html>
"@

# SMTP Server Config
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
$mailMessage.Subject = "Resumo de Processamento Via Varejo"
$mailMessage.Body = $emailBody
$mailMessage.IsBodyHtml = $true

# Send E-mail
$smtpClient.Send($mailMessage)

# Release the resources
$mailMessage.Dispose()
$smtpClient.Dispose()