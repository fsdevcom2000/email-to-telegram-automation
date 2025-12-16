param(
    [switch]$Fallback
)

# === Settings ===
$senderEmail   = "user@example.com"
$savePath      = "C:\Temp"
$botToken      = "your-token"
$chatId        = "your-chat-id"
$logPath       = "$savePath\auto_log.txt"

# === Create directory if it doesn't exist ===
if (-not (Test-Path $savePath)) {
    New-Item -ItemType Directory -Path $savePath | Out-Null
}

# === Enable TLS 1.2 for Telegram API ===
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# === Logging Function ===
function Write-Log {
    param ($message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $fullMessage = "$timestamp :: $message"
    Add-Content -Path $logPath -Value $fullMessage
    Write-Host $fullMessage
}

# === Telegram Send Function ===
function Send-Telegram {
    param(
        [string]$text = $null,
        [string]$filePath = $null
    )

    try {
        if ($filePath) {
            $client = New-Object System.Net.Http.HttpClient
            try {
                $multipart = New-Object System.Net.Http.MultipartFormDataContent
                $fileStream = [System.IO.File]::OpenRead($filePath)
                $fileContent = New-Object System.Net.Http.StreamContent($fileStream)
                $fileContent.Headers.Add("Content-Type","application/pdf")
                $fileName = [System.IO.Path]::GetFileName($filePath)
                $encodedFileName = [System.Uri]::EscapeDataString($fileName)
                $multipart.Add($fileContent,"document",$encodedFileName)
                $multipart.Add([System.Net.Http.StringContent]::new($chatId),"chat_id")
                $apiUrl = "https://api.telegram.org/bot$botToken/sendDocument"
                $response = $client.PostAsync($apiUrl,$multipart).Result
                if ($response.IsSuccessStatusCode) {
                    Write-Log "Sent to Telegram: $filePath"
                } else {
                    $errorBody = $response.Content.ReadAsStringAsync().Result
                    Write-Log "Telegram error: HTTP $($response.StatusCode) - $errorBody"
                }
            } finally {
                if ($fileStream) { $fileStream.Close() }
                if ($client) { $client.Dispose() }
            }
        } elseif ($text) {
            # Simple text message (fallback)
            Invoke-RestMethod -Uri "https://api.telegram.org/bot$botToken/sendMessage" `
                -Method Post `
                -Body @{ chat_id = $chatId; text = $text } `
                -ContentType "application/x-www-form-urlencoded"
            Write-Log "Sent Telegram text message."
        }
    } catch {
        Write-Log "Telegram send exception: $_"
    }
}

# === Fallback Mode ===
if ($Fallback) {
    $computerName = $env:COMPUTERNAME
    $userName     = $env:USERNAME
    $msg = "Microsoft Outlook is NOT running on computer '${computerName}', user '${userName}'. Email processing was skipped."
    Write-Log $msg
    Send-Telegram -text $msg
    exit
}

# === Main Script ===
try {
    Write-Log "Script started."

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6) # Inbox
    $messages = $inbox.Items | Where-Object {
        $_.UnRead -eq $true -and $_.SenderEmailAddress -eq $senderEmail
    }

    if ($messages.Count -eq 0) {
        Write-Log "No new emails from ${senderEmail}."
    }

    foreach ($message in $messages) {
        Write-Log "New email from ${senderEmail}: '$($message.Subject)'"

        if ($message.Attachments.Count -gt 0) {
            for ($i = 1; $i -le $message.Attachments.Count; $i++) {
                $attachment = $message.Attachments.Item($i)
                if ($attachment.FileName -like "*.pdf") {
                    $filePath = Join-Path $savePath $attachment.FileName
                    if (Test-Path $filePath) {
                        $filePath = Join-Path $savePath ("{0}_{1}" -f ([Guid]::NewGuid(), $attachment.FileName))
                    }
                    $attachment.SaveAsFile($filePath)
                    Write-Log "Saved file: $filePath"
                    Send-Telegram -filePath $filePath
                } else {
                    Write-Log "Skipped non-PDF attachment: $($attachment.FileName)"
                }
            }
        } else {
            Write-Log "Email has no attachments."
        }

        # Mark as read
        $message.UnRead = $false
        $message.Save()
        Write-Log "Marked email as read."
    }

    Write-Log "Script finished."
} catch {
    Write-Log "Script error: $_"
}
