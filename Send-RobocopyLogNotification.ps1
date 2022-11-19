# Programm/Skript: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
# Argumente: -file "C:\Backup-Script\Send-RobocopyLogNotification.ps1"
# Kennwort nicht speichern
# NICHT mit hoechsten Privilegien ausfuehren

#region variables
$logPath = 'C:\Backup-Script\Logs'
$logCount = 8 #number of logfiles written in interval
$logInterval = 1 #interval in days
$deletedPercentage = 0.05 #warn if more than $deletedPercentage files are deleted

$mailFrom = 'from@domain.com'
$mailTo = "recipient1@domainA.com", "recipient2@domainB.com"
$mailSubject = '[Customer] Fehler bei Dateibackup'
$mailSmtpServer = 'smtp.domain.com'
#if smtp auth required run next line once, remove # in line with Send-MailMessage
#Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File -FilePath $env:LOCALAPPDATA\$mailFrom".securestring"
#endregion variables


#region code
$localComputername = $env:COMPUTERNAME
if ($env:USERDNSDOMAIN) { $localComputername += "." + ($env:USERDNSDOMAIN) }


$problemFound = $false
$mailContent = ""

# check if LastWriteTime of the expected number of logfiles is within the interval
Get-ChildItem -Path $logPath\*.txt | Sort-Object -Property LastWriteTime | Select-Object -Last $logCount | `
	ForEach-Object { if ($_.LastWriteTime -lt (Get-Date).AddDays(-$logInterval)) { $problemFound = $true } }

if ($problemFound) {
    $mailContent += "Datei Backup NICHT erfolgreich auf $localComputername`n`n"
    $mailContent += "Mindestens ein Logfile ist nicht aktuell"
    Send-MailMessage -From $mailFrom -To $mailTo -Subject $mailSubject -Body $mailContent -SmtpServer $mailSmtpServer #-Port 587 -UseSsl -Credential (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $mailFrom,(Get-Content -Path $env:LOCALAPPDATA\$mailFrom".securestring" | ConvertTo-SecureString))
} else {
    # check for FAILED entries and lookup summary
    $Dateien = Get-ChildItem -Path $logPath\*.txt | Sort-Object -Property LastWriteTime | Select-Object -Last $logCount
	ForEach ($Datei in $Dateien) {
		$summaryFound = $false
        $summary = @()
        $failedFound = $false
        $failed = @()
        $failedTemp = ""

        Get-Content -Path $Datei | ForEach-Object {
            #check line for summary
			if ($_ -match "           Insgesamt   Kopiert*") { $summaryFound = $true }
            if ($summaryFound) { $summary += $_ }
            #check for FAILED lines with regex
            if ($_ -cmatch "^(\d{4}/\d{2}/\d{2}\s\d{2}:\d{2}:\d{2})\s(FEHLER\s.+)$") {
                $failedTemp = ""
                $failedTemp += $Matches[2]
                $failedTemp += "`n"
                $failedFound = $true
            #if FAILED line found next line has the error message
            } elseif ($failedFound) {
                $failedTemp += $_
                $failed += $failedTemp
                $failedFound = $false
            }
		}

        if ($summaryFound) {
            #get FAILED count and Extras count
            $summary | Select-Object -Skip 1 -First 2 | `
            ForEach-Object {
                #Text:  Insgesamt  Kopiert  Uebersprungen  Keine Uebereinstimmung  FEHLER  Extras
                $summaryFound = $_ -match "^*:\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)$"
                if ($summaryFound) {
                    #if FAILED count > 0 or Extras count / Total count * 100 > $deletedPercentage send notification
                    if (([convert]::ToInt32($matches[5], 10) -gt 0) -or ([convert]::ToInt32($matches[6], 10) / [convert]::ToInt32($matches[1], 10) * 100 -gt $deletedPercentage)) { $problemFound = $true }
                }
            }
        } else {
            #summary not found, send notification
            $problemFound = $true
        }

		if ($problemFound) {
            $problemFound = $false
            $mailContent += "Datei Backup NICHT erfolgreich auf $localComputername`n`n"
            $mailContent += "Protokolldatei: " + $Datei.Name + "`n`n"
            if (@($failed).Length -gt 0) {
                $mailContent += $failed | Select-Object -Unique | Out-String
                $mailContent += "`n"
            }
            if (@($summary).Length -gt 0) {
                $mailContent += $summary | Out-String
            } else {
                #summary not found
                $mailContent += "Zusammenfassung nicht gefunden`n`n"
            }
            $mailContent += "`n"
        }
	}

    if ($mailContent -ne "") {
    	#[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Send-MailMessage -From $mailFrom -To $mailTo -Subject $mailSubject -Body $mailContent -SmtpServer $mailSmtpServer #-Port 587 -UseSsl -Credential (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $mailFrom,(Get-Content -Path $env:LOCALAPPDATA\$mailFrom".securestring" | ConvertTo-SecureString))
    }
}
#endregion code
