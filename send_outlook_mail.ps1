# Send a file through the desktop Outlook on this PC. Used by daily_report.py --email.
# With no -To, the message goes to the signed-in Outlook account's own address.
param(
    [string]$To = "",
    [Parameter(Mandatory = $true)][string]$Subject,
    [Parameter(Mandatory = $true)][string]$BodyFile,
    [Parameter(Mandatory = $true)][string]$Attachment
)
$ErrorActionPreference = "Stop"
$outlook = New-Object -ComObject Outlook.Application
if (-not $To) {
    $To = ($outlook.Session.Accounts | Select-Object -First 1).SmtpAddress
}
$mail = $outlook.CreateItem(0)
$mail.To = $To
$mail.Subject = $Subject
$mail.Body = Get-Content -Raw -Encoding UTF8 $BodyFile
[void]$mail.Attachments.Add((Resolve-Path $Attachment).Path)
$mail.Send()
Write-Output "Sent $Subject to $To"
