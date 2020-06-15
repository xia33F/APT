param (
    [string]$p = ""
)

$Payload = $p

[Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$Payload"))

#[Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes("$Payload"))
#[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Payload))