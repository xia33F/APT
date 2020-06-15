$Payload = Get-Content "./p.txt"

[System.Text.Encoding]::Unicode.GetBytes([Convert]::FromBase64String("$Payload"))