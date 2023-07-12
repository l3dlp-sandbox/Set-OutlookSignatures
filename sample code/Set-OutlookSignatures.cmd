REM Start Set-OutlookSignatures in a hidden, non-blocking window
REM A PowerShell windows will pop up for a second.
REM For a completely hidden method, read FAQ 'Start Set-OutlookSignatures in hidden/invisible mode' in 'README' file
REM Write a default PowerShell transcript file
start powershell.exe -WindowStyle hidden -Command "Start-Transcript; & '\\server.example.com\share\folder\Set-OutlookSignatures\set-outlooksignatures.ps1' -SignatureTemplatePath '\\server.example.com\share\folder\templates\signatures docx with ini' -SignatureIniPath '\\server.example.com\share\folder\signatures docx with ini\_.ini' -OOFTemplatePath '\\server.example.com\share\folder\templates\out of office docx' -OOFIniPath '\\server.example.com\share\folder\templates\out of office docx with ini\_.ini'; Stop-Transcript"
