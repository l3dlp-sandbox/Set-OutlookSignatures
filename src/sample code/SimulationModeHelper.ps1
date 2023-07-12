Write-Host 'Set-OutlookSignatures simulation mode helper'


# Admin part
## Define parameters to use for Set-OutlookSignatures
## It is sufficient to only list parameters where the default value should not be used
$params = [ordered]@{
	SignatureTemplatePath         = '.\sample templates\Signatures DOCX'
	SignatureIniPath              = '.\sample templates\Signatures DOCX\_Signatures.ini'
	OOFTemplatePath               = '.\sample templates\Out of Office DOCX'
	OOFIniPath                    = '.\sample templates\Out of Office DOCX\_OOF.ini'
	ReplacementVariableConfigFile = '.\config\default replacement variables.ps1'
	GraphConfigFile               = '.\config\default graph config.ps1'
	GraphOnly                     = $false
}


## User part

Write-Host
Write-Host '  Please enter the login name of the user to simulate'
Write-Host '    Allowed formats:'
Write-Host '      user.x@example.com (UPN, a.k.a. User Principal Name)'
Write-Host '      EXAMPLE\User (NT4 style NetBIOS domain name and logon name)'

do {
	$tempSimulateUser = Read-Host '    Your input'
} until (
	$(
		$tempSimulateUser = $tempSimulateUser.trim()
		if ($tempSimulateUser -match '^\S+@\S+$|^\S+\\\S+$') {
			Write-Host "      Simulate user: $($tempSimulateUser)"
			$params['SimulateUser'] = $tempSimulateUser
			$true
		} else {
			Write-Host '      Wrong format. Please try again.' -ForegroundColor yellow
		}
	)
)


Write-Host
Write-Host '  Please enter the e-mail addresses of the mailboxes to simulate'
Write-Host '    Separate multiple mailboxes by spaces, commas or semicolons'
Write-Host '      Example: user.x@domain.com, user.a@domain.com, sharedmailbox.y@domain.com'

do {
	$tempSimulateMailboxes = Read-Host '    Your input'
} until (
	$(
		try {
			[mailaddress[]] $tempSimulateMailboxes = @(@(($tempSimulateMailboxes -replace '\s+', ',' -replace ';+', ',' -replace ',+', ',') -split ',') | Where-Object { $_ })
			Write-Host "      Simulate mailboxes: $($tempSimulateMailboxes -join ', ')"
			$params['SimulateMailboxes'] = $tempSimulateMailboxes
			$true
		} catch {
			Write-Host '      Wrong format. Please try again.' -ForegroundColor yellow
			$false
		}
	)
)




Write-Host
Write-Host '  Please enter the time use for simulation mode'
Write-Host '    Keep blank to use current date and time'
Write-Host '    Input must be in the international format yyyyMMddHHmm'
Write-Host '      yyyy = year in 4 digits'
Write-Host '      MM = month in 2 digits'
Write-Host '      dd = day in 2 digits'
Write-Host '      HH = hour in 2 digits, using 24-hour-format'
Write-Host '      mm = minute in 2 digits'
Write-Host '    Examples:'
Write-Host "      202303152249 is March 15th 2023 at 22:49 o'clock (22:49 is 10:49 p.m.)"

do {
	$tempTime = Read-Host '    Your input'
} until (
	$(
		if ($tempTime) {
			try {
				[DateTime]::ParseExact($tempTime, 'yyyyMMddHHmm', $null)
				Write-Host "      In local time: $([DateTime]::ParseExact($tempTime, 'yyyyMMddHHmm', $null))"
				$params['SimulateTime'] = $tempTime
				$true
			} catch {
				Write-Host '      Time is not valid. Please try again.' -ForegroundColor yellow
				$false
			}
		} else {
			$true
		}
	)
)


Write-Host
Write-Host '  Please enter the file path to use for simulation mode'
Write-Host '    The folder must already exist'
Write-Host '      Examples:'
Write-Host '        c:\users\userx\documents\Set-OutlookSignatures simulation folder'
Write-Host '        \\server\share\folder'
Write-Host '        https://server.example.com/site/library/folder'

do {
	$tempPath = Read-Host '    Your input'
} until (
	$(
		try {
			$tempPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($tempPath)
			if (Test-Path $tempPath) {
				Write-Host "      Path: $($tempPath)"
				$params['AdditionalSignaturePath'] = $tempPath
				$true
			} else {
				throw
			}
		} catch {
			Write-Host '      Folder does not exist. Please try again.' -ForegroundColor yellow
			$false
		}
	)
)


$paramsString = @(
	$params.GetEnumerator() | Where-Object { $_.Value } | ForEach-Object {
		if ($_.Value -is [array]) {
			"-$($_.Name) '$($_.Value -join "', '")'"
		} elseif ($_.value.tostring().startswith("'") -or $_.value.tostring().startswith('"')) {
			"-$($_.Name) $($_.Value)"
		} else {
			"-$($_.Name) '$($_.Value)'"
		}
	}
) -join ' '


Write-Host
Write-Host 'Resulting commands'
Write-Host '  For starting Set-OutlookSignatures directly from within this script'
Write-Host '    & ..\Set-OutlookSignatures.ps1 @params'
Write-Host '  For the command line'
Write-Host "    powershell.exe -command `"& ..\Set-OutlookSignatures.ps1 $($paramsString)`""


Write-Host
Write-Host 'Thank you for using Set-OutlookSignatures!'

# SIG # Begin signature block
# MIIwTwYJKoZIhvcNAQcCoIIwQDCCMDwCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJZwjk9C3J+F82
# kxZ5eyMWJUwexNgKDttid6sJXtPy36CCFCkwggWQMIIDeKADAgECAhAFmxtXno4h
# MuI5B72nd3VcMA0GCSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNV
# BAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0xMzA4MDExMjAwMDBaFw0z
# ODAxMTUxMjAwMDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# AL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/z
# G6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZ
# anMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7s
# Wxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL
# 2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfb
# BHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3
# JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3c
# AORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqx
# YxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0
# viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aL
# T8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjQjBAMA8GA1Ud
# EwEB/wQFMAMBAf8wDgYDVR0PAQH/BAQDAgGGMB0GA1UdDgQWBBTs1+OC0nFdZEzf
# Lmc/57qYrhwPTzANBgkqhkiG9w0BAQwFAAOCAgEAu2HZfalsvhfEkRvDoaIAjeNk
# aA9Wz3eucPn9mkqZucl4XAwMX+TmFClWCzZJXURj4K2clhhmGyMNPXnpbWvWVPjS
# PMFDQK4dUPVS/JA7u5iZaWvHwaeoaKQn3J35J64whbn2Z006Po9ZOSJTROvIXQPK
# 7VB6fWIhCoDIc2bRoAVgX+iltKevqPdtNZx8WorWojiZ83iL9E3SIAveBO6Mm0eB
# cg3AFDLvMFkuruBx8lbkapdvklBtlo1oepqyNhR6BvIkuQkRUNcIsbiJeoQjYUIp
# 5aPNoiBB19GcZNnqJqGLFNdMGbJQQXE9P01wI4YMStyB0swylIQNCAmXHE/A7msg
# dDDS4Dk0EIUhFQEI6FUy3nFJ2SgXUE3mvk3RdazQyvtBuEOlqtPDBURPLDab4vri
# RbgjU2wGb2dVf0a1TD9uKFp5JtKkqGKX0h7i7UqLvBv9R0oN32dmfrJbQdA75PQ7
# 9ARj6e/CVABRoIoqyc54zNXqhwQYs86vSYiv85KZtrPmYQ/ShQDnUBrkG5WdGaG5
# nLGbsQAe79APT0JsyQq87kP6OnGlyE0mpTX9iV28hWIdMtKgK1TtmlfB2/oQzxm3
# i0objwG2J5VT6LaJbVu8aNQj6ItRolb58KaAoNYes7wPD1N1KarqE3fk3oyBIa0H
# EEcRrYc9B9F1vM/zZn4wggawMIIEmKADAgECAhAIrUCyYNKcTJ9ezam9k67ZMA0G
# CSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0zNjA0MjgyMzU5NTla
# MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UE
# AxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEz
# ODQgMjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDVtC9C
# 0CiteLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0JAfhS0/TeEP0F9ce
# 2vnS1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJrQ5qZ8sU7H/Lvy0da
# E6ZMswEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhFLqGfLOEYwhrMxe6T
# SXBCMo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+FLEikVoQ11vkunKoA
# FdE3/hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh3K3kGKDYwSNHR7Oh
# D26jq22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJwZPt4bRc4G/rJvmM
# 1bL5OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQayg9Rc9hUZTO1i4F4z
# 8ujo7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbIYViY9XwCFjyDKK05
# huzUtw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchApQfDVxW0mdmgRQRNY
# mtwmKwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRroOBl8ZhzNeDhFMJlP
# /2NPTLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IBWTCCAVUwEgYDVR0T
# AQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHwYD
# VR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNV
# HR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRU
# cnVzdGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATAN
# BgkqhkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql+Eg08yy25nRm95Ry
# sQDKr2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFFUP2cvbaF4HZ+N3HL
# IvdaqpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1hmYFW9snjdufE5Btf
# Q/g+lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3RywYFzzDaju4ImhvTnh
# OE7abrs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5UbdldAhQfQDN8A+KVssIh
# dXNSy0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw8MzK7/0pNVwfiThV
# 9zeKiwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnPLqR0kq3bPKSchh/j
# wVYbKyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatEQOON8BUozu3xGFYH
# Ki8QxAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bnKD+sEq6lLyJsQfmC
# XBVmzGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQjiWQ1tygVQK+pKHJ6l
# /aCnHwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbqyK+p/pQd52MbOoZW
# eE4wggfdMIIFxaADAgECAhAKaypbp7cyIFa+lR7OVPAvMA0GCSqGSIb3DQEBCwUA
# MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UE
# AxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEz
# ODQgMjAyMSBDQTEwHhcNMjMwNzExMDAwMDAwWhcNMjYwNzEwMjM1OTU5WjCB5TET
# MBEGCysGAQQBgjc8AgEDEwJBVDEVMBMGCysGAQQBgjc8AgECEwRXaWVuMRUwEwYL
# KwYBBAGCNzwCAQETBFdpZW4xHTAbBgNVBA8MFFByaXZhdGUgT3JnYW5pemF0aW9u
# MRAwDgYDVQQFEwc2MDcwMTN0MQswCQYDVQQGEwJBVDENMAsGA1UECBMEV2llbjEN
# MAsGA1UEBxMEV2llbjEhMB8GA1UEChMYRXhwbGljSVQgQ29uc3VsdGluZyBHbWJI
# MSEwHwYDVQQDExhFeHBsaWNJVCBDb25zdWx0aW5nIEdtYkgwggIiMA0GCSqGSIb3
# DQEBAQUAA4ICDwAwggIKAoICAQDxdNfDY8ulBB2NIOYzd2mVQRhjMBAzNgvJEjXs
# VACQyjesfJfvXZ3gMnUT8M5HkohWjHvhftCFkL5cCck+4XuEGiLisV3hilLL4p8z
# 6L+tbvPnVSWML7VOV835/de+hM/mKdFhqRG+fYNQ1ceFlggiwqfHjIoXLweZACRD
# 3bLwRLYk7w5IEDCtHa0Hit+SpqbZ4MDcEhfS8krG5ha0FqOLkVLAhPfkZ4sOB32V
# dUfQPknxYnhWZVyGVH/ypTYnEY4oo3CFO0f8k4fNc8fGDwNAoxHJwGKYjxeEasgm
# a2EZMHKkZyJpwJKSdZ9FPp4qldYVt/NiCoXzdrLRta0M/Vg5E+XKVtC0EOhY2w6u
# lgFx0Qog/hfC3w2imATDt7Fv5R+ZQ8v3BXzn2pH2DZ1sGI7JZjH0NCxXdY8kaDuZ
# fCQRcDCej/5otpuDxu7l6bBUTBe2ao+ZwCBuN0PWdbyxunii1W/Q3t1bU2Hmu/97
# 4hQOWJNDBuWrPNOlr2qHVqFNCOpHtuddTHMGt9bGwr9FXXe5gTIrAk2CCX+vnDhw
# zgi8UuLWJy+H1b1Y2hUt2oX2izyAjDrXdA6wgGNr3YtIgUt+4BBRz0Zhw6/KQdpN
# wCTnofcgezhz0OS4WMB+ZARaMNK4DpzVwlGrg9NF/nCuQ0sJzt913ndIRl5FXJ71
# GwgCKwIDAQABo4ICAjCCAf4wHwYDVR0jBBgwFoAUaDfg67Y7+F8Rhvv+YXsIiGX0
# TkIwHQYDVR0OBBYEFDz+YL61H9M50y8W+urzdKxOSpf4MA4GA1UdDwEB/wQEAwIH
# gDATBgNVHSUEDDAKBggrBgEFBQcDAzCBtQYDVR0fBIGtMIGqMFOgUaBPhk1odHRw
# Oi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRDb2RlU2lnbmlu
# Z1JTQTQwOTZTSEEzODQyMDIxQ0ExLmNybDBToFGgT4ZNaHR0cDovL2NybDQuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29kZVNpZ25pbmdSU0E0MDk2U0hB
# Mzg0MjAyMUNBMS5jcmwwPQYDVR0gBDYwNDAyBgVngQwBAzApMCcGCCsGAQUFBwIB
# FhtodHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwgZQGCCsGAQUFBwEBBIGHMIGE
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXAYIKwYBBQUH
# MAKGUGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRH
# NENvZGVTaWduaW5nUlNBNDA5NlNIQTM4NDIwMjFDQTEuY3J0MAkGA1UdEwQCMAAw
# DQYJKoZIhvcNAQELBQADggIBAISWy98G7WUbOBA3S0odwfltQ3YZmuNgNZDoIdLQ
# YFnB43wgnClFuPIPaKJGYeRH90iioYKsnGDOYvUgr+b+XbIDRRqkHoYYZB+jDYUJ
# f1LS6eD79GAsLEomY/VzyRY9LEbYsmDmHi/riDWDiKWL0YYQmVuxU6NSLz4JZADA
# VsC7bZovRJnL9XFQo0QQxz9jymHH1UVBOAUUojrs7IznXBtQza/PYg+285kCoR/U
# ToA+Bc7j/mwon0tKlNCKyPn04viwjHRSIr8VlCH+qXU+nw6eSH7PVJWargv2sX/h
# t9zJ4JK843KRtd2mEXMUVcS2AUnmuwBSrxXhFQguR5nfrZBUHb4epiAMreGfidEl
# bmxEpzLaBegF8A+C7mCambjhnQ1p9b6JKuV1aS9qyfRf6AYF+OKLzBBbIAKLOmSx
# aHoJdn65B50/Gq5zUIxkoa8lKjEw4xtIBto4xYnFOLQJmiNeyAJeRLHbPGpHm6M+
# tTorAVDdGPQbhDlQT2RHn9pJDiJxFIbPdsNoEgtzAQee5US4QCng1qySpsvhQEoX
# JHh3jq62djlgx2GmVGOsysBfhcqjJROeo0+B32YQRHST/RBEaesZ6SFfXGaO3bBt
# onaU0JOQ9LOioHOuhGVNPjrcKT/NE99Bs2JF1Z8XJfPcDt5R0c10eRY1fiLJLvU5
# GNmrMYIbfDCCG3gCAQEwfTBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgQ29kZSBTaWdu
# aW5nIFJTQTQwOTYgU0hBMzg0IDIwMjEgQ0ExAhAKaypbp7cyIFa+lR7OVPAvMA0G
# CWCGSAFlAwQCAQUAoIIBjjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgaF9dYW1B
# oSj8CvMxnh9u3mght+wlXbrCciJdBbCiJR8wggEgBgorBgEEAYI3AgEMMYIBEDCC
# AQyggcWAgcIAUABvAHcAZQByAGUAZAAgAGIAeQAgAEUAeABwAGwAaQBjAEkAVAAg
# AEMAbwBuAHMAdQBsAHQAaQBuAGcALgAgAFUAbgBsAG8AYwBrACAAYQBsAGwAIABm
# AGUAYQB0AHUAcgBlAHMAIAB3AGkAdABoACAAUwBlAHQALQBPAHUAdABsAG8AbwBr
# AFMAaQBnAG4AYQB0AHUAcgBlAHMAIABCAGUAbgBlAGYAYQBjAHQAbwByACAAQwBp
# AHIAYwBsAGUALqFCgEBodHRwczovL2V4cGxpY2l0Y29uc3VsdGluZy5hdC9vcGVu
# LXNvdXJjZS9zZXQtb3V0bG9va3NpZ25hdHVyZXMgMA0GCSqGSIb3DQEBAQUABIIC
# ALy1bJGXRzv0MVxcdKdJnoJn+GA2wS+IssrmmgOPR+YXURxhrCn+2vE/N9PrlquT
# hAf1DkCKVSc9nrdQaGVp4p/7FBX3tdOWw3Odtq3XAcfPwqvDuy7To0FEIZNG/9j+
# SHoo8UA8+nnwtShQpLV4133Gh4c0Pg1K2JugnOpEVOcdnrOrD3r2JrhWTTjTHsv/
# fOB0JNPf3BqN31iGjmdSFgfft6I+Abwvz9UlNeqeHKMJR1yBmujtTRPdJQGx/H4v
# 9bW3y9YTnf80zm4lrRpHjvxAMDrfO+1vDaNbjpyZz+qFsZgvD1iyZ3fDYlEk9ZW3
# +7Rh7Rw5wZSeY/804Ef3e+qQUTnsZpSLYq5vnPIpJGhxhJwZK/B+KDHa+cA/mE9S
# fOx4DoqsJUBPxddnIAN/MqM1/Hp1IEDppi4+2DHf2ftAWp5hwe1vEF62OoBxi/yS
# b+QZvolglHjykxsG700p8a42sZUPNGtWillqVSTg3GCIVHyNsOI0VUYyrjWIDCQW
# q9+ytTOsDGTEZa/IjIJFkWT1CyEk8B8BXWtvOSa6nXc/BdcYw8aIAY9EHoKhzsWy
# p5H7y7IXJy/asNUyJjMO493kEXUjXNs9EDJkwtMY44PNnI03y3ViExYwwiEzd51u
# OUVWb5YY1Esvb2uWosmW2h6zsj6hZECf9bTKVRbSloXxoYIXPjCCFzoGCisGAQQB
# gjcDAwExghcqMIIXJgYJKoZIhvcNAQcCoIIXFzCCFxMCAQMxDzANBglghkgBZQME
# AgEFADB4BgsqhkiG9w0BCRABBKBpBGcwZQIBAQYJYIZIAYb9bAcBMDEwDQYJYIZI
# AWUDBAIBBQAEIHBXw83qz50+gMci+cnXN1nZPJw/wVbGF3UX3337lKNPAhEA4VNb
# IY+JnSchK2H1cOy13hgPMjAyMzA3MTIxNzA5NTVaoIITBzCCBsAwggSooAMCAQIC
# EAxNaXJLlPo8Kko9KQeAPVowDQYJKoZIhvcNAQELBQAwYzELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVz
# dGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTAeFw0yMjA5MjEw
# MDAwMDBaFw0zMzExMjEyMzU5NTlaMEYxCzAJBgNVBAYTAlVTMREwDwYDVQQKEwhE
# aWdpQ2VydDEkMCIGA1UEAxMbRGlnaUNlcnQgVGltZXN0YW1wIDIwMjIgLSAyMIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAz+ylJjrGqfJru43BDZrboegU
# hXQzGias0BxVHh42bbySVQxh9J0Jdz0Vlggva2Sk/QaDFteRkjgcMQKW+3KxlzpV
# rzPsYYrppijbkGNcvYlT4DotjIdCriak5Lt4eLl6FuFWxsC6ZFO7KhbnUEi7iGkM
# iMbxvuAvfTuxylONQIMe58tySSgeTIAehVbnhe3yYbyqOgd99qtu5Wbd4lz1L+2N
# 1E2VhGjjgMtqedHSEJFGKes+JvK0jM1MuWbIu6pQOA3ljJRdGVq/9XtAbm8WqJqc
# lUeGhXk+DF5mjBoKJL6cqtKctvdPbnjEKD+jHA9QBje6CNk1prUe2nhYHTno+EyR
# EJZ+TeHdwq2lfvgtGx/sK0YYoxn2Off1wU9xLokDEaJLu5i/+k/kezbvBkTkVf82
# 6uV8MefzwlLE5hZ7Wn6lJXPbwGqZIS1j5Vn1TS+QHye30qsU5Thmh1EIa/tTQznQ
# ZPpWz+D0CuYUbWR4u5j9lMNzIfMvwi4g14Gs0/EH1OG92V1LbjGUKYvmQaRllMBY
# 5eUuKZCmt2Fk+tkgbBhRYLqmgQ8JJVPxvzvpqwcOagc5YhnJ1oV/E9mNec9ixezh
# e7nMZxMHmsF47caIyLBuMnnHC1mDjcbu9Sx8e47LZInxscS451NeX1XSfRkpWQNO
# +l3qRXMchH7XzuLUOncCAwEAAaOCAYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNV
# HRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATAfBgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGog
# j57IbzAdBgNVHQ4EFgQUYore0GH8jzEU7ZcLzT0qlBTfUpwwWgYDVR0fBFMwUTBP
# oE2gS4ZJaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0
# UlNBNDA5NlNIQTI1NlRpbWVTdGFtcGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMw
# gYAwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEF
# BQcwAoZMaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3Rl
# ZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsF
# AAOCAgEAVaoqGvNG83hXNzD8deNP1oUj8fz5lTmbJeb3coqYw3fUZPwV+zbCSVEs
# eIhjVQlGOQD8adTKmyn7oz/AyQCbEx2wmIncePLNfIXNU52vYuJhZqMUKkWHSphC
# K1D8G7WeCDAJ+uQt1wmJefkJ5ojOfRu4aqKbwVNgCeijuJ3XrR8cuOyYQfD2DoD7
# 5P/fnRCn6wC6X0qPGjpStOq/CUkVNTZZmg9U0rIbf35eCa12VIp0bcrSBWcrduv/
# mLImlTgZiEQU5QpZomvnIj5EIdI/HMCb7XxIstiSDJFPPGaUr10CU+ue4p7k0x+G
# AWScAMLpWnR1DT3heYi/HAGXyRkjgNc2Wl+WFrFjDMZGQDvOXTXUWT5Dmhiuw8nL
# w/ubE19qtcfg8wXDWd8nYiveQclTuf80EGf2JjKYe/5cQpSBlIKdrAqLxksVStOY
# kEVgM4DgI974A6T2RUflzrgDQkfoQTZxd639ouiXdE4u2h4djFrIHprVwvDGIqhP
# m73YHJpRxC+a9l+nJ5e6li6FV8Bg53hWf2rvwpWaSxECyIKcyRoFfLpxtU56mWz0
# 6J7UWpjIn7+NuxhcQ/XQKujiYu54BNu90ftbCqhwfvCXhHjjCANdRyxjqCU4lwHS
# Pzra5eX25pvcfizM/xdMTQCi2NYBDriL7ubgclWJLCcZYfZ3AYwwggauMIIElqAD
# AgECAhAHNje3JFR82Ees/ShmKl5bMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2Vy
# dC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yMjAz
# MjMwMDAwMDBaFw0zNzAzMjIyMzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQK
# Ew5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBS
# U0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQDGhjUGSbPBPXJJUVXHJQPE8pE3qZdRodbSg9GeTKJtoLDM
# g/la9hGhRBVCX6SI82j6ffOciQt/nR+eDzMfUBMLJnOWbfhXqAJ9/UO0hNoR8XOx
# s+4rgISKIhjf69o9xBd/qxkrPkLcZ47qUT3w1lbU5ygt69OxtXXnHwZljZQp09ns
# ad/ZkIdGAHvbREGJ3HxqV3rwN3mfXazL6IRktFLydkf3YYMZ3V+0VAshaG43IbtA
# rF+y3kp9zvU5EmfvDqVjbOSmxR3NNg1c1eYbqMFkdECnwHLFuk4fsbVYTXn+149z
# k6wsOeKlSNbwsDETqVcplicu9Yemj052FVUmcJgmf6AaRyBD40NjgHt1biclkJg6
# OBGz9vae5jtb7IHeIhTZgirHkr+g3uM+onP65x9abJTyUpURK1h0QCirc0PO30qh
# HGs4xSnzyqqWc0Jon7ZGs506o9UD4L/wojzKQtwYSH8UNM/STKvvmz3+DrhkKvp1
# KCRB7UK/BZxmSVJQ9FHzNklNiyDSLFc1eSuo80VgvCONWPfcYd6T/jnA+bIwpUzX
# 6ZhKWD7TA4j+s4/TXkt2ElGTyYwMO1uKIqjBJgj5FBASA31fI7tk42PgpuE+9sJ0
# sj8eCXbsq11GdeJgo1gJASgADoRU7s7pXcheMBK9Rp6103a50g5rmQzSM7TNsQID
# AQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUuhbZbU2F
# L3MpdpovdYxqII+eyG8wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08w
# DgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEB
# BGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsG
# AQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVz
# dGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAH1ZjsCTtm+Y
# qUQiAX5m1tghQuGwGC4QTRPPMFPOvxj7x1Bd4ksp+3CKDaopafxpwc8dB+k+YMjY
# C+VcW9dth/qEICU0MWfNthKWb8RQTGIdDAiCqBa9qVbPFXONASIlzpVpP0d3+3J0
# FNf/q0+KLHqrhc1DX+1gtqpPkWaeLJ7giqzl/Yy8ZCaHbJK9nXzQcAp876i8dU+6
# WvepELJd6f8oVInw1YpxdmXazPByoyP6wCeCRK6ZJxurJB4mwbfeKuv2nrF5mYGj
# VoarCkXJ38SNoOeY+/umnXKvxMfBwWpx2cYTgAnEtp/Nh4cku0+jSbl3ZpHxcpzp
# SwJSpzd+k1OsOx0ISQ+UzTl63f8lY5knLD0/a6fxZsNBzU+2QJshIUDQtxMkzdwd
# eDrknq3lNHGS1yZr5Dhzq6YBT70/O3itTK37xJV77QpfMzmHQXh6OOmc4d0j/R0o
# 08f56PGYX/sr2H7yRp11LB4nLCbbbxV7HhmLNriT1ObyF5lZynDwN7+YAN8gFk8n
# +2BnFqFmut1VwDophrCYoCvtlUG3OtUVmDG0YgkPCr2B2RP+v6TR81fZvAT6gt4y
# 3wSJ8ADNXcL50CN/AAvkdgIm2fBldkKmKYcJRyvmfxqkhQ/8mJb2VVQrH4D6wPIO
# K+XW+6kvRBVK5xMOHds3OBqhK/bt1nz8MIIFjTCCBHWgAwIBAgIQDpsYjvnQLefv
# 21DiCEAYWjANBgkqhkiG9w0BAQwFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQD
# ExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMjIwODAxMDAwMDAwWhcN
# MzExMTA5MjM1OTU5WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2Vy
# dCBUcnVzdGVkIFJvb3QgRzQwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQC/5pBzaN675F1KPDAiMGkz7MKnJS7JIT3yithZwuEppz1Yq3aaza57G4QNxDAf
# 8xukOBbrVsaXbR2rsnnyyhHS5F/WBTxSD1Ifxp4VpX6+n6lXFllVcq9ok3DCsrp1
# mWpzMpTREEQQLt+C8weE5nQ7bXHiLQwb7iDVySAdYyktzuxeTsiT+CFhmzTrBcZe
# 7FsavOvJz82sNEBfsXpm7nfISKhmV1efVFiODCu3T6cw2Vbuyntd463JT17lNecx
# y9qTXtyOj4DatpGYQJB5w3jHtrHEtWoYOAMQjdjUN6QuBX2I9YI+EJFwq1WCQTLX
# 2wRzKm6RAXwhTNS8rhsDdV14Ztk6MUSaM0C/CNdaSaTC5qmgZ92kJ7yhTzm1EVgX
# 9yRcRo9k98FpiHaYdj1ZXUJ2h4mXaXpI8OCiEhtmmnTK3kse5w5jrubU75KSOp49
# 3ADkRSWJtppEGSt+wJS00mFt6zPZxd9LBADMfRyVw4/3IbKyEbe7f/LVjHAsQWCq
# sWMYRJUadmJ+9oCw++hkpjPRiQfhvbfmQ6QYuKZ3AeEPlAwhHbJUKSWJbOUOUlFH
# dL4mrLZBdd56rF+NP8m800ERElvlEFDrMcXKchYiCd98THU/Y+whX8QgUWtvsauG
# i0/C1kVfnSD8oR7FwI+isX4KJpn15GkvmB0t9dmpsh3lGwIDAQABo4IBOjCCATYw
# DwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU7NfjgtJxXWRM3y5nP+e6mK4cD08w
# HwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDgYDVR0PAQH/BAQDAgGG
# MHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNl
# cnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MEUGA1UdHwQ+MDwwOqA4oDaGNGh0
# dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5j
# cmwwEQYDVR0gBAowCDAGBgRVHSAAMA0GCSqGSIb3DQEBDAUAA4IBAQBwoL9DXFXn
# OF+go3QbPbYW1/e/Vwe9mqyhhyzshV6pGrsi+IcaaVQi7aSId229GhT0E0p6Ly23
# OO/0/4C5+KH38nLeJLxSA8hO0Cre+i1Wz/n096wwepqLsl7Uz9FDRJtDIeuWcqFI
# tJnLnU+nBgMTdydE1Od/6Fmo8L8vC6bp8jQ87PcDx4eo0kxAGTVGamlUsLihVo7s
# pNU96LHc/RzY9HdaXFSMb++hUD38dglohJ9vytsgjTVgHAIDyyCwrFigDkBjxZgi
# wbJZ9VVrzyerbHbObyMt9H5xaiNrIv8SuFQtJ37YOtnwtoeW/VvRXKwYw02fc7cB
# qZ9Xql4o4rmUMYIDdjCCA3ICAQEwdzBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# RGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNB
# NDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBAhAMTWlyS5T6PCpKPSkHgD1aMA0G
# CWCGSAFlAwQCAQUAoIHRMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAcBgkq
# hkiG9w0BCQUxDxcNMjMwNzEyMTcwOTU1WjArBgsqhkiG9w0BCRACDDEcMBowGDAW
# BBTzhyJNhjOCkjWplLy9j5bp/hx8czAvBgkqhkiG9w0BCQQxIgQgGG1vg8Aj9rMD
# wsVBRzlU8I/yjOg39KURk9PTqNYm8CYwNwYLKoZIhvcNAQkQAi8xKDAmMCQwIgQg
# x/ThvjIoiSCr4iY6vhrE/E/meBwtZNBMgHVXoCO1tvowDQYJKoZIhvcNAQEBBQAE
# ggIArl6AYn9SzljQOvU9kFz4cQKBubc5bBS2E+njeuoqNCtXl7MAJ38fb8ePLenW
# YReKiVdMwjaN4yUjscAEE0K3JCIvHJXQtqS4orlbS/302+eH4XybF86RZnU8ojYy
# z2v2Q3zLCSYi+ITCmFuqv/7ieptBUZsOgJG16PhTjdbFavKUaRtInYTasgJZQsNS
# f9bts/InlpfvfOxzQkRphiUZe4ohDbFX7cYOxguvcj2+T1gLCipYOAom1mn1+W1q
# uD9TxcojObjcVUU2LkTSFkXl7AiIQeOA1RapHqfFSile9Qv3Q0X/WqQUUE3xV/01
# 9b4WDV7YMY17CqDcgPF8kzIAeCVP/KO8xyi5D6EBdpP728LpmsnVvtZ4PfzFbeQn
# Wg8Dl2mhcTi6KsLyxU/sSRjXoZTjldUXBBmRj4rxXIMnAN89RPH8+EY7n7ztyDZf
# C4uvj2UB0FA7EoWMyMdhUDFvAmO1B3FYE1p5N5D78VlEURNZkjibyYdSs6Lk4ICW
# kxy3Wr43HbhwOA0splZzIAWZLtVEPa/xUjgIXuoX1+pQpezvq8tgLTFaukemASqE
# nu/i3o4BE5Rmq8N5jimGNPA0lENhOjNFf/nl6xOvWMQWkD3kYlhZM5EpHuMXeMZy
# L5L3QXeUpEuJPuDxwuqHaMczovTRW/T4eXkv/5UNuYrGZYc=
# SIG # End signature block
