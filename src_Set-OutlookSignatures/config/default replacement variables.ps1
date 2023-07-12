# This file allows defining custom replacement variables for Set-OutlookSignatures
#
# This script is executed as a whole once for each mailbox.
# It allows for complex replacement variable handling (complex string transformations, retrieving information from web services and databases, etc.).
#
# Attention: The configuration file is executed as part of Set-OutlookSignatures.ps1 and is not checked for any harmful content. Please only allow qualified technicians write access to this file, only use it to to define replacement variables and test it thoroughly.
#
# Replacement variable names are case sensitive.
# It is required to use full uppercase replacement variable names.
#
# Active Directory property names are case sensitive.
# It is required to use full lowercase Active Directory property names.
#
# A variable defined in this file overrides the definition of the same variable defined earlier in the script.
#
#
# What is the recommended approach for custom configuration files?
# You should not change the default configuration file '.\config\default replacement variable.ps1', as it might be changed in a future release of Set-OutlookSignatures. In this case, you would have to sort out the changes yourself.
#
# The following steps are recommended:
# 1. Create a new custom configuration file in a separate folder.
# 2. The first step in the new custom configuration file should be to load the default configuration file:
#    # Loading default replacement variables shipped with Set-OutlookSignatures
#    . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $(Join-Path -Path $(Get-Location).path -ChildPath '\config\default replacement variables.ps1') -Raw)))
# 3. After importing the default configuration file, existing replacement variables can be altered with custom definitions and new replacement variables can be added.
# 4. Instead of altering existing replacement variables, it is recommended to create new replacement variables with modified content.
# 5. Start Set-OutlookSignatures with the parameter 'ReplacementVariableConfigFile' pointing to the new custom configuration file.


# Currently logged in user
$ReplaceHash['$CURRENTUSERGIVENNAME$'] = [string]$ADPropsCurrentUser.givenname
$ReplaceHash['$CURRENTUSERSURNAME$'] = [string]$ADPropsCurrentUser.sn
$ReplaceHash['$CURRENTUSERDEPARTMENT$'] = [string]$ADPropsCurrentUser.department
$ReplaceHash['$CURRENTUSERTITLE$'] = [string]$ADPropsCurrentUser.title
$ReplaceHash['$CURRENTUSERSTREETADDRESS$'] = [string]$ADPropsCurrentUser.streetaddress
$ReplaceHash['$CURRENTUSERPOSTALCODE$'] = [string]$ADPropsCurrentUser.postalcode
$ReplaceHash['$CURRENTUSERLOCATION$'] = [string]$ADPropsCurrentUser.l
$ReplaceHash['$CURRENTUSERCOUNTRY$'] = [string]$ADPropsCurrentUser.co
$ReplaceHash['$CURRENTUSERSTATE$'] = [string]$ADPropsCurrentUser.st
$ReplaceHash['$CURRENTUSERTELEPHONE$'] = [string]$ADPropsCurrentUser.telephonenumber
$ReplaceHash['$CURRENTUSERFAX$'] = [string]$ADPropsCurrentUser.facsimiletelephonenumber
$ReplaceHash['$CURRENTUSERMOBILE$'] = [string]$ADPropsCurrentUser.mobile
$ReplaceHash['$CURRENTUSERMAIL$'] = [string]$ADPropsCurrentUser.mail
$ReplaceHash['$CURRENTUSERPHOTO$'] = $ADPropsCurrentUser.thumbnailphoto
$ReplaceHash['$CURRENTUSERPHOTODELETEEMPTY$'] = $ADPropsCurrentUser.thumbnailphoto
$ReplaceHash['$CURRENTUSEREXTATTR1$'] = [string]$ADPropsCurrentUser.extensionattribute1
$ReplaceHash['$CURRENTUSEREXTATTR2$'] = [string]$ADPropsCurrentUser.extensionattribute2
$ReplaceHash['$CURRENTUSEREXTATTR3$'] = [string]$ADPropsCurrentUser.extensionattribute3
$ReplaceHash['$CURRENTUSEREXTATTR4$'] = [string]$ADPropsCurrentUser.extensionattribute4
$ReplaceHash['$CURRENTUSEREXTATTR5$'] = [string]$ADPropsCurrentUser.extensionattribute5
$ReplaceHash['$CURRENTUSEREXTATTR6$'] = [string]$ADPropsCurrentUser.extensionattribute6
$ReplaceHash['$CURRENTUSEREXTATTR7$'] = [string]$ADPropsCurrentUser.extensionattribute7
$ReplaceHash['$CURRENTUSEREXTATTR8$'] = [string]$ADPropsCurrentUser.extensionattribute8
$ReplaceHash['$CURRENTUSEREXTATTR9$'] = [string]$ADPropsCurrentUser.extensionattribute9
$ReplaceHash['$CURRENTUSEREXTATTR10$'] = [string]$ADPropsCurrentUser.extensionattribute10
$ReplaceHash['$CURRENTUSEREXTATTR11$'] = [string]$ADPropsCurrentUser.extensionattribute11
$ReplaceHash['$CURRENTUSEREXTATTR12$'] = [string]$ADPropsCurrentUser.extensionattribute12
$ReplaceHash['$CURRENTUSEREXTATTR13$'] = [string]$ADPropsCurrentUser.extensionattribute13
$ReplaceHash['$CURRENTUSEREXTATTR14$'] = [string]$ADPropsCurrentUser.extensionattribute14
$ReplaceHash['$CURRENTUSEREXTATTR15$'] = [string]$ADPropsCurrentUser.extensionattribute15
$ReplaceHash['$CURRENTUSEROFFICE$'] = [string]$ADPropsCurrentUser.physicaldeliveryofficename
$ReplaceHash['$CURRENTUSERCOMPANY$'] = [string]$ADPropsCurrentUser.company
$ReplaceHash['$CURRENTUSERMAILNICKNAME$'] = [string]$ADPropsCurrentUser.mailnickname
$ReplaceHash['$CURRENTUSERDISPLAYNAME$'] = [string]$ADPropsCurrentUser.displayname


# Manager of currently logged in user
$ReplaceHash['$CURRENTUSERMANAGERGIVENNAME$'] = [string]$ADPropsCurrentUserManager.givenname
$ReplaceHash['$CURRENTUSERMANAGERSURNAME$'] = [string]$ADPropsCurrentUserManager.sn
$ReplaceHash['$CURRENTUSERMANAGERDEPARTMENT$'] = [string]$ADPropsCurrentUserManager.department
$ReplaceHash['$CURRENTUSERMANAGERTITLE$'] = [string]$ADPropsCurrentUserManager.title
$ReplaceHash['$CURRENTUSERMANAGERSTREETADDRESS$'] = [string]$ADPropsCurrentUserManager.streetaddress
$ReplaceHash['$CURRENTUSERMANAGERPOSTALCODE$'] = [string]$ADPropsCurrentUserManager.postalcode
$ReplaceHash['$CURRENTUSERMANAGERLOCATION$'] = [string]$ADPropsCurrentUserManager.l
$ReplaceHash['$CURRENTUSERMANAGERCOUNTRY$'] = [string]$ADPropsCurrentUserManager.co
$ReplaceHash['$CURRENTUSERMANAGERSTATE$'] = [string]$ADPropsCurrentUserManager.st
$ReplaceHash['$CURRENTUSERMANAGERTELEPHONE$'] = [string]$ADPropsCurrentUserManager.telephonenumber
$ReplaceHash['$CURRENTUSERMANAGERFAX$'] = [string]$ADPropsCurrentUserManager.facsimiletelephonenumber
$ReplaceHash['$CURRENTUSERMANAGERMOBILE$'] = [string]$ADPropsCurrentUserManager.mobile
$ReplaceHash['$CURRENTUSERMANAGERMAIL$'] = [string]$ADPropsCurrentUserManager.mail
$ReplaceHash['$CURRENTUSERMANAGERPHOTO$'] = $ADPropsCurrentUserManager.thumbnailphoto
$ReplaceHash['$CURRENTUSERMANAGERPHOTODELETEEMPTY$'] = $ADPropsCurrentUserManager.thumbnailphoto
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR1$'] = [string]$ADPropsCurrentUserManager.extensionattribute1
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR2$'] = [string]$ADPropsCurrentUserManager.extensionattribute2
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR3$'] = [string]$ADPropsCurrentUserManager.extensionattribute3
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR4$'] = [string]$ADPropsCurrentUserManager.extensionattribute4
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR5$'] = [string]$ADPropsCurrentUserManager.extensionattribute5
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR6$'] = [string]$ADPropsCurrentUserManager.extensionattribute6
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR7$'] = [string]$ADPropsCurrentUserManager.extensionattribute7
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR8$'] = [string]$ADPropsCurrentUserManager.extensionattribute8
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR9$'] = [string]$ADPropsCurrentUserManager.extensionattribute9
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR10$'] = [string]$ADPropsCurrentUserManager.extensionattribute10
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR11$'] = [string]$ADPropsCurrentUserManager.extensionattribute11
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR12$'] = [string]$ADPropsCurrentUserManager.extensionattribute12
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR13$'] = [string]$ADPropsCurrentUserManager.extensionattribute13
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR14$'] = [string]$ADPropsCurrentUserManager.extensionattribute14
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR15$'] = [string]$ADPropsCurrentUserManager.extensionattribute15
$ReplaceHash['$CURRENTUSERMANAGEROFFICE$'] = [string]$ADPropsCurrentUserManager.physicaldeliveryofficename
$ReplaceHash['$CURRENTUSERMANAGERCOMPANY$'] = [string]$ADPropsCurrentUserManager.company
$ReplaceHash['$CURRENTUSERMANAGERMAILNICKNAME$'] = [string]$ADPropsCurrentUserManager.mailnickname
$ReplaceHash['$CURRENTUSERMANAGERDISPLAYNAME$'] = [string]$ADPropsCurrentUserManager.displayname


# Current mailbox
$ReplaceHash['$CURRENTMAILBOXGIVENNAME$'] = [string]$ADPropsCurrentMailbox.givenname
$ReplaceHash['$CURRENTMAILBOXSURNAME$'] = [string]$ADPropsCurrentMailbox.sn
$ReplaceHash['$CURRENTMAILBOXDEPARTMENT$'] = [string]$ADPropsCurrentMailbox.department
$ReplaceHash['$CURRENTMAILBOXTITLE$'] = [string]$ADPropsCurrentMailbox.title
$ReplaceHash['$CURRENTMAILBOXSTREETADDRESS$'] = [string]$ADPropsCurrentMailbox.streetaddress
$ReplaceHash['$CURRENTMAILBOXPOSTALCODE$'] = [string]$ADPropsCurrentMailbox.postalcode
$ReplaceHash['$CURRENTMAILBOXLOCATION$'] = [string]$ADPropsCurrentMailbox.l
$ReplaceHash['$CURRENTMAILBOXCOUNTRY$'] = [string]$ADPropsCurrentMailbox.co
$ReplaceHash['$CURRENTMAILBOXSTATE$'] = [string]$ADPropsCurrentMailbox.st
$ReplaceHash['$CURRENTMAILBOXTELEPHONE$'] = [string]$ADPropsCurrentMailbox.telephonenumber
$ReplaceHash['$CURRENTMAILBOXFAX$'] = [string]$ADPropsCurrentMailbox.facsimiletelephonenumber
$ReplaceHash['$CURRENTMAILBOXMOBILE$'] = [string]$ADPropsCurrentMailbox.mobile
$ReplaceHash['$CURRENTMAILBOXMAIL$'] = [string]$ADPropsCurrentMailbox.mail
$ReplaceHash['$CURRENTMAILBOXPHOTO$'] = $ADPropsCurrentMailbox.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXPHOTODELETEEMPTY$'] = $ADPropsCurrentMailbox.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXEXTATTR1$'] = [string]$ADPropsCurrentMailbox.extensionattribute1
$ReplaceHash['$CURRENTMAILBOXEXTATTR2$'] = [string]$ADPropsCurrentMailbox.extensionattribute2
$ReplaceHash['$CURRENTMAILBOXEXTATTR3$'] = [string]$ADPropsCurrentMailbox.extensionattribute3
$ReplaceHash['$CURRENTMAILBOXEXTATTR4$'] = [string]$ADPropsCurrentMailbox.extensionattribute4
$ReplaceHash['$CURRENTMAILBOXEXTATTR5$'] = [string]$ADPropsCurrentMailbox.extensionattribute5
$ReplaceHash['$CURRENTMAILBOXEXTATTR6$'] = [string]$ADPropsCurrentMailbox.extensionattribute6
$ReplaceHash['$CURRENTMAILBOXEXTATTR7$'] = [string]$ADPropsCurrentMailbox.extensionattribute7
$ReplaceHash['$CURRENTMAILBOXEXTATTR8$'] = [string]$ADPropsCurrentMailbox.extensionattribute8
$ReplaceHash['$CURRENTMAILBOXEXTATTR9$'] = [string]$ADPropsCurrentMailbox.extensionattribute9
$ReplaceHash['$CURRENTMAILBOXEXTATTR10$'] = [string]$ADPropsCurrentMailbox.extensionattribute10
$ReplaceHash['$CURRENTMAILBOXEXTATTR11$'] = [string]$ADPropsCurrentMailbox.extensionattribute11
$ReplaceHash['$CURRENTMAILBOXEXTATTR12$'] = [string]$ADPropsCurrentMailbox.extensionattribute12
$ReplaceHash['$CURRENTMAILBOXEXTATTR13$'] = [string]$ADPropsCurrentMailbox.extensionattribute13
$ReplaceHash['$CURRENTMAILBOXEXTATTR14$'] = [string]$ADPropsCurrentMailbox.extensionattribute14
$ReplaceHash['$CURRENTMAILBOXEXTATTR15$'] = [string]$ADPropsCurrentMailbox.extensionattribute15
$ReplaceHash['$CURRENTMAILBOXOFFICE$'] = [string]$ADPropsCurrentMailbox.physicaldeliveryofficename
$ReplaceHash['$CURRENTMAILBOXCOMPANY$'] = [string]$ADPropsCurrentMailbox.company
$ReplaceHash['$CURRENTMAILBOXMAILNICKNAME$'] = [string]$ADPropsCurrentMailbox.mailnickname
$ReplaceHash['$CURRENTMAILBOXDISPLAYNAME$'] = [string]$ADPropsCurrentMailbox.displayname


# Manager of current mailbox
$ReplaceHash['$CURRENTMAILBOXMANAGERGIVENNAME$'] = [string]$ADPropsCurrentMailboxManager.givenname
$ReplaceHash['$CURRENTMAILBOXMANAGERSURNAME$'] = [string]$ADPropsCurrentMailboxManager.sn
$ReplaceHash['$CURRENTMAILBOXMANAGERDEPARTMENT$'] = [string]$ADPropsCurrentMailboxManager.department
$ReplaceHash['$CURRENTMAILBOXMANAGERTITLE$'] = [string]$ADPropsCurrentMailboxManager.title
$ReplaceHash['$CURRENTMAILBOXMANAGERSTREETADDRESS$'] = [string]$ADPropsCurrentMailboxManager.streetaddress
$ReplaceHash['$CURRENTMAILBOXMANAGERPOSTALCODE$'] = [string]$ADPropsCurrentMailboxManager.postalcode
$ReplaceHash['$CURRENTMAILBOXMANAGERLOCATION$'] = [string]$ADPropsCurrentMailboxManager.l
$ReplaceHash['$CURRENTMAILBOXMANAGERCOUNTRY$'] = [string]$ADPropsCurrentMailboxManager.co
$ReplaceHash['$CURRENTMAILBOXMANAGERSTATE$'] = [string]$ADPropsCurrentMailboxManager.st
$ReplaceHash['$CURRENTMAILBOXMANAGERTELEPHONE$'] = [string]$ADPropsCurrentMailboxManager.telephonenumber
$ReplaceHash['$CURRENTMAILBOXMANAGERFAX$'] = [string]$ADPropsCurrentMailboxManager.facsimiletelephonenumber
$ReplaceHash['$CURRENTMAILBOXMANAGERMOBILE$'] = [string]$ADPropsCurrentMailboxManager.mobile
$ReplaceHash['$CURRENTMAILBOXMANAGERMAIL$'] = [string]$ADPropsCurrentMailboxManager.mail
$ReplaceHash['$CURRENTMAILBOXMANAGERPHOTO$'] = $ADPropsCurrentMailboxManager.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$'] = $ADPropsCurrentMailboxManager.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR1$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute1
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR2$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute2
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR3$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute3
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR4$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute4
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR5$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute5
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR6$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute6
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR7$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute7
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR8$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute8
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR9$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute9
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR10$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute10
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR11$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute11
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR12$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute12
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR13$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute13
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR14$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute14
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR15$'] = [string]$ADPropsCurrentMailboxManager.extensionattribute15
$ReplaceHash['$CURRENTMAILBOXMANAGEROFFICE$'] = [string]$ADPropsCurrentMailboxManager.physicaldeliveryofficename
$ReplaceHash['$CURRENTMAILBOXMANAGERCOMPANY$'] = [string]$ADPropsCurrentMailboxManager.company
$ReplaceHash['$CURRENTMAILBOXMANAGERMAILNICKNAME$'] = [string]$ADPropsCurrentMailboxManager.mailnickname
$ReplaceHash['$CURRENTMAILBOXMANAGERDISPLAYNAME$'] = [string]$ADPropsCurrentMailboxManager.displayname


# $CURRENTUSERNAMEWITHTITLES$, $CURRENTUSERMANAGERNAMEWITHTITLES$
# $CURRENTMAILBOXNAMEWITHTITLES$, $CURRENTMAILBOXMANAGERNAMEWITHTITLES$
# Academic titles according to standards in German speaking countries
# <custom AD attribute 'svstitelvorne'> <standard AD attribute 'givenname'> <standard AD attribute 'surname'>, <custom AD attribute 'svstitelhinten'>
# If one or more attributes are not set, unnecessary whitespaces and commas are avoided by using '-join'
# Examples:
#   Mag. Dr. John Doe, BA MA PhD
#   Dr. John Doe
#   John Doe, PhD
#   John Doe
$ReplaceHash['$CURRENTUSERNAMEWITHTITLES$'] = (((((([string]$ADPropsCurrentUser.svstitelvorne, [string]$ADPropsCurrentUser.givenname, [string]$ADPropsCurrentUser.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUser.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CURRENTUSERMANAGERNAMEWITHTITLES$'] = (((((([string]$ADPropsCurrentUserManager.svstitelvorne, [string]$ADPropsCurrentUserManager.givenname, [string]$ADPropsCurrentUserManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUserManager.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CURRENTMAILBOXNAMEWITHTITLES$'] = (((((([string]$ADPropsCurrentMailbox.svstitelvorne, [string]$ADPropsCurrentMailbox.givenname, [string]$ADPropsCurrentMailbox.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailbox.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')
$ReplaceHash['$CURRENTMAILBOXMANAGERNAMEWITHTITLES$'] = (((((([string]$ADPropsCurrentMailboxManager.svstitelvorne, [string]$ADPropsCurrentMailboxManager.givenname, [string]$ADPropsCurrentMailboxManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailboxManager.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', ')


# SIG # Begin signature block
# MIIwTgYJKoZIhvcNAQcCoIIwPzCCMDsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDp4LkEvZ4F5/hW
# XE7aVjDS02J+q6xm7UZwgg9VOvgMtKCCFCkwggWQMIIDeKADAgECAhAFmxtXno4h
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
# GNmrMYIbezCCG3cCAQEwfTBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgQ29kZSBTaWdu
# aW5nIFJTQTQwOTYgU0hBMzg0IDIwMjEgQ0ExAhAKaypbp7cyIFa+lR7OVPAvMA0G
# CWCGSAFlAwQCAQUAoIIBjjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg+T8O+QiS
# 80y3aTZN8af/nYlntUTKqzciN/g6JXtiZxIwggEgBgorBgEEAYI3AgEMMYIBEDCC
# AQyggcWAgcIAUABvAHcAZQByAGUAZAAgAGIAeQAgAEUAeABwAGwAaQBjAEkAVAAg
# AEMAbwBuAHMAdQBsAHQAaQBuAGcALgAgAFUAbgBsAG8AYwBrACAAYQBsAGwAIABm
# AGUAYQB0AHUAcgBlAHMAIAB3AGkAdABoACAAUwBlAHQALQBPAHUAdABsAG8AbwBr
# AFMAaQBnAG4AYQB0AHUAcgBlAHMAIABCAGUAbgBlAGYAYQBjAHQAbwByACAAQwBp
# AHIAYwBsAGUALqFCgEBodHRwczovL2V4cGxpY2l0Y29uc3VsdGluZy5hdC9vcGVu
# LXNvdXJjZS9zZXQtb3V0bG9va3NpZ25hdHVyZXMgMA0GCSqGSIb3DQEBAQUABIIC
# ANG7gnAQ3fKOCim+u2nrpDHkmZhCJcm2fSIZvhbOH9BBTQnaQkI6bnxSx/k5n8WS
# GSBpxObtaJS4hmoyoaIpNf8tJBjtb42Q9duPSnXcu7eCEiCWpDhyZY8+cOEuj2sM
# pqMr/Z0f/rkwv2/8zZJjOOQq4xMfREAT2umb2RvUXUUlBf6aC3Cyzk/VS1wnlrpb
# UYY3EBfSZII3/qWLS2aLW1vMK7R6SEIZBSHTDLRGSuxNyEeQNDOGr/HXKHUMxDEN
# BHB2SJjea3nyEQGCXYXz61JQDGf1WrAzWuKhHxNYhMFPadUrkTk80ijkksWE/akS
# V6XarexjovHmrHk8SSiTOlvLpMuSQalnY2tHtOeBiaNrZW9i9R0NK6yuqos1GH7S
# a0qhDF7xGy3tkkPnFelA1U8OoSnvc8JXuqX1bUQ/WXFldRulbo8wvRE+5cRuizBV
# NqIB+nRVFJkWE++BcTPBZrIuN3xSHh0ULl9i1Y3rTU5Y0NLGnt0pKw0LRynrZXFS
# mtGNkkaAXQsOuOZ1c4amIrbIzsAG0yP0Tt+F98ndzjUrNO2w2lrPoIGgh+ptKDeS
# +Ost5KrB0nBgjTulu0hTmjVPVYhf67++Ubt57N8+u1vj/1gT2EhOsq+MDfm/VoEW
# AbytGTLHANzz6bHVL5BjxRw+sYDmVYEVIhE6CAaII7uAoYIXPTCCFzkGCisGAQQB
# gjcDAwExghcpMIIXJQYJKoZIhvcNAQcCoIIXFjCCFxICAQMxDzANBglghkgBZQME
# AgEFADB3BgsqhkiG9w0BCRABBKBoBGYwZAIBAQYJYIZIAYb9bAcBMDEwDQYJYIZI
# AWUDBAIBBQAEIF1DNckbSSQ41G17Zs5qsev1xv3YBkkGbq8yIXlk9y28AhBFEEyi
# VDcN0nbTtb3FS319GA8yMDIzMDcxMjE3MDk1NVqgghMHMIIGwDCCBKigAwIBAgIQ
# DE1pckuU+jwqSj0pB4A9WjANBgkqhkiG9w0BAQsFADBjMQswCQYDVQQGEwJVUzEX
# MBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0
# ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBMB4XDTIyMDkyMTAw
# MDAwMFoXDTMzMTEyMTIzNTk1OVowRjELMAkGA1UEBhMCVVMxETAPBgNVBAoTCERp
# Z2lDZXJ0MSQwIgYDVQQDExtEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMiAtIDIwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDP7KUmOsap8mu7jcENmtuh6BSF
# dDMaJqzQHFUeHjZtvJJVDGH0nQl3PRWWCC9rZKT9BoMW15GSOBwxApb7crGXOlWv
# M+xhiummKNuQY1y9iVPgOi2Mh0KuJqTku3h4uXoW4VbGwLpkU7sqFudQSLuIaQyI
# xvG+4C99O7HKU41Agx7ny3JJKB5MgB6FVueF7fJhvKo6B332q27lZt3iXPUv7Y3U
# TZWEaOOAy2p50dIQkUYp6z4m8rSMzUy5Zsi7qlA4DeWMlF0ZWr/1e0BubxaompyV
# R4aFeT4MXmaMGgokvpyq0py2909ueMQoP6McD1AGN7oI2TWmtR7aeFgdOej4TJEQ
# ln5N4d3CraV++C0bH+wrRhijGfY59/XBT3EuiQMRoku7mL/6T+R7Nu8GRORV/zbq
# 5Xwx5/PCUsTmFntafqUlc9vAapkhLWPlWfVNL5AfJ7fSqxTlOGaHUQhr+1NDOdBk
# +lbP4PQK5hRtZHi7mP2Uw3Mh8y/CLiDXgazT8QfU4b3ZXUtuMZQpi+ZBpGWUwFjl
# 5S4pkKa3YWT62SBsGFFguqaBDwklU/G/O+mrBw5qBzliGcnWhX8T2Y15z2LF7OF7
# ucxnEweawXjtxojIsG4yeccLWYONxu71LHx7jstkifGxxLjnU15fVdJ9GSlZA076
# XepFcxyEftfO4tQ6dwIDAQABo4IBizCCAYcwDgYDVR0PAQH/BAQDAgeAMAwGA1Ud
# EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwIAYDVR0gBBkwFzAIBgZn
# gQwBBAIwCwYJYIZIAYb9bAcBMB8GA1UdIwQYMBaAFLoW2W1NhS9zKXaaL3WMaiCP
# nshvMB0GA1UdDgQWBBRiit7QYfyPMRTtlwvNPSqUFN9SnDBaBgNVHR8EUzBRME+g
# TaBLhklodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRS
# U0E0MDk2U0hBMjU2VGltZVN0YW1waW5nQ0EuY3JsMIGQBggrBgEFBQcBAQSBgzCB
# gDAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMFgGCCsGAQUF
# BzAChkxodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVk
# RzRSU0E0MDk2U0hBMjU2VGltZVN0YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEBCwUA
# A4ICAQBVqioa80bzeFc3MPx140/WhSPx/PmVOZsl5vdyipjDd9Rk/BX7NsJJUSx4
# iGNVCUY5APxp1MqbKfujP8DJAJsTHbCYidx48s18hc1Tna9i4mFmoxQqRYdKmEIr
# UPwbtZ4IMAn65C3XCYl5+QnmiM59G7hqopvBU2AJ6KO4ndetHxy47JhB8PYOgPvk
# /9+dEKfrALpfSo8aOlK06r8JSRU1NlmaD1TSsht/fl4JrXZUinRtytIFZyt26/+Y
# siaVOBmIRBTlClmia+ciPkQh0j8cwJvtfEiy2JIMkU88ZpSvXQJT657inuTTH4YB
# ZJwAwuladHUNPeF5iL8cAZfJGSOA1zZaX5YWsWMMxkZAO85dNdRZPkOaGK7DycvD
# +5sTX2q1x+DzBcNZ3ydiK95ByVO5/zQQZ/YmMph7/lxClIGUgp2sCovGSxVK05iQ
# RWAzgOAj3vgDpPZFR+XOuANCR+hBNnF3rf2i6Jd0Ti7aHh2MWsgemtXC8MYiqE+b
# vdgcmlHEL5r2X6cnl7qWLoVXwGDneFZ/au/ClZpLEQLIgpzJGgV8unG1TnqZbPTo
# ntRamMifv427GFxD9dAq6OJi7ngE273R+1sKqHB+8JeEeOMIA11HLGOoJTiXAdI/
# Otrl5fbmm9x+LMz/F0xNAKLY1gEOuIvu5uByVYksJxlh9ncBjDCCBq4wggSWoAMC
# AQICEAc2N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMy
# MzAwMDAwMFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoT
# DkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJT
# QTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQAD
# ggIPADCCAgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD
# +Vr2EaFEFUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz
# 7iuAhIoiGN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp
# 39mQh0YAe9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0Cs
# X7LeSn3O9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OT
# rCw54qVI1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4
# EbP29p7mO1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEc
# azjFKfPKqpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUo
# JEHtQr8FnGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfp
# mEpYPtMDiP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSy
# Px4JduyrXUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMB
# AAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUv
# cyl2mi91jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAO
# BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEE
# azBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYB
# BQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ip
# RCIBfmbW2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL
# 5Vxb122H+oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU
# 1/+rT4osequFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa
# 96kQsl3p/yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNW
# hqsKRcnfxI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlL
# AlKnN36TU6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14
# OuSereU0cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjT
# x/no8Zhf+yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7
# YGcWoWa63VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLf
# BInwAM1dwvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r
# 5db7qS9EFUrnEw4d2zc4GqEr9u3WfPwwggWNMIIEdaADAgECAhAOmxiO+dAt5+/b
# UOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMT
# G0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0z
# MTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
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
# T8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAP
# BgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAf
# BgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYw
# eQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# bDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4
# X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc4
# 7/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0
# mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk
# 1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLB
# sln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGp
# n1eqXijiuZQxggN2MIIDcgIBATB3MGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0
# MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0ECEAxNaXJLlPo8Kko9KQeAPVowDQYJ
# YIZIAWUDBAIBBQCggdEwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwGCSqG
# SIb3DQEJBTEPFw0yMzA3MTIxNzA5NTVaMCsGCyqGSIb3DQEJEAIMMRwwGjAYMBYE
# FPOHIk2GM4KSNamUvL2Plun+HHxzMC8GCSqGSIb3DQEJBDEiBCAphtnkJkjqkYHB
# v7AxQI6ai9vP0YVO18RNdJHPb1kJxzA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCDH
# 9OG+MiiJIKviJjq+GsT8T+Z4HC1k0EyAdVegI7W2+jANBgkqhkiG9w0BAQEFAASC
# AgBjhxRx9m8HomRFOwdmP+zQHMt1sNpY9QVeCndeP7c4P162SysX1S83tTVjIA7Y
# ZMlD36lk2R+7sMuUeN8bNnv/ECE9hL9r3Y07bLd3RjiTOnKTyvRo0Q2s1WxKUX6h
# Ytk5O7gObXGc50V6yEd5cirHLtqI5OqIUycBVTSKVrC4kgVE4T+1gmSdYQ/yv2Vj
# i07Z8HhG0IdxqjSZ879+RkVhLKJNOZfml385rxOYmcmKyNcZqWCYrur5dI7DJg4C
# YvXWr4SAnjiQJreRrExuQzW2iZtj3ncxBZAHeceLQ4mhAfR9yCbHcTs4Oelbpc16
# PKqubR5UBCWHFfJEFwYlGd7nhiLJsU1LwdpAPpest3x0h836Z2VCkbrLzAGazkMB
# Q4Q1fIp+4JP8KgHXS1wFUwohK3J4hJWOjkQv4s7A+bZtPuXSw2naU+yu0O5iPU4y
# W02fV3yGsoX5SWLnoiCnrEylLpM7PqfGLqQiO0PL4ucLiH/TwYExk1+6JuDfN7d7
# G++gFKesGnoZw+t+jpojHY4pG5HrMEoAqWbdXAVBDMOV4vshvUYN/0v5Yp5AHblW
# QRAJSm0JY8544wsATlpMleHBCA6A94DFWGbLC9cgtY6Yyqlzu5KA1c5lhdQPlXOn
# j0a66okU341gg1KlSCDaUiWZgc/N/5z1U5wUyUPvXezaEg==
# SIG # End signature block
