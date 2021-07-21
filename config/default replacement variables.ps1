﻿# This file allows defining custom replacement variables for Set-OutlookSignatures.ps1
#   
# This script is executed as a whole once for each mailbox.
# This allows for complex replacement variable handling (complex string transformations, retrieving information from web services and databases, etc.).
#
# Attention: The configuration file is executed as part of Set-OutlookSignatures.ps1 and is not checked for any harmful content. Please only allow qualified technicians write access to this file, only use it to to define replacement variables and test it thoroughly.
#
# Replacement variable names are case sensitive.
# It is required to use full uppercase replacement variables.
#
# A variable defined in this file overrides the definition of the same variable defined earlier in the script.


# Currently logged on user
$ReplaceHash['$CURRENTUSERGIVENNAME$'] = [string]$ADPropsCurrentUser.givenname
$ReplaceHash['$CURRENTUSERSURNAME$'] = [string]$ADPropsCurrentUser.sn
$ReplaceHash['$CURRENTUSERDEPARTMENT$'] = [string]$ADPropsCurrentUser.department
$ReplaceHash['$CURRENTUSERTITLE$'] = [string]$ADPropsCurrentUser.title
$ReplaceHash['$CURRENTUSERSTREETADDRESS$'] = [string]$ADPropsCurrentUser.streetaddress
$ReplaceHash['$CURRENTUSERPOSTALCODE$'] = [string]$ADPropsCurrentUser.postalcode
$ReplaceHash['$CURRENTUSERLOCATION$'] = [string]$ADPropsCurrentUser.l
$ReplaceHash['$CURRENTUSERCOUNTRY$'] = [string]$ADPropsCurrentUser.co
$ReplaceHash['$CURRENTUSERTELEPHONE$'] = [string]$ADPropsCurrentUser.telephonenumber
$ReplaceHash['$CURRENTUSERFAX$'] = [string]$ADPropsCurrentUser.facsimiletelephonenumber
$ReplaceHash['$CURRENTUSERMOBILE$'] = [string]$ADPropsCurrentUser.mobile
$ReplaceHash['$CURRENTUSERMAIL$'] = [string]$ADPropsCurrentUser.mail
$ReplaceHash['$CURRENTUSERPHOTO$'] = $ADPropsCurrentUser.thumbnailphoto
$ReplaceHash['$CURRENTUSERPHOTODELETEEMPTY$'] = $ADPropsCurrentUser.thumbnailphoto
$ReplaceHash['$CURRENTUSEREXTATTR1$'] = [string]$ADPropsCurrentUser.extensionAttribute1
$ReplaceHash['$CURRENTUSEREXTATTR2$'] = [string]$ADPropsCurrentUser.extensionAttribute2
$ReplaceHash['$CURRENTUSEREXTATTR3$'] = [string]$ADPropsCurrentUser.extensionAttribute3
$ReplaceHash['$CURRENTUSEREXTATTR4$'] = [string]$ADPropsCurrentUser.extensionAttribute4
$ReplaceHash['$CURRENTUSEREXTATTR5$'] = [string]$ADPropsCurrentUser.extensionAttribute5
$ReplaceHash['$CURRENTUSEREXTATTR6$'] = [string]$ADPropsCurrentUser.extensionAttribute6
$ReplaceHash['$CURRENTUSEREXTATTR7$'] = [string]$ADPropsCurrentUser.extensionAttribute7
$ReplaceHash['$CURRENTUSEREXTATTR8$'] = [string]$ADPropsCurrentUser.extensionAttribute8
$ReplaceHash['$CURRENTUSEREXTATTR9$'] = [string]$ADPropsCurrentUser.extensionAttribute9
$ReplaceHash['$CURRENTUSEREXTATTR10$'] = [string]$ADPropsCurrentUser.extensionAttribute10
$ReplaceHash['$CURRENTUSEREXTATTR11$'] = [string]$ADPropsCurrentUser.extensionAttribute11
$ReplaceHash['$CURRENTUSEREXTATTR12$'] = [string]$ADPropsCurrentUser.extensionAttribute12
$ReplaceHash['$CURRENTUSEREXTATTR13$'] = [string]$ADPropsCurrentUser.extensionAttribute13
$ReplaceHash['$CURRENTUSEREXTATTR14$'] = [string]$ADPropsCurrentUser.extensionAttribute14
$ReplaceHash['$CURRENTUSEREXTATTR15$'] = [string]$ADPropsCurrentUser.extensionAttribute15


# Manager of currently logged on user
$ReplaceHash['$CURRENTUSERMANAGERGIVENNAME$'] = [string]$ADPropsCurrentUserManager.givenname
$ReplaceHash['$CURRENTUSERMANAGERSURNAME$'] = [string]$ADPropsCurrentUserManager.sn
$ReplaceHash['$CURRENTUSERMANAGERDEPARTMENT$'] = [string]$ADPropsCurrentUserManager.department
$ReplaceHash['$CURRENTUSERMANAGERTITLE$'] = [string]$ADPropsCurrentUserManager.title
$ReplaceHash['$CURRENTUSERMANAGERSTREETADDRESS$'] = [string]$ADPropsCurrentUserManager.streetaddress
$ReplaceHash['$CURRENTUSERMANAGERPOSTALCODE$'] = [string]$ADPropsCurrentUserManager.postalcode
$ReplaceHash['$CURRENTUSERMANAGERLOCATION$'] = [string]$ADPropsCurrentUserManager.l
$ReplaceHash['$CURRENTUSERMANAGERCOUNTRY$'] = [string]$ADPropsCurrentUserManager.co
$ReplaceHash['$CURRENTUSERMANAGERTELEPHONE$'] = [string]$ADPropsCurrentUserManager.telephonenumber
$ReplaceHash['$CURRENTUSERMANAGERFAX$'] = [string]$ADPropsCurrentUserManager.facsimiletelephonenumber
$ReplaceHash['$CURRENTUSERMANAGERMOBILE$'] = [string]$ADPropsCurrentUserManager.mobile
$ReplaceHash['$CURRENTUSERMANAGERMAIL$'] = [string]$ADPropsCurrentUserManager.mail
$ReplaceHash['$CURRENTUSERMANAGERPHOTO$'] = $ADPropsCurrentUserManager.thumbnailphoto
$ReplaceHash['$CURRENTUSERMANAGERPHOTODELETEEMPTY$'] = $ADPropsCurrentUserManager.thumbnailphoto
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR1$'] = [string]$ADPropsCurrentUserManager.extensionAttribute1
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR2$'] = [string]$ADPropsCurrentUserManager.extensionAttribute2
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR3$'] = [string]$ADPropsCurrentUserManager.extensionAttribute3
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR4$'] = [string]$ADPropsCurrentUserManager.extensionAttribute4
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR5$'] = [string]$ADPropsCurrentUserManager.extensionAttribute5
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR6$'] = [string]$ADPropsCurrentUserManager.extensionAttribute6
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR7$'] = [string]$ADPropsCurrentUserManager.extensionAttribute7
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR8$'] = [string]$ADPropsCurrentUserManager.extensionAttribute8
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR9$'] = [string]$ADPropsCurrentUserManager.extensionAttribute9
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR10$'] = [string]$ADPropsCurrentUserManager.extensionAttribute10
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR11$'] = [string]$ADPropsCurrentUserManager.extensionAttribute11
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR12$'] = [string]$ADPropsCurrentUserManager.extensionAttribute12
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR13$'] = [string]$ADPropsCurrentUserManager.extensionAttribute13
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR14$'] = [string]$ADPropsCurrentUserManager.extensionAttribute14
$ReplaceHash['$CURRENTUSERMANAGEREXTATTR15$'] = [string]$ADPropsCurrentUserManager.extensionAttribute15


# Current mailbox
$ReplaceHash['$CURRENTMAILBOXGIVENNAME$'] = [string]$ADPropsCurrentMailbox.givenname
$ReplaceHash['$CURRENTMAILBOXSURNAME$'] = [string]$ADPropsCurrentMailbox.sn
$ReplaceHash['$CURRENTMAILBOXDEPARTMENT$'] = [string]$ADPropsCurrentMailbox.department
$ReplaceHash['$CURRENTMAILBOXTITLE$'] = [string]$ADPropsCurrentMailbox.title
$ReplaceHash['$CURRENTMAILBOXSTREETADDRESS$'] = [string]$ADPropsCurrentMailbox.streetaddress
$ReplaceHash['$CURRENTMAILBOXPOSTALCODE$'] = [string]$ADPropsCurrentMailbox.postalcode
$ReplaceHash['$CURRENTMAILBOXLOCATION$'] = [string]$ADPropsCurrentMailbox.l
$ReplaceHash['$CURRENTMAILBOXCOUNTRY$'] = [string]$ADPropsCurrentMailbox.co
$ReplaceHash['$CURRENTMAILBOXTELEPHONE$'] = [string]$ADPropsCurrentMailbox.telephonenumber
$ReplaceHash['$CURRENTMAILBOXFAX$'] = [string]$ADPropsCurrentMailbox.facsimiletelephonenumber
$ReplaceHash['$CURRENTMAILBOXMOBILE$'] = [string]$ADPropsCurrentMailbox.mobile
$ReplaceHash['$CURRENTMAILBOXMAIL$'] = [string]$ADPropsCurrentMailbox.mail
$ReplaceHash['$CURRENTMAILBOXPHOTO$'] = $ADPropsCurrentMailbox.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXPHOTODELETEEMPTY$'] = $ADPropsCurrentMailbox.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXEXTATTR1$'] = [string]$ADPropsCurrentMailbox.extensionAttribute1
$ReplaceHash['$CURRENTMAILBOXEXTATTR2$'] = [string]$ADPropsCurrentMailbox.extensionAttribute2
$ReplaceHash['$CURRENTMAILBOXEXTATTR3$'] = [string]$ADPropsCurrentMailbox.extensionAttribute3
$ReplaceHash['$CURRENTMAILBOXEXTATTR4$'] = [string]$ADPropsCurrentMailbox.extensionAttribute4
$ReplaceHash['$CURRENTMAILBOXEXTATTR5$'] = [string]$ADPropsCurrentMailbox.extensionAttribute5
$ReplaceHash['$CURRENTMAILBOXEXTATTR6$'] = [string]$ADPropsCurrentMailbox.extensionAttribute6
$ReplaceHash['$CURRENTMAILBOXEXTATTR7$'] = [string]$ADPropsCurrentMailbox.extensionAttribute7
$ReplaceHash['$CURRENTMAILBOXEXTATTR8$'] = [string]$ADPropsCurrentMailbox.extensionAttribute8
$ReplaceHash['$CURRENTMAILBOXEXTATTR9$'] = [string]$ADPropsCurrentMailbox.extensionAttribute9
$ReplaceHash['$CURRENTMAILBOXEXTATTR10$'] = [string]$ADPropsCurrentMailbox.extensionAttribute10
$ReplaceHash['$CURRENTMAILBOXEXTATTR11$'] = [string]$ADPropsCurrentMailbox.extensionAttribute11
$ReplaceHash['$CURRENTMAILBOXEXTATTR12$'] = [string]$ADPropsCurrentMailbox.extensionAttribute12
$ReplaceHash['$CURRENTMAILBOXEXTATTR13$'] = [string]$ADPropsCurrentMailbox.extensionAttribute13
$ReplaceHash['$CURRENTMAILBOXEXTATTR14$'] = [string]$ADPropsCurrentMailbox.extensionAttribute14
$ReplaceHash['$CURRENTMAILBOXEXTATTR15$'] = [string]$ADPropsCurrentMailbox.extensionAttribute15


# Manager of current mailbox
$ReplaceHash['$CURRENTMAILBOXMANAGERGIVENNAME$'] = [string]$ADPropsCurrentMailboxManager.givenname
$ReplaceHash['$CURRENTMAILBOXMANAGERSURNAME$'] = [string]$ADPropsCurrentMailboxManager.sn
$ReplaceHash['$CURRENTMAILBOXMANAGERDEPARTMENT$'] = [string]$ADPropsCurrentMailboxManager.department
$ReplaceHash['$CURRENTMAILBOXMANAGERTITLE$'] = [string]$ADPropsCurrentMailboxManager.title
$ReplaceHash['$CURRENTMAILBOXMANAGERSTREETADDRESS$'] = [string]$ADPropsCurrentMailboxManager.streetaddress
$ReplaceHash['$CURRENTMAILBOXMANAGERPOSTALCODE$'] = [string]$ADPropsCurrentMailboxManager.postalcode
$ReplaceHash['$CURRENTMAILBOXMANAGERLOCATION$'] = [string]$ADPropsCurrentMailboxManager.l
$ReplaceHash['$CURRENTMAILBOXMANAGERCOUNTRY$'] = [string]$ADPropsCurrentMailboxManager.co
$ReplaceHash['$CURRENTMAILBOXMANAGERTELEPHONE$'] = [string]$ADPropsCurrentMailboxManager.telephonenumber
$ReplaceHash['$CURRENTMAILBOXMANAGERFAX$'] = [string]$ADPropsCurrentMailboxManager.facsimiletelephonenumber
$ReplaceHash['$CURRENTMAILBOXMANAGERMOBILE$'] = [string]$ADPropsCurrentMailboxManager.mobile
$ReplaceHash['$CURRENTMAILBOXMANAGERMAIL$'] = [string]$ADPropsCurrentMailboxManager.mail
$ReplaceHash['$CURRENTMAILBOXMANAGERPHOTO$'] = $ADPropsCurrentMailboxManager.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$'] = $ADPropsCurrentMailboxManager.thumbnailphoto
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR1$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute1
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR2$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute2
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR3$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute3
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR4$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute4
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR5$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute5
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR6$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute6
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR7$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute7
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR8$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute8
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR9$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute9
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR10$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute10
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR11$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute11
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR12$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute12
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR13$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute13
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR14$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute14
$ReplaceHash['$CURRENTMAILBOXMANAGEREXTATTR15$'] = [string]$ADPropsCurrentMailboxManager.extensionAttribute15


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
