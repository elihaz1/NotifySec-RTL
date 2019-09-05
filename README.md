# NotifySec-RTL
NotifySec is an Outlook add-in used to help users report a suspicious e-mail to security team. the addin is designed to support right to left notifications and ready to build setup projects for building MSI installer.

It is based on the work of https://github.com/certsocietegenerale/NotifySecurity and NightWizzard's way for adding message headers, (https://www.codeproject.com/Questions/1074498/Outlook-add-in-in-Csharp-get-message-header) with several modification and new features:

1. the add-in is designed to support notifications in right to left languges  (e.g Hebrew, Arab) 
2. the Solution include 2 wixtool setup project for building MSI installer for 32bit and 64bit MS office outlook (see howto section) which allow easy deployment. 
3.the add-in button is in the inbox toolbar tabMail and in TabReadMessage so users can report old e-mails located in any folder.
4. New icon 

Usage
The solution was coded on visual studio 2017. 
The add-in was tested on office 2010, 2013, 2016 with 

© 2019 GitHub, Inc.
