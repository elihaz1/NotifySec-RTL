# NotifySec-RTL
NotifySec is an Outlook add-in used to help users report a suspicious e-mail to security team. this Outlook add-in is designed to support Hebrew (right to left notifications) and building MSI installer for fast deployment.

It is based on the work of https://github.com/certsocietegenerale/NotifySecurity and NightWizzard's way for adding message headers, (https://www.codeproject.com/Questions/1074498/Outlook-add-in-in-Csharp-get-message-header) with several modification as well as new features:

1. the add-in is designed to support notifications in right to left languges  (e.g Hebrew, Arab) 
2. the Solution include 2 wixtool setup project for building MSI installer for 32bit and 64bit MS office outlook (see howto section) which allow easy deployment. 
3.the add-in button is in the inbox toolbar tabMail and in TabReadMessage so users can report old e-mails located in any folder.
4. New icon 
 ![](NotifySec Addin.png)

**Usage**
The solution was coded on visual studio 2017. 
The add-in was successfully tested on office 2010, 2013, 2016 on x86 and x64 versions.


**Prerequisites**
1. visual studio 2017
2. wix toolset - https://wixtoolset.org/releases/v3.11.1/stable

**howto**
1. Open NotifySecOutlook2010.sln in VisualStudio
2. Using solution Explorer - click "ribbon.cs" to edit, look for the comments asking for security team mail and help desk mail:



© 2019 GitHub, Inc.
