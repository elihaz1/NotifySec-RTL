# NotifySec-RTL
NotifySec is an Outlook add-in used to help users report a suspicious e-mail to security team. this Outlook add-in is designed to support Hebrew (right to left notifications) and building MSI installer for fast deployment.

It is based on the work of https://github.com/certsocietegenerale/NotifySecurity and NightWizzard's way for adding message headers, (https://www.codeproject.com/Questions/1074498/Outlook-add-in-in-Csharp-get-message-header) with several modification as well as new features:
1. the add-in is designed to support notifications in right to left languges  (e.g Hebrew, Arab) 
2. the Solution include 2 wixtool setup project for building MSI installer for 32bit and 64bit MS office outlook (see howto section) which allow easy deployment. 
3.the add-in button is in the inbox toolbar tabMail and in TabReadMessage so users can report old e-mails located in any folder.
4. New icon <br />
![NotifySec Addin](https://user-images.githubusercontent.com/29439567/64485811-9fed4880-d22d-11e9-9fc6-5dbcd65986ca.png)

**Usage** <br />
The solution was coded on visual studio 2017. 
The add-in was successfully tested on office 2010, 2013, 2016 on x86 and x64 versions.

**Prerequisites** <br />
1. visual studio 2017
2. wix toolset - https://wixtoolset.org/releases/v3.11.1/stable

**howto** <br />
1. Open NotifySecOutlook2010.sln in VisualStudio
2. Using solution Explorer - double click "ribbon.cs" to edit mode, look for the comments asking you to enter security team mail and help desk mail as foloow: <br />
![enteryoursecteammail](https://user-images.githubusercontent.com/29439567/64485848-0bcfb100-d22e-11e9-81a6-c36aa5a08114.png)
3. SAVE. 
4. Test the solution by closing OUTLOOK.EXE and under the "DEBUG" menu click "START DEBUGGING".
5. Before building the Outlook add-in and the MSI files. You'll need to Check the Configuration Manger (under "Build" Menu) as follow:  
![configurationManagment](https://user-images.githubusercontent.com/29439567/64486017-62d68580-d230-11e9-95a4-0b6758375787.png)
6. It's best to build each project separatly, starting from the "NotifySecOutllok2010" project. In the solution explorer window, right click it and choose build. 
![build](https://user-images.githubusercontent.com/29439567/64486644-2eff5e00-d238-11e9-87f8-8a098818c818.png)
7. Reapet this for SetupProject32 (build MSI installer for 32 bit offfice) and SetupProject64 (build MSI installer for 32 bit offfice). *_Before building the SetupProject you need, Please:_*<br/>
 a. Add Wixtoolset DLL file to the SetupProject32 and SetupProject64 (In soulotion explorer right click the "References Folder" that's under the SetupProject and choose "Add Reference".Locate WixUIExtension.dll and WixNetFxExtension.dll aand Add them (usally in C:\Program Files (x86)\WiX Toolset v3.11\bin)
![WixDependencies](https://user-images.githubusercontent.com/29439567/64486635-1bec8e00-d238-11e9-912d-c0cc1ad96c19.png)
 b.that the configuration mangment definition are as requiered.<br/>

8. The Add-in is installed on HKLM (Local Machine) so it should be done with a user that is a local admin on the target pc.
9. Run Command is:<br/>
msiexec /q /i \\Remote_Server_Address\Folder_Name\InstallNotifysecOutlookAddin32bit.msi
msiexec /q /i \\Remote_Server_Address\Folder_Name\InstallNotifysecOutlookAddin64bit.msi

**Support** <br />
Eli Hazan - elihaz@gmail.com

**Thanks** <br />
Dvir Dahan & Alon Ekroni for QA and Testing. 

 


© 2019 GitHub, Inc.
