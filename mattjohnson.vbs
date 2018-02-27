Option Explicit
Err.Clear

'*** Comment out for testing, leave in for production (in case logon/logoff is out of sequence) ***'
ON ERROR RESUME NEXT

'***********************************'
'*** BEGIN VARIABLE DECLARATIONS ***'
'***********************************'
                                   
Dim wshNetwork, wshShell 
Dim strUserTimestamp, strWksTimestamp
Dim objDomain, DomainString, UserString, UserStrLen, UserStringPrint, UserObj, Path, objNetwork, strDesktopPath, objShortcutURL
Dim filedate, logType, strUserFile, strWksFile, strWksType, strDriveLetter
Dim strDesktop, strNetwork, WinDir, strComputer, GroupObj
Dim iCnt, sArg, oArgs, objFSO, objFile1, objFile2

'*** Login string variables for event logging ***'
Dim strLogLogin, strLogApp, strLogDriveQ

'*** Desktop / Network Icons (Projects) ***'
Dim oShellLink1, oShellLink2, oShellLink3, oShellLink4

'*** Departmental drives / shortcuts ***'
Dim strCreateShortcut, strTargetPath, strWindowStyle, strIconLocation, strDescription, strWorkingDirectory
Dim oShellLinkTT, oShellLinkQDrive

Const ForAppending = 8
Const strUserFileLocation = "\\rtix.com\admin\UserLogins"
Const strWksFileLocation = "\\rtix.com\admin\WksLogins"

'**** BIND TO DOMAIN AND DEFINE VARIABLES ****'

Set WSHShell = CreateObject("WScript.Shell")
Set WSHNetwork = CreateObject("WScript.Network")
Set objDomain = getObject("LDAP://rootDse")

WinDir = WshShell.ExpandEnvironmentStrings("%WinDir%")

DomainString = objDomain.Get("dnsHostName")
UserString = WSHNetwork.UserName
strComputer = WSHNetwork.ComputerName

UserStrLen = len(UserString)
if UserStrLen < 12 THEN
                UserStringPrint = UCase(UserString & String(12-UserStrLen," "))
else
                UserStringPrint = UCase(UserString)
end if

set objFSO = CreateObject("Scripting.FileSystemObject")
Set UserObj = GetObject("WinNT://" & DomainString & "/" & UserString)
strDesktop = WshShell.SpecialFolders("Desktop")
strNetwork = WshShell.SpecialFolders("Nethood")

strUserFile = strUserFileLocation & "\" & UserString & ".txt"
strWksFile = strWksFileLocation & "\" & strComputer & ".txt"
fileDate = "18/october/2015 - 1502"
strWksType = "-"

'***********************************'
'**** END VARIABLE DECLARATIONS ****'
'***********************************'

'***********************************'
'****** BEGIN SCRIPTING BLOCK ******'
'***********************************'

'**** Select logon or logoff ****'

Set oArgs = WScript.Arguments
For iCnt = 0 to oArgs.Count - 1
sArg = oArgs(iCnt)
Next

Select Case UCase(sArg)
Case UCase("/logonAlachua")

                strWksTimestamp = UserStringPrint & Chr(9) & Now & Chr(9) & "LOGON"
                strUserTimestamp = strComputer & Chr(9) & Now & Chr(9) & "LOGON"
                logType = "logon"

                '*** Checks to see if the workstation is a laptop or a desktop ***'
                
                'if InStr(1,strComputer,"rtix-nt",1) = 1 OR InStr(1,strComputer,"rtix-ws",1) = 1 OR InStr(1,strComputer,"rtix-cc",1) = 1 OR InStr(1,strComputer,"eng",1) = 1 OR InStr(1,strComputer,"dstx",1) = 1 OR InStr(1,strComputer,"dswi",1) = 1 OR InStr(1,strComputer,"rti-ws",1) = 1 then
                '               strWksType = "desktop"
                'end if
                
                if InStr(1,strComputer,"rtix-lp",1) = 1 OR InStr(1,strComputer,"lp-",1) = 1 OR InStr(1,strComputer,"rtix-tspr",1) = 1 then
                                strWksType = "laptop"
                else
                                strWksType = "desktop"
                end if

                For Each GroupObj In UserObj.Groups
                                Select Case GroupObj.Name

                                                '**************************'
                                                '******** Alachua *********'
                                                '**************************'

                                                Case "Global Communications"
                                                                fGlobalCommDrive
                                                Case "Information Technology"
                                                                fITDrive                
                                                Case "Human Resources"
                                                                fHRDrive                              
                                                Case "Legal"
                                                                fLegalDrive
                                                Case "R&D"
                                                                fRandDDrive

                                                Case "Rtix.com_Alachua_Clinical Projects"
                                                                fClinicalDrive


                
                                                '**************************'
                                                '****** APPLICATIONS ******'
                                                '**************************'

                                                Case "Solomon 6 Users"
                                                                fApplySolomon 
                                                Case "TTrack Users"
                                                                if strWksType = "desktop" then
                                                                                'msgbox "Install TTrack icon"
                                                                                fUpdateTTrackIcon         
                                                                end if
                                                Case "FRx Drill Down"
                                                                fApplySolomon 

                                                Case else
                                                                'nothing goes here

                                End Select
                Next

                fAddShortcutsNetwork
                fAddShortcutsDesktop
                fWriteLogEntries


Case UCase("/logonMarquette")

                strWksTimestamp = UserStringPrint & Chr(9) & Now & Chr(9) & "LOGON"
                strUserTimestamp = strComputer & Chr(9) & Now & Chr(9) & "LOGON"
                logType = "logon"
                fAddDrivesMarquette

                For Each GroupObj In UserObj.Groups
                                Select Case GroupObj.Name

                                                '**************************'
                                                '******* Marquette ********'
                                                '**************************'

                                                Case "Shop"
                                                                fShopDrive
                                                Case "Product Managers"
                                                                fShopDrive
                                                Case "Engineering Full Control"
                                                                fEngDrive
                                                Case "Process_Engineering"
                                                                fEngDrive
                                                Case "ShopEng"
                                                                fEngDrive
                                                                fNewShopDrive
                                                Case "Quality Full"
                                                                fEngDrive
                                                Case "ShopFloor Full"
                                                                fNewShopDrive
                                                Case "ShopFloor Read"
                                                                fNewShopDrive
                                                Case "FileShares-DeptShares_Intellectual Property_IP_M"
                                                                fIPdrive

                                                '**************************'
                                                '******* Greenville *******'
                                                '**************************'

                                                Case "GRVAdministration"
                                                                fGrvAdminDrive
                                                Case "GRVMiddlemgnt"
                                                                fGrvMiddleDrive
                                                Case "GRVUppermgnt"
                                                                fGrvUpperDrive
                                                Case "GRVDocControl"
                                                                fGrvCalDrive
                                                Case "GRVPurchasing"
                                                                fGrvPurchDrive
                                                Case "GRVUsers"
                                                                fGrvPubDrive
                                                                fGrvDocsDrive
                                                                fGrvMaxDrive
                                                                fGrvOrthoDrive

                                                '**************************'
                                                '********* Other **********'
                                                '**************************'

                                                Case "Engineering Full - Austin"
                                                                fAustinEngDrive
                                                Case "Doc Control - Austin"
                                                                fAustinEngDrive
                                                Case "AustinJDrive"
                                                                fAustinEngDrive
                                                Case "WBNUsers"
                                                                fWoburnDrive
                                                Case "EuropeMap"
                                                                fEuroDrive
                                                Case "Sales Full - Austin"
                                                                fSalesAustinDrive
                                                Case "Logon Script - Europe"
                                                                fData1Drive
                                                                fNetherlandsDrive

                                                Case else
                                                                'Nothing goes here

                                End Select
                Next

                fAddShortcutsNetwork
                fAddShortcutsDesktop
                fWriteLogEntries

Case UCase("/logoff")

                strWksTimestamp = UserStringPrint & Chr(9) & Now & Chr(9) & "LOGOFF"
                strUserTimestamp = strComputer & Chr(9) & Now & Chr(9) & "LOGOFF"
                logType = "logoff"
                fDeleteShortcutsCommon
                fDeleteShortcutsDepartment
                fDeleteNetworkDrives
                fWriteLogEntries

Case Else
                msgbox "Manual script execution - use /logonAlachua, /logonMarquette, or /logoff to test."
End Select


'***********************************'
'******* END SCRIPTING BLOCK *******'
'***********************************'

'***********************************'
'*** BEGIN FUNCTION DECLARATIONS ***'
'***********************************'

'********************************************************************************'
'*** Windowstyle 1 = Normal / Windowstyle 3 = Maximize                                                 ***'
'*** SHELL32.dll, 88 refers to the standard icon we use for network shortcuts ***'
'*** pnidui.dll, 26 refers to the hi-res Windows 7 icon for network shortcuts ***'
'********************************************************************************'

Function fAddShortcutsDesktop

                set oShellLink1 = WshShell.CreateShortcut(strDesktop & "\RTI Projects.lnk")
                                oShellLink1.TargetPath = "\\rtix.com\alachua\projects"
                                oShellLink1.WindowStyle = 3
                                oShellLink1.IconLocation = "SHELL32.dll, 88"
                                oShellLink1.Description = "RTI Projects"
                                oShellLink1.WorkingDirectory = "\\rtix.com\alachua\projects"
                oShellLink1.Save

                set oShellLink2 = WshShell.CreateShortcut(strDesktop & "\Projects West.lnk")
                                oShellLink2.TargetPath = "\\rtix.com\alachua\projects west"
                                oShellLink2.WindowStyle = 3
                                oShellLink2.IconLocation = "SHELL32.dll, 88"
                                oShellLink2.Description = "Projects West"
                                oShellLink2.WorkingDirectory = "\\rtix.com\alachua\projects"
                oShellLink2.Save

end Function

Function fAddShortcutsNetwork

                set oShellLink3 = WshShell.CreateShortcut(strNetwork & "\RTI Projects.lnk")
                                oShellLink3.TargetPath = "\\rtix.com\alachua\projects"
                                oShellLink3.WindowStyle = 1
                                oShellLink3.IconLocation = "SHELL32.dll, 88"
                                oShellLink3.Description = "RTI Projects"
                                oShellLink3.WorkingDirectory = "\\rtix.com\alachua\projects"
                oShellLink3.Save

                set oShellLink4 = WshShell.CreateShortcut(strNetwork & "\Projects West.lnk")
                                oShellLink4.TargetPath = "\\rtix.com\alachua\projects west"
                                oShellLink4.WindowStyle = 1
                                oShellLink4.IconLocation = "SHELL32.dll, 88"
                                oShellLink4.Description = "Projects West"
                                oShellLink4.WorkingDirectory = "\\rtix.com\alachua\projects west"
                oShellLink4.Save

End Function

sub CreateDriveShortcut()

                WSHNetwork.MapNetworkDrive strDriveLetter, strTargetPath, True

                set oShellLinkQDrive = WshShell.CreateShortcut(strCreateShortcut)
                                oShellLinkQDrive.TargetPath = strTargetPath
                                oShellLinkQDrive.WindowStyle = "3"
                                oShellLinkQDrive.IconLocation = "pnidui.dll, 26"
                                oShellLinkQDrive.Description = strDescription
                                oShellLinkQDrive.WorkingDirectory = strTargetPath
                oShellLinkQDrive.Save

                strLogDriveQ = UserString & " has mapped the " & strDescription & " - (" & strDriveLetter & ")."
                WshShell.LogEvent 4, strLogDriveQ

end sub

sub FixTTrackLocalFile()

                If objFSO.FileExists("C:\Program Files\TTrack2000\ProductionDB.udl") Then
                                objFSO.DeleteFile("C:\Program Files\TTrack2000\ProductionDB.udl")
                                msgbox "Local file deleted"
                else
                                msgbox "Local file doesn't exist"
                End If

                If objFSO.FileExists("\\rtix.com\rti\TTrack Downloads\ProductionDB.udl") Then
                                objFSO.CopyFile "\\rtix.com\rti\TTrack Downloads\ProductionDB.udl", "C:\Program Files\TTrack2000\"
                                msgbox "Network file copied to local"
                else
                                msgbox "Network file doesn't exist"
                End If

end sub

Function fWriteLogEntries

                '*** Write Windows Event Log entries (Logon / Logoff and Departmental Drive) ***'
                strLogLogin = UserString & " has completed the " & fileDate & " version of the RTI Surgical " & LogType & " script."
                WshShell.LogEvent 4, strLogLogin
                
                '*** Write logon/logoff entries to text files on network drives ***'
                set objFile1 = objFSO.OpenTextFile(strUserFile, ForAppending, True)
                objFile1.WriteLine strUserTimestamp
                objFile1.Close
                
                set objFile2 = objFSO.OpenTextFile(strWksFile, ForAppending, True)
                objFile2.WriteLine strWksTimestamp
                objFile2.Close

End Function

Function fDeleteShortcutsCommon

                If objFSO.FileExists(strDesktop & "\Work.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Work.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\RTI Projects.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\RTI Projects.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Projects West.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Projects West.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\ARW Projects.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\ARW Projects.lnk")
                End If

                If objFSO.FileExists(strNetwork & "\RTI Projects.lnk") Then
                                objFSO.DeleteFile(strNetwork & "\RTI Projects.lnk")
                End If

                If objFSO.FileExists(strNetwork & "\Projects West.lnk") Then
                                objFSO.DeleteFile(strNetwork & "\Projects West.lnk")
                End If

                If objFSO.FileExists(strNetwork & "\ARW Projects.lnk") Then
                                objFSO.DeleteFile(strNetwork & "\ARW Projects.lnk")
                End If

End Function

Function fDeleteShortcutsDepartment

                If objFSO.FileExists(strDesktop & "\Corporate Communications.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Corporate Communications.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\IT Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\IT Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Human Resources.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Human Resources.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Shop Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Shop Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Engineering Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Engineering Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\New Shop Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\New Shop Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Admin Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Admin Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Middle Mgmt Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Middle Mgmt Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Upper Mgmt Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Upper Mgmt Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Calman3 Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Calman3 Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Purchasing Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Purchasing Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Public Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Public Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Doc Control Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Doc Control Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Maxtor Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Maxtor Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Orthobiologics Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Orthobiologics Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Spine Fusion Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Spine Fusion Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Woburn Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Woburn Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Europe Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Europe Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Sales Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Sales Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Solomon Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Solomon Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Data1.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Data1 Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\Netherlands Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Netherlands Drive.lnk")
                End If

                If objFSO.FileExists(strDesktop & "\RandD Drive.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\RandD Drive.lnk")
                End If

        If objFSO.FileExists(strDesktop & "\Clinical Projects.lnk") Then
                                objFSO.DeleteFile(strDesktop & "\Clinical Projects.lnk")
                End If

End Function

Function fDeleteNetworkDrives

                If (objFSO.DriveExists("i:") = True) then
                                wshNetwork.RemoveNetworkDrive "i:", True, True
                end if

                If (objFSO.DriveExists("j:") = True) then
                                wshNetwork.RemoveNetworkDrive "j:", True, True
                end if

                If (objFSO.DriveExists("k:") = True) then
                                wshNetwork.RemoveNetworkDrive "k:", True, True
                end if

                If (objFSO.DriveExists("m:") = True) then
                                wshNetwork.RemoveNetworkDrive "m:", True, True
                end if

                If (objFSO.DriveExists("n:") = True) then
                                wshNetwork.RemoveNetworkDrive "n:", True, True
                end if

                If (objFSO.DriveExists("p:") = True) then
                                wshNetwork.RemoveNetworkDrive "p:", True, True
                end if

                If (objFSO.DriveExists("q:") = True) then
                                wshNetwork.RemoveNetworkDrive "q:", True, True
                end if

                If (objFSO.DriveExists("r:") = True) then
                                wshNetwork.RemoveNetworkDrive "r:", True, True
                end if

                If (objFSO.DriveExists("s:") = True) then
                                wshNetwork.RemoveNetworkDrive "s:", True, True
                end if

                If (objFSO.DriveExists("u:") = True) then
                                wshNetwork.RemoveNetworkDrive "u:", True, True
                end if

                If (objFSO.DriveExists("w:") = True) then
                                wshNetwork.RemoveNetworkDrive "w:", True, True
                end if

                If (objFSO.DriveExists("x:") = True) then
                                wshNetwork.RemoveNetworkDrive "x:", True, True
                end if

                If (objFSO.DriveExists("y:") = True) then
                                wshNetwork.RemoveNetworkDrive "y:", True, True
                end if

                If (objFSO.DriveExists("z:") = True) then
                                wshNetwork.RemoveNetworkDrive "z:", True, True
                end if

end Function

Function fApplySolomon

                strCreateShortcut = strDesktop & "\Solomon Drive.lnk"
                strTargetPath = "\\rtix.com\RTI\Solomon"
                strDescription = "Solomon Drive"
                strDriveLetter = "S:"
                call CreateDriveShortcut()

End Function

Function fUpdateTTrackIcon

                strCreateShortcut = strDesktop & "\Work.lnk"

                set oShellLinkTT = WshShell.CreateShortcut(strCreateShortcut)
                                oShellLinkTT.TargetPath = "\\rtix.com\RTI\TTrack Downloads\AppShell\AppShell.exe"
                                oShellLinkTT.WindowStyle = "1"
                                oShellLinkTT.IconLocation = "\\rtix.com\RTI\TTrack Downloads\AppShell\AppShell.exe, 0"
                                oShellLinkTT.Description = "Work"
                                oShellLinkTT.WorkingDirectory = "\\rtix.com\RTI\TTrack Downloads\AppShell\"
                oShellLinkTT.Save
                
                'call CreateDriveShortcut()
                'call FixTTrackLocalFile()

                strLogApp = UserString & " has the updated shortcut for TTrack. (Runs from network)"
                WshShell.LogEvent 4, strLogApp

End Function








'*** Marquette ***'






Function fAddDrivesMarquette

                WSHNetwork.MapNetworkDrive "S:", "\\rtix.com\Marquette\Site Share", True
                WSHNetwork.MapNetworkDrive "W:", "\\rtix.com\Marquette\Department Shares", True
                WSHNetwork.MapNetworkDrive "Y:", "\\rtix.com\Marquette\Public Share", True

End function

Function fShopDrive

                strTargetPath = "\\rtix.com\Marquette\Shop Floor"
                strCreateShortcut = strDesktop & "\Shop Drive.lnk"
                strDescription = "Shop Drive"
                strDriveLetter = "K:"
                call CreateDriveShortcut()

End Function

Function fEngDrive

                strTargetPath = "\\rtix.com\Marquette\Engineering"
                strCreateShortcut = strDesktop & "\Engineering Drive.lnk"
                strDescription = "Engineering Drive"
                strDriveLetter = "R:"
                call CreateDriveShortcut()

End Function

Function fNewShopDrive

                strTargetPath = "\\rtix.com\Marquette\New Shop Floor"
                strCreateShortcut = strDesktop & "\New Shop Drive.lnk"
                strDescription = "New Shop Drive"
                strDriveLetter = "M:"
                call CreateDriveShortcut()

End Function

Function fIPdrive

                strTargetPath = "\\rtix.com\Marquette\Intellectual Property"
                strCreateShortcut = strDesktop & "\IP Drive.lnk"
                strDescription = "IP Drive"
                strDriveLetter = "Z:"
                call CreateDriveShortcut()

End Function

'*** Alachua ***'

Function fGlobalCommDrive

                strCreateShortcut = strDesktop & "\Corporate Communications.lnk"
                strTargetPath = "\\rtix.com\general admin\Corporate Communications"
                strDescription = "Corporate Communications Drive"
                strDriveLetter = "Q:"
                call CreateDriveShortcut()

End Function

Function fITDrive

                strCreateShortcut = strDesktop & "\IT Drive.lnk"
                strTargetPath = "\\rtix.com\general admin\IT"
                strDescription = "IT Drive"
                strDriveLetter = "Q:"
                call CreateDriveShortcut()

End Function

Function fHRDrive

                strCreateShortcut = strDesktop & "\Human Resources.lnk"
                strTargetPath = "\\rtix.com\general admin\Human Resources"
                strDescription = "Human Resources Drive"
                strDriveLetter = "Q:"
                call CreateDriveShortcut()

End Function

Function fLegalDrive

                strCreateShortcut = strDesktop & "\Legal Drive.lnk"
                strTargetPath = "\\rtix.com\general admin\Legal"
                strDescription = "Legal Drive"
                strDriveLetter = "Q:"
                call CreateDriveShortcut()

End Function

Function fRandDDrive

                strCreateShortcut = strDesktop & "\RandD Drive.lnk"
                strTargetPath = "\\rtix.com\alachua\research and development"
                strDescription = "R&D Drive"
                strDriveLetter = "Q:"
                call CreateDriveShortcut()

End Function



Function fClinicalDrive

                strCreateShortcut = strDesktop & "\Clinical Projects.lnk"
                strTargetPath = "\\rtix.com\Alachua\Clinical Projects"
                strDescription = "Clinical Projects Drive"
                strDriveLetter = "N:"
                call CreateDriveShortcut()

End Function



'*** Greenville ***'

Function fGrvAdminDrive

                strCreateShortcut = strDesktop & "\Admin Drive.lnk"
                strTargetPath = "\\rtix.com\Greenville\Administration"
                strDescription = "Admin Drive"
                strDriveLetter = "N:"
                call CreateDriveShortcut()

End Function

Function fGrvMiddleDrive

                strCreateShortcut = strDesktop & "\Middle Mgmt Drive.lnk"
                strTargetPath = "\\rtix.com\Greenville\Middle Management"
                strDescription = "Middle Mgmt Drive"
                strDriveLetter = "Z:"
                call CreateDriveShortcut()

End Function

Function fGrvUpperDrive

                strCreateShortcut = strDesktop & "\Upper Mgmt Drive.lnk"
                strTargetPath = "\\rtix.com\Greenville\Upper Management"
                strDescription = "Upper Mgmt Drive"
                strDriveLetter = "U:"
                call CreateDriveShortcut()

End Function

Function fGrvCalDrive

                strCreateShortcut = strDesktop & "\Calman3 Drive.lnk"
                strTargetPath = "\\rtix.com\Greenville\Calman3"
                strDescription = "Calman3 Drive"
                strDriveLetter = "U:"
                call CreateDriveShortcut()

End Function

Function fGrvPurchDrive

                strCreateShortcut = strDesktop & "\Purchasing Drive.lnk"
                strTargetPath = "\\rtix.com\Greenville\Purchasing"
                strDescription = "Purchasing Drive"
                strDriveLetter = "I:"
                call CreateDriveShortcut()

End Function

Function fGrvPubDrive

                strCreateShortcut = strDesktop & "\Public Drive.lnk"
                strTargetPath = "\\rtix.com\Greenville\Public Share"
                strDescription = "Public Drive"
                strDriveLetter = "P:"
                call CreateDriveShortcut()

End Function

Function fGrvDocsDrive

                strCreateShortcut = strDesktop & "\Doc Control Drive.lnk"
                strTargetPath = "\\rtix.com\Greenville\Document Control"
                strDescription = "Document Control Drive"
                strDriveLetter = "Q:"
                call CreateDriveShortcut()

End Function

Function fGrvMaxDrive

                strCreateShortcut = strDesktop & "\Maxtor Drive.lnk"
                strTargetPath = "\\rtix.com\Greenville\Maxtor"
                strDescription = "Maxtor Drive"
                strDriveLetter = "M:"
                call CreateDriveShortcut()

End Function


'*** Other ***'

Function fAustinEngDrive

                strCreateShortcut = strDesktop & "\Spine Fusion Drive.lnk"
                strTargetPath = "\\rtix.com\Austin\Engineering"
                strDescription = "Spine Fusion Drive"
                strDriveLetter = "J:"
                call CreateDriveShortcut()

End Function

Function fWoburnDrive

                strCreateShortcut = strDesktop & "\Woburn Drive.lnk"
                strTargetPath = "\\rtix.com\siteshares\Woburn"
                strDescription = "Woburn Drive"
                strDriveLetter = "Z:"
                call CreateDriveShortcut()

End Function

Function fEuroDrive

                strCreateShortcut = strDesktop & "\Europe Drive.lnk"
                strTargetPath = "\\10.1.20.3\Data1"
                strDescription = "Europe Drive"
                strDriveLetter = "Z:"
                call CreateDriveShortcut()

End Function

Function fSalesAustinDrive

                strCreateShortcut = strDesktop & "\Sales Drive.lnk"
                strTargetPath = "\\rtix.com\Austin\Sales and Marketing"
                strDescription = "Sales and Marketing Drive"
                strDriveLetter = "K:"
                call CreateDriveShortcut()

End Function

Function fData1Drive

                strCreateShortcut = strDesktop & "\Data1 Drive.lnk"
                strTargetPath = "\\rtix.com\europe\Data1"
                strDescription = "Data1 Drive"
                strDriveLetter = "Z:"
                call CreateDriveShortcut()

End Function

function fNetherlandsDrive

                strCreateShortcut = strDesktop & "\Netherlands Drive.lnk"
                strTargetPath = "\\rtix.com\europe\Netherlands"
                strDescription = "Netherlands Drive"
                strDriveLetter = "X:"
                call CreateDriveShortcut()

End Function

'***********************************'
'**** END FUNCTION DECLARATIONS ****'
'***********************************'

'Clean up 
Set oArgs = Nothing
Set objFile1 = Nothing
Set objFile2 = Nothing
Set objFSO = Nothing
set UserObj = Nothing
set GroupObj = Nothing
set WSHNetwork = Nothing
set DomainString = Nothing
set WSHSHell = Nothing
set WinDir = Nothing
set UserString = Nothing
set strComputer = Nothing
set strDesktop = Nothing
set strNetwork = Nothing
set strUserTimestamp = Nothing
set strWksTimestamp = Nothing

'Quit the Script
wscript.quit
