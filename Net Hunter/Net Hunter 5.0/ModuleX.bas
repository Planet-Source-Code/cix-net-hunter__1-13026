Attribute VB_Name = "ModuleX"
' ------------------------------------------------------------------------------------------------------------
' If you are Looking for some Help to this codes, then may god be with you.
' I am sorry i did't comment my codes.
' But i made a littel Eksample on how to use theis Functions.
' Good luck with the codes.
'
' Cix -Virusteam
' http://www.virusteam.cjb.net
'
' ------------------------------------------------------------------------------------------------------------
'#############################################################################################################
'
' Here are just a few declaretions.

Private Declare Function mciSendString Lib "winmm.dll" Alias _
"mciSendStringA" (ByVal lpstrCommand As String, ByVal _
lpstrReturnString As String, ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const EXIT_LOGOFF = 0
Private Const EXIT_SHUTDOWN = 1
Private Const EXIT_REBOOT = 2
Private Const conSwNormal = 1
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Const MAX_PATH = 260
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public ProcessLineMaxValue As Long
'#############################################################################################################
' Function ReadINIFile
' This Function is to List all Settings in a ini file under det given section.
'
' To use this function do like this
' ReadIniFile "C:\windows\system\win.ini", "Network", List1
'
'#############################################################################################################

Function ReadIniFile(Filename As String, Section As String, Settings As ListBox)
On Error Resume Next
Dim Tmp As String
Dim FileNumber As Integer
Dim ReadLinjer As Boolean
FileNumber = FreeFile
Open Filename For Input As #FileNumber
ReadLinjer = False
Do While Not EOF(FileNumber)
 Line Input #FileNumber, Tmp
If Loc(FileNumber) = LOF(FileNumber) Then
' Close #FileNumber
 'Exit Function
Else
 If ReadLinjer = True Then
  If InStr(1, LCase(Tmp), "[", vbTextCompare) Then
   ReadLinjer = False
  Else
   Settings.AddItem Tmp
  End If
 Else
  If InStr(1, LCase(Tmp), "[" & LCase(Section), vbTextCompare) Then
   ReadLinjer = True
  Else

  End If
 End If
End If
Loop
Close #FileNumber
End Function

'#############################################################################################################
' Function ListSectionsInINIFile
' This Function is to List all Sections in a IniFile.
'
' To use this function do like this
' ListSectionsInINIFile "C:\windows\system\win.ini", List1
'
'#############################################################################################################

Function ListSectionsInINIFile(Filename As String, Sections As ListBox)
Dim Tmp As String
Dim FileNumber As Integer
Dim I As Integer, J As Integer
FileNumber = FreeFile
Open Filename For Input As #FileNumber
Do While Not EOF(FileNumber)
 Line Input #FileNumber, Tmp
If Loc(FileNumber) = LOF(FileNumber) Then
Close #FileNumber
Exit Function
 Else
  If InStr(1, LCase(Tmp), "[", vbTextCompare) Then
   I = InStr(1, LCase(Tmp), "[", vbTextCompare) + 1
   J = InStr(I, LCase(Tmp), "]", vbTextCompare)
   Tmp = Mid$(Tmp, I, (J - I))
   Sections.AddItem Tmp
 Else

 End If
End If
Loop
Close #FileNumber
End Function

'#############################################################################################################
' Function FindFiles
' This Function is to List all Files on your/(the kliens) Harddrive.
' with the given filter such as (*.jpg)
' To use this function do like this
' FindFiles "C:\", "*.ini", List1
'
'#############################################################################################################

Function FindFiles(StartPath As String, FileFilter As String, OutputList As ListBox)
On Error Resume Next
Dim Tmp As String
Dim I As Long, J As Long

If Right$(StartPath, 1) = "\" Then
 StartPath = Mid$(StartPath, 1, Len(StartPath) - 1)
End If
Tmp = Dir$(StartPath & "\" & FileFilter, vbNormal)
If Tmp = "" Then
Else
 OutputList.AddItem StartPath & "\" & Tmp
 Do While Len(Tmp)
  Tmp = Dir$
  If Tmp = "" Then
  Else
  OutputList.AddItem StartPath & "\" & Tmp
  End If
  
 Loop
End If
Tmp = Dir$(StartPath & "\*.*", vbDirectory)
If Tmp = "" Then
Else
If Tmp = "." Or Tmp = ".." Then
 Tmp = Dir$
 If Tmp = "." Or Tmp = ".." Then
  Tmp = Dir$
 End If
End If
 I = 1
 Do While Len(Tmp)
  Tmp = Dir$
  I = I + 1
 Loop
DoEvents
 I = I - 1
ReDim dirFolders(1 To (I)) As String
Tmp = Dir$(StartPath & "\*.*", vbDirectory)
If Tmp = "." Then
Tmp = Dir$
dirFolders(1) = Dir$
Else
End If
If I > 1 Then
 For J = 2 To I
  dirFolders(J) = Dir$
 Next J
Else
End If
 For J = 1 To I
 FindFiles StartPath & "\" & dirFolders(J), FileFilter, OutputList
 Next J
End If

End Function

'#############################################################################################################
' Function ProcessLine_ValueChange
' This Function is to draw your own Progessbar, It nice to make it your own color.
'
' To use this function do like this
'
' Dim Tmp as String
' Open "somefile" for input as #1
' ProcessLineMaxValue = Lof(1)
' Do while not EOF(1)
'  Line Input#1, Tmp
'  ProcessLine_ValueChange Picture1, Loc(1)
' Loop
'
'#############################################################################################################


Function ProcessLine_ValueChange(Pic As PictureBox, ProcessValue As Long) As String

Dim PicLen As Long
Dim PicDrawWidth As Long
Dim Procents As Long
Dim PicDraw As Long
Pic.AutoRedraw = -1
PicLen = Pic.ScaleWidth
Pic.Cls
Pic.DrawMode = 10
PicDrawWidth = PicLen / 105
Pic.DrawWidth = PicDrawWidth
Procents = Int((ProcessValue * 100) / ProcessLineMaxValue)
ProcessLine_ValueChange = Procents & "%"
If Procents = 100 Or ProcessValue = ProcessLineMaxValue Then
 Pic.Cls
 Exit Function
End If
PicDraw = Int((PicDrawWidth * Procents))
Pic.CurrentX = (Pic.ScaleWidth - Pic.TextWidth(Procents & "%")) / 2
Pic.CurrentY = (Pic.ScaleHeight - Pic.TextHeight(Procents & "%")) / 2
Pic.Print Procents & "%"
Pic.Line (0, 0)-(PicDraw, Pic.ScaleHeight), , BF
Pic.Refresh
End Function

'#############################################################################################################
' Function Filesize
' This Function is to get the size of a file in kbytes.
'
' To use this function do like this
'
' msgbox "the file C:\autoexec.bat's size is " & FileSize("C:\autoexec.bat")
'
'#############################################################################################################

Function FileSize(File As String) As String
Dim LSize As String
If File = "" Then
FileSize = ""
Exit Function
End If
LSize = FileLen(File)
FileSize = LSize / 100
FileSize = FileSize & " KB"
End Function

'#############################################################################################################
' Sub OpenCD
' This Sub Opens your CD-drive
'
' To use this function do like this
'Private Sub Form1_Click()
' OpenCD
'End Sub
'#############################################################################################################

Sub OpenCD()
retvalue = mciSendString("set CDAudio door open", _
returnstring, 127, 0)
End Sub

'#############################################################################################################
' Sub CloseCD
' This Sub Closes your CD-drive
'
' To use this function do like this
'Private Sub Form1_Click()
' CloseCD
'End Sub
'#############################################################################################################
Sub CloseCD()
retvalue = mciSendString("set CDAudio door closed", _
returnstring, 127, 0)
End Sub

'#############################################################################################################
' Sub SystemReboot
' This Sub Reboots Your PC
'
' To use this function do like this
'Private Sub Form1_Click()
' SystemReboot
'End Sub
'#############################################################################################################

Sub SystemREBOOT()
   Call ExitWindowsEx(2, 0)
End Sub

'#############################################################################################################
' Sub SystemShutDown
' This Sub will Shutdown Your PC
'
' To use this function do like this
'Private Sub Form1_Click()
' SystemShutDown
'End Sub
'#############################################################################################################

Sub SystemSHUTDOWN()
   Call ExitWindowsEx(1, 0)
End Sub

'#############################################################################################################
' Sub SystemLogOff
' This Sub will Logoff Your current User
'
' To use this function do like this
'Private Sub Form1_Click()
' SystemLogOff
'End Sub
'#############################################################################################################

Sub SystemLOGOFF()
   Call ExitWindowsEx(0, 0)
End Sub

'#############################################################################################################
' Function GotoUrl
' This function will Launce your web browser and goto the given URL
'
' To use this function do like this
'Private Sub Form1_Click()
' GotoUrl "http://www.hotmail.com"
'End Sub
'#############################################################################################################

Function GotoURL(URL As String) As String
ShellExecute hwnd, "open", URL, vbNullString, vbNullString, conSwNormal
End Function

'#############################################################################################################
' Function PlayWave
' This function will Play a wavefile on your pc.
'
' To use this function do like this
'Private Sub Form1_Click()
' PlayWave "C:\Media.wav"
'End Sub
'#############################################################################################################

Function PlayWave(File As String) As String
SoundFile$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
X% = sndPlaySound(SoundFile$, wFlags%)
End Function

'#############################################################################################################
' Function GetWindowsDir
' This function will Give you the dir of windows
'
' To use this function do like this
'Private Sub Form1_Click()
' dim Tmp as String
'GetWindowsDir(tmp)
'MsgBox "Your Windows's Path is " & tmp
'End Sub
'#############################################################################################################
Function GetWindowsDir(OutputDir As String) As String
Dim strBuffer As String
Dim lngReturn As Long
Dim strWindowsDirectory As String
strBuffer = Space$(MAX_PATH)
lngReturn = GetWindowsDirectory(strBuffer, MAX_PATH)
strWindowsDirectory = Left$(strBuffer, Len(strBuffer) - 1)
OutputDir = strWindowsDirectory
End Function

'#############################################################################################################
' Function TellComputerName
' This function will Give you the computers name
'
' To use this function do like this
'Private Sub Form1_Click()
' dim Tmp as String
'TellComputername(tmp)
'MsgBox "Your Computers name is " & tmp
'End Sub
'#############################################################################################################

Function TellComputerName(Output As String) As String
Dim strBuffer As String
  Dim lngBufSize As Long
  Dim lngStatus As Long
  lngBufSize = 255
  strBuffer = String$(lngBufSize, " ")
  lngStatus = GetComputerName(strBuffer, lngBufSize)
  If lngStatus <> 0 Then
    Output = Left(strBuffer, lngBufSize)
  End If
End Function
'#############################################################################################################
' Function ChangeComputerName
' This function will change the computername.
'
' To use this function do like this
'Private Sub Form1_Click()
' Dim Tmp as String
' Tmp =" My PC"
'ChangeComputername(tmp)
'MsgBox "Your Computer's name is set to " & tmp
'End Sub
'#############################################################################################################


Function ChangeComputerName(ComputerName As String) As String
Dim strNewComputerName As String
Dim lngReturn As Long
If ComputerName = "" Then
Exit Function
Else
strNewComputerName = ComputerName
lngReturn = SetComputerName(strNewComputerName)
End If
End Function

'#############################################################################################################
' Function DisableCTRLALTDELETE
' This function will Disable/Enable the Keys CTRL+ALT+DELETE.
'
' To use this function do like this
'Private Sub Form1_Click()
'DisableCTRLALTDELETE = True
'MsgBox "Your CTRL+ALT+DELETE Is Not Working "
'DisableCTRLALTDELETE = False
'MsgBox "Your CTRL+ALT+DELETE Is Now Working "
'End Sub
'#############################################################################################################

Function DisableCTRLALTDELETE(Enable As Boolean)
GD = SystemParametersInfo(97, Enable, CStr(1), 0)
End Function

