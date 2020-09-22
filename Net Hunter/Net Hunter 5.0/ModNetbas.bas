Attribute VB_Name = "ModNetbas"
Option Explicit
Dim NodeX As Node
Function GetShortname2(SName As String) As String
Dim I As Long
I = InstrRev(SName, "/", 1)
GetShortname2 = Mid$(SName, 1, I)
End Function
Function GetShortname(SName As String) As String
Dim I As Long
I = InstrRev(SName, "/", 1) + 1
GetShortname = Mid$(SName, I)
End Function
Function GetRealAdress(Site As String) As String
Dim Tmpsite As String
Dim I As Long
HaltErr:

If Left(Site, 1) = "/" Or Left(Site, 2) = "./" Then
 If Tmpsite = "" Then
  GetRealAdress = GetAliasAdress(Hunter.Text1.Text) & Mid$(Site, 2)
  Tmpsite = ""
 Else
  GetRealAdress = Tmpsite & Mid$(Site, 2)
  Tmpsite = ""
 End If
 
ElseIf Left(Site, 3) = "../" Then
 I = InstrRev(Hunter.Text1.Text, "/", 1)
 I = InstrRev(Hunter.Text1.Text, "/", I)
 Tmpsite = Mid$(Hunter.Text1.Text, 1, I)
 Site = Mid$(Site, 3)
 GoTo HaltErr
ElseIf LCase(Left$(Site, 5)) = "http:" Then
 GetRealAdress = Site
Else
 
 GetRealAdress = GetShortname2(Hunter.Text1.Text) & Site
End If
End Function

Function GetRealAdress2(Site As String, OldSite As String) As String
Dim Tmpsite As String
Dim I As Long
HaltErr:

If Left(Site, 1) = "/" Or Left(Site, 2) = "./" Then
 If Tmpsite = "" Then
  GetRealAdress2 = GetAliasAdress(OldSite) & Mid$(Site, 2)
  Tmpsite = ""
 Else
  GetRealAdress2 = Tmpsite & Mid$(Site, 2)
  Tmpsite = ""
 End If
 
ElseIf Left(Site, 3) = "../" Then
 I = InstrRev(OldSite, "/", 1)
 I = InstrRev(OldSite, "/", I)
 Tmpsite = Mid$(OldSite, 1, I)
 Site = Mid$(Site, 3)
 GoTo HaltErr
ElseIf LCase(Left$(Site, 5)) = "http:" Then
 GetRealAdress2 = Site
Else
If InStr(1, LCase(OldSite), ".htm") Or InStr(1, LCase(OldSite), ".shtm") Or InStr(1, LCase(OldSite), ".asp") Or InStr(1, LCase(OldSite), ".php") Then
 GetRealAdress2 = GetShortname2(OldSite) & Site
Else
 If Right$(OldSite, 1) = "/" Then
  GetRealAdress2 = GetShortname2(OldSite) & Site
 Else
  OldSite = OldSite & "/"
  GetRealAdress2 = GetShortname2(OldSite) & Site
 End If
End If
End If
End Function

Function GetAliasAdress(Site As String) As String
Dim I As Long
If Left$(Site, 4) = "http" Then
 I = InStr(8, Site, "/", vbTextCompare)
 GetAliasAdress = Mid$(Site, 1, I)
Else
I = InStr(1, Site, "/", vbTextCompare)
 GetAliasAdress = Mid$(Site, 1, I)
End If
End Function

Function RipHtmlForMedia(URL As String, OutputList As ListBox, Net As Inet, StartPoint As Long, SecondFile As String, Frm As Form, Capline As String) As String
On Error Resume Next
Dim Download As String, Tmp As String
Dim EndPoint As Long
Dim I As Long
If CancelOrder = True Then Exit Function
'ProcesBar.Visible = False
Capline = "Connecting to : " & URL
Capline = "Seaching for URL's in (" & URL & ")"

DoEvents
If StartPoint > 1 Then
'SecondFile = Space$(Len(Download))
Download = ""
Download = Space$(Len(SecondFile))
Download = SecondFile
SecondFile = ""
Else
 Download = Net.OpenURL(URL, icString) ': RipHtmlForUrLs = "Downloading The HTML File"
  'If Len(Download) > 6400000 Then Download = ""
 If Download = "" Then
  RippingDone = True
  Net.Cancel
  Download = ""
  Tmp = ""
  StartPoint = 1
  EndPoint = 1
  Exit Function
 End If
  
' Form1.Caption = "URL downloaded"
 'MsgBox Download
End If
ProcessLineMaxValue = Len(Download)
DoEvents
If InStr(StartPoint, LCase(Download), "<img src", vbTextCompare) Then
  StartPoint = InStr(StartPoint, LCase(Download), "<img src", vbTextCompare) + 10
If InStr(StartPoint, LCase(Download), """", vbTextCompare) Then
 EndPoint = InStr(StartPoint, LCase(Download), """", vbTextCompare)
Else
 If InStr(StartPoint, LCase(Download), ">", vbTextCompare) Then
  EndPoint = InStr(StartPoint, LCase(Download), ">", vbTextCompare)
 End If
  'EndPoint = InStr(StartPoint, LCase(Download), """", vbTextCompare)
 End If
  ProcessLine_ValueChange Hunter.Picture1, StartPoint
  Tmp = Mid$(Download, StartPoint, (EndPoint - StartPoint))
  Tmp = Trim$(Tmp)
  Tmp = TrimUrl(Tmp, 1)
  
   If InStr(1, LCase(Tmp), "mailto", vbTextCompare) Then
   Else
     Tmp = GetRealAdress2(Tmp, URL)
    OutputList.AddItem Tmp
    
   Set NodeX = Hunter.Tree1.Nodes.Add("images", tvwChild, Tmp, Tmp, 3)
   End If
  
  Tmp = ""
  StartPoint = EndPoint + 1
  DoEvents
  RipHtmlForMedia URL, OutputList, Net, StartPoint, Download, Frm, Capline
Else
  Capline = "Done..."
  Download = ""
  Tmp = ""
  StartPoint = 1
  EndPoint = 1
  ProcessLine_ValueChange Hunter.Picture1, ProcessLineMaxValue
  RippingDone = True
  Net.Cancel
End If
End Function
Function InstrRev(RealText As String, FindText As String, Start As Long) As Long
Dim I As Long
Dim StrTemp As String

For I = Len(RealText) To 1 Step -1
StrTemp = StrTemp & Mid$(RealText, I, 1)
DoEvents
Next I

InstrRev = InStr(Start, StrTemp, FindText)
InstrRev = (Len(RealText) - InstrRev) + 1
End Function

