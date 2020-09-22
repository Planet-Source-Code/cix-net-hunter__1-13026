Attribute VB_Name = "modRipHTML"
Global RippingDone As Boolean
Global TimeOutS As Integer
Global MellemRum As String
Dim NodeX As Node
Dim K As Long
Function RipHtmlForUrLs(URL As String, OutputList As ListBox, Net As Inet, StartPoint As Long, SecondFile As String, Frm As Form, Capline As String) As String
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
' Her downloader den fra internettet
Download = ""
Download = Space$(Len(SecondFile))
Download = SecondFile
SecondFile = ""
Else
 Download = Net.OpenURL(URL, icString) ': RipHtmlForUrLs = "Downloading The HTML File"
 Close #9
 Open App.Path & "\sites\site" & K & ".htm" For Output As 9
 K = K + 1
 Print #9, Download
 Close #9
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
If InStr(StartPoint, LCase(Download), "href", vbTextCompare) Then
  StartPoint = InStr(StartPoint, LCase(Download), "href", vbTextCompare) + 6
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
  If StopAll = True Then Exit Function
   If InStr(1, LCase(Tmp), "mailto", vbTextCompare) Then
   Else
   If LCase(Right$(Tmp, 4)) = ".jpg" Or LCase(Right$(Tmp, 4)) = ".gif" Or LCase(Right$(Tmp, 4)) = ".bmp" Then
    Tmp = GetRealAdress2(Tmp, URL)
    Set NodeX = Hunter.Tree1.Nodes.Add("images", tvwChild, Tmp, Tmp, 3)
 
    Hunter.Download.AddItem Tmp
   Else
    Tmp = GetRealAdress2(Tmp, URL)
    Set NodeX = Hunter.Tree1.Nodes.Add("urls", tvwChild, Tmp, Tmp, 5)
    OutputList.AddItem Tmp
   End If
   End If
  
  Tmp = ""
  StartPoint = EndPoint + 1
  DoEvents
  RipHtmlForUrLs URL, OutputList, Net, StartPoint, Download, Frm, Capline
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


'#############################################################################################################
' Function TrimUrl
' This Function is to Trim the urls form the function RipHtmlForUrLs.
'
' You have no need for this function, since it only here for use in the function RipHtmlForUrLs
' But if you do, you can call it like this.
' Dim URL as String
' Url = TrimUrl(Url, 1)
'
'#############################################################################################################

Function TrimUrl(URL As String, StartPoint As Long) As String
Dim X As Long, Y As Long
If InStr(StartPoint, URL, ">", vbTextCompare) Then
X = InStr(StartPoint, URL, ">", vbTextCompare) + 1
URL = Mid$(URL, X)
If InStr(StartPoint, URL, "</", vbTextCompare) Then
Y = InStr(StartPoint, URL, "</", vbTextCompare)
URL = Mid$(URL, Y)
End If
TrimUrl URL, X + 1
Else
TrimUrl = URL
End If
End Function

