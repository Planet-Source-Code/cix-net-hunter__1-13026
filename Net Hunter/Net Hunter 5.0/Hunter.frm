VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Hunter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " VirusTeam 2000 - Net Hunter Version 1.2"
   ClientHeight    =   3444
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   5952
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Hunter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3444
   ScaleWidth      =   5952
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   840
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":1036
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":1488
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":18DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Hunter.frx":1D2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tree1 
      Height          =   2172
      Left            =   0
      TabIndex        =   21
      Top             =   3480
      Width           =   6012
      _ExtentX        =   10605
      _ExtentY        =   3831
      _Version        =   393217
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.ListBox lstpicdown2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.ListBox Download 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.ListBox lstPicdown 
      Height          =   264
      Left            =   6120
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   135
      Left            =   50
      ScaleHeight     =   108
      ScaleWidth      =   5880
      TabIndex        =   12
      Top             =   3045
      Width           =   5900
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   -40
      TabIndex        =   1
      Top             =   840
      Width           =   6160
      Begin VB.CommandButton Command5 
         Caption         =   "&More"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   5040
         TabIndex        =   22
         Top             =   1680
         Width           =   852
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use realname"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Top             =   1440
         Width           =   1815
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   5400
         Top             =   720
         _ExtentX        =   995
         _ExtentY        =   995
         _Version        =   393216
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   5400
         Top             =   720
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Download"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   5400
         Picture         =   "Hunter.frx":2046
         Top             =   240
         Width           =   384
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   6000
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   0
         Top             =   0
         Width           =   6015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Url :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Level(s) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Output Path :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Label lblsites 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   50
      TabIndex        =   20
      Top             =   3200
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.virusteam.cjb.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   3480
      MouseIcon       =   "Hunter.frx":2350
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   0
      Width           =   2400
   End
   Begin VB.Label lblAdress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   60
   End
   Begin VB.Label Statusline 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Downloading :"
      Height          =   225
      Left            =   45
      TabIndex        =   11
      Top             =   2790
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   0
      Top             =   2760
      Width           =   6015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Hunter "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "Hunter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Xnode As Node
Dim I As Long, J As Long
Dim Levels As Integer
Dim Nrlevel As Integer
Dim Usename As Integer
Dim StopAll As Boolean


Private Sub Check1_Click()
If Check1.Value = Checked Then
 Usename = 1
Else
 Usename = 2
End If
End Sub

Private Sub Command1_Click()
Folders.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
StopAll = False
Timer1.Enabled = True
If Right$(Text3.Text, 1) = "\" Then Text3.Text = Mid$(Text3.Text, 1, Len(Text3.Text) - 1)
lstPicdown.Clear
lstpicdown2.Clear
Download.Clear
Nrlevel = 0
lstpicdown2.AddItem Text1.Text
Statusline.Caption = "Makeing DownList.. at level (" & (Nrlevel) + 1 & ")"
lblAdress.Caption = GetAliasAdress(Text1.Text)
Set Xnode = Tree1.Nodes.Add(, , "site", GetAliasAdress(Text1.Text), 4, 4)
Set Xnode = Tree1.Nodes.Add("site", tvwChild, "urls", "HTML", 6)
Set Xnode = Tree1.Nodes.Add("site", tvwChild, "images", "Images", 6)
Levels = Val(Text2.Text)
RipHtmlForUrLs Text1.Text, lstPicdown, Inet1, 1, 0, Me, Me.Caption
WaitforIT:
Statusline.Caption = "Filtering list.."
RemoveDupes lstPicdown
For I = 0 To lstPicdown.ListCount - 1
 If StopAll = True Then Exit Sub
 lstPicdown.ListIndex = I
 If lstPicdown.Text = "" Then
  lstPicdown.RemoveItem I
 Else
  lstpicdown2.AddItem GetRealAdress(lstPicdown.Text)
 End If
 DoEvents
Next I
 lstPicdown.Clear
If Nrlevel = Levels Then
 Nrlevel = Nrlevel + 1
 RemoveDupes lstpicdown2
 Statusline.Caption = "Finding Image(s).."
 RippingDone = False
 lstpicdown2.AddItem Text1.Text
 For I = 0 To lstpicdown2.ListCount - 1
  If StopAll = True Then Exit Sub
  lstpicdown2.ListIndex = I
  Statusline.Caption = "Finding Image(s) in (Page " & I & ")"
  RipHtmlForMedia lstpicdown2.Text, Download, Inet1, 1, 0, Me, Me.Caption
  DoEvents
  Do While RippingDone = False
   DoEvents
  Loop
  Next I
  Statusline.Caption = "Filtering list.."
  RemoveDupes Download
  Statusline.Caption = "Download Image(s) number(" & J & ") "
  Dim FilByte() As Byte
  ProcessLineMaxValue = Download.ListCount
  For J = 0 To Download.ListCount - 1
  If StopAll = True Then Exit Sub
  Download.ListIndex = J
  Statusline.Caption = "Download Image(s) number(" & J & ") "
  ProcessLine_ValueChange Hunter.Picture1, (J + 1)
   DoEvents
   If Download.Text = "" Then
   Else
    FilByte() = Inet1.OpenURL(Download.Text, icByteArray)
   
    Close #9
    If Check1.Value = Checked Then
     Open Text3.Text & "\" & GetShortname(Download.Text) For Binary As #9
    Else
     Open Text3.Text & "\Image" & J & Right$(Download.Text, 4) For Binary As #9
    End If
     Put #9, , FilByte()
    Close #9
   End If
   Next J
   Timer1.Enabled = False
 Statusline.Caption = "Download done.."
 Exit Sub
Else
On Error Resume Next
 For I = 0 To lstpicdown2.ListCount - 1
 If StopAll = True Then Exit Sub
  lstpicdown2.ListIndex = I
  If lstpicdown2.Text = "" Then
     lstpicdown2.RemoveItem I
  ElseIf LCase(Right$(lstpicdown2.Text, 4)) = ".jpg" Or LCase(Right$(lstpicdown2.Text, 4)) = ".gif" Or LCase(Right$(lstpicdown2.Text, 4)) = ".bmp" Then
  Download.AddItem lstpicdown2.Text
  lstpicdown2.RemoveItem I
  Else
  lstPicdown.AddItem lstpicdown2.Text
  End If
  DoEvents
 Next I
 
Statusline.Caption = "Makeing DownList.. at level (" & (Nrlevel) + 1 & ")"
For I = 0 To lstpicdown2.ListCount - 1
   If StopAll = True Then Exit Sub
  Statusline.Caption = "Makeing DownList from (Page " & I & ") at Level (" & (Nrlevel) + 1 & ")"
  lstpicdown2.ListIndex = I
  RipHtmlForUrLs lstpicdown2.Text, lstPicdown, Inet1, 1, 0, Me, Me.Caption
  DoEvents
  Do While RippingDone = False
   DoEvents
  Loop
Next I
lstpicdown2.Clear
 Nrlevel = Nrlevel + 1
End If
If Nrlevel <= Levels Then GoTo WaitforIT

End Sub

Private Sub Command3_Click()
StopAll = True

End Sub

Private Sub Command4_Click()
SaveSetting "Net Hunter", "Settings", "LastUrl", Text1.Text
SaveSetting "Net Hunter", "Settings", "Levels", Text2.Text
SaveSetting "Net Hunter", "Settings", "SaveDir", Text3.Text
SaveSetting "Net Hunter", "Settings", "Names", Usename
Inet1.Cancel
End
End Sub

Private Sub Command5_Click()
If Command5.Caption = "&Less" Then
 Command5.Caption = "&More"
 Me.Height = 3744
Else
 Me.Height = 5964
 Command5.Caption = "&Less"
End If
End Sub

Private Sub Form_Load()
Text1.Text = GetSetting("Net Hunter", "Settings", "LastUrl", "http://www.site.com/pictures.htm")
Text2.Text = GetSetting("Net Hunter", "Settings", "Levels", 1)
Text3.Text = GetSetting("Net Hunter", "Settings", "SaveDir", "c:\temp")
Usename = GetSetting("Net Hunter", "Settings", "Names", 2)
If Usename = 1 Then
 Check1.Value = Checked
Else
  Check1.Value = Unchecked
End If
End Sub

Private Sub Label5_Click()
GotoURL Label5.Caption
End Sub

Private Sub lstPicdown_Validate(Cancel As Boolean)
lblsites.Caption = "Pages : " & lstPicdown.ListCount & " images : " & Download.ListCount
End Sub

Private Sub lstpicdown2_Validate(Cancel As Boolean)
lblsites.Caption = "Pages : " & lstpicdown2.ListCount & " images : " & Download.ListCount
End Sub

Private Sub Timer1_Timer()
DoEvents
If lstpicdown2.ListCount = 0 Then
 lblsites.Caption = "Pages : " & lstPicdown.ListCount & " images : " & Download.ListCount
Else
 lblsites.Caption = "Pages : " & lstpicdown2.ListCount & " images : " & Download.ListCount
End If
End Sub

