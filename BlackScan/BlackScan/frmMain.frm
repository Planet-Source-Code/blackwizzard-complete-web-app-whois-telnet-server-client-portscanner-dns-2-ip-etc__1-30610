VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "..:: BlackScan V2.0 ::.."
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   HelpContextID   =   16761024
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox rtf 
      Height          =   30
      Left            =   5880
      TabIndex        =   19
      Top             =   7785
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":1CFA
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2520
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   6720
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A86
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9278
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin BlackScan.XPDesign XP 
      Height          =   7095
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   12515
      Begin VB.TextBox txtUpperBound 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3720
         TabIndex        =   9
         Text            =   "32676"
         ToolTipText     =   "Scanner jusqu'au port..."
         Top             =   1485
         Width           =   495
      End
      Begin VB.TextBox txtLowerBound 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2640
         TabIndex        =   8
         Text            =   "1"
         ToolTipText     =   "scanner a partir du port..."
         Top             =   1485
         Width           =   495
      End
      Begin VB.TextBox txtMaxConnections 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2640
         TabIndex        =   7
         Text            =   "100"
         ToolTipText     =   "Plus le nombre est elevé, plus la connection est rapide. Valeur entre 1 et 500."
         Top             =   1125
         Width           =   1575
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2640
         TabIndex        =   6
         Text            =   "000.000.000.000"
         ToolTipText     =   "IP de la machine a scanner"
         Top             =   765
         Width           =   1575
      End
      Begin MSComctlLib.ListView l 
         Height          =   3735
         Left            =   480
         TabIndex        =   5
         ToolTipText     =   "Ports ouverts et type d'utlisation."
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Port"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Utilisation"
            Object.Width           =   4613
         EndProperty
      End
      Begin VB.Label dnstool 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Net Tool"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   6270
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   960
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Label portspy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Port Spy"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   6495
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   960
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label about 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "a propos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   6495
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   2280
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label export 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exporter"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   6255
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2280
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Image Image12 
         Height          =   720
         Left            =   240
         Picture         =   "frmMain.frx":F512
         ToolTipText     =   "Scanner!"
         Top             =   720
         Width           =   720
      End
      Begin VB.Label scan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scan"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   2280
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label pingT 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ping Tool"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   16
         ToolTipText     =   "Teste de connection, ping, messages, etc..."
         Top             =   6015
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   4140
         Index           =   5
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2160
         TabIndex        =   14
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "à"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   3360
         TabIndex        =   13
         Top             =   1440
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thread:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP a Scanner:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   4
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   2
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   1
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   0
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   1320
         Picture         =   "frmMain.frx":14CF4
         ToolTipText     =   "Scanner!"
         Top             =   3000
         Width           =   720
      End
   End
   Begin MSWinsockLib.Winsock w 
      Index           =   0
      Left            =   4560
      Top             =   10200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClearList 
      Appearance      =   0  'Flat
      Caption         =   "&Clear List"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "S&top"
      Height          =   495
      Left            =   7800
      TabIndex        =   0
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Timer timTimer 
      Interval        =   100
      Left            =   8640
      Top             =   8640
   End
   Begin MSWinsockLib.Winsock wskSocket 
      Index           =   0
      Left            =   6480
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1920
      Picture         =   "frmMain.frx":1A4D6
      Top             =   9720
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFAA00&
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      Top             =   9600
      Width           =   3495
   End
   Begin VB.Image Image10 
      Height          =   765
      Left            =   7560
      Picture         =   "frmMain.frx":1ADA0
      Top             =   10440
      Width           =   30
   End
   Begin VB.Label lblTo 
      Caption         =   "To"
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   9600
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bol As Boolean
Dim one As Boolean


Private Sub about_Click()
Form1.Show
End Sub


Private Sub cmdScan_Click()

   Dim intI As Integer

   lngNextPort = Val(Me.txtLowerBound)

   For intI = 1 To Val(Me.txtMaxConnections)

      Load Me.wskSocket(intI)

      lngNextPort = lngNextPort + 1

      Me.wskSocket(intI).Connect Me.txtIP, lngNextPort

   Next intI

End Sub

Private Sub cmdStop_Click()
On Error Resume Next
   Dim intI As Integer
   

   For intI = 1 To Val(Me.txtMaxConnections)

      Me.wskSocket(intI).Close

      Unload Me.wskSocket(intI)

   Next intI

End Sub

Private Sub dnstool_Click()
Form3.Show
End Sub

Private Sub export_Click()
Dim i As Long

rtf.text = "..:: BlackScan ::.." & Chr(13) & Chr(10) & "IP:   " & txtIP.text & Chr(13) & Chr(10) & "Date: " & Day(Now) & " " & Month(Now) & " " & Year(Now) & vbCrLf
For i = 1 To l.ListItems.Count
rtf.text = rtf.text & l.ListItems.Item(i).text & " " & l.ListItems(i).SubItems(1) & Chr(13) & Chr(10)
Next i
rtf.text = rtf.text & vbCrLf & i - 2 & " ports ouverts"
cd.DialogTitle = "Exporter le rapport"
cd.Filter = "Text (*.txt)|*.txt|RTF (*.rtf)|*.rtf|autre|*.*"
cd.ShowSave
If cd.FileName <> "" Then
Open cd.FileName For Output As #1
Print #1, rtf.text
Close #1
End If
End Sub

Private Sub Form_Load()
bol = False
one = False
XP.Initialise = Me
XP.icon = Me
XP.Caption = Me.Caption
txtIP.text = W(0).LocalIP
XP.Top = 0
XP.Left = 0
Dim style As Long
   Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(l.hWnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   style = style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If style Then
      Call SetWindowLong(hHeader, GWL_STYLE, style)
      Call SetWindowPos(l.hWnd, Form1.hWnd, 0, 0, 0, 0, SWP_FLAGS)
   End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
DragForm Me
End Sub

Private Sub Image1_Click()
Form1.Show
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
DragForm Me
End Sub

Private Sub Form_Resize()
XP.Width = Me.Width
XP.Height = Me.Height
End Sub

Private Sub Image2_Click()
If bol = False Then
l.ListItems.Clear
cmdScan_Click
Image2.Top = Image2.Top + 255
Image2.Picture = Image3.Picture
Image2.ToolTipText = "Cliquez pour arreter!"
scan.Caption = "Stop"
bol = True
Else
cmdStop_Click
Image2.Top = Image2.Top - 255
Image2.Picture = Image12.Picture
Image2.ToolTipText = "Commencer a scanner!"
scan.Caption = "Scan"
bol = False
End If
End Sub

Private Sub Image4_Click()
End
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
DragForm Me
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
DragForm Me
End Sub

Private Sub Image8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
DragForm Me
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
DragForm Me
End Sub

Private Sub pingT_Click()
Form2.Show
End Sub

Private Sub portspy_Click()
Spy.Show
End Sub

Private Sub scan_Click()
Image2_Click
End Sub

Private Sub timTimer_Timer()

   Label2(0).Caption = "Current Port: " + Str(lngNextPort)
   Label2(1).Caption = Label2(0).Caption
   XP.Stexte = Label2(1).Caption
'If Label2(1).Caption >= txtUpperBound.text Then scan.Caption = "scan": scan.Enabled = True
End Sub

Private Sub w_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
W(Index).GetData strData
If strData <> "" Then
addport strData, Index
Else
Dim l_list As ListItem
Set l_list = l.ListItems.Add(, , Index)
l_list.SubItems(1) = portname(Index)
l_list.SubItems(2) = "pas d'info"
Try_Next_Port (Index)
End If
End Sub


Private Sub wskSocket_Connect(Index As Integer)
'Dim ind As Integer
'ind = wskSocket(Index).RemotePort
'Load w(ind)
'w(ind).Connect Me.txtIP, wskSocket(Index).RemotePort
'w(ind).SendData "Ping"
Dim l_list As ListItem
If one = False Then
Set l_list = l.ListItems.Add(, , "")
l_list.SubItems(1) = ""
one = True
End If
Set l_list = l.ListItems.Add(, , wskSocket(Index).RemotePort)
If portname(wskSocket(Index).RemotePort) <> "" Then
l_list.SubItems(1) = portname(wskSocket(Index).RemotePort)
Else
l_list.SubItems(1) = "Unknow"
End If
Try_Next_Port (Index)

End Sub


Public Sub addport(data, Index)

Dim l_list As ListItem
Set l_list = l.ListItems.Add(, , Index)
l_list.SubItems(1) = portname(Index)
l_list.SubItems(2) = data
Try_Next_Port (Index)

End Sub

Public Sub addport2(Index)



End Sub

Private Sub wskSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim strData As String
   wskSocket(Index).GetData strData
   addport2 Index
    'wskSocket(index).Close
End Sub

Private Sub wskSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

   Try_Next_Port (Index)

End Sub

Private Sub Try_Next_Port(Index As Integer)

   Me.wskSocket(Index).Close

   If lngNextPort < Val(Me.txtUpperBound) Then

      Me.wskSocket(Index).Connect , lngNextPort

      lngNextPort = lngNextPort + 1

   Else

      Unload Me.wskSocket(Index)

   End If

End Sub




Public Sub texte(text As String, color As Long, vbcr As Boolean)
rtf.text = rtf.text & text
If vbcr = True Then
rtf.text = rtf.text & vbCrLf
End If
HLW rtf, text, color
End Sub


Private Function HLW(rtb As RichTextBox, sFindString As String, Lcolor As Long)
Dim LfoundPos As Long
Dim LfindLenght
Dim LorigSelStart
Dim LorigSelLenght
Dim ImatchCount As Integer
LorigSelStart = rtb.SelStart
LorigSelLenght = rtb.SelLength
LfindLenght = Len(sFindString)
LfoundPos = rtb.Find(sFindString, 0, , rtfNoHighlight)
While LfoundPos > 0
ImatchCount = ImatchCount + 1
rtb.SelStart = LfoundPos
rtb.SelLength = LfindLenght
rtb.SelColor = Lcolor
LfoundPos = rtb.Find(sFindString, LfoundPos + LfindLenght, , rtfNoHighlight)
Wend
rtb.SelStart = LorigSelStart
rtb.SelLength = LorigSelLenght
HLW = ImatchCount
End Function



