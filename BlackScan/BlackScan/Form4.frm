VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Spy 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Port Spy"
   ClientHeight    =   9270
   ClientLeft      =   1350
   ClientTop       =   1185
   ClientWidth     =   4935
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin BlackScan.XPDesign XP 
      Height          =   9255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   16325
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   20
         Text            =   "Form4.frx":1CFA
         Top             =   7440
         Width           =   4695
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1785
         Left            =   1680
         TabIndex        =   10
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox port2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   9
         Text            =   "500"
         ToolTipText     =   "Scanner jusqu'au port..."
         Top             =   645
         Width           =   495
      End
      Begin VB.TextBox port1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2040
         TabIndex        =   8
         Text            =   "1"
         ToolTipText     =   "scanner a partir du port..."
         Top             =   645
         Width           =   495
      End
      Begin RichTextLib.RichTextBox msg 
         Height          =   975
         Left            =   1680
         TabIndex        =   7
         Top             =   3720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1720
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form4.frx":1E10
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Accept mesage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   3270
         Width           =   2895
      End
      Begin RichTextLib.RichTextBox msg_r 
         Height          =   975
         Left            =   1680
         TabIndex        =   5
         Top             =   3720
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1720
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form4.frx":1EC4
      End
      Begin RichTextLib.RichTextBox Rmsg 
         Height          =   1695
         Left            =   1680
         TabIndex        =   4
         Top             =   4920
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2990
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form4.frx":1F8E
      End
      Begin VB.Label Command1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Connect"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   7035
         Width           =   1335
      End
      Begin VB.Label quit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fermer"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   7035
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   1560
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   3360
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display msg:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   4800
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1980
         Index           =   3
         Left            =   1560
         Shape           =   4  'Rounded Rectangle
         Top             =   4800
         Width           =   3135
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1140
         Index           =   1
         Left            =   1560
         Shape           =   4  'Rounded Rectangle
         Top             =   3600
         Width           =   3135
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   0
         Left            =   1560
         Shape           =   4  'Rounded Rectangle
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "When user is connected:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   1320
         TabIndex        =   14
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   2760
         TabIndex        =   13
         Top             =   600
         Width           =   210
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
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   4
         Left            =   3000
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   2
         Left            =   1920
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   1980
         Index           =   5
         Left            =   1560
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   465
      End
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1620
      Left            =   9960
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.ListBox e 
      Height          =   255
      Left            =   9600
      TabIndex        =   1
      Top             =   5280
      Width           =   135
   End
   Begin VB.ListBox l 
      Height          =   255
      Left            =   9600
      TabIndex        =   0
      Top             =   5580
      Width           =   135
   End
   Begin MSWinsockLib.Winsock w 
      Index           =   0
      Left            =   11160
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   11520
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Spy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API Declaration
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Form level variables
Dim UserID As String
Dim Password As String
Dim AcceptedId As Boolean
Dim SuccessLoging As Boolean
Dim UserCommand As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
msg.Enabled = False
msg.Visible = False
msg_r.Visible = True
Else
msg.Enabled = True
msg_r.Visible = False
msg.Visible = True
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim i As Integer
    For i = port1 To port2
    Load w(i)
    w(i).LocalPort = i
    Next i
For i = port1 To port2
w(i).Listen
Next i
Command1.Enabled = False
End Sub


Private Sub Form_Resize()
XP.Width = Me.Width
XP.Height = Me.Height
End Sub

Private Sub Form_Load()
Dim i As Integer
On Error Resume Next
    'Set the telnet port

    'Set the server to listen for a client request

    'Initialisation of the telnet server variables
    UserID = ""
    Password = ""
    UserCommand = ""
    AcceptedId = False
    SuccessLoging = True

XP.Top = 0
XP.Left = 0
XP.Initialise = Me
XP.icon = Me
XP.Caption = Me.Caption
XP.Stexte = "By BlackWizzard"
XP.QuitButton = False
XP.hlpbutton = False
End Sub

Private Sub quit_Click()
Unload Me
End Sub

Private Sub w_Close(Index As Integer)
On Error Resume Next
    'When user wants to close the telnet connection
    l.AddItem w(Index).RemoteHostIP & " disconnected"
    w(Index).Close 'Close the telnet port
    w(Index).LocalPort = 23 'set telnet port
    w(Index).Listen 'Listen for the new user
    
    
    'Initialisation of the telnet server variables
    UserID = ""
    Password = ""
    UserCommand = ""
    AcceptedId = False
    SuccessLoging = True
End Sub

Private Sub w_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'User wants to connect to the server
    If w(Index).State <> sckClosed Then w(Index).Close
    w(Index).Accept requestID
    
    'Send him the accepted message and ask him to logon to the server
    If Check1.Value = 1 Then
    w(Index).SendData "Welcome <" & w(Index).RemoteHostIP & ">." & vbCrLf & msg_r.text & vbCrLf & "To send the message, type 'send' on a single line." & vbCrLf & vbCrLf
    Else
    w(Index).SendData "Welcome <" & w(Index).RemoteHostIP & ">." & vbCrLf & msg.text & vbCrLf
    End If
    List1.AddItem w(Index).RemoteHostIP & " on port " & Index
End Sub

Private Sub w_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim str1 As String
    Dim MyName As String
    Dim DirList() As String
    Dim DirPointer As Integer
    Dim ii As Integer
    Dim Fullmsg As String
Fullmsg = w(Index).RemoteHostIP
    w(Index).GetData str1

If Check1.Value = 1 Then
    If Asc(str1) = 13 Then
        If UserCommand Like "send" Then
            List2.AddItem UserCommand
            Rmsg.text = Rmsg.text & vbCrLf & "### User " & w(Index).RemoteHostIP & " deconnected at " & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
        w(Index).SendData "User " & w(Index).RemoteHostIP & " deconnected." & vbCrLf
        Sleep 4000
        w(Index).Close
        w(Index).LocalPort = Index
        w(Index).Listen
        Else
            Rmsg.text = Rmsg.text & vbCrLf & w(Index).RemoteHostIP & "> " & UserCommand
            UserCommand = ""
        End If
    Else

    UserCommand = UserCommand & str1

    End If
Else

    w(Index).Close
    w(Index).LocalPort = Index
    w(Index).Listen
End If
End Sub

Private Sub w_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
e.AddItem "error - " & Index & " - " & Description & " - " & Number
End Sub


