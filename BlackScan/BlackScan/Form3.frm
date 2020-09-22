VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Whois, DNS, IP, ..."
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox text5 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form3.frx":1CFA
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Text            =   "vbfrance.com"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Text            =   "whois.internic.net"
      Top             =   1320
      Width           =   1935
   End
   Begin BlackScan.WhoIs WhoIs 
      Left            =   7680
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      Server          =   ""
      Query           =   ""
   End
   Begin BlackScan.DNS DNS 
      Left            =   7680
      Top             =   600
      _ExtentX        =   1244
      _ExtentY        =   1244
   End
   Begin MSWinsockLib.Winsock W 
      Left            =   7680
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Query"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1815
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result:"
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
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Query:"
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
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WhoIs Server:"
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
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DSN:"
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
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
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
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   285
   End
   Begin VB.Label portspy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DNS 2 IP"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   735
      Width           =   975
   End
   Begin VB.Label dnstool 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IP 2 DNS"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   735
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   8
      Top             =   2715
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   7
      Top             =   3435
      Width           =   1335
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   3075
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   7080
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   7080
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   7080
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   2640
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   1680
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   2160
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Function UNIXtoDOS(ByVal s As String) As String
    UNIXtoDOS = Replace(s, Chr(10), Chr(13) & Chr(10))
End Function


Private Sub dnstool_Click()
On Error Resume Next
Text2.text = DNS.AddressToName(Text1)
End Sub

Private Sub Form_Load()
Text1.text = w.LocalIP
End Sub

Private Sub Label4_Click()
'On Error Resume Next
WhoIs.Server = Text3.text
WhoIs.Query = Text4.text
WhoIs.Connect
text5.text = "en attente..."
text5.text = UNIXtoDOS(WhoIs.Result)
End Sub

Private Sub portspy_Click()
On Error Resume Next
Text1.text = DNS.NameToAddress(Text2)
End Sub

Private Sub WhoIs_CloseWhoIs()
'On Error Resume Next
    text5.text = UNIXtoDOS(WhoIs.Result)
End Sub


