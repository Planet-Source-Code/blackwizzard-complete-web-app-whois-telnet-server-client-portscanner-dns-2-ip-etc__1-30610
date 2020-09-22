VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Ping Tool"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1920
      Top             =   4560
   End
   Begin MSWinsockLib.Winsock W 
      Left            =   1080
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin BlackScan.XPDesign XP 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8070
      Begin VB.TextBox texte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1320
         TabIndex        =   8
         Text            =   "Ping"
         Top             =   1485
         Width           =   1575
      End
      Begin VB.TextBox receive 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1035
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2205
         Width           =   3735
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Text            =   "193.253.53.44"
         Top             =   765
         Width           =   1575
      End
      Begin VB.TextBox port 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1320
         TabIndex        =   1
         Text            =   "21"
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label disco 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Disconnect"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   1815
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   3120
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label quit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fermer"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   3920
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   2760
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label sendmsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Send MSG"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   1815
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2160
         Top             =   1800
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   3
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texte:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label ping 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Connect"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1815
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   1200
         Top             =   1800
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   1140
         Index           =   1
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   2160
         Width           =   3975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receive:"
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
         TabIndex        =   6
         Top             =   1800
         Width           =   945
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   0
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   300
         Index           =   2
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   735
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
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   285
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
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub disco_Click()
W.Close
receive.text = "Deconnection de l'utilisateur"
End Sub

Private Sub Form_Load()
XP.Top = 0
XP.Left = 0
XP.Initialise = Me
XP.icon = Me
XP.Caption = Me.Caption
XP.Stexte = "Ping Tool"
XP.QuitButton = False
XP.hlpbutton = False
txtIP.text = frmMain.txtIP.text
End Sub

Private Sub Form_Resize()
XP.Width = Me.Width
XP.Height = Me.Height
End Sub

Private Sub ping_Click()
On Error Resume Next
W.RemoteHost = txtIP.text
W.RemotePort = port.text
W.Connect
ping.Enabled = False
Timer1.Enabled = True
receive.text = ""
End Sub

Private Sub quit_Click()
Unload Me
End Sub

Private Sub sendmsg_Click()
On Error Resume Next
receive.text = ""
W.SendData texte.text
receive.text = "message envoyé" & vbCrLf & "En attente d'une reponse"
End Sub

Private Sub Timer1_Timer()
XP.Stexte = getState
End Sub

'Private Sub w_Close()
'ping.Enabled = True
'W.Close
'If receive.text = "" Then
'receive.text = "Pas de reponse du server sur le port " & port.text
'End If
'Timer1.Enabled = False
'End Sub

Private Sub W_Connect()
W.SendData texte.text
End Sub

Private Sub w_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
   W.GetData strData
    receive.text = strData
'W.Close
ping.Enabled = True
End Sub

Function getState() As String
Dim buffer As String
Select Case W.State
    Case 0:
        buffer = "Fermé (valeur par défaut)"
        ping.Enabled = True
        If receive.text = "" Then
        receive.text = "Pas de reponse du server sur le port " & port.text
        End If
        sendmsg.Enabled = False
        ping.Enabled = True
        disco.Enabled = False
    Case 1:
        buffer = "Ouvert"
        sendmsg.Enabled = True
        ping.Enabled = False
        disco.Enabled = True
        
    Case 2:
        buffer = "À l'écoute"
        sendmsg.Enabled = True
        ping.Enabled = False
        disco.Enabled = True
    Case 3:
        buffer = "Connexion en attente"
        sendmsg.Enabled = False
        ping.Enabled = False
        disco.Enabled = True
    Case 4:
        buffer = "Hôte en cours de résolution"
        sendmsg.Enabled = False
        ping.Enabled = False
        disco.Enabled = True
    Case 5:
        buffer = "Hôte résolu"
        sendmsg.Enabled = False
        ping.Enabled = False
        disco.Enabled = True
    Case 6:
        buffer = "En cours de connexion"
        sendmsg.Enabled = False
        ping.Enabled = False
        disco.Enabled = True
    Case 7:
        buffer = "Connecté"
        sendmsg.Enabled = True
        ping.Enabled = False
         'cmdconnect.Enabled = False
       ' cmdSend.Enabled = True
       disco.Enabled = True
    Case 8:
        buffer = "en cours de fermeture"
        sendmsg.Enabled = False
        ping.Enabled = False
        disco.Enabled = True
    Case 9:
        buffer = "Connexion en cours de fermeture par l'homologue"
        sendmsg.Enabled = False
        ping.Enabled = False
        disco.Enabled = False
        W.Close
        'cmdconnect.Enabled = True
       ' cmdSend.Enabled = False
    Case sckError:
        buffer = "Erreur"
End Select
getState = buffer
End Function
