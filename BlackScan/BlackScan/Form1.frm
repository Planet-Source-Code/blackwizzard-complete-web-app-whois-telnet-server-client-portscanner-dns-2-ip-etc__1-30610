VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "About BlackScan"
   ClientHeight    =   4950
   ClientLeft      =   6900
   ClientTop       =   5400
   ClientWidth     =   4800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin BlackScan.XPDesign XP 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8705
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3135
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "Form1.frx":1CFA
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label quit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fermer"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   4275
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   1800
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF7700&
         BorderWidth     =   3
         FillStyle       =   0  'Solid
         Height          =   3540
         Index           =   1
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
XP.Top = 0
XP.Left = 0
XP.Initialise = Me
XP.icon = Me
XP.Caption = Me.Caption
XP.Stexte = "About BlackScan"
XP.QuitButton = False
XP.hlpbutton = False
End Sub

Private Sub quit_Click()
Unload Me
End Sub
