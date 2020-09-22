VERSION 5.00
Begin VB.UserControl XPDesign 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   11655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   ControlContainer=   -1  'True
   ScaleHeight     =   11655
   ScaleWidth      =   10995
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.PictureBox Bbar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   3855
      TabIndex        =   2
      Top             =   11280
      Width           =   3855
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
         Left            =   120
         TabIndex        =   3
         Top             =   45
         Width           =   3495
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   4
         Top             =   80
         Width           =   3495
      End
      Begin VB.Image Bbar2 
         Height          =   285
         Left            =   0
         Picture         =   "UserControl1.ctx":0312
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3735
      End
   End
   Begin VB.Image EndButton 
      Height          =   480
      Left            =   3240
      Picture         =   "UserControl1.ctx":0391
      Stretch         =   -1  'True
      ToolTipText     =   "Quitter"
      Top             =   20
      Width           =   480
   End
   Begin VB.Image icon 
      Height          =   375
      Left            =   120
      Top             =   160
      Width           =   375
   End
   Begin VB.Image HelpButton 
      Height          =   600
      Left            =   2760
      Picture         =   "UserControl1.ctx":661B
      Stretch         =   -1  'True
      ToolTipText     =   "About"
      Top             =   -60
      Width           =   600
   End
   Begin VB.Image Image6 
      Height          =   500
      Left            =   0
      Picture         =   "UserControl1.ctx":C8A5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   165
   End
   Begin VB.Image RRoundBar 
      Height          =   495
      Left            =   3570
      Picture         =   "UserControl1.ctx":CAC1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   165
   End
   Begin VB.Image LSideBar 
      Height          =   6405
      Left            =   0
      Picture         =   "UserControl1.ctx":CCD5
      Stretch         =   -1  'True
      Top             =   285
      Width           =   60
   End
   Begin VB.Label Label2 
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
      Index           =   3
      Left            =   600
      TabIndex        =   1
      Top             =   215
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   570
      TabIndex        =   0
      Top             =   225
      Width           =   2175
   End
   Begin VB.Image titlebar 
      Height          =   495
      Left            =   165
      Picture         =   "UserControl1.ctx":CD54
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3400
   End
   Begin VB.Image RsideBar 
      Height          =   6165
      Left            =   0
      Picture         =   "UserControl1.ctx":CDD3
      Stretch         =   -1  'True
      Top             =   120
      Width           =   60
   End
End
Attribute VB_Name = "XPDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2










Dim Frm As Long

Private Sub Bbar2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call ReleaseCapture: Call SendMessage(Frm, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub EndButton_Click()
End
End Sub

Private Sub HelpButton_Click()
Form1.Show
End Sub

Private Sub icon_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call ReleaseCapture: Call SendMessage(Frm, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call ReleaseCapture: Call SendMessage(Frm, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call ReleaseCapture: Call SendMessage(Frm, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub LSideBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call ReleaseCapture: Call SendMessage(Frm, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub RRoundBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call ReleaseCapture: Call SendMessage(Frm, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub RsideBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call ReleaseCapture: Call SendMessage(Frm, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub titlebar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call ReleaseCapture: Call SendMessage(Frm, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call ReleaseCapture: Call SendMessage(Frm, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Dim W
Dim h
W = UserControl.Width
h = UserControl.Height
UserControl.RsideBar.Height = h - 150
UserControl.LSideBar.Height = h - 150
UserControl.Bbar.Top = h - UserControl.Bbar.Height
UserControl.Bbar.Width = W
UserControl.Bbar2.Width = W
UserControl.titlebar.Width = W - (165 * 2)
UserControl.RsideBar.Left = W - UserControl.RsideBar.Width
UserControl.RRoundBar.Left = W - UserControl.RRoundBar.Width
UserControl.Label2(0).Width = W
UserControl.Label2(1).Width = W
UserControl.Label2(2).Width = W - 975 - UserControl.icon.Width
UserControl.Label2(3).Width = W - 975 - UserControl.icon.Width
UserControl.HelpButton.Left = W - 975
UserControl.EndButton.Left = W - 495
'UserControl.icon.Picture = Frm.icon
'UserControl.Label2(3).Caption = Frm.Caption
'UserControl.Label2(2).Caption = Frm.Caption
End Sub

Public Property Get Stexte() As Variant
Stexte = UserControl.Label2(0).Caption
End Property

Public Property Let Stexte(ByVal vNewValue As Variant)
UserControl.Label2(0).Caption = vNewValue
UserControl.Label2(1).Caption = vNewValue
End Property

Public Property Get Caption() As Variant
Caption = UserControl.Label2(0).Caption
End Property

Public Property Let Caption(ByVal text As Variant)
UserControl.Label2(2).Caption = text
UserControl.Label2(3).Caption = text
End Property


Public Property Get Initialise() As Form
'Initialise = Frm
End Property

Public Property Let Initialise(ByVal ActiveForm As Form)
Frm = ActiveForm.hWnd
End Property

Public Property Get icon() As Form
'icon = UserControl.icon.Picture
End Property

Public Property Let icon(ByVal IconPicture As Form)
UserControl.icon.Picture = IconPicture.icon
End Property

Public Property Get QuitButton() As Boolean
QuitButton = EndButton.Visible
End Property

Public Property Let QuitButton(ByVal bool As Boolean)
EndButton.Visible = bool
End Property

Public Property Get hlpbutton() As Boolean
hlpbutton = HelpButton.Visible
End Property

Public Property Let hlpbutton(ByVal bool As Boolean)
HelpButton.Visible = bool
End Property
