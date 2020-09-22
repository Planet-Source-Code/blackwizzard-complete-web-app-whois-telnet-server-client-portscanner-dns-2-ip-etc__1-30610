VERSION 5.00
Begin VB.UserControl WhoIs 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ocxWhoIs.ctx":0000
   PropertyPages   =   "ocxWhoIs.ctx":05A1
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "ocxWhoIs.ctx":05B2
End
Attribute VB_Name = "WhoIs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Dim strServ As String
Attribute strServ.VB_VarProcData = "Connection"
Dim strQuery As String
Attribute strQuery.VB_VarProcData = "Connection"
Dim strResult As String
Attribute strResult.VB_VarProcData = "Connection"
Dim strStatus As String

Dim wsTCP As oswinsck.TCP
Attribute wsTCP.VB_VarHelpID = -1
Public Event ConnectWhoIs()
Public Event CloseWhoIs()
Public Event ErrorWhoIs(ByVal Number As Integer, Description As String)

Private sCommand As String

Public Property Let Server(ByVal strTemp As String)
Attribute Server.VB_Description = "Returns/sets mailserver"
  strServ = strTemp
  PropertyChanged "Server"
End Property

Public Property Let Query(ByVal strTemp As String)
Attribute Query.VB_Description = "Returns/sets recepient e-mail address"
  strQuery = strTemp
  PropertyChanged "Query"
End Property

Public Property Get Server() As String
  Server = strServ
End Property

Public Property Get Query() As String
  Query = strQuery
End Property

Public Property Get Result() As String
Attribute Result.VB_Description = "Returns/sets sender e-mail address"
  Result = strResult
End Property

Public Property Get Status() As String
  Status = strStatus
End Property

Public Sub Connect()
  On Error Resume Next
  Set wsTCP = CreateObject("oswinsck.TCP")
  strStatus = "Connecting to WhoIs server"
  If wsTCP.Connect(strServ, 43) = 0 Then
    strStatus = "Connected to WhoIs server"
    RaiseEvent ConnectWhoIs
    wsTCP.SendData strQuery & vbCrLf
    strResult = wsTCP.GetData
    wsTCP.Disconnect
    Set wsTCP = Nothing
    strStatus = "WhoIs session closed"
    RaiseEvent CloseWhoIs
  End If
  Exit Sub
  
ErrHandler:
  If Err.Number <> 0 Then
    strStatus = "WhoIs control error"
    RaiseEvent ErrorWhoIs(Err.Number, Err.Description)
    Err.Clear
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  'Server = PropBag.ReadProperty("Server")
  'Query = PropBag.ReadProperty("Query")
End Sub

Private Sub UserControl_Resize()
With UserControl
.Width = 480
.Height = 480
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Server", Server
  PropBag.WriteProperty "Query", Query
End Sub
