VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmSQLConnection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Server Connection"
   ClientHeight    =   5415
   ClientLeft      =   7185
   ClientTop       =   6315
   ClientWidth     =   3900
   Icon            =   "frmSQLConnection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   3900
   Begin VB.Frame Frame1 
      Caption         =   "SQL/MSDE Properties"
      Height          =   4125
      Left            =   90
      TabIndex        =   3
      Top             =   720
      Width           =   3735
      Begin VB.ComboBox cmbServerName 
         Height          =   315
         Left            =   210
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   600
         Width           =   2475
      End
      Begin VB.ComboBox cmbDatabaseName 
         Height          =   315
         Left            =   210
         TabIndex        =   8
         Top             =   1350
         Width           =   2475
      End
      Begin VB.TextBox txtUID 
         Height          =   345
         Left            =   210
         TabIndex        =   7
         Top             =   2760
         Width           =   3315
      End
      Begin VB.TextBox txtPWD 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   210
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   3390
         Width           =   3315
      End
      Begin VB.ComboBox cmbSecurity 
         Height          =   315
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2130
         Width           =   2505
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test"
         Height          =   405
         Left            =   2820
         TabIndex        =   4
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Server name"
         Height          =   285
         Left            =   210
         TabIndex        =   14
         Top             =   330
         Width           =   2205
      End
      Begin VB.Label Label2 
         Caption         =   "Database name"
         Height          =   255
         Left            =   210
         TabIndex        =   13
         Top             =   1110
         Width           =   1965
      End
      Begin VB.Label Label3 
         Caption         =   "Security"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   1830
         Width           =   1965
      End
      Begin VB.Label Label4 
         Caption         =   "User name"
         Height          =   255
         Left            =   210
         TabIndex        =   11
         Top             =   2520
         Width           =   1965
      End
      Begin VB.Label Label5 
         Caption         =   "Password"
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   3150
         Width           =   1965
      End
      Begin VB.Image imgSrvConnect 
         Height          =   480
         Left            =   2910
         Picture         =   "frmSQLConnection.frx":08CA
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image imgSrvNotConnect 
         Height          =   480
         Left            =   2910
         Picture         =   "frmSQLConnection.frx":1194
         Top             =   1890
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCommit 
      Caption         =   "&Commit"
      Height          =   315
      Left            =   2460
      TabIndex        =   2
      Top             =   4920
      Width           =   645
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3180
      TabIndex        =   1
      Top             =   4920
      Width           =   645
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3090
      Top             =   4800
   End
   Begin VB.CheckBox chkIs2000 
      Caption         =   "Use SQL-2000 DMO"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   4950
      Width           =   2175
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   585
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   1032
      _Version        =   196609
      BackColor       =   16777215
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmSQLConnection.frx":1A5E
         Top             =   60
         Width           =   330
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reassign the default SQL Server/MSDE database used by the application."
         Height          =   435
         Left            =   570
         TabIndex        =   16
         Top             =   90
         Width           =   3255
      End
   End
   Begin VB.Label lblNoDMO 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No DMO - Manual Config"
      Height          =   255
      Left            =   90
      TabIndex        =   17
      Top             =   5100
      Width           =   2175
   End
End
Attribute VB_Name = "frmSQLConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcFuncLib As New FuncLib_Lib
Private mcDMO As New DMO_Lib
Private mcReg As New Registry32
Private mbooLoaded As Boolean
Private mbooDBListIsDirty As Boolean
Private mbooServerListDirty As Boolean
Private mbooManualConfig As Boolean
Private Enum eTag
    UserFire
    CodeFire
End Enum

Private Enum eSecurityLI
    NTSecurity = 0
    SQLServer = 1
End Enum

Private Sub EB()
MsgBox Err.description
End Sub

Public Property Get x_ManualConfig() As Boolean
x_ManualConfig = mbooManualConfig
End Property

Private Sub chkIs2000_Click()
On Error GoTo errh
If chkIs2000.Value = 1 Then
    puFormSQLConnection.uSQLConnection.SQL_Engine_Ver = SQL_2000
Else
    puFormSQLConnection.uSQLConnection.SQL_Engine_Ver = SQL_7
End If
If Not mbooLoaded Then Exit Sub
puFormSQLConnection.uSQLConnection.IsValid = False
Call ImageShow
DoEvents#
Exit Sub
errh:
EB
End Sub

Private Sub cmbDatabaseName_Change()
On Error Resume Next
If Me.cmbDatabaseName.Tag = CStr(eTag.CodeFire) Then Exit Sub
puFormSQLConnection.uSQLConnection.IsValid = False
Call ImageShow
DoEvents
End Sub

Private Sub cmbDatabaseName_Click()
On Error Resume Next
If Me.cmbServerName.Tag = CStr(eTag.CodeFire) Then Exit Sub
puFormSQLConnection.uSQLConnection.IsValid = False
Call ImageShow
End Sub

Private Sub cmbSecurity_Click()
On Error Resume Next
If Me.cmbSecurity.Tag = CStr(eTag.CodeFire) Then Exit Sub
puFormSQLConnection.uSQLConnection.IsValid = False
Call ImageShow
DoEvents
End Sub

Private Sub cmbServerName_Change()
On Error Resume Next
mbooServerListDirty = True
End Sub

Private Sub cmbServerName_Click()
Static sLastServ$
On Error Resume Next
If Me.cmbServerName.Tag = CStr(eTag.CodeFire) Then Exit Sub
If sLastServ = cmbServerName.Text Then Exit Sub
sLastServ = cmbServerName.Text
Me.cmbDatabaseName.Clear
DoEvents
mbooDBListIsDirty = True
End Sub

Private Sub cmbServerName_LostFocus()
On Error Resume Next
If mbooServerListDirty Then
    Me.cmbDatabaseName.Clear
    DoEvents
    mbooServerListDirty = False
    mbooDBListIsDirty = True
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
puFormSQLConnection.Cancelled = True
Unload Me
End Sub

Private Sub cmdCommit_Click()
Dim b As Boolean
On Error GoTo errh
If Not puFormSQLConnection.uSQLConnection.IsValid Then
    If MsgBox("The connection details have not been validated. You may not be able to connect to the application data." & _
    vbCrLf & vbCrLf & "Do you wish to continue anyway?", vbYesNo) = vbNo Then Exit Sub
End If
puFormSQLConnection.Cancelled = False
Unload Me
Exit Sub
errh:
Resume e
e:
EB
Exit Sub
End Sub


Private Sub cmdTest_Click()
Dim Cnn As New ADODB.Connection
Dim cADO As New ADO_LIB
On Error GoTo errh
MPH
Select Case Me.cmbSecurity.ListIndex
    Case eSecurityLI.SQLServer
        Set Cnn = cADO.GetOLEDBADOConn(Me.cmbServerName.Text, Me.cmbDatabaseName.Text, Me.txtUID.Text, Me.txtPWD.Text, , , False)
    Case Else
        Set Cnn = cADO.GetOLEDBADOConn(Me.cmbServerName.Text, Me.cmbDatabaseName.Text)
End Select

If Cnn.State And adStateOpen Then
    puFormSQLConnection.uSQLConnection.IsValid = True
    MPD
    MsgBox "Connection details verified.", vbInformation
Else
    puFormSQLConnection.uSQLConnection.IsValid = False
    MPD
    MsgBox "Connection details are invalid.", vbInformation
End If
Call ImageShow
GoSub LocalClean
Exit Sub

errh:
Resume e
e:
On Error Resume Next
puFormSQLConnection.uSQLConnection.IsValid = False
GoSub LocalClean
MsgBox "Test Failed! Please redefine your connection parameters."
Err.Clear
Exit Sub

LocalClean:
On Error Resume Next
D Cnn
Err.Clear
MPD
Return
End Sub

Private Sub D(obj As Object)
On Error Resume Next
If Not obj Is Nothing Then
    If TypeOf obj Is ADODB.Recordset Or TypeOf obj Is ADODB.Connection Then
        obj.Close
    End If
    Set obj = Nothing
End If
Err.Clear
End Sub

Private Sub Form_Activate()
Dim b As Boolean
Dim l&
Dim lVer As eSQLEngineVerBitwise
Dim ServList() As String
On Error GoTo errh
If mbooLoaded Then Exit Sub
MPH
Me.lblNoDMO.Top = Me.chkIs2000.Top
Me.lblNoDMO.Left = Me.chkIs2000.Left
DoEvents
lVer = mcDMO.DMOVersionsAvailable
If lVer = bwNone Then
    MPD
    MsgBox "SQL Server DMO objects not available on this machine. Unable to enumerate sql servers and database." & CR2 & "Connection information can be manually ammended.", vbCritical, "No DMO Objects Installed"
    Me.chkIs2000.Enabled = False
    Me.chkIs2000.Visible = False
    Me.lblNoDMO.Visible = True
    Me.lblNoDMO.Enabled = True
    mbooManualConfig = True
Else
    Me.chkIs2000.Enabled = True
    Me.chkIs2000.Visible = True
    Me.lblNoDMO.Visible = False
    Me.lblNoDMO.Enabled = False
    mbooManualConfig = False
    With chkIs2000
        If CBool(lVer And bwSQL_2000) And CBool(lVer And bwSQL_7) Then
            .Value = 1
            .Enabled = True
        ElseIf lVer And bwSQL_2000 Then
            .Value = 1
            .Enabled = False
        Else
            .Value = 0
            .Enabled = False
        End If
    End With
End If

mbooLoaded = True
mbooDBListIsDirty = True
puFormSQLConnection.Cancelled = False

Me.txtPWD.Text = puFormSQLConnection.uSQLConnection.PWD
Me.txtUID.Text = puFormSQLConnection.uSQLConnection.UID

Me.cmbSecurity.Clear
Me.cmbSecurity.AddItem "NT Security"
Me.cmbSecurity.ItemData(Me.cmbSecurity.NewIndex) = eSQLSecurity.NTOnly
Me.cmbSecurity.AddItem "SQL Server"
Me.cmbSecurity.ItemData(Me.cmbSecurity.NewIndex) = eSQLSecurity.SQLServer

Me.cmbSecurity.Tag = CStr(eTag.CodeFire)
l = ComboItemDataExists(Me.cmbSecurity, puFormSQLConnection.uSQLConnection.Security, -1)
If l > -1 Then Me.cmbSecurity.ListIndex = l
Me.cmbSecurity.Tag = CStr(eTag.UserFire)

b = mcDMO.SQLServer2000ServerList(ServList)
If b Then
    Me.cmbServerName.Tag = CStr(eTag.CodeFire)
    b = ComboFillWithSingleArray(Me.cmbServerName, ServList, puFormSQLConnection.uSQLConnection.ServerName)
    Me.cmbServerName.Tag = CStr(eTag.UserFire)
    If b Then
        b = PopulateDBCombo()
    End If
End If

Me.cmbDatabaseName.Tag = CStr(eTag.CodeFire)
l = ComboTextItemExists(Me.cmbDatabaseName, puFormSQLConnection.uSQLConnection.DatabaseName, -1)
If l > -1 Then Me.cmbDatabaseName.ListIndex = l
Me.cmbDatabaseName.Tag = CStr(eTag.UserFire)

Me.cmbServerName.Tag = CStr(eTag.CodeFire)
l = ComboTextItemExists(Me.cmbServerName, puFormSQLConnection.uSQLConnection.ServerName, -1)
If l > -1 Then Me.cmbServerName.ListIndex = l
Me.cmbServerName.Tag = CStr(eTag.UserFire)

Me.Timer1.Enabled = True
GoSub LC
DoEvents
Exit Sub

LC:
 On Error Resume Next
 MPD
Return

errh:
 EB
 On Error Resume Next
 GoSub LC
End Sub


Private Function PopulateDBCombo() As Boolean
Dim sDB() As String
Dim dmoSrv As Object
Dim b As Boolean
On Error GoTo errh

If puFormSQLConnection.uSQLConnection.SQL_Engine_Ver = SQL_2000 Then
    Set dmoSrv = CreateObject("SQLDMO.SQLServer2")
ElseIf puFormSQLConnection.uSQLConnection.SQL_Engine_Ver = SQL_7 Then
    Set dmoSrv = CreateObject("SQLDMO.SQLServer")
Else
    Err.Raise vbObjectError, "frmSQLConnection:PopulateDBCombo", "Invalid SQL Version"
End If

Select Case Me.cmbSecurity.ListIndex
    Case eSecurityLI.NTSecurity
        Set dmoSrv = mcDMO.ServerConnect(Me.cmbServerName.Text, , , "", "", puFormSQLConnection.uSQLConnection.SQL_Engine_Ver)
    Case eSecurityLI.SQLServer
        Set dmoSrv = mcDMO.ServerConnect(Me.cmbServerName.Text, False, 15, Me.txtUID.Text, Me.txtPWD.Text, puFormSQLConnection.uSQLConnection.SQL_Engine_Ver)
End Select
If Not dmoSrv Is Nothing Then
    b = mcDMO.DatabasesEnum(dmoSrv, sDB, puFormSQLConnection.uSQLConnection.SQL_Engine_Ver)
    If b Then
        b = ComboFillWithSingleArray(Me.cmbDatabaseName, sDB)
        If b Then
            PopulateDBCombo = True
            puFormSQLConnection.uSQLConnection.IsValid = True
        Else
            PopulateDBCombo = False
            puFormSQLConnection.uSQLConnection.IsValid = False
        End If
    Else
        PopulateDBCombo = False
        puFormSQLConnection.uSQLConnection.IsValid = False
    End If
Else
    PopulateDBCombo = False
    puFormSQLConnection.uSQLConnection.IsValid = False
End If
mbooDBListIsDirty = False
Call ImageShow
Exit Function

errh:
PopulateDBCombo = False
Err.Clear
Exit Function
End Function

Private Function ComboItemDataExists(cmb As ComboBox, lItemData As Long, Optional ByVal lRetNoItem As Long = vbObjectError) As Long
Dim cnt&
Dim l&
On Error GoTo errh
l = 0
For cnt = 0 To cmb.ListCount - 1
    If cmb.ItemData(cnt) = lItemData Then
        l = cnt
    End If
Next cnt
ComboItemDataExists = l
GoSub LocalClean
Exit Function

errh:
On Error Resume Next
ComboItemDataExists = False
GoSub LocalClean
Err.Clear
Exit Function

LocalClean:
On Error Resume Next
Err.Clear
Return
End Function

Private Function ComboTextItemExists(cmb As ComboBox, strItem As String, Optional ByVal lOutNoItem As Long = vbObjectError) As Long
Dim cnt&
Dim l As Long
On Error GoTo errh
l = lOutNoItem
For cnt = 0 To cmb.ListCount - 1
    If LCase$(cmb.List(cnt)) = LCase$(strItem) Then
        l = cnt
        Exit For
    End If
Next cnt
ComboTextItemExists = l
GoSub LocalClean
Exit Function
errh:
On Error Resume Next
ComboTextItemExists = False
GoSub LocalClean
Err.Clear
Exit Function
LocalClean:
On Error Resume Next
Err.Clear
Return
End Function

Private Sub ImageShow()
On Error Resume Next
If puFormSQLConnection.uSQLConnection.IsValid Then
    Me.imgSrvConnect.Visible = True
    Me.imgSrvNotConnect.Visible = False
Else
    Me.imgSrvConnect.Visible = False
    Me.imgSrvNotConnect.Visible = True
End If
DoEvents
End Sub

Private Sub CF()
On Error Resume Next
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
DoEvents
End Sub

Private Sub Form_Load()
Dim b As Boolean
Dim x&, y&
Dim cFL As New FuncLib_Lib
On Error GoTo errh
Me.Timer1.Enabled = False
b = cFL.CursorGetPosInTwips(x, y)
If b Then Me.Top = y: Me.Left = x
MPH
mbooLoaded = False
Call ImageShow
mbooDBListIsDirty = True
Me.imgSrvConnect.Top = 1200
Me.imgSrvConnect.Left = 2910
Me.imgSrvNotConnect.Top = Me.imgSrvConnect.Top
Me.imgSrvNotConnect.Left = Me.imgSrvConnect.Left
Me.cmbServerName.Clear
Me.cmbServerName.Text = ""
GoSub LC
Exit Sub

errh:
Resume e
e:
On Error Resume Next
GoSub LC
Err.Clear
Exit Sub

LC:
On Error Resume Next
D cFL
MPD
Return
End Sub


Private Sub MPH()
Screen.MousePointer = vbHourglass
DoEvents
End Sub

Private Sub MPD()
Screen.MousePointer = vbDefault
DoEvents
End Sub

Private Function ComboFillWithSingleArray(cmb As ComboBox, sArray() As String, Optional ByVal SelArrayVal As Variant = -1) As Boolean
Dim b As Boolean
Dim cnt&
Dim SelArrayInd&, lLI&
Dim sSel$
On Error GoTo errh
cmb.Clear
SelArrayInd = -1
If VarType(SelArrayVal) = vbLong Or VarType(SelArrayVal) = vbInteger Or VarType(SelArrayVal) = vbByte Then
    SelArrayInd = CLng(SelArrayVal)
ElseIf VarType(SelArrayVal) = vbString Then
    sSel = CStr(SelArrayVal)
Else
    GoTo errh
End If

For cnt = 0 To UBound(sArray)
    cmb.AddItem sArray(cnt)
    If LCase$(sArray(cnt)) = LCase$(sSel) Then SelArrayInd = cnt
Next cnt

If SelArrayInd > -1 Then
    cmb.ListIndex = SelArrayInd
ElseIf cmb.ListCount > 0 Then
    cmb.ListIndex = 0
End If

ComboFillWithSingleArray = True
Exit Function
errh:
ComboFillWithSingleArray = False
Err.Clear
Exit Function
End Function

Private Sub Form_Unload(Cancel As Integer)
Dim cPWD As New Password_LIB
On Error Resume Next
With puFormSQLConnection.uSQLConnection
    .DatabaseName = Me.cmbDatabaseName.Text
    .PWD = Me.txtPWD.Text
    cPWD.Key = App.ExeName
    .PWD_Encoded = cPWD.XOr_CodeFunc(.PWD)
    .UID = Me.txtUID.Text
    .ServerName = Me.cmbServerName.Text
    .Security = Me.cmbSecurity.ItemData(Me.cmbSecurity.ListIndex)
End With

D mcFuncLib
D mcDMO
D mcReg
D cPWD
End Sub


Private Sub Timer1_Timer()
Dim b As Boolean
On Error Resume Next
If cmbSecurity.ListIndex > -1 Then
    Select Case cmbSecurity.ListIndex
        Case eSecurityLI.NTSecurity
            Me.txtPWD.Enabled = False
            Label4.Enabled = False
            Label5.Enabled = False
            Me.txtUID.Enabled = False
        Case Else
            Me.txtPWD.Enabled = True
            Me.txtUID.Enabled = True
            Label4.Enabled = True
            Label5.Enabled = True
    End Select
End If
If Not mbooDBListIsDirty Then Exit Sub

Screen.MousePointer = vbHourglass
DoEvents
If Not Me.x_ManualConfig Then b = PopulateDBCombo
puFormSQLConnection.uSQLConnection.IsValid = False
Call ImageShow
Screen.MousePointer = vbDefault
End Sub

Private Sub txtPWD_change()
On Error Resume Next
puFormSQLConnection.uSQLConnection.IsValid = False
Call ImageShow
End Sub

Private Sub txtUID_change()
On Error Resume Next
puFormSQLConnection.uSQLConnection.IsValid = False
Call ImageShow
End Sub
