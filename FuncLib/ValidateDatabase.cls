VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidateDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Private mdblVersion#
Private mCnn As New ADODB.Connection
Private mbooValidateRequired As Boolean
Private mstrVersion$
Private mdtUpdateDate As Date
Private mbooMessageTheUser As Boolean
Private mstrManualVersion As String

Private Const DIVIDER$ = "."
Private Const GENERAL_ERRORS = vbObjectError - 1000
Private Const ADO_ERRORS = vbObjectError - 3000
Private Const CRR2 = vbCrLf & vbCrLf
Private Const VALIDATE_REQUIRED_MSG = "Because of an update to the application it is STRONGLY recommended that the database structures are validated AS SOON AS POSSIBLE." & CRR2 & "This can be done from the Tools menu, as an administrative user."

'For this to work the underlying SQL Server permissions for the user must have the admin privs to the database
'because we are making potentially making structure changes.
'Also your project versioning must be on autoincrement, and when you compile you project ensure
'you have the .vbp file checked out, otherwise the autoincrement will not increment!

Public Property Let ManualVersion(ByVal Manual As String)
mstrManualVersion = Manual
mstrVersion = mstrManualVersion
End Property
Public Property Get ManualVersion() As String
ManualVersion = mstrManualVersion
End Property

Public Property Let MessageTheUser(ByVal MessageUser As Boolean)
mbooMessageTheUser = MessageUser
End Property
Public Property Get MessageTheUser() As Boolean
MessageTheUser = mbooMessageTheUser
End Property

Public Property Get LastUpdateDate() As Date
LastUpdateDate = mdtUpdateDate
End Property

Public Property Get VersionAsDouble() As Double
VersionAsDouble = mdblVersion
End Property

Public Property Set Connection(ByRef Cnn As ADODB.Connection)
Set mCnn = Cnn
End Property

Public Property Get ValidateRequired() As Boolean
ValidateRequired = mbooValidateRequired
End Property

Private Function AppGetVersion() As String
Dim s$
s = CStr(App.Major) & DIVIDER & CStr(App.Minor) & DIVIDER & CStr(App.Revision)
AppGetVersion = s
End Function

Private Sub Class_Initialize()
On Error Resume Next
mstrVersion = AppGetVersion
mbooMessageTheUser = False
End Sub

Public Sub Init()
If mstrManualVersion = "" Then
    mdblVersion = AppVersionConvertToNumber(mstrVersion, DIVIDER)
Else
    mdblVersion = AppVersionConvertToNumber(mstrManualVersion, DIVIDER)
End If
Call CreateStructure
Call ExeAdd
Call CheckVersion
End Sub

Private Sub MessageUser()
On Error Resume Next
If Me.ValidateRequired Then MsgBox VALIDATE_REQUIRED_MSG, vbExclamation, "Validate Required"
End Sub

Private Sub CheckVersion()
Dim dblDBVer#
Dim sql$
Dim rs As New ADODB.Recordset
Dim lErr&, sErr$
On Error GoTo errh
sql = "SELECT * FROM ValidateDatabase WHERE ApplicationExeName='" & VB.App.ExeName & "'"
rs.Open sql, mCnn, adOpenForwardOnly, adLockReadOnly, adCmdText
If RP(rs) Then
    dblDBVer = rs("Version").Value
    mbooValidateRequired = (dblDBVer < mdblVersion)
Else
    mbooValidateRequired = True
End If
mdtUpdateDate = IIf(Format(rs("updatedate").Value, "") = "", "1/Jan/1900", rs("updatedate").Value)
GoSub LC
If mbooMessageTheUser Then Call MessageUser
Exit Sub
LC:
 On Error Resume Next
 D rs
Return

errh:
 lErr = Err.Number: sErr = Err.description
 Resume e
e:
 On Error Resume Next
 GoSub LC
 On Error GoTo errh
 Err.Raise lErr, "ValidateDatabase#CheckVersion", sErr
End Sub

'Call after structures successfully validated
Public Sub Refresh()
Dim sql$
'these next 2 calls are incase init hasnt first been called
Call CreateStructure
Call ExeAdd
sql = "UPDATE ValidateDatabase SET UpdateDate='" & Format$(Now, "yyyy mmm dd HH:nn:ss") & "',Version=" & CStr(mdblVersion) & _
    " WHERE ApplicationExeName='" & VB.App.ExeName & "'"
mCnn.Execute sql, , adCmdText
mdtUpdateDate = Now
mbooValidateRequired = False
End Sub

Private Sub CreateStructure()
Dim sql$
Dim lErr&
If Not SQL_TableExists(mCnn, "ValidateDatabase", lErr) Then
    sql = ""
    sql = "CREATE TABLE [dbo].[ValidateDatabase] (" & _
        "[Version] [varchar] (30) NOT NULL ," & _
        "[ApplicationExeName] [varchar] (255) NOT NULL," & _
        "[UpdateDate] [datetime] NULL) ON [PRIMARY]"
    mCnn.Execute sql
    
    sql = "CREATE  UNIQUE  CLUSTERED  INDEX [IX_ValidateDatabase] ON [dbo].[ValidateDatabase]([ApplicationExeName]) ON [PRIMARY]"
    mCnn.Execute sql
    
    sql = "ALTER TABLE [dbo].[ValidateDatabase] WITH NOCHECK ADD" & _
        " CONSTRAINT [DF_ValidateDatabase_Version] DEFAULT (0) FOR [Version]," & _
        " CONSTRAINT [DF_ValidateDatabase_ApplicationExeName] DEFAULT ('') FOR [ApplicationExeName]"
    mCnn.Execute sql
End If
End Sub

Private Function AppVersionConvertToNumber(sVersion As String, Optional DIVIDER As String = ".") As Double
Dim aV() As String
Dim cnt&
Dim Maj&, Min&, Rev&
Dim sMaj$, sMin$, sRev$
Dim s$
On Error GoTo errh
aV = Split(sVersion, DIVIDER, , vbTextCompare)
If UBound(aV) = -1 Then GoTo errh
If UBound(aV) <> 2 Then GoTo errh
Maj = CLng(aV(0))
sMaj = Format$(aV(0), "0000")

Min = CLng(aV(1))
sMin = Format$(aV(1), "00000")

Rev = CLng(aV(2))
sRev = Format$(aV(2), "000000")

s = sMaj & sMin & sRev
AppVersionConvertToNumber = CDbl(s)

Exit Function
errh:
AppVersionConvertToNumber = 0
End Function

Private Sub ExeAdd()
Dim sql$
sql = "IF NOT EXISTS (SELECT * FROM ValidateDatabase WHERE ApplicationExeName='" & App.ExeName & "')" & _
            "INSERT INTO ValidateDatabase (ApplicationExeName) VALUES('" & VB.App.ExeName & "')"
mCnn.Execute sql, , adCmdText
End Sub

Private Function SQL_TableExists(ByRef cn As ADODB.Connection, ByVal TableName As String, ByRef lOutErr&) As Boolean
Dim rs As New ADODB.Recordset
Dim sql$
On Error GoTo errh
sql = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" & TableName & "]')" & _
    " and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
    " Select 1 as Ret else Select 0 as Ret"
Set rs = cn.Execute(sql)
If RP(rs) Then
    If rs("Ret").Value = 1 Then
        SQL_TableExists = True
    Else
        SQL_TableExists = False
    End If
Else
    SQL_TableExists = False
End If
GoSub LC
Exit Function

errh:
 lOutErr = Err.Number
 On Error Resume Next
 GoSub LC
 SQL_TableExists = False
Exit Function

LC:
 On Error Resume Next
 D rs
Return
End Function

Private Sub D(obj As Object)
On Error Resume Next
If Not obj Is Nothing Then
    If TypeOf obj Is VB.Form Then Unload obj
    If TypeOf obj Is ADODB.Recordset Or TypeOf obj Is ADODB.Connection Or TypeOf obj Is DAO.Database Then
        If Not TypeOf obj Is DAO.Database Then
            If obj.State <> ObjectStateEnum.adStateClosed Then obj.Close
        Else
            obj.Close
        End If
    End If
    Set obj = Nothing
End If
Err.Clear
End Sub

Private Function RP(rs As ADODB.Recordset) As Boolean
On Error GoTo errh
If rs Is Nothing Then RP = False: Exit Function
If rs.EOF And rs.BOF Then
    RP = False
Else
    RP = True
End If
Exit Function
errh:
 On Error Resume Next
 RP = False
 Err.Clear
End Function
