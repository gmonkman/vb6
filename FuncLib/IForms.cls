VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmSQLConnection As frmSQLConnection
Public Event frmSQLConnectionClosed(uFrmSQLCnn As udtFrmSQLConnection)

Private mfrmSQLQuery As frmSQLQuery
Private mlfrmSQLQueryHwnd&


'******************
'* SQL Connection *
'******************
Public Property Get Form_SQLConnection() As udtFrmSQLConnection
Form_SQLConnection = puFormSQLConnection
End Property

Public Sub frmSQLConnection_Init(uSQLDataCnn As udtSQLDataConnection)
On Error GoTo errh
Set mfrmSQLConnection = New frmSQLConnection
puFormSQLConnection.uSQLConnection = uSQLDataCnn
puFormSQLConnection.Cancelled = False
frmSQLConnection.Show vbModal
RaiseEvent frmSQLConnectionClosed(puFormSQLConnection)
GoSub LC
Exit Sub
errh:
Resume e
e:
On Error Resume Next
Err.Clear
Exit Sub

LC:
On Error Resume Next
D mfrmSQLConnection
Return

End Sub



'******************
'* SQL Query Tool *
'******************
Public Property Get frmSQLQueryHWnd() As Long
frmSQLQueryHWnd = mlfrmSQLQueryHwnd
End Property

Public Sub frmSQLQueryOpen(ByRef Cnn As adodb.Connection, Optional ByVal Mode As VBRUN.FormShowConstants = vbModeless)
If mlfrmSQLQueryHwnd = 0 Then
    Set mfrmSQLQuery = New frmSQLQuery
End If
Load mfrmSQLQuery
mfrmSQLQuery.x_Connection = Cnn
mfrmSQLQuery.Show Mode
mlfrmSQLQueryHwnd = mfrmSQLQuery.hWnd
End Sub

Public Sub frmSQLQueryClose()
If mlfrmSQLQueryHwnd <> 0 And Not mfrmSQLQuery Is Nothing Then
    Unload mfrmSQLQuery
End If
D mfrmSQLQuery
mlfrmSQLQueryHwnd = 0
End Sub

Private Sub D(obj As Object)
On Error Resume Next
If TypeOf obj Is adodb.Recordset Or TypeOf obj Is adodb.Connection Then
    obj.Close
End If
If Not obj Is Nothing Then Set obj = Nothing
Err.Clear
End Sub

