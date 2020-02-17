VERSION 5.00
Object = "{343F59D0-FE0F-11D0-A89A-0000C02AC6DB}#1.0#0"; "SSTBars.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSQLQuery 
   Caption         =   "SQL Query Tool (Advanced Users Only!)"
   ClientHeight    =   5115
   ClientLeft      =   6045
   ClientTop       =   3840
   ClientWidth     =   7170
   Icon            =   "frmSQLQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   7170
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   9022
      _Version        =   196609
      AutoSize        =   1
      PaneTree        =   "frmSQLQuery.frx":27A2
      Begin VB.TextBox txtResult 
         Height          =   1980
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Text            =   "frmSQLQuery.frx":27F4
         Top             =   3105
         Width           =   7110
      End
      Begin VB.TextBox txtQuery 
         Height          =   2985
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Text            =   "frmSQLQuery.frx":27FA
         Top             =   30
         Width           =   7110
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   6480
      Top             =   4470
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   65541
      ToolBarsCount   =   1
      ToolsCount      =   7
      Tools           =   "frmSQLQuery.frx":2800
      ToolBars        =   "frmSQLQuery.frx":80E0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".txt"
      DialogTitle     =   "Open Query"
      Filter          =   "Text File (*.txt)|*.txt|SQL (*.sql)|*.sql|Rich text (*.rtf)|*.rtf"
   End
End
Attribute VB_Name = "frmSQLQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mCnn As ADODB.Connection
Private mrs As New ADODB.Recordset

Private mcFSO As New FSO_LIB

Private Type udtErr
    Number As Long
    description As String
    Source As String
End Type
Private muErr As udtErr

Private Sub PopErr()
If muErr.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.description
End Sub
Private Sub PushErr()
muErr.description = Err.description
muErr.Number = Err.Number
muErr.Source = Err.Source
End Sub
Private Sub ClearErr()
With muErr
    .description = ""
    .Source = ""
    .Number = 0
End With
End Sub

Public Property Let x_Connection(ByRef Cnn As ADODB.Connection)
Set mCnn = Cnn
End Property

Private Sub Form_Activate()
Static Loaded As Boolean
Dim bInv As Boolean
On Error GoTo errh
If Loaded Then Exit Sub
Loaded = True
Me.SSActiveToolBars1.Tools("ID_SaveResults").Enabled = False
If mCnn Is Nothing Then
    bInv = True
ElseIf mCnn.State = ObjectStateEnum.adStateClosed Then
    bInv = True
End If
If bInv = True Then
    MsgBox "No valid database connection found." & vbCrLf & vbCrLf & "This form will now close.", vbCritical, "Invalid Connection"
    Unload Me
End If

With Me.CommonDialog1
    .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNShareAware
    .InitDir = App.Path
End With
Exit Sub

errh:
 MsgBox Err.description
 Err.Clear
End Sub


Private Sub MPD()
On Error Resume Next
Screen.MousePointer = vbDefault
DoEvents
End Sub

Private Sub MPH()
On Error Resume Next
Screen.MousePointer = vbHourglass
DoEvents
End Sub

Private Sub Form_Load()
On Error Resume Next
CF
Me.txtQuery.Text = ""
Me.txtResult.Text = ""
Err.Clear
End Sub

Private Sub CF()
On Error Resume Next
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
DoEvents
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.SSSplitter1.Height = Me.ScaleHeight
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
D mrs
Err.Clear
End Sub

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

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
Dim l&, m&, n&
Dim s$, s1$
Dim aSQL() As String
Dim dbl#
On Error GoTo errh
MPH
Select Case Tool.ID
    Case "ID_Paste"    '(Button)
        Me.txtQuery.Text = VB.Clipboard.GetText
    Case "ID_Copy"    '(Button)
        VB.Clipboard.SetText Me.txtQuery.SelText
    Case "ID_Cut"    '(Button)
        VB.Clipboard.SetText Me.txtQuery.SelText
        l = Me.txtQuery.SelLength
        m = Me.txtQuery.SelStart + 1
        s = Mid$(Me.txtQuery.Text, m, l)
        s1 = Me.txtQuery.Text
        s1 = Replace(Me.txtQuery.Text, s, "", m, 1)
        s1 = Left$(Me.txtQuery.Text, m - 1) & s1
        Me.txtQuery.Text = s1
    Case "ID_Test"    '(Button)
        If Me.txtQuery.Text <> "" Then
            If Not mrs Is Nothing Then D mrs
            Me.txtResult.Text = ""
            Me.SSActiveToolBars1.Tools("ID_SaveResults").Enabled = False
            s = Me.txtQuery.Text
            s = Replace(s, vbCrLf, " ", , , vbTextCompare)
            s = s & " "
            aSQL = Split(s, " go ", , vbTextCompare)
            mCnn.BeginTrans
            On Error Resume Next
            For n = 0 To UBound(aSQL)
                mCnn.Execute aSQL(n), l
                If Err.Number = 0 Then
                    Me.txtResult.Text = CStr(l) & " record(s) will be affected by query " & CStr(l) & vbCrLf & vbCrLf
                Else
                    Me.txtResult.Text = "Invalid syntax." & vbCrLf & vbCrLf & "The error was " & vbCrLf & Err.description
                    Exit For
                End If
            Next n
            mCnn.RollbackTrans
            On Error GoTo 0
        End If
    Case "ID_Run"    '(Button)
        On Error Resume Next
        If Me.txtQuery.Text <> "" Then
            If mrs Is Nothing Then Set mrs = New ADODB.Recordset
            If mrs.State <> adStateClosed Then mrs.Close
            Me.txtResult.Text = ""
            Me.SSActiveToolBars1.Tools("ID_SaveResults").Enabled = False
            s = Me.txtQuery.Text
            s = Replace(s, vbCrLf, " ", , , vbTextCompare)
            s = s & " "
            aSQL = Split(s, " go ", , vbTextCompare)
            mCnn.BeginTrans
            On Error Resume Next
            For n = 0 To UBound(aSQL)
                If Trim$(aSQL(n)) = "" Then GoTo ENDLOOP
                mCnn.Execute aSQL(n), l
                If l = -1 Then l = 0
                If Err.Number = 0 Then
                    Me.txtResult.Text = Me.txtResult.Text & CStr(l) & " record(s) affected by query " & CStr(l) & vbCrLf & vbCrLf
                Else
                    Me.txtResult.Text = Me.txtResult.Text & "Invalid syntax." & vbCrLf & vbCrLf & "The error was " & vbCrLf & Err.description
                    Exit For
                End If
ENDLOOP:
            DoEvents
            Next n
            If Err Then
                mCnn.RollbackTrans
            Else
                mCnn.CommitTrans
            End If
            On Error GoTo errh
        End If
    Case "ID_SaveResults"
        If RP(mrs) Then
            dbl = Rnd
            s = "~SQL" & Format$(Date, "yyyymmmdd") & "_" & Replace(Format$(dbl, "#.0000"), ".", "") & ".xml"
            s = EnvironGetTempDir & "\" & s
            mrs.Save s, adPersistXML
            Me.txtResult.Text = Me.txtResult.Text & vbCrLf & vbCrLf & "Results saved to file " & s
        Else
            MsgBox "No records to dump to disk."
        End If
    Case "ID_open"
            Me.CommonDialog1.ShowOpen
            If Me.CommonDialog1.FileName = "" Then GoSub LC: Exit Sub
            Me.txtQuery.Text = mcFSO.FileGetText(Me.CommonDialog1.FileName)
    Case Else
End Select
GoSub LC
Exit Sub

LC:
 On Error Resume Next
 MPD
Return

errh:
 MsgBox Err.description
 On Error Resume Next
 GoSub LC
End Sub

Private Function EnvironGetTempDir() As String
Dim strTmp$
On Error GoTo errh
strTmp = Environ$("tmp")
If strTmp = "" Then strTmp = Environ$("temp")
EnvironGetTempDir = strTmp
Exit Function
errh:
Resume e
e:
On Error Resume Next
EnvironGetTempDir = ""
Err.Clear
Exit Function
End Function

Private Function RP(rs As Object) As Boolean
On Error GoTo errh
If rs Is Nothing Then RP = False: Exit Function
If TypeOf rs Is ADODB.Recordset Or TypeOf rs Is DAO.Recordset Then
    If rs.EOF And rs.BOF Then
        RP = False
    Else
        RP = True
    End If
Else
    RP = False
End If
Exit Function
errh:
 On Error Resume Next
 RP = False
 Err.Clear
End Function


