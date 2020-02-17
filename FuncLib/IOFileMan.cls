VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOFileMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Private mintFileID As Integer
Private mstrFilePath As String
Private mstrFileName As String
Private mlngCurrentRow As Long
Private mlngRowCount As Long
Private mbooEOF As Boolean

Public Enum enmWriteType
    AppendData = 0
    InsertData = 1
    ReplaceData = 2
End Enum

Public Enum eIOSearchType
    ioForwards
    ioBackwards
    ioFromStart
    ioRange
End Enum

Public Property Get FilePath() As String
FilePath = mstrFilePath
End Property

Public Property Get FileName() As String
FileName = mstrFileName
End Property

Private Sub Class_Initialize()
mintFileID = FreeFile(0)
End Sub

Public Sub OpenFile(strFileLocation As String, pstrFileName As String)
    mintFileID = FreeFile(0)
    If Right(strFileLocation, 1) = "\" Then
        mstrFilePath = strFileLocation
    Else
        mstrFilePath = strFileLocation & "\"
    End If
    mstrFileName = pstrFileName
    If Len(Dir(mstrFilePath & mstrFileName)) > 0 Then SetRowCount
End Sub

Public Function OpenFile_F(strFileLocation As String, pstrFileName As String) As Boolean
On Error GoTo errh
mintFileID = FreeFile(0)

If Right(strFileLocation, 1) = "\" Then
    mstrFilePath = strFileLocation
Else
    mstrFilePath = strFileLocation & "\"
End If
mstrFileName = pstrFileName

If Len(Dir$(mstrFilePath & mstrFileName)) > 0 Then
    SetRowCount
    OpenFile_F = True
Else
    OpenFile_F = False
End If
Exit Function
errh:
 On Error Resume Next
 OpenFile_F = False
End Function

Public Property Get EOF() As Boolean
EOF = mbooEOF
End Property

Public Sub OpenFile2(ByVal FileName As String)
Dim cSt As New Strings_LIB
mintFileID = FreeFile(0)
mstrFilePath = cSt.GetPathOf_F(FileName) & "\"
mstrFileName = cSt.GetFileOf(FileName)
SetRowCount
End Sub

Public Function OpenFile2_F(ByVal FileName As String) As Boolean
Dim cSt As New Strings_LIB
On Error GoTo errh
mintFileID = FreeFile(0)
If Dir$(FileName) <> "" Then
    mstrFilePath = cSt.GetPathOf_F(FileName) & "\"
    mstrFileName = cSt.GetFileOf(FileName)
    SetRowCount
    OpenFile2_F = True
Else
    OpenFile2_F = False
End If
Exit Function
errh:
 On Error Resume Next
 OpenFile2_F = False
End Function

Public Sub MoveFirst()
    Close mintFileID
    If mintFileID > 0 Then Open mstrFilePath & mstrFileName For Input As mintFileID
    mlngCurrentRow = 0
End Sub

Private Sub SetRowCount()
Dim strTemp As String
Dim s$
Close mintFileID
mlngRowCount = 0
s = mstrFilePath & mstrFileName
If Dir$(s) = "" Then 'create file if doesnt exist
    Open s For Random As mintFileID
    Close mintFileID
End If

mintFileID = FreeFile
Open s For Input As mintFileID
Do While Not VBA.FileSystem.EOF(mintFileID)
    Line Input #mintFileID, strTemp
    mlngRowCount = mlngRowCount + 1
Loop
MoveFirst
End Sub

Public Property Get Value() As String
Value = Me.GetDataFromRow(Me.CurrentRow)
End Property

Public Function RowCount() As Long
    RowCount = mlngRowCount
End Function

Public Function CurrentRow() As Long
    CurrentRow = mlngCurrentRow
End Function

Public Function SearchRows(ByVal SearchType As eIOSearchType, ByVal SearchText As String, Optional ByVal AllowPartialMatch As Boolean = False, Optional ByVal StartRow& = 0, Optional ByVal EndRow& = 0, Optional ByVal ReverseRangeSearch As Boolean = False) As Long
Dim b As Boolean
Dim cnt&
Dim lStart&
On Error GoTo errh

If StartRow > 0 Then Call Me.MoveToRow(StartRow)
If Me.RowCount = 0 Then GoTo errh
lStart = Me.CurrentRow
Select Case SearchType
    Case eIOSearchType.ioBackwards
        If Not Me.CurrentRow > 0 Then GoTo errh
        For cnt = Me.CurrentRow To 1 Step -1
            Me.MoveToRow cnt
            GoSub SEARCH
        Next cnt
    Case eIOSearchType.ioForwards
        For cnt = Me.CurrentRow To Me.RowCount
            Me.MoveToRow cnt
            GoSub SEARCH
        Next cnt
    Case eIOSearchType.ioFromStart
            Me.MoveToRow (1)
        For cnt = 1 To Me.RowCount
            Me.MoveToRow cnt
            GoSub SEARCH
        Next cnt
    Case eIOSearchType.ioRange
        If StartRow > EndRow Then GoTo errh
        If EndRow > Me.RowCount Then EndRow = Me.RowCount
        If ReverseRangeSearch Then
            For cnt = EndRow To StartRow Step -1
                Me.MoveToRow cnt
                GoSub SEARCH
            Next cnt
        Else
            For cnt = StartRow To EndRow
                Me.MoveToRow cnt
                GoSub SEARCH
            Next cnt
        End If
End Select
Exit Function

SEARCH:
 If AllowPartialMatch Then
    If InStr(1, Me.GetDataFromRow(Me.CurrentRow), SearchText, vbTextCompare) Then
        SearchRows = cnt
        Exit Function
    End If
 Else
    If Me.GetDataFromRow(Me.CurrentRow) = SearchText Then
        SearchRows = cnt
        Exit Function
    End If
 End If
Return

errh:
 On Error Resume Next
 SearchRows = 0
 Me.MoveToRow lStart
 Err.Clear
End Function

Public Sub MoveToRow(lngRowID As Long)
Dim lngTemp As Long
Dim strTemp As String

If (lngRowID > mlngRowCount) Or (mintFileID = 0) Then mbooEOF = True: Exit Sub
mbooEOF = False
lngTemp = 0

MoveFirst
Do While (Not VBA.FileSystem.EOF(mintFileID)) And (lngTemp < lngRowID)
    Line Input #mintFileID, strTemp
    lngTemp = lngTemp + 1
Loop

mlngCurrentRow = lngTemp
    
End Sub

Public Sub MoveNext()
    MoveToRow mlngCurrentRow + 1
End Sub

Public Sub MovePrevious()
    MoveToRow mlngCurrentRow - 1
End Sub

Public Sub MoveLast()
    MoveToRow mlngRowCount
End Sub

Public Function GetDataFromRow(Optional lngRowID As Long = 0) As String
Dim strTemp As String
If (lngRowID > mlngRowCount) Or (mintFileID = 0) Then GetDataFromRow = "": Exit Function
mlngCurrentRow = lngRowID + Abs((mlngCurrentRow * (lngRowID = 0)))
MoveToRow mlngCurrentRow - 1
Line Input #mintFileID, strTemp
mlngCurrentRow = mlngCurrentRow + 1
GetDataFromRow = strTemp
End Function

Public Sub DeleteRow(Optional lngRowID As Long = 0)
    If mlngRowCount = 0 Then Exit Sub
    WriteData ReplaceData, "", lngRowID
    MoveToRow (Abs((mlngCurrentRow - 1) = 0) + (mlngCurrentRow - 1))
End Sub

'append only
Public Sub WriteData2(strData As String)
Dim cFSO As New FuncLib.FSO_LIB
Dim sFile$
Dim b As Boolean
sFile = mstrFilePath & mstrFileName
b = cFSO.WriteText(sFile, strData)
If Not b Then Err.Raise 2, "IOFileman#WriteData2", "Failed to write text to " & sFile
End Sub

Public Sub WriteData(intWriteType As enmWriteType, strData As String, Optional lngRowID As Long = 0)
    Dim intWriteFileID As Integer
    Dim lngCounter As Long
    Dim lngWriteToRowID As Long
    Dim strTemp As String
    
    If (lngRowID > mlngRowCount) Or (mintFileID = 0) Then Exit Sub
    
    lngCounter = 0
    
    intWriteFileID = FreeFile
    
    Open mstrFilePath & "tmp" & mstrFileName For Output As intWriteFileID
    
    If intWriteType = AppendData Then
        lngWriteToRowID = mlngRowCount
    Else
        lngWriteToRowID = lngRowID + Abs((mlngCurrentRow * (lngRowID = 0))) - 1
    End If
    
    MoveFirst
    
    Do While lngCounter < lngWriteToRowID
        Line Input #mintFileID, strTemp
        Print #intWriteFileID, strTemp
        lngCounter = lngCounter + 1
    Loop

    If Len(Trim(strData)) > 0 Then Print #intWriteFileID, strData
    
    If intWriteType = ReplaceData Then Line Input #mintFileID, strTemp
    
    Do While Not VBA.FileSystem.EOF(mintFileID)
        Line Input #mintFileID, strTemp
        Print #intWriteFileID, strTemp
    Loop

    Close
    
    Kill mstrFilePath & mstrFileName
    Name mstrFilePath & "tmp" & mstrFileName As mstrFilePath & mstrFileName
    
    OpenFile mstrFilePath, mstrFileName
    
    mlngCurrentRow = lngWriteToRowID + 1
    
    MoveToRow lngWriteToRowID + 1

End Sub
Private Sub Class_Terminate()
On Error Resume Next
 Close
End Sub
