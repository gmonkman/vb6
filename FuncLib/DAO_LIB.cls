VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAO_LIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eDAOFind
    FindFirst = 1
    FindLast = 2
    FindNext = 3
    FindPrevious = 4
End Enum

Public Enum eDAOTypes
    BigInt = dbBigInt
    Binary = dbBinary
    Boolean_ = dbBoolean
    Byte_ = dbByte
    char = dbChar
    Currency_ = dbCurrency
    Date_ = dbDate
    Decimal_ = dbDecimal
    Double_ = dbDouble
    Float = dbFloat
    GUID_ = dbGUID
    Integer_ = dbInteger
    Long_ = dbLong
    LongBinary = dbLongBinary
    Memo = dbMemo
    Numeric = dbNumeric
    Single_ = dbSingle
    Text_ = dbText
    Time_ = dbTime
    TimeStamp = dbTimeStamp
End Enum

Public Enum edaoSysObjTypes
    sotStandard = 1
End Enum

Public Enum edaoSysObjFlags
    sotUserTable = 0
End Enum

Private Const BLOCK_SIZE = 32768

Private Function DBTableExists(ByRef Cnn As ADODB.Connection, ByVal TableName As String, ByRef lOutErr As Long) As Boolean
Dim sql$
Dim rs As ADODB.Recordset
On Error GoTo errh
sql = "SELECT name FROM MSysObjects WHERE Name='" & TableName & "';"
rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
DBTableExists = RP(rs)
GoSub LC
Exit Function

LC:
 On Error Resume Next
 d rs
Return

errh:
 lOutErr = Err.Number
 On Error Resume Next
 GoSub LC
End Function

Public Function FieldsEnum(db As DAO.Database, Table As String) As String()
Dim a() As String
Dim cnt&
Dim td As DAO.TableDef
Dim fld As DAO.field
Set td = db.TableDefs(Table)

For Each fld In td
    ReDim Preserve a(cnt)
    a(cnt) = td.Name
    cnt = cnt + 1
Next
FieldsEnum = a()
End Function

Public Function TablesEnum(db As DAO.Database) As String()
Dim a() As String
Dim cnt&
Dim tb As DAO.TableDef
For Each tb In db.TableDefs
    If tb.Attributes = 0 Then
        ReDim Preserve a(cnt)
        a(cnt) = tb.Name
        cnt = cnt + 1
    End If
Next
TablesEnum = a()
End Function

Public Function DatabaseRepair(db As DAO.Database, ByRef bInsuffSpace As Boolean, Optional ByRef lOutErr As Long) As Boolean
Dim cFSO As New FSO_LIB
Dim cSys As New SysInfo_Lib
Dim cStrings As New Strings_LIB
Dim sTemp$, sTmpPath$, sDBCopy$
Dim scrFile As Scripting.File
Dim dblSizeInBytes#, dblSpaceInBytes#
Dim uDisk As udtDiskInformation
Dim b As Boolean
On Error GoTo errh

b = cFSO.GetFileInfo(db.Name, scrFile)
If Not b Then
    DatabaseRepair = False
    lOutErr = vbObjectError
    GoSub LC
    Exit Function
End If

dblSizeInBytes = CDbl(scrFile.Size)
b = cFSO.GetDiskSpace(sTemp, uDisk)
If Not b Then
    DatabaseRepair = False
    lOutErr = vbObjectError
    GoSub LC
    Exit Function
End If

dblSpaceInBytes = uDisk.NumFreeBytes
If (4 * dblSizeInBytes) < dblSpaceInBytes Then
    DatabaseRepair = False
    bInsuffSpace = True
    GoSub LC
    Exit Function
End If

Do
    sTemp = cFSO.GenRndFileName(".mdb")
    sTmpPath = cSys.EnvironGetTempDir
    sDBCopy = cStrings.AddRightSlash_F(sTmpPath) & sTemp
    If cFSO.FileExists(sDBCopy) Then sDBCopy = ""
Loop Until sDBCopy <> ""

b = cFSO.CopyFile(db.Name, sTemp)
If Not b Then
    DatabaseRepair = False
    lOutErr = vbObjectError
    GoSub LC
    Exit Function
End If

DBEngine.RepairDatabase sDBCopy

b = cFSO.RenameFile(db.Name, db.Name & ".tmp")
If Not b Then
    DatabaseRepair = False
    lOutErr = vbObjectError
    GoSub LC
    Exit Function
End If

b = cFSO.CopyFile(sDBCopy, db.Name)
If Not b Then
    b = cFSO.RenameFile(db.Name & ".tmp", db.Name)
    DatabaseRepair = False
    lOutErr = vbObjectError
    GoSub LC
    Exit Function
End If
DatabaseRepair = True
GoSub LC
Exit Function

errh:
lOutErr = Err.Number
Resume e
e:
On Error Resume Next
GoSub LC
DatabaseRepair = False
Err.Clear
Exit Function

LC:
On Error Resume Next
b = cFSO.DeleteFile(sDBCopy)
d cFSO
d cSys
d cStrings
Return

End Function

Private Sub d(obj As Object)
On Error Resume Next
If TypeOf obj Is ADODB.Recordset Or TypeOf obj Is ADODB.Connection Then
    obj.Close
ElseIf TypeOf obj Is VB.Form Then
    Unload obj
    DoEvents
End If
If Not obj Is Nothing Then Set obj = Nothing
Err.Clear
End Sub


Private Function Exist_(obj As Object) As Boolean
On Error GoTo errh
Exist_ = IIf(obj Is Nothing, False, True)
Exit Function
errh:
Exist_ = False
Exit Function
End Function

Public Function ErrorIsCorrupt(lErr As Long) As Boolean
On Error GoTo errh
Select Case lErr
    Case 3049, 3340, 3428, 3070, 3078
        ErrorIsCorrupt = True
    Case Else
        ErrorIsCorrupt = False
End Select
Exit Function

errh:
Resume e
e:
On Error Resume Next
ErrorIsCorrupt = False
Err.Clear
Exit Function

End Function


Public Function RelationExists(ByVal relname, TestDB As DAO.Database, Optional ByRef vErr As Variant) As Boolean
Dim i%, n%
Dim RetVal As Boolean
On Error GoTo errh
Do
   n = TestDB.Relations.Count
   RetVal = False
   For i = 0 To n - 1
      If UCase(TestDB.Relations(i).Name) = UCase(relname) Then
         RetVal = True
         Exit For
      End If
   Next
Loop Until True

GoSub LocalClean

errh:
GoSub LocalClean
Exit Function

LocalClean:
If Not IsMissing(vErr) Then vErr = Err.Number
RelationExists = RetVal
Return

End Function

Public Function FieldExists(ByVal TableName As String, ByVal field As String, TestDB As DAO.Database, Optional ByRef vErr As Variant) As Boolean
Dim i%, n%
Dim RetVal As Boolean
Dim mydef As TableDef

On Error GoTo errh
RetVal = False
Do
   Set mydef = TestDB.TableDefs(TableName)
   n = mydef.Fields.Count
   For i = 0 To n - 1
      If UCase(TestDB.TableDefs(TableName).Fields(i).Name) = UCase(field) Then
         RetVal = True
      End If
   Next
Loop Until True

FieldExists = RetVal
GoSub LocalClean
Exit Function

errh:
FieldExists = False
GoSub LocalClean
Exit Function

LocalClean:
If Not IsMissing(vErr) Then vErr = Err.Number
On Error Resume Next
If Not mydef Is Nothing Then Set mydef = Nothing
Return

End Function

'Microsoft Jet Databases Only
Public Function CreateForeignRelation(strRelationName$, strPrimaryTable$, strPrimaryField$, strForeignTable, strForeignField$, strDatabase$, Optional vErr As Variant) As Boolean
Dim relRelation As DAO.Relation
Dim dbDatabase As DAO.Database
Dim fldPrimary As DAO.field, fldRelPrimary As DAO.field
Dim fldForeign As DAO.field, fldRelForeign As DAO.field

On Error GoTo errh
Set dbDatabase = OpenDatabase(strDatabase, False, False)

With dbDatabase
    Set relRelation = New Relation
    Set relRelation = dbDatabase.CreateRelation(strRelationName, strPrimaryTable, strForeignTable, dbRelationInherited)
    Set fldPrimary = dbDatabase.TableDefs(strPrimaryTable).Fields(strPrimaryField)
    Set fldForeign = dbDatabase.TableDefs(strForeignTable).Fields(strForeignField)
End With

With relRelation
    Set fldRelPrimary = .CreateField(fldPrimary.Name)
    Set fldRelForeign = .CreateField(fldForeign.Name)
    fldRelPrimary.Name = fldPrimary.Name
    fldRelForeign.ForeignName = fldForeign.Name
    .Fields.Append fldRelPrimary
    .Fields.Append fldRelForeign
End With

dbDatabase.Relations.Append relRelation

GoSub LocalClean
CreateForeignRelation = True

Exit Function


errh:
CreateForeignRelation = False
GoSub LocalClean
Exit Function


LocalClean:
If Not IsMissing(vErr) Then vErr = Err.Number
On Error Resume Next
If Not fldPrimary Is Nothing Then Set fldPrimary = Nothing
If Not fldForeign Is Nothing Then Set fldForeign = Nothing
If Not fldRelForeign Is Nothing Then Set fldRelForeign = Nothing
If Not fldRelPrimary Is Nothing Then Set fldRelPrimary = Nothing
If Not relRelation Is Nothing Then Set relRelation = Nothing
If Not dbDatabase Is Nothing Then dbDatabase.Close: Set dbDatabase = Nothing
Return
End Function


Public Function TableExists(ByVal TableName As String, TestDB As DAO.Database, Optional vErr As Variant) As Boolean
Dim i%, n%, RetVal%
On Error GoTo errh

n = TestDB.TableDefs.Count

RetVal = False
For i = 0 To n - 1
   If UCase$(TestDB.TableDefs(i).Name) = UCase$(TableName) Then
      RetVal = True
      Exit For
   End If
Next

TableExists = RetVal
If Not IsMissing(vErr) Then vErr = Err.Number

Exit Function

errh:
TableExists = False
If Not IsMissing(vErr) Then vErr = Err.Number
Exit Function

End Function


'# VB 5 will missreport recordcounts, especially with
'# non forwardonly recordsets
Public Function RecordsPresent(rs As DAO.Recordset) As Boolean

On Error Resume Next
  
If rs.EOF And rs.BOF Then
  RecordsPresent = False
  Exit Function
End If

If rs.EOF Then
  rs.MovePrevious
ElseIf rs.BOF Then
  rs.MoveNext
End If

If Err Then Err.Clear
On Error GoTo 0

If rs.EOF And rs.BOF Then
  RecordsPresent = False
Else
  RecordsPresent = True
End If
  
End Function


Public Function GetSQLFieldList(strDatabase As String, strTableName As String) As String
Dim tdfTable As DAO.TableDef
Dim dbDatabase As DAO.Database
Dim strSQLBuild$, strTemp$
Dim lngI&

On Error GoTo errGSQLFL

Set dbDatabase = OpenDatabase(strDatabase, False, True)
Set tdfTable = dbDatabase.TableDefs(strTableName)

For lngI = 0 To tdfTable.Fields.Count - 1
    strTemp = tdfTable.Fields(lngI).Name
 '   If pIsKeyName(strTemp) Then strTemp = pAddSquareBrackets(strTemp)
    If lngI > 0 Then
        strSQLBuild = strSQLBuild & ", " & strTemp
    Else
        strSQLBuild = strTemp
    End If
Next lngI
GetSQLFieldList = strSQLBuild
GoSub LocalClean
Exit Function

errGSQLFL:
GetSQLFieldList = False
GoSub LocalClean
Exit Function

LocalClean:
On Error Resume Next
If Not dbDatabase Is Nothing Then dbDatabase.Close: Set dbDatabase = Nothing
Return
End Function

Private Function RP(rs As Object) As Boolean
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
Exit Function
End Function

Private Function ReadField(ByRef vIn As Variant, ByVal vRetVal As Variant) As Variant
On Error GoTo errh
If IsNull(vIn) Then
    ReadField = vRetVal
Else
    ReadField = vIn
End If
Exit Function
errh:
ReadField = vIn
Exit Function
End Function


Public Function GetRefFromFieldString(strRefToGet$, strFieldWithKnownValue$, strValue$, strTableName$, ByRef varReturn As Variant, dbDatabase As DAO.Database) As Boolean
Dim strSQL$
Dim rsTarget As DAO.Recordset
On Error GoTo errh
strSQL = "Select DISTINCTROW " & strRefToGet & ", " & strFieldWithKnownValue
strSQL = strSQL & " FROM " & strTableName
strSQL = strSQL & " WHERE " & strFieldWithKnownValue & "=" & Chr$(34) & strValue & Chr$(34) & ";"

Set rsTarget = dbDatabase.OpenRecordset(strSQL, dbOpenForwardOnly)

If Not RP(rsTarget) Then
    GetRefFromFieldString = False: varReturn = 0
    GoSub LC
    Exit Function
End If
varReturn = rsTarget(strRefToGet)
GoSub LC
GetRefFromFieldString = True
Exit Function

errh:
Resume e
e:
GoSub LC
GetRefFromFieldString = False
Exit Function

LC:
On Error Resume Next
d rsTarget
Return

End Function


Public Function FindAny(eFindChoice As eDAOFind, rstTemp As DAO.Recordset, strFind As String) As Boolean    ' Use Find method based on user input.
On Error GoTo errh
   Select Case eFindChoice
    Case eDAOFind.FindFirst
        rstTemp.FindFirst strFind
    Case eDAOFind.FindLast
        rstTemp.FindLast strFind
    Case eDAOFind.FindNext
         rstTemp.FindNext strFind
    Case eDAOFind.FindPrevious
         rstTemp.FindPrevious strFind
    End Select
   ' Set return value based on NoMatch property.
   FindAny = IIf(rstTemp.NoMatch, False, True)
Exit Function

errh:
MsgBox Err.description
FindAny = False
Exit Function
End Function


Public Function AttachTable(TableName$, todata$, fromdata$) As Boolean
Dim tbl() As New DAO.TableDef
Dim ToDB As DAO.Database, FromDB As DAO.Database

On Error GoTo errAT

Set ToDB = OpenDatabase(todata, False, False)
Set FromDB = OpenDatabase(fromdata, False, False)

On Error Resume Next
ToDB.TableDefs.Delete TableName
If Err = 3265 Then
    Err = 0
ElseIf Err Then
    MsgBox Error$
    AttachTable = False
    GoSub LocalClean
    Exit Function
End If
On Error GoTo errAT

ReDim tbl(0) As New TableDef

tbl(0).Connect = ";DATABASE=" & FromDB.Name
tbl(0).SourceTableName = TableName
tbl(0).Name = TableName
ToDB.TableDefs.Append tbl(0)

GoSub LocalClean
AttachTable = True

Exit Function

errAT:
AttachTable = False
GoSub LocalClean
Exit Function


LocalClean:
On Error Resume Next
    If Not ToDB Is Nothing Then ToDB.Close: Set ToDB = Nothing
    If Not FromDB Is Nothing Then FromDB.Close: Set FromDB = Nothing
Return

End Function


Function RetypeField(db As DAO.Database, sTableName As String, sFieldName As String, sCopyToField As String, dbDataType As Long, Optional iFieldSize As Integer, Optional ByRef vErr As Variant) As Boolean
Dim td As DAO.TableDef
Dim sql$
Dim fldSource As DAO.field
Dim lStartType As Long
On Error GoTo errh

Set td = db.TableDefs(sTableName)

Set fldSource = db.TableDefs(sTableName).Fields(sFieldName)
fldSource.Name = sCopyToField
db.TableDefs.Refresh

lStartType = fldSource.Type

If iFieldSize > 0 Then
    td.Fields.Append td.CreateField(sFieldName, dbDataType, iFieldSize)
Else
    td.Fields.Append td.CreateField(sFieldName, dbDataType)
End If

sql = "UPDATE " & sTableName & " SET " & sFieldName & " = " & sCopyToField & ";"
db.Execute sql, dbFailOnError

db.TableDefs(sTableName).Fields.Delete sCopyToField

RetypeField = True
GoSub LocalClean

Exit Function

errh:
If Not IsMissing(vErr) Then vErr = Err.Number
GoSub LocalClean
RetypeField = False
Exit Function

LocalClean:
On Error Resume Next
Set fldSource = Nothing
Set td = Nothing
Return

End Function


Public Function QueryFieldType(db As DAO.Database, sTable As String, sfield As String) As Long
Dim td As DAO.TableDef
Set td = db.TableDefs(sTable)
QueryFieldType = td.Fields(sfield).Type
Set td = Nothing
End Function


Public Function WriteBLOBToDB(strSourceFile As String, strSQL As String, strFieldName As String, dbSource As DAO.Database) As Long
Dim rsDAO As DAO.Recordset
Dim intNumBlocks As Integer, strSourceFileFile As Integer, i As Integer
Dim lngFileLength As Long, lngLeftOver As Long
Dim strFileData As String

' Open the strSourceFile file.
strSourceFileFile = FreeFile
Open strSourceFile For Binary Access Read As strSourceFileFile

' Get the length of the file.
lngFileLength = LOF(strSourceFileFile)
If lngFileLength = 0 Then
    WriteBLOBToDB = 0
    Exit Function
End If

' Calculate the number of blocks to read and lngLeftOver bytes.
intNumBlocks = lngFileLength \ BLOCK_SIZE
lngLeftOver = lngFileLength Mod BLOCK_SIZE

Set rsDAO = dbSource.OpenRecordset(strSQL, dbOpenDynaset)
rsDAO.Edit

' Read the lngLeftOver data, writing it to the table.
strFileData = String$(lngLeftOver, 32)
Get strSourceFileFile, , strFileData
rsDAO(strFieldName).AppendChunk (strFileData)

' Read the remaining blocks of data, writing them to the table.
strFileData = String$(BLOCK_SIZE, 32)
For i = 1 To intNumBlocks
    Get strSourceFileFile, , strFileData
    rsDAO(strFieldName).AppendChunk (strFileData)
Next i
rsDAO.Update
DBEngine.Idle
Close strSourceFileFile

WriteBLOBToDB = lngFileLength

End Function


Public Function WriteBLOBToDisk(strSQL As String, strFieldName As String, strFileName As String, dbSource As DAO.Database) As Long
Dim rsDAO As DAO.Recordset
Dim intNumBlocks As Integer, intDestFile As Integer, i As Integer
Dim lngFileLength As Long, lngLeftOver As Long
Dim strFileData As String

Set rsDAO = dbSource.OpenRecordset(strSQL, dbOpenForwardOnly)
' Get the size of the field.
lngFileLength = rsDAO(strFieldName).FieldSize()
If lngFileLength = 0 Then
    WriteBLOBToDisk = 0
    Exit Function
End If

' Calculate number of blocks to write and LeftOver bytes.
intNumBlocks = lngFileLength \ BLOCK_SIZE
lngLeftOver = lngFileLength Mod BLOCK_SIZE

' Remove any existing file.
intDestFile = FreeFile
Open strFileName For Binary Access Write As intDestFile


' Write the lngLeftOver data to the output file.
strFileData = rsDAO(strFieldName).GetChunk(0, lngLeftOver)
Put intDestFile, , strFileData

' Write the remaining blocks of data to the output file.
For i = 1 To intNumBlocks
    ' Reads a chunk and writes it to output file.
    strFileData = rsDAO(strFieldName).GetChunk((i - 1) * BLOCK_SIZE _
       + lngLeftOver, BLOCK_SIZE)
    Put intDestFile, , strFileData
Next i

Close intDestFile
WriteBLOBToDisk = lngFileLength
End Function

Private Sub DestroyObj(obj As Object)
On Error Resume Next
obj.Close
Set obj = Nothing
End Sub

Public Function ColumnGetDataTypeDesc(ByVal TypeEnum As DAO.DatabaseTypeEnum) As String
Dim s$
Select Case TypeEnum
    Case dbBigInt, dbInteger, dbLong
        s = "Interger"
    Case dbBinary, dbLongBinary, dbVarBinary
        s = "Binary"
    Case dbBoolean
        s = "Boolean"
    Case dbChar, dbText
        s = "Text"
    Case dbCurrency, dbDecimal, dbDouble, dbFloat, dbNumeric, dbSingle
        s = "Float"
    Case dbDate, dbTime
        s = "Date/Time"
    Case dbGUID
        s = "GUID"
    Case dbMemo
        s = "Memo"
    Case dbTimeStamp
        s = "Timestamp"
End Select
ColumnGetDataTypeDesc = s
End Function

