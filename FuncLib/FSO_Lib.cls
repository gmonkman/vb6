VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FSO_LIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mbooLetterOnly As Boolean

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias _
      "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
      lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
      lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) _
    As Long

Public Enum eDriveTypes
    Unknown = 0
    NoSuchRootDir = 1
    Removeable = 2
    Fixed = 3
    Remote = 4
    CDRom = 5
    RamDisk = 6
End Enum

Public Function FileCreate(ByVal FileName As String, Optional OverWrite As Boolean = True, Optional Unicode As Boolean = False) As Boolean
Dim fso As New FileSystemObject

Dim TextStream As TextStream
On Error GoTo errh
Set TextStream = fso.CreateTextFile(FileName, OverWrite, Unicode)
GoSub LC
FileCreate = True
Exit Function
LC:
 On Error Resume Next
 D TextStream
 D fso
Return

errh:
 On Error Resume Next
 GoSub LC
 FileCreate = False
End Function

Private Function IsInvalidASCII(ByVal char As String) As Boolean
On Error GoTo errh
If char = "\" Or char = "/" Or char = ":" Or char = "*" Or char = "?" Or Asc(char) = 33 Or char = "<" Or char = ">" Or char = "|" Then
    IsInvalidASCII = True
Else
    IsInvalidASCII = False
End If

Exit Function

errh:
Resume e
e:
On Error Resume Next
IsInvalidASCII = True
Err.Clear
Exit Function

End Function


Public Function GenRndFileName(Optional ByVal AddExtension As String = "") As String
Dim X&
Dim sRet$
Dim arrIn(0 To 4) As Byte, arrOut() As Byte
Dim cPassword As Password_LIB
Dim cFSO As FSO_LIB
Set cPassword = New Password_LIB
Set cFSO = New FSO_LIB
On Error GoTo errh
Randomize
For X = 0 To 4
    arrIn(X) = Int((255 - 0 + 1) * Rnd + 0)
Next X
cPassword.EncodeB64 arrIn(), arrOut()
cPassword.ByteArray2Str arrOut(), sRet
If AddExtension <> "" Then
    If Left$(AddExtension, 1) <> "." Then AddExtension = "." & AddExtension
End If
sRet = cFSO.FileRemoveInvalidASCII(sRet)
GenRndFileName = "~APD" & Left$(sRet, 4) & AddExtension
GoSub LocalClean
Exit Function
errh:
GoSub LocalClean
GenRndFileName = ""
Exit Function

LocalClean:
On Error Resume Next
D cPassword
D cFSO
Err.Clear
Return

End Function

Public Function FileRemoveInvalidASCII(ByVal FileName As String, Optional Substitute As String = "~") As String
Dim cnt&
Dim sFN$
On Error GoTo errh
sFN$ = FileName
If sFN = "" Then
    FileRemoveInvalidASCII = ""
    Exit Function
End If

For cnt = 1 To Len(FileName)
    If IsInvalidASCII(Mid$(sFN, cnt, 1)) Then
        Mid$(sFN, cnt, 1) = Substitute
    End If
Next cnt
FileRemoveInvalidASCII = sFN
Exit Function

errh:
Resume e
e:
On Error Resume Next
FileRemoveInvalidASCII = ""
Err.Clear
Exit Function

End Function


Public Function WriteBlankLines(ByVal strFile As String, Optional ByVal lBlankLines As Long = 1) As Long
Dim fso As FileSystemObject
Dim fsoTxtStr As Scripting.TextStream
On Error GoTo errh

Set fso = New Scripting.FileSystemObject

Set fsoTxtStr = fso.OpenTextFile(strFile, ForAppending, False, TristateUseDefault)
fsoTxtStr.WriteBlankLines (lBlankLines)

WriteBlankLines = lBlankLines

errh:
WriteBlankLines = 0
GoSub LocalClean
Exit Function

LocalClean:
On Error Resume Next
Set fsoTxtStr = Nothing
Set fso = Nothing
Return

End Function


Public Function GetFileInfo(ByVal FileTarget As String, ByRef scrOutFile As Scripting.File) As Boolean
Dim FileOut As Scripting.File
Dim fs As New Scripting.FileSystemObject
On Error GoTo errh
Set FileOut = fs.GetFile(FileTarget)
Set scrOutFile = FileOut
GoSub LC
GetFileInfo = True
Exit Function

errh:
Resume e
e:
On Error Resume Next
Set scrOutFile = Nothing
GetFileInfo = False
GoSub LC
Err.Clear
Exit Function

LC:
On Error Resume Next
D fs
Err.Clear
Return

End Function


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

Public Function FileExists(ByVal strFile$) As Boolean
Dim fsObj As Scripting.FileSystemObject
On Error GoTo errh
Set fsObj = New Scripting.FileSystemObject
FileExists = fsObj.FileExists(strFile)
GoSub LocalClean
Exit Function

errh:
Resume e
e:
GoSub LocalClean
FileExists = False
Exit Function

LocalClean:
On Error Resume Next
D fsObj
Return

End Function

Public Function DirExists(ByVal strDir$) As Boolean
Dim fsObj As Scripting.FileSystemObject
On Error GoTo errh
Set fsObj = New Scripting.FileSystemObject
DirExists = fsObj.FolderExists(strDir)
GoSub LocalClean
Exit Function

errh:
GoSub LocalClean
DirExists = False
Exit Function

LocalClean:
On Error Resume Next
D fsObj
Return

End Function


Public Function FolderExists(ByVal strFile$) As Boolean
Dim fsObj As Scripting.FileSystemObject
On Error GoTo errh
Set fsObj = New Scripting.FileSystemObject
FolderExists = fsObj.FolderExists(strFile)
GoSub LocalClean
Exit Function
errh:
GoSub LocalClean
FolderExists = False
Exit Function
LocalClean:
On Error Resume Next
Set fsObj = Nothing
Return
End Function

Public Function GetDriveFromPath(ByVal sPath As String) As String
Dim arr() As String
On Error GoTo errh
arr = Split(sPath, ":", 2)
GetDriveFromPath = arr(0) & ":\"
Exit Function
errh:
GetDriveFromPath = ""
Err.Clear
Exit Function
End Function

Public Function WriteText(ByVal strFile As String, ByVal strText As String, Optional ByVal bOverwrite As Boolean = False) As Boolean
Dim fso As Scripting.FileSystemObject
Dim FileStr As Scripting.TextStream
On Error GoTo errh
Set fso = New Scripting.FileSystemObject

If bOverwrite Then
    Set FileStr = fso.CreateTextFile(strFile, bOverwrite)
Else
    If Me.FileExists(strFile) Then
        Set FileStr = fso.OpenTextFile(strFile, ForAppending, False, TristateUseDefault)
    Else
        Set FileStr = fso.CreateTextFile(strFile, bOverwrite)
    End If
End If

FileStr.WriteLine strText

WriteText = True

Exit Function

LocalClean:
On Error Resume Next
If Not FileStr Is Nothing Then FileStr.Close: Set FileStr = Nothing
If Not fso Is Nothing Then Set fso = Nothing
Return

errh:
GoSub LocalClean
WriteText = False
Exit Function

End Function

Function CopyFile(ByVal sSource$, ByVal sTarget$, Optional ByVal bOverwrite As Boolean = True) As Boolean
Dim fso As FileSystemObject
Set fso = New FileSystemObject
On Error GoTo errh
fso.CopyFile sSource, sTarget, bOverwrite
CopyFile = True
Exit Function

errh:
On Error Resume Next
CopyFile = False
GoSub LocalClean
Exit Function

LocalClean:
On Error Resume Next
D fso
Err.Clear
Return
End Function


Function RenameFile(ByVal sSource$, ByVal sTarget$, Optional ByVal bOverwrite As Boolean = True) As Boolean
Dim fso As FileSystemObject
Set fso = New FileSystemObject
On Error GoTo errh
fso.MoveFile sSource, sTarget
RenameFile = True
Exit Function

errh:
On Error Resume Next
RenameFile = False
GoSub LocalClean
Exit Function

LocalClean:
On Error Resume Next
D fso
Err.Clear
Return

End Function

Public Function FileRemoveAttrib(ByVal sFile As String, Attrib As VBA.VbFileAttribute) As Boolean
Dim fso As New FileSystemObject
Dim fl As File
On Error GoTo errh
Set fl = fso.GetFile(sFile)
If fl Is Nothing Then GoTo errh
If fl.Attributes And Attrib Then
    fl.Attributes = fl.Attributes - Attrib
End If
FileRemoveAttrib = True
Exit Function
errh:
FileRemoveAttrib = False
End Function

Public Function FileGetAttr(ByRef sFile As String) As Long
Dim fso As New FileSystemObject
Dim fl As File
On Error GoTo errh
Set fl = fso.GetFile(sFile)
FileGetAttr = fl.Attributes
Exit Function
errh:
FileGetAttr = 0
End Function

Public Function DeleteFile(ByVal strFile$, Optional ByVal DeleteReadOnly As Boolean = True) As Boolean
Dim fso As New FileSystemObject
Dim b As Boolean
On Error GoTo errh
If fso.FileExists(strFile) Then
    fso.DeleteFile strFile, DeleteReadOnly
End If
DeleteFile = True
GoSub LocalClean
Exit Function

errh:
Debug.Print "Couldnt delete " & strFile
On Error Resume Next
DeleteFile = False
GoSub LocalClean
Exit Function

LocalClean:
On Error Resume Next
D fso
Return

End Function

Public Sub MoveFile(ByVal strSource$, ByVal strDest$)
Dim fso As New FileSystemObject
fso.MoveFile strSource, strDest
End Sub

Public Function CreateFolder(ByVal strFolder$) As Boolean
Dim fso As FileSystemObject, fsoFold As Folder
Set fso = New Scripting.FileSystemObject
On Error GoTo errh
If FolderExists(strFolder) Then
    CreateFolder = True
Else
    Set fsoFold = fso.CreateFolder(strFolder)
End If
CreateFolder = True
GoSub LocalClean
Exit Function

errh:
On Error Resume Next
GoSub LocalClean
Exit Function

LocalClean:
DestroyObj fso
Return

End Function

Private Sub DestroyObj(obj As Object)
On Error Resume Next
Set obj = Nothing
Err.Clear
End Sub


Private Function CropRightSlash(ByVal strArg$) As String
If Right$(strArg, 1) = "\" Then
    CropRightSlash = Left$(strArg, Len(strArg) - 1)
Else
    CropRightSlash = strArg
End If
End Function


Function QueryCDRomDrives(ByRef lOutErr As Long) As String()
Dim bytDrv As Byte, bytI As Byte
Dim arrDrive() As String
Dim strTemp$
On Error GoTo errh
GetRoot strTemp
ReDim arrDrive(0)
For bytDrv = 65 To 90
    If GetDriveType(Chr(bytDrv) & ":\") = CDRom Then
        bytI = UBound(arrDrive)
        If bytI = 0 And arrDrive(0) = "" Then
            arrDrive(0) = Chr(bytDrv) & strTemp
        Else
            ReDim Preserve arrDrive(bytI + 1)
            arrDrive(bytI + 1) = Chr(bytDrv) & strTemp
        End If
    End If
Next bytDrv

QueryCDRomDrives = arrDrive
Exit Function
errh:
lOutErr = Err.Number
Err.Clear
Exit Function

End Function


Function QueryFixedDrives(ByRef lOutErr As Long) As String()
Dim bytDrv As Byte, bytI As Byte
Dim arrDrive() As String
Dim strTemp$
On Error GoTo errh
GetRoot strTemp
ReDim arrDrive(0)
For bytDrv = 65 To 90
    If GetDriveType(Chr(bytDrv) & ":\") = Fixed Then
        bytI = UBound(arrDrive)
        If bytI = 0 And arrDrive(0) = "" Then
            arrDrive(0) = Chr(bytDrv) & strTemp
        Else
            ReDim Preserve arrDrive(bytI + 1)
            arrDrive(bytI + 1) = Chr(bytDrv) & strTemp
        End If
    End If
Next bytDrv
QueryFixedDrives = arrDrive
Exit Function
errh:
lOutErr = Err.Number
Err.Clear
Exit Function

End Function


Function QueryRemoteDrives(ByRef lOutErr As Long) As String()
Dim bytDrv As Byte, bytI As Byte
Dim arrDrive() As String
Dim strTemp$
On Error GoTo errh
ReDim arrDrive(0)
GetRoot strTemp
For bytDrv = 65 To 90
    If GetDriveType(Chr(bytDrv) & ":\") = Remote Then
        bytI = UBound(arrDrive)
        If bytI = 0 And arrDrive(0) = "" Then
            arrDrive(0) = Chr(bytDrv) & strTemp
        Else
            ReDim Preserve arrDrive(bytI + 1)
            arrDrive(bytI + 1) = Chr(bytDrv) & strTemp
        End If
    End If
Next bytDrv

QueryRemoteDrives = arrDrive

Exit Function
errh:
lOutErr = Err.Number
Err.Clear
Exit Function
End Function


Function QueryRemoveableDrives(ByRef lOutErr As Long) As String()
Dim bytDrv As Byte, bytI As Byte
Dim arrDrive() As String
Dim strTemp$
On Error GoTo errh
ReDim arrDrive(0)
GetRoot strTemp
For bytDrv = 65 To 90
    If GetDriveType(Chr(bytDrv) & ":\") = Removeable Then
        bytI = UBound(arrDrive)
        If bytI = 0 And arrDrive(0) = "" Then
            arrDrive(0) = Chr(bytDrv) & strTemp
        Else
            ReDim Preserve arrDrive(bytI + 1)
            arrDrive(bytI + 1) = Chr(bytDrv) & strTemp
        End If
    End If
Next bytDrv

QueryRemoveableDrives = arrDrive

Exit Function
errh:
lOutErr = Err.Number
Err.Clear
Exit Function

End Function

Public Function FileGetText(ByVal File As String) As String
Dim fso As New FileSystemObject
Dim TxtStr As TextStream
On Error GoTo errh
Set TxtStr = fso.OpenTextFile(File, ForReading, False, TristateUseDefault)
FileGetText = TxtStr.ReadAll
GoSub LC
Exit Function

LC:
 On Error Resume Next
 D TxtStr
 D fso
Return

errh:
 Resume e
e:
 On Error Resume Next
 GoSub LC
End Function

Function QueryUnknownDrives(ByRef lOutErr As Long) As String()
Dim bytDrv As Byte, bytI As Byte
Dim arrDrive() As String
Dim strTemp$
On Error GoTo errh
ReDim arrDrive(0)
GetRoot strTemp
For bytDrv = 65 To 90
    If GetDriveType(Chr(bytDrv) & ":\") = Unknown Then
        bytI = UBound(arrDrive)
        If bytI = 0 And arrDrive(0) = "" Then
            arrDrive(0) = Chr(bytDrv) & strTemp
        Else
            ReDim Preserve arrDrive(bytI + 1)
            arrDrive(bytI + 1) = Chr(bytDrv) & strTemp
        End If
    End If
Next bytDrv

QueryUnknownDrives = arrDrive

Exit Function
errh:
lOutErr = Err.Number
Exit Function

End Function


Function QueryRAMDiskDrives(ByRef lOutErr As Long) As String()
Dim bytDrv As Byte, bytI As Byte
Dim arrDrive() As String
Dim strTemp$
On Error GoTo errh
ReDim arrDrive(0)
GetRoot strTemp
For bytDrv = 65 To 90
    If GetDriveType(Chr(bytDrv) & ":\") = RamDisk Then
        bytI = UBound(arrDrive)
        If bytI = 0 And arrDrive(0) = "" Then
            arrDrive(0) = Chr(bytDrv) & strTemp
        Else
            ReDim Preserve arrDrive(bytI + 1)
            arrDrive(bytI + 1) = Chr(bytDrv) & strTemp
        End If
    End If
Next bytDrv

QueryRAMDiskDrives = arrDrive

Exit Function
errh:
lOutErr = Err.Number
Exit Function
End Function


Function GetDiskSpace(ByVal RootPathName As String, ByRef OutDiskInfo As udtDiskInformation) As Boolean
Dim lAnswer As Long
Dim lpSectorsPerCluster As Long
Dim lpBytesPerSector As Long
Dim lpNumberOfFreeClusters As Long
Dim lpTotalNumberOfClusters As Long
On Error GoTo errh

If Not Me.DirExists(RootPathName) Then
    GetDiskSpace = False
    Exit Function
End If

RootPathName = Me.GetDriveFromPath(RootPathName)

lAnswer = GetDiskFreeSpace(RootPathName, lpSectorsPerCluster, lpBytesPerSector, lpNumberOfFreeClusters, lpTotalNumberOfClusters)
With OutDiskInfo
    .SectorsPerCluster = lpSectorsPerCluster
    .BytesPerSector = lpBytesPerSector
    .NumberOfFreeClusters = lpNumberOfFreeClusters
    .TotalNumberOfClusters = lpTotalNumberOfClusters
    .BytesPerCluster = lpSectorsPerCluster * lpBytesPerSector
    .NumFreeBytes = .BytesPerCluster * lpNumberOfFreeClusters
    .NumFreeKiloBytes = .NumFreeBytes / 1024
    .FreeMegabytes = (.NumFreeBytes / 1024) / 1024
End With
GetDiskSpace = True
Exit Function

errh:
Resume e
e:
On Error Resume Next
GetDiskSpace = False
Err.Clear
Exit Function

End Function

Private Sub GetRoot(ByRef str As String)
If mbooLetterOnly Then
    str = ""
Else
    str = ":\"
End If
End Sub


Public Property Get ReturnLetterOnly() As Boolean
ReturnLetterOnly = mbooLetterOnly
End Property
Public Property Let ReturnLetterOnly(ByVal booRLO As Boolean)
mbooLetterOnly = booRLO
End Property

Private Sub Class_Initialize()
mbooLetterOnly = True
End Sub


