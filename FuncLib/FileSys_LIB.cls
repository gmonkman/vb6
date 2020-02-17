VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSys_LIB"
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

Public Type udtDiskInformation
    SectorsPerCluster As Long
    BytesPerSector As Long
    NumberOfFreeClusters As Long
    TotalNumberOfClusters As Long
    BytesPerCluster As Long
    NumFreeBytes As Long
    NumFreeKiloBytes As Long
    FreeMegabytes As Double
End Type
   
Public Function QueryFileExists(strPath) As Boolean
On Error GoTo errh
strPath = CropRightSlash(strPath)
If Len(Dir$(strPath)) > 0 Then
    QueryFileExists = True
Else
    QueryFileExists = False
End If

Exit Function

errh:
QueryFileExists = False
Exit Function
End Function

Public Function QueryDirExists(strPath) As Boolean
On Error GoTo errh
If Len(Dir$(strPath, vbDirectory)) > 0 Then
    QueryDirExists = True
Else
    QueryDirExists = False
End If
Exit Function
errh:
QueryDirExists = False
Exit Function
End Function

Private Function CropRightSlash(ByVal strArg$) As String
If Right$(strArg, 1) = "\" Then
    CropRightSlash = Left$(strArg, Len(strArg) - 1)
Else
    CropRightSlash = strArg
End If
End Function


Function QueryCDRomDrives() As String()
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
Err.Raise 512 + vbObjectError, "FileSysLib::QueryCDRomDrives", "Unknown Error"
Exit Function
End Function


Function QueryFixedDrives() As String()
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
Err.Raise 512 + vbObjectError, "FileSys_LIB::QueryFixedDrives", "Unknown Error"
Exit Function

End Function


Function QueryRemoteDrives() As String()
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
Err.Raise 512 + vbObjectError, "FileSys_LIB::QueryRemoteDrives", "Unknown Error"
Exit Function

End Function


Function QueryRemoveableDrives() As String()
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
Err.Raise 512 + vbObjectError, "FileSys_LIB::QueryRemoveableDrives", "Unknown Error"
Exit Function

End Function


Function QueryUnknownDrives() As String()
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
Err.Raise 512 + vbObjectError, "FileSys_LIB::QueryUnknownDrives", "Unknown Error"
Exit Function

End Function


Function QueryRAMDiskDrives() As String()
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
Err.Raise 512 + vbObjectError, "FileSys_LIB::QueryRAMDiskDrives", "Unknown Error"
Exit Function
End Function


Function GetDiskSpace(ByVal RootPathName As String, ByRef OutDiskInfo As FileSys_LIB.udtDiskInformation) As Boolean
Dim lAnswer As Long
Dim lpSectorsPerCluster As Long
Dim lpBytesPerSector As Long
Dim lpNumberOfFreeClusters As Long
Dim lpTotalNumberOfClusters As Long
Dim lBytesPerCluster As Long
On Error GoTo errh

If Not Me.QueryDirExists(RootPathName) Then
    GetDiskSpace = False
    Exit Function
End If

lAnswer = GetDiskFreeSpace(RootPathName, lpSectorsPerCluster, lpBytesPerSector, lpNumberOfFreeClusters, lpTotalNumberOfClusters)
With OutDiskInfo
    .SectorsPerCluster = lpSectorsPerCluster
    .BytesPerSector = lpBytesPerSector
    .NumberOfFreeClusters = lpNumberOfFreeClusters
    .TotalNumberOfClusters = lpTotalNumberOfClusters
    .BytesPerCluster = lpSectorsPerCluster * lpBytesPerSector
    .NumFreeBytes = lBytesPerCluster * lpNumberOfFreeClusters
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

