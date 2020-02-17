VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Types"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Type udtTemplateDB
    SimpleName As String
    ServerPath As String
    ClientPath As String
    DoClientCopy As Boolean
End Type

Public Type udtRect
    x1 As Long
    x2 As Long
    y1 As Long
    y2 As Long
End Type


Public Type udtSQLDataConnection
    Security As eSQLSecurity
    ServerName As String
    UID As String
    PWD As String
    PWD_Encoded As String
    DatabaseName As String
    ConnTimeoutShort As Long
    ConnTimeoutMed As Long
    ConnTimeoutLong As Long
    IsValid As Boolean
    SQL_Engine_Ver As eSQLEngineVer
End Type

Public Enum udtErrorSeverity
    Critical
    Major
    Minor
    Recoverable
    Fatal
    Unrecoverable
End Enum

Public Type udtErrorObject
    VBDescription As String
    description As String
    Number As Long
    Location As String
    Severity As udtErrorSeverity
    Advice As String
End Type


Public Type udtAppEnvironment
    App_ExeName As String
    App_LocalLogFile As String
    App_StartUpArgs As Scripting.Dictionary
    App_DefaultDocument As String
    App_ShowOpenMenuOnStart As Boolean
    App_HelpFileFull As String
    
    DB_SQLSrv As udtSQLDataConnection
    DB_Jet() As udtTemplateDB
    
    Dir_System As String
    Dir_Temp As String
    Dir_AppPath As String
    Dir_Cursors As String
    
    PC_FixedDrives() As udtDisk
    PC_CDROMDrives() As udtDisk
    PC_RAMDrives() As udtDisk
    PC_RemoteDrives() As udtDisk
    
    PC_Name As String
    PC_Processor As String
    PC_OS As String
    
    User_LogonServer As String
    User_NetworkName As String
    
    Reg_BaseKey As String
    Reg_FormsKey As String
    Reg_MRUKey As String
End Type


Public Type udtFrmSQLConnection
    uSQLConnection As udtSQLDataConnection
    Cancelled As Boolean
End Type

Public Enum eTag
    UserFire
    CodeFire
End Enum


Public Enum eResourceType
    Cursor = 1
    Bitmap = 2 ' Bitmap resource
    Icon = 3 ' Icon resource
    Menu = 4 ' Menu resource
    DialogBox = 5 ' Dialog box
    String_ = 6 ' String resource
    FontDirectoryRes = 7 ' Font directory resource
    FontRes = 8 '     Font resource
    AcceleratorTable = 9 ' Accelerator Table
    UserDefined = 10 ' User-defined resource
    GroupCursor = 12 ' Group Cursor
    GroupIcon = 14 ' Group Icon
End Enum


Public Enum eSQLSecurity
    SQLServer = 0
    Mixed = 1
    NTOnly = 4
End Enum


Public Enum eFormMode
    Add
    Edit
End Enum

Public Enum ePermissions
    Read = 1
    Write_ = 2
    Delete = 4
    Add = 8
    
    mixFull = 15
    mixEdit = 3
    mixAdd = 9
    
    NoAccess = 0
End Enum

Public Enum eMouse
    ButtonLeft = 1
    ButtonRight = 2
    ButtonMiddle = 4
End Enum

