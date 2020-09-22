VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsISOFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" ( _
    lpFileTime As FILETIME, _
    lpLocalFileTime As FILETIME _
) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
    lpFileTime As FILETIME, _
    lpSystemTime As SYSTEMTIME _
) As Long

Private Declare Function GetFileTimeAPI Lib "kernel32" _
Alias "GetFileTime" ( _
    ByVal hFile As Long, _
    ByRef lpCreationTime As Any, _
    ByRef lpLastAccessTime As Any, _
    ByRef lpLastWriteTime As Any _
) As Long

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type SYSTEMTIME
  wYear                 As Integer
  wMonth                As Integer
  wDayOfWeek            As Integer
  wDay                  As Integer
  wHour                 As Integer
  wMinute               As Integer
  wSecond               As Integer
  wMilliseconds         As Integer
End Type

Private clsFiles()      As clsISOFile
Private lngFileCnt      As Long

Private lngBlocks       As Long
Private lngBlocksJoliet As Long
Private lngLBA          As Long
Private lngLBAJoliet    As Long

Public Function FileExists( _
    ByVal name As String _
) As Boolean

    Dim i   As Long

    For i = 0 To lngFileCnt - 1
        If StrComp(clsFiles(i).name, name, vbTextCompare) = 0 Then
            FileExists = True
            Exit Function
        End If
    Next
End Function

Public Property Get Blocks( _
) As Long

    Blocks = lngBlocks
End Property

Public Property Let Blocks( _
    ByVal lngNewVal As Long _
)

    lngBlocks = lngNewVal
End Property

Public Property Get BlocksJoliet( _
) As Long

    BlocksJoliet = lngBlocksJoliet
End Property

Public Property Let BlocksJoliet( _
    ByVal lngNewVal As Long _
)


    lngBlocksJoliet = lngNewVal
End Property

Public Property Get LBAJoliet( _
) As Long

    LBAJoliet = lngLBAJoliet
End Property

Public Property Let LBAJoliet( _
    ByVal lngNewVal As Long _
)

    lngLBAJoliet = lngNewVal
End Property

Public Property Get LBA( _
) As Long

    LBA = lngLBA
End Property

Public Property Let LBA( _
    ByVal lngNewVal As Long _
)

    lngLBA = lngNewVal
End Property

Public Sub Clear( _
)

    lngFileCnt = 0
    Erase clsFiles
End Sub

Public Function Count( _
) As Long

    Count = lngFileCnt
End Function

Public Property Get File( _
    ByVal index As Long _
) As clsISOFile

    Set File = clsFiles(index)
End Property

Public Sub Remove( _
    ByVal index As Long _
)

    Dim i   As Long

    Set clsFiles(index) = Nothing

    For i = index + 1 To lngFileCnt - 1
        Set clsFiles(i - 1) = clsFiles(i)
    Next

    lngFileCnt = lngFileCnt - 1
End Sub

Public Function Add( _
    ByVal localfile As String, _
    Optional ByVal name As String _
) As clsISOFile

    If localfile = "" Then Exit Function
    If InStr(name, "\") > 0 Then Exit Function

    ReDim Preserve clsFiles(lngFileCnt) As clsISOFile

    If FileExists(IIf(name = "", GetFilename(localfile), name)) Then
        Exit Function
    End If

    Set clsFiles(lngFileCnt) = New clsISOFile
    With clsFiles(lngFileCnt)
        .LocalPath = localfile
        .name = IIf(name = "", GetFilename(localfile), name)
        .Size = FileLen(localfile)
        .CreationDate = GetFileTime(localfile)
    End With

    Set Add = clsFiles(lngFileCnt)

    lngFileCnt = lngFileCnt + 1
End Function

Private Function GetFilename( _
    ByVal strPath As String _
) As String

    GetFilename = Mid$(strPath, InStrRev(strPath, "\") + 1)
End Function

Private Function GetFileTime( _
    ByVal strFile As String _
) As Date

    Dim FTCreationTime      As FILETIME
    Dim FTLastAccessTime    As FILETIME
    Dim FTLastWriteTime     As FILETIME
    Dim SysTime             As SYSTEMTIME
    Dim hF                  As hFile

    hF = FileOpen(strFile)
    If hF.handle = INVALID_HANDLE Then Exit Function
    GetFileTimeAPI hF.handle, FTCreationTime, FTLastAccessTime, FTLastWriteTime
    FileClose hF

    FileTimeToLocalFileTime FTCreationTime, FTCreationTime
    FileTimeToSystemTime FTCreationTime, SysTime

    With SysTime
        GetFileTime = DateSerial(.wYear, .wMonth, .wDay) + _
                      TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function
