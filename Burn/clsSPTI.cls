VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DeviceIoControl Lib "kernel32" ( _
    ByVal hDevice As Long, _
    ByVal dwIoControlCode As Long, _
    ByRef lpInBuffer As Any, _
    ByVal nInBufferSize As Long, _
    ByRef lpOutBuffer As Any, _
    ByVal nOutBufferSize As Long, _
    ByRef lpBytesReturned As Long, _
    lpOverlapped As Any _
) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long _
) As Long

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" ( _
    ByVal nDrive As String _
) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( _
    lpVersionInformation As OSVERSIONINFOEX _
) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    source As Any, _
    ByVal Length As Long _
)

Private Type SCSI_ADDRESS
    Length                          As Long
    PortNumber                      As Byte
    PathId                          As Byte
    TargetID                        As Byte
    LUN                             As Byte
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize             As Long
    dwMajorVersion                  As Long
    dwMinorVersion                  As Long
    dwBuildNumber                   As Long
    dwPlatformId                    As Long
    szCSDVersion                    As String * 128
End Type

Private Type SPTD
    Length                          As Integer
    ScsiStatus                      As Byte
    PathId                          As Byte
    TargetID                        As Byte
    LUN                             As Byte
    CdbLength                       As Byte
    SenseInfoLength                 As Byte
    DataIn                          As Byte
    DataTransferLength              As Long
    TimeOutValue                    As Long
    DataBuffer                      As Long
    SenseInfoOffset                 As Long
    cdb(15)                         As Byte
    Fill(2)                         As Byte
End Type

Private Type SCSI_PASS_THROUGH_DIRECT_WITH_BUFFER
    SPT                             As SPTD
    'Fill(3)                         As Byte
    SenseBuffer(35)                 As Byte
End Type

Private Enum SPTIDirection
    SCSI_IOCTL_DATA_OUT = 0
    SCSI_IOCTL_DATA_IN = 1
    SCSI_IOCTL_DATA_UNSPECIFIED = 2
End Enum

Private Const IOCTL_SCSI_BASE       As Long = &H4

Private Const METHOD_BUFFERED       As Long = &H0
Private Const METHOD_IN_DIRECT      As Long = &H1
Private Const METHOD_OUT_DIRECT     As Long = &H2
Private Const METHOD_NEITHER        As Long = &H3

Private Const FILE_ANY_ACCESS       As Long = &H0
Private Const FILE_READ_ACCESS      As Long = &H1
Private Const FILE_WRITE_ACCESS     As Long = &H2

Private Const INVALID_HANDLE_VALUE  As Long = -1
Private Const OPEN_EXISTING         As Long = &H3
Private Const GENERIC_READ          As Long = &H80000000
Private Const GENERIC_WRITE         As Long = &H40000000
Private Const FILE_SHARE_READ       As Long = &H1
Private Const FILE_SHARE_WRITE      As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

Private Const VER_PLATFORM_WIN32_NT As Long = &H2

Private IOCTL_SCSI_PASS_THROUGH_DIRECT As Long
Private IOCTL_SCSI_GET_ADDRESS         As Long

Private Const DTYPE_CDROM           As Long = 5

Private colDrives                   As Collection

Private blnW2K                      As Boolean

Private lngPower2(31)               As Long
Private lngHandles(25)              As Long

Private btLastSK                    As Byte
Private btLastASC                   As Byte
Private btLastASCQ                  As Byte

Implements ISCSI

Private Sub FindDrives()
    Dim i           As Integer
    Dim strDrive    As String

    For i = 1 To 26
        strDrive = Chr$(i + 64)
        If GetDriveType(strDrive & ":") = DTYPE_CDROM Then
            colDrives.Add strDrive
        End If
    Next
End Sub

Private Function IsW2K() As Boolean
    Dim sys As OSVERSIONINFOEX

    sys.dwOSVersionInfoSize = Len(sys)
    GetVersionEx sys

    If sys.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        IsW2K = sys.dwMajorVersion >= 5
    End If
End Function

Private Function GetDriveHandle(ByVal drv As String, ByRef fh As Long) As Boolean
    Static Init As Boolean
    Dim flags   As Long
    Dim i       As Integer

    ' Already opened all CD/DVD-ROM devices?
    If Not Init Then

        ' open all devices for performance

        ' on Windows 2000 and higher you need the
        ' GENERIC_WRITE flag as well
        flags = GENERIC_READ
        If IsW2K Then flags = flags Or GENERIC_WRITE

        For i = 1 To 26

            If GetDriveType(Chr$(i + 64) & ":") = DTYPE_CDROM Then
                    ' \\.\X: for devices
                    fh = CreateFile("\\.\" & Chr$(i + 64) & ":", flags, _
                                    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
                                    OPEN_EXISTING, 0, 0)

                    lngHandles(i - 1) = fh

            Else
                lngHandles(i - 1) = -1
            End If

        Next
        Init = True
    End If

    fh = lngHandles(Asc(UCase$(drv)) - 65)
    GetDriveHandle = fh <> -1
End Function

Private Function CTL_CODE(ByVal lDevType As Long, _
                          ByVal lFunction As Long, _
                          ByVal lMethod As Long, _
                          ByVal lAccess As Long) As Long

    CTL_CODE = LShift(lDevType, 16) Or _
               LShift(lAccess, 14) Or _
               LShift(lFunction, 2) Or _
               lMethod
End Function

' >> Operator
' from VB-Accelerator
Private Function RShift(ByVal lThis As Long, ByVal lBits As Long) As Long

    Static Init As Boolean

    If Not Init Then InitShifting: Init = True

    If (lBits <= 0) Then
        RShift = lThis
    ElseIf (lBits > 63) Then
        Exit Function
    ElseIf (lBits > 31) Then
        RShift = 0
    Else
        If (lThis And lngPower2(31)) = lngPower2(31) Then
            RShift = (lThis And &H7FFFFFFF) \ lngPower2(lBits) Or lngPower2(31 - lBits)
        Else
            RShift = lThis \ lngPower2(lBits)
        End If
    End If

End Function

' << Operator
' from VB-Accelerator
Private Function LShift(ByVal lThis As Long, ByVal lBits As Long) As Long

    Static Init As Boolean

    If Not Init Then InitShifting: Init = True

    If (lBits <= 0) Then
        LShift = lThis
    ElseIf (lBits > 63) Then
        Exit Function
    ElseIf (lBits > 31) Then
        LShift = 0
    Else
        If (lThis And lngPower2(31 - lBits)) = lngPower2(31 - lBits) Then
            LShift = (lThis And (lngPower2(31 - lBits) - 1)) * lngPower2(lBits) Or lngPower2(31)
        Else
            LShift = (lThis And (lngPower2(31 - lBits) - 1)) * lngPower2(lBits)
        End If
    End If

End Function

' powers of 2
Private Sub InitShifting()
    Dim i   As Long
    For i = 0 To 30: lngPower2(i) = 2& ^ i: Next
    lngPower2(31) = &H80000000
End Sub

Private Sub Class_Initialize()
    Set colDrives = New Collection

    FindDrives

    'SPTD Control Code
    IOCTL_SCSI_PASS_THROUGH_DIRECT = CTL_CODE(IOCTL_SCSI_BASE, _
                                              &H405, _
                                              METHOD_BUFFERED, _
                                              FILE_READ_ACCESS Or _
                                              FILE_WRITE_ACCESS)

    'SGA Control Code
    IOCTL_SCSI_GET_ADDRESS = CTL_CODE(IOCTL_SCSI_BASE, _
                                      &H406, _
                                      METHOD_BUFFERED, _
                                      FILE_ANY_ACCESS)
End Sub

Private Sub Class_Terminate()
    Dim i   As Long
    For i = 0 To 25: CloseHandle lngHandles(i): Next
End Sub

Private Property Get Iscsi_DriveChar(handle As String) As Variant
    Iscsi_DriveChar = handle
End Property

Private Property Get ISCSI_DriveCount() As Integer
    ISCSI_DriveCount = colDrives.Count
End Property

Private Property Get ISCSI_DriveHandle(index As Integer) As String
    ISCSI_DriveHandle = colDrives(index)
End Property

Private Function ISCSI_ExecCMD(ByVal drive As String, cdb() As Byte, CDBLen As Byte, direction As DataDirection, ByVal buffer As Long, ByVal bufferlen As Long, Optional timeout As Integer = 5) As Status
    Dim lngHandle   As Long
    Dim BytesRet    As Long
    Dim lngStatus   As Long
    Dim SPT         As SCSI_PASS_THROUGH_DIRECT_WITH_BUFFER

    If Not GetDriveHandle(drive, lngHandle) Then
        ISCSI_ExecCMD = STATUS_RESERV_CONF
        Exit Function
    End If

    ' Wait For Ever = 1 hour
    If timeout = 0 Then timeout = 3600000

    With SPT.SPT
        .Length = Len(SPT.SPT)
        .TimeOutValue = timeout

        .SenseInfoLength = UBound(SPT.SenseBuffer) - 4
        .SenseInfoOffset = Len(SPT.SPT) + 4

        If direction = DIR_IN Then
            .DataIn = SCSI_IOCTL_DATA_IN
        Else
            .DataIn = SCSI_IOCTL_DATA_OUT
        End If

        CopyMemory .cdb(0), cdb(0), CDBLen
        .CdbLength = CDBLen

        .DataBuffer = buffer
        .DataTransferLength = bufferlen
    End With

    lngStatus = DeviceIoControl(lngHandle, IOCTL_SCSI_PASS_THROUGH_DIRECT, _
                                SPT, Len(SPT), SPT, Len(SPT), _
                                BytesRet, ByVal 0&)

    If lngStatus <> 1 Then
        ' probably wrong most of the time
        ISCSI_ExecCMD = STATUS_TIMEOUT
    Else
        ISCSI_ExecCMD = SPT.SPT.ScsiStatus
    End If

    btLastSK = SPT.SenseBuffer(2)
    btLastASC = SPT.SenseBuffer(12)
    btLastASCQ = SPT.SenseBuffer(13)
End Function

Private Property Get ISCSI_HostAdapter(handle As String) As Byte
    Dim lngHandle   As Long
    Dim addr        As SCSI_ADDRESS
    Dim dwRead      As Long

    addr.Length = Len(addr)

    If GetDriveHandle(handle, lngHandle) Then
        If DeviceIoControl(lngHandle, IOCTL_SCSI_GET_ADDRESS, addr, Len(addr), addr, Len(addr), dwRead, ByVal 0&) = 1 Then
            ISCSI_HostAdapter = addr.PortNumber
        End If
    End If
End Property

Private Property Get ISCSI_Initialized() As Boolean
    ISCSI_Initialized = True
End Property

Private Property Get ISCSI_Installed() As Boolean
    Dim DrvID   As String
    Dim handle  As Long
    DrvID = ISCSI_DriveHandle(1)
    ISCSI_Installed = GetDriveHandle(DrvID, handle)
End Property

Private Property Get ISCSI_Interface() As String
    ISCSI_Interface = "SPTI"
End Property

Private Property Get ISCSI_LastASC() As Byte
    ISCSI_LastASC = btLastASC
End Property

Private Property Get ISCSI_LastASCQ() As Byte
    ISCSI_LastASCQ = btLastASCQ
End Property

Private Property Get ISCSI_LastSK() As Byte
    ISCSI_LastSK = btLastSK
End Property

Private Property Get ISCSI_LUN(handle As String) As Byte
    Dim lngHandle   As Long
    Dim addr        As SCSI_ADDRESS
    Dim dwRead      As Long

    addr.Length = Len(addr)

    If GetDriveHandle(handle, lngHandle) Then
        If DeviceIoControl(lngHandle, IOCTL_SCSI_GET_ADDRESS, addr, Len(addr), addr, Len(addr), dwRead, ByVal 0&) = 1 Then
            ISCSI_LUN = addr.LUN
        End If
    End If
End Property

Private Property Get ISCSI_TargetID(handle As String) As Byte
    Dim lngHandle   As Long
    Dim addr        As SCSI_ADDRESS
    Dim dwRead      As Long

    addr.Length = Len(addr)

    If GetDriveHandle(handle, lngHandle) Then
        If DeviceIoControl(lngHandle, IOCTL_SCSI_GET_ADDRESS, addr, Len(addr), addr, Len(addr), dwRead, ByVal 0&) = 1 Then
            ISCSI_TargetID = addr.TargetID
        End If
    End If
End Property
