VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsASPI"
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

Private Declare Sub ZeroMemory Lib "kernel32.dll" _
Alias "RtlZeroMemory" ( _
    Destination As Any, _
    ByVal Length As Long _
)

Private Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal ByteLen As Long _
)

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMS As Long _
) As Long

Private Declare Function CreateEvent Lib "kernel32" _
Alias "CreateEventA" ( _
    ByVal lpEventAttributes As Long, _
    ByVal bManualReset As Long, _
    ByVal bInitialState As Long, _
    ByVal lpname As String _
) As Long

Private Declare Function ResetEvent Lib "kernel32" ( _
    ByVal hEvent As Long _
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

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Const WAIT_FAILED            As Long = &HFFFFFFFF
Private Const WAIT_OBJECT_0          As Long = &H0&
Private Const WAIT_ABANDONED         As Long = &H80&
Private Const WAIT_TIMEOUT           As Long = &H102&

Private Const INVALID_HANDLE_VALUE   As Long = -1
Private Const OPEN_EXISTING          As Long = &H3
Private Const GENERIC_READ           As Long = &H80000000
Private Const GENERIC_WRITE          As Long = &H40000000
Private Const FILE_SHARE_READ        As Long = &H1
Private Const FILE_SHARE_WRITE       As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL  As Long = &H80

Private Const VER_PLATFORM_WIN32_NT  As Long = &H2
Private Const IOCTL_SCSI_GET_ADDRESS As Long = &H41018

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Type t_SCSI_ADDRESS
    Length              As Long
    PortNumber          As Byte
    PathId              As Byte
    TargetID            As Byte
    LUN                 As Byte
End Type

Private Type t_HAID
    HA                  As Byte
    ID                  As Byte
    LUN                 As Byte
End Type

Private Type SRB
    SRB_Cmd             As Byte
    SRB_Status          As Byte
    SRB_HAID            As Byte
    SRB_Flags           As Byte
    SRB_Hdr_Rsvd        As Long
End Type

Private Type SRB_HAInquiry
    HA_Hdr              As SRB
    HA_Count            As Byte
    HA_Id               As Byte
    HA_MgrId            As String * 16
    HA_Ident            As String * 16
    HA_Unique(15)       As Byte
    HA_Rsvd             As Integer
    HA_Pad(19)          As Byte
End Type

Private Type SRB_GetDevType
    SRB_Hdr             As SRB
    SRB_Target          As Byte
    SRB_LUN             As Byte
    DEV_DeviceType      As Byte
    DEV_Rsvd1           As Byte
    DEV_Pad(67)         As Byte
End Type

Private Type SRB_GetDiskInfo
    SRB_Hdr             As SRB
    SRB_Target          As Byte
    SRB_LUN             As Byte
    SRB_DriveFlags      As Byte
    SRB_Int13DrvInfo    As Byte
    SRB_Heads           As Byte
    SRB_Sectors         As Byte
    SRB_Rsvd1(9)        As Byte
End Type

Private Type SRB_ExecuteIO
    SRB_Hdr             As SRB
    SRB_Target          As Byte
    SRB_LUN             As Byte
    SRB_Rsvd1           As Integer
    SRB_BufLen          As Long
    SRB_BufPointer      As Long
    SRB_SenseLen        As Byte
    SRB_CDBLen          As Byte
    SRB_HaStat          As Byte
    SRB_TargStat        As Byte
    SRB_PostProc        As Long
    SRB_Rsvd2(19)       As Byte
    SRB_CDBByte(15)     As Byte
    SRB_SenseData(15)   As Byte
End Type

Private Type SRB_Abort
    SRB_Hdr             As SRB
    SRB_SRBAbort        As Long
End Type

Private Type SRB_BusDecviceReset
    SRB_Hdr             As SRB
    SRB_Target          As Byte
    SRB_LUN             As Byte
    SRB_Rsvd1(11)       As Byte
    SRB_HaStat          As Byte
    SRB_TargStat        As Byte
    SRB_PostProc        As Long
    SRB_Rsvd2(35)       As Byte
End Type

Private Type SRB_GetSetTimeouts
    SRB_Hdr             As SRB
    SRB_Target          As Byte
    SRB_LUN             As Byte
    SRB_Timeout         As Long
End Type

Private Enum SRB_Command
    SC_HA_INQUIRY = &H0
    SC_GET_DEV_TYPE = &H1
    SC_EXEC_SCSI_CMD = &H2
    SC_ABORT_SRB = &H3
    SC_RESET_DEV = &H4
    SC_SET_HA_PARMS = &H5
    SC_GET_DISK_INFO = &H6
    SC_RESCAN_SCSI_BUS = &H7
    SC_GETSET_TIMEOUTS = &H8
End Enum

Private Enum HA_Status
    HASTAT_OK = &H0
    HASTAT_SEL_TO = &H11
    HASTAT_DO_DU = &H12
    HASTAT_BUS_FREE = &H13
    HASTAT_PHASE_ERR = &H14
    HASTAT_TIMEOUT = &H9
    HASTAT_COMMAND_TIMEOUT = &HB
    HASTAT_MESSAGE_REJECT = &HD
    HASTAT_BUS_RESET = &HE
    HASTAT_PARITY_ERROR = &HF
    HASTAT_REQUEST_SENSE_FAILED = &H10
End Enum

Private Enum SRB_Status
    SS_PENDING = &H0
    SS_COMP = &H1
    SS_ABORTED = &H2
    SS_ABORT_FAIL = &H3
    SS_ERR = &H4
    SS_INVALID_CMD = &H80
    SS_INVALID_HA = &H81
    SS_NO_DEVICE = &H82
    SS_INVALID_SRB = &HE0
    SS_BUFFER_ALIGN = &HE1
    SS_ILLEGAL_MODE = &HE2
    SS_NO_ASPI = &HE3
    SS_FAILED_INIT = &HE4
    SS_ASPI_IS_BUSY = &HE5
    SS_BUFFER_TO_BIG = &HE6
    SS_MISMATCHED_COMPONENTS = &HE7
    SS_NO_ADAPTERS = &HE8
    SS_INSUFFICIENT_RESOURCES = &HE9
    SS_ASPI_IS_SHUTDOWN = &HEA
    SS_BAD_INSTALL = &HBE
End Enum

Private Enum SRB_Flags
    SRB_POSTING = &H1
    SRB_ENABLE_RESIDUAL_COUNT = &H4
    SRB_DIR_IN = &H8
    SRB_DIR_OUT = &H10
    SRB_DIR_NOTIFY = &H40
End Enum

Private Enum DevTypes
    DTYPE_DASD = 0              ' direct access device
    DTYPE_SEQD = 1              ' sequential-access device
    DTYPE_PRNT = 2              ' Printer
    DTYPE_PROC = 3              ' Processor
    DTYPE_WORM = 4              ' Write-once device
    DTYPE_CDROM = 5             ' CD/DVD-ROM
    DTYPE_SCAN = 6              ' Scanner
    DTYPE_OPTI = 7              ' Optical Memory Device
    DTYPE_JUKE = 8              ' Changer
    DTYPE_COMM = 9              ' Communication device
    DTYPE_UNKNOWN = &H1F        ' unknown device
End Enum

Private Const SENSE_LEN As Long = 14&
Private Const MAX_SRB_TIMEOUT As Long = 108000
Private Const DEFAULT_SRB_TIMEOUT As Long = 108000

Private ASPILib         As clsCDECL
Private blnASPIInst     As Boolean

Private Const FNC_INFO  As String = "GetASPI32SupportInfo"
Private Const FNC_CMD   As String = "SendASPI32Command"


Private colDrives       As Collection

Private btLastSK        As Byte
Private btLastASC       As Byte
Private btLastASCQ      As Byte

Implements ISCSI

Private Function GetDriveHandle(ByVal drv As String, ByRef fh As Long) As Boolean
    Dim flags   As Long

    ' starting from win 2k you also need GENERIC_WRITE
    flags = GENERIC_READ
    If IsW2K Then flags = flags Or GENERIC_WRITE

    ' get the handle with CreateFile().
    ' you can access drives with "\\.\X:", where X is the drive's char.
    fh = CreateFile("\\.\" & Left$(drv, 1) & ":", flags, _
                    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, _
                    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)

    GetDriveHandle = fh <> -1

End Function

Private Function DriveCharWin9x(ByVal HA As Byte, ByVal ID As Byte, ByVal LUN As Byte) As String
    Dim diskinfo As SRB_GetDiskInfo

    diskinfo.SRB_Hdr.SRB_Cmd = SC_GET_DISK_INFO
    diskinfo.SRB_Hdr.SRB_Flags = SRB_DIR_IN
    diskinfo.SRB_Hdr.SRB_HAID = HA
    diskinfo.SRB_Target = ID
    diskinfo.SRB_LUN = LUN

    SRBDiskInfo diskinfo

    DriveCharWin9x = Chr$(diskinfo.SRB_Int13DrvInfo + 65)

End Function

Private Function DriveCharWinNT(ByVal HA As Byte, ByVal ID As Byte, ByVal LUN As Byte) As String

    Dim haid    As t_HAID
    Dim i       As Long

    'go through all 26 possible drives
    For i = 1 To 26

        'CD-ROM?
        If GetDriveType(Chr(i + 64) & ":") = 5 Then

            'compare to the parameters
            If NTGetHAID(Chr$(i + 64), haid) Then
                If haid.HA = HA And _
                   haid.ID = ID And _
                   haid.LUN = LUN Then
                    'found it!
                    DriveCharWinNT = Chr$(i + 64)
                End If
            End If

        End If

    Next

End Function

Private Function NTGetHAID(ByVal strDrv As String, ByRef haid As t_HAID) As Boolean

    Dim returned    As Long, Status As Long
    Dim fh          As Long, i      As Long

    Dim pscsiAddr   As t_SCSI_ADDRESS

    'get drive handle
    If GetDriveHandle(Left$(strDrv, 1), fh) Then

        'get SCSI address
        pscsiAddr.Length = Len(pscsiAddr)
        Status = DeviceIoControl(fh, IOCTL_SCSI_GET_ADDRESS, _
                                 pscsiAddr, Len(pscsiAddr), _
                                 pscsiAddr, Len(pscsiAddr), _
                                 returned, ByVal 0&)

        CloseHandle fh

        'success?
        If Status = 1 Then
            With pscsiAddr
                haid.HA = .PortNumber
                haid.ID = .TargetID
                haid.LUN = .LUN
                NTGetHAID = True
                Exit Function
            End With
        End If

    End If

End Function

Private Sub FindDrives()
    Dim HACnt   As Integer
    Dim IDCnt   As Integer
    Dim LUNCnt  As Integer
    Dim HAInq   As SRB_HAInquiry
    Dim DevTyp  As SRB_GetDevType

    ' Host Adapters
    HACnt = LoByte(LoWord(ASPILib.CallFunc(FNC_INFO)))

    For HACnt = 0 To HACnt

        For IDCnt = 0 To 7

            For LUNCnt = 0 To 7

                ZeroMemory DevTyp, LenB(DevTyp)
                DevTyp.SRB_Hdr.SRB_Cmd = SC_GET_DEV_TYPE
                DevTyp.SRB_Hdr.SRB_HAID = HACnt
                DevTyp.SRB_Target = IDCnt
                DevTyp.SRB_LUN = LUNCnt
                SRBGetDev DevTyp

                If DevTyp.SRB_Hdr.SRB_Status = SS_COMP Then

                    Debug.Print "ASPI: found device (" & DevTyp.DEV_DeviceType & ")", _
                                " HA: " & HACnt & " " & _
                                " ID: " & IDCnt & " " & _
                                " LUN: " & LUNCnt
 
                    If DevTyp.DEV_DeviceType = DTYPE_CDROM Then
                        colDrives.Add Chr$(HACnt) & Chr$(IDCnt) & Chr$(LUNCnt)
                    End If

                End If

            Next

        Next

    Next
End Sub

Private Property Get Iscsi_DriveChar(handle As String) As Variant
    Dim btHA As Byte
    Dim btID As Byte
    Dim btLUN As Byte

     btHA = Asc(Mid$(handle, 1, 1))
     btID = Asc(Mid$(handle, 2, 1))
    btLUN = Asc(Mid$(handle, 3, 1))

    If IsNT Then
        Iscsi_DriveChar = DriveCharWinNT(btHA, btID, btLUN)
    Else
        Iscsi_DriveChar = DriveCharWin9x(btHA, btID, btLUN)
    End If
End Property

Private Property Get ISCSI_DriveCount() As Integer
    ISCSI_DriveCount = colDrives.Count
End Property

Private Property Get ISCSI_DriveHandle(index As Integer) As String
    ISCSI_DriveHandle = colDrives.Item(index)
End Property

Private Function ISCSI_ExecCMD(ByVal drive As String, _
            cdb() As Byte, _
            CDBLen As Byte, _
            direction As DataDirection, _
            ByVal buffer As Long, _
            ByVal bufferlen As Long, _
            Optional timeout As Integer = 5) As Status

    Dim SRB_Exec    As SRB_ExecuteIO
    Dim SRB_Timeout As SRB_GetSetTimeouts
    Dim blnTimeout  As Boolean
    Dim hEvent      As Long
    Dim lngRet      As Long
    Dim btHA        As Byte
    Dim btID        As Byte
    Dim btLUN       As Byte

    ' Bus Address
    btHA = Asc(Mid$(drive, 1, 1))
    btID = Asc(Mid$(drive, 2, 1))
   btLUN = Asc(Mid$(drive, 3, 1))

    ' Event for scsi status
    hEvent = CreateEvent(0, 1, 0, 0)
    ResetEvent hEvent

    With SRB_Exec
        .SRB_Hdr.SRB_Cmd = SC_EXEC_SCSI_CMD
        .SRB_PostProc = hEvent

        .SRB_Hdr.SRB_HAID = btHA
        .SRB_Target = btID
        .SRB_LUN = btLUN

        .SRB_BufPointer = buffer
        .SRB_BufLen = bufferlen

        .SRB_SenseLen = SENSE_LEN

        If direction = DIR_IN Then
            .SRB_Hdr.SRB_Flags = SRB_DIR_IN
        Else
            .SRB_Hdr.SRB_Flags = SRB_DIR_OUT
        End If

        .SRB_CDBLen = CDBLen
        CopyMemory .SRB_CDBByte(0), cdb(0), CDBLen
    End With

    SRBExec SRB_Exec

    If Not blnTimeout Then
        ' wait for ever?
        If timeout = 0 Then
            Do While SRB_Exec.SRB_Hdr.SRB_Status = SS_PENDING
                DoEvents
            Loop
        Else
            ' wait for a specific amount of time for completion
            Sleep 30
            If SRB_Exec.SRB_Hdr.SRB_Status = SS_PENDING Then
                lngRet = WaitForSingleObject(hEvent, timeout * 1000)
            End If
        End If
    Else
        Do While SRB_Exec.SRB_Hdr.SRB_Status = SS_PENDING
            DoEvents
        Loop
    End If

    If blnTimeout Then
        If SRB_Exec.SRB_Hdr.SRB_Status = SS_ABORTED Then
            ISCSI_ExecCMD = STATUS_TIMEOUT
        Else
            ISCSI_ExecCMD = SRB_Exec.SRB_TargStat
        End If
    Else
        If lngRet = WAIT_TIMEOUT And SRB_Exec.SRB_Hdr.SRB_Status <> SS_COMP Then
            ISCSI_ExecCMD = STATUS_TIMEOUT
        Else
            ISCSI_ExecCMD = SRB_Exec.SRB_TargStat
        End If
    End If

    btLastSK = SRB_Exec.SRB_SenseData(2) And &HF
    btLastASC = SRB_Exec.SRB_SenseData(12)
    btLastASCQ = SRB_Exec.SRB_SenseData(13)

    CloseHandle hEvent
End Function

Private Function SRBInq(udt As SRB_HAInquiry) As Long
    SRBInq = ASPILib.CallFunc(FNC_CMD, VarPtr(udt))
End Function

Private Function SRBGetDev(udt As SRB_GetDevType) As Long
    SRBGetDev = ASPILib.CallFunc(FNC_CMD, VarPtr(udt))
End Function

Private Function SRBDiskInfo(udt As SRB_GetDiskInfo) As Long
    SRBDiskInfo = ASPILib.CallFunc(FNC_CMD, VarPtr(udt))
End Function

Private Function SRBExec(udt As SRB_ExecuteIO) As Long
    SRBExec = ASPILib.CallFunc(FNC_CMD, VarPtr(udt))
End Function

Private Function SRBSetTimeout(udt As SRB_GetSetTimeouts) As Long
    SRBSetTimeout = ASPILib.CallFunc(FNC_CMD, VarPtr(udt))
End Function

Private Property Get ISCSI_HostAdapter(handle As String) As Byte
    ISCSI_HostAdapter = Asc(Mid$(handle, 1, 1))
End Property

Private Property Get ISCSI_Initialized() As Boolean
    ISCSI_Initialized = HiByte(LoWord(ASPILib.CallFunc(FNC_INFO))) = SS_COMP
End Property

Private Property Get ISCSI_Installed() As Boolean
    ISCSI_Installed = blnASPIInst
End Property

Private Sub Class_Initialize()
    Set ASPILib = New clsCDECL
    Set colDrives = New Collection
    blnASPIInst = ASPILib.DllLoad("WNASPI32.DLL")
    FindDrives
End Sub

Private Sub Class_Terminate()
    ASPILib.DllUnload
End Sub

Private Function IsW2K() As Boolean
    Dim sys As OSVERSIONINFOEX

    sys.dwOSVersionInfoSize = Len(sys)
    GetVersionEx sys

    If sys.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        IsW2K = sys.dwMajorVersion >= 5
    End If
End Function

Private Function IsNT() As Boolean
    Dim sys As OSVERSIONINFOEX

    sys.dwOSVersionInfoSize = Len(sys)
    GetVersionEx sys

    IsNT = sys.dwPlatformId = VER_PLATFORM_WIN32_NT
End Function

Private Function LoWord(ByVal DWord As Long) As Long
  LoWord = DWord And &HFFFF&
End Function

Private Function HiWord(ByVal DWord As Long) As Long
  HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Private Function LoByte(ByRef Word As Integer) As Byte
  LoByte = Word And &HFF
End Function

Private Function HiByte(ByRef Word As Integer) As Byte
  HiByte = (Word And &HFF00&) \ &H100
End Function

Private Property Get ISCSI_Interface() As String
    ISCSI_Interface = "ASPI"
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
    ISCSI_LUN = Asc(Mid$(handle, 3, 1))
End Property

Private Property Get ISCSI_TargetID(handle As String) As Byte
    ISCSI_TargetID = Asc(Mid$(handle, 2, 1))
End Property
