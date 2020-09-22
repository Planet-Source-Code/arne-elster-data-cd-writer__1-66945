VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISCSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum SenseKeys
    KEY_NO_SENSE = &H0
    KEY_RECOVERABLE = &H1
    KEY_NOT_READY = &H2
    KEY_MEDIUM_ERROR = &H3
    KEY_HARDWARE_ERROR = &H4
    KEY_ILLEGAL_REQUEST = &H5
    KEY_UNIT_ATTENTION = &H6
    KEY_WRITE_PROTECT = &H7
    KEY_BLANK_CHECK = &H8
    KEY_VENDOR_UNIQUE = &H9
    KEY_COPY_ABORTED = &HA
    KEY_ABORTED_COMMAND = &HB
    KEY_EQUAL = &HC
    KEY_VOLUME_OVERFLOW = &HD
    KEY_MISCOMPARE_ERROR = &HE
    KEY_RESERVED = &HF
End Enum

Public Enum OpCodes
    OP_TEST_UNIT_READY = &H0
    OP_INQUIRY = &H12
    OP_READTOC = &H43
    OP_BLANK = &HA1
    OP_CLOSE_TRACK_SESSION = &H5B
    OP_ERASE10 = &H2C
    OP_FORMAT_UNIT = &H4
    OP_GET_CONFIGURATION = &H46
    OP_GET_EVENT_STATUS_NOTIFICATION = &H4A
    OP_GET_PERFORMANCE = &HAC
    OP_LOUNLOAD_MEDIUM = &HA6
    OP_MECHANISM_STATUS = &HBD
    OP_MODE_SELECT10 = &H55
    OP_MODE_SENSE10 = &H5A
    OP_PAUSE_RESUME = &H4B
    OP_PLAY_AUDIO10 = &H45
    OP_PLAY_AUDIO12 = &HA5
    OP_PLAY_AUDIOMSF = &H47
    OP_PLAY_CD = &HBC
    OP_PREVENT_MEDIUM_REMOVAL = &H1E
    OP_READ10 = &H28
    OP_READ12 = &HA8
    OP_READ_BUFFER = &H3C
    OP_READ_BUFFER_CAPACITY = &H5C
    OP_READ_CD = &HBE
    OP_READ_CD_MSF = &HB9
    OP_READ_CAPACITY = &H25
    OP_READ_DISC_INFORMATION = &H51
    OP_READ_DVD_STRUCTURE = &HAD
    OP_READ_FORMAT_CAPCITIES = &H23
    OP_READ_HEADER = &H44
    OP_READ_MASTER_CUE = &H59
    OP_READ_SUB_CHANNEL = &H42
    OP_READ_TRACK_INFORMATION = &H52
    OP_REPAIR_TRACK = &H58
    OP_REPORT_KEY = &HA4
    OP_REQUEST_SENSE = &H3
    OP_RESERVE_TRACK = &H53
    OP_SCAN = &HBA
    OP_SEEK10 = &H2B
    OP_SEND_CUE_SHEET = &H5D
    OP_SEND_DVD_STRUCTURE = &HAD
    OP_SEND_EVENT = &HA2
    OP_SEND_KEY = &HA3
    OP_SEND_OPC_INFORMATION = &H54
    OP_SET_CD_SPEED = &HBB
    OP_SET_READ_AHEAD = &HA7
    OP_SET_STREAMING = &HB6
    OP_START_STOP_UNIT = &H1B
    OP_STOP_PLAY_SCAN = &H4E
    OP_SYNCHRONIZE_CACHE = &H35
    OP_VERIFY10 = &H2F
    OP_WRITE10 = &H2A
    OP_WRITE_AND_VERIFY10 = &H2E
    OP_WRITE_BUFFER = &H3B
End Enum

Public Enum DataDirection
    DIR_IN
    DIR_OUT
End Enum

Public Enum Status
    STATUS_GOOD = &H0
    STATUS_CHKCOND = &H2
    STATUS_CONDMET_GOOD = &H4
    STATUS_BUSY = &H8
    STATUS_INTERMED_GOOD = &H10
    STATUS_INTERMED_CONDMET_GOOD = &H14
    STATUS_RESERV_CONF = &H18
    STATUS_CMD_TERMINATED = &H22
    STATUS_QUEUE_FULL = &H28
    STATUS_TIMEOUT = &HFF
End Enum

Public Property Get Interface() As String: End Property

Public Property Get LastASCQ() As Byte: End Property

Public Property Get LastASC() As Byte: End Property

Public Property Get LastSK() As Byte: End Property

Public Property Get Initialized() As Boolean: End Property

Public Property Get Installed() As Boolean: End Property

Public Property Get DriveCount() As Integer: End Property

Public Property Get DriveChar(handle As String): End Property

Public Property Get DriveHandle(index As Integer) As String: End Property

Public Property Get HostAdapter(handle As String) As Byte: End Property

Public Property Get TargetID(handle As String) As Byte: End Property

Public Property Get LUN(handle As String) As Byte: End Property

Public Function ExecCMD(ByVal drive As String, cdb() As Byte, CDBLen As Byte, direction As DataDirection, ByVal buffer As Long, ByVal bufferlen As Long, Optional timeout As Integer = 5) As Status: End Function
