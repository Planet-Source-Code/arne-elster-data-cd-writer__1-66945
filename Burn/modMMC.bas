Attribute VB_Name = "modMMC"
Option Explicit

Public scsi                         As ISCSI
Private lngPower2(31)               As Long

Public Type page_header
    datalen(1)                      As Byte
    rsvd(3)                         As Byte
    blockdescriptor(1)              As Byte
End Type

Public Type page_capabilities
    hdr                             As page_header
    pagecode                        As Byte
    pagelen                         As Byte
    readcaps                        As Byte
    writecaps                       As Byte
    audiomsess                      As Byte
    audiosubs                       As Byte
    mechanism                       As Byte
    mechpw                          As Byte
    maxreadspeed(1)                 As Byte
    vollevels(1)                    As Byte
    buffersize(1)                   As Byte
    currreadspeed(1)                As Byte
    rsvd3                           As Byte
    BCKF                            As Byte
    maxwritespeed(1)                As Byte
    currwritespeedobs(1)            As Byte ' MMC-1/2
    cpymngmt(1)                     As Byte
    rsvd6                           As Byte
    rsvd7                           As Byte
    rotctrl                         As Byte
    currwritespeed(1)               As Byte ' MMC-3
End Type

Public Type discinformation
    infolen(1)                      As Byte
    discstat                        As Byte
    firsttrack                      As Byte
    sessionsLSB                     As Byte
    firsttracklastsessionLSB        As Byte
    lasttracklastsessionLSB         As Byte
    uru                             As Byte
    DiscType                        As Byte
    sessionsMSB                     As Byte
    firsttracklastsessionMSB        As Byte
    lasttracklastsessionMSB         As Byte
    discID(3)                       As Byte
    lastsessionleadinstartMSF(3)    As Byte
    lastpossibleleadoutstartMSF(3)  As Byte
    discbarcode(7)                  As Byte
    reserved                        As Byte
    OPCEntries                      As Byte
End Type

Public Type toc_track
    rsvd1                           As Byte
    ADR                             As Byte
    track                           As Byte
    rsvd2                           As Byte
    addr(3)                         As Byte
End Type

Public Type formated_toc
    TocLen(1)                       As Byte
    firsttrack                      As Byte
    lasttrack                       As Byte
    TocTrack(99)                    As toc_track
End Type

Public Type inquiry
    qualifier                       As Byte
    rsvd1                           As Byte
    version                         As Byte
    respfmt                         As Byte
    addlen                          As Byte
    rsvd2                           As Byte
    stuff(1)                        As Byte
    vendor(7)                       As Byte
    product(15)                     As Byte
    revision(3)                     As Byte
    rsvd3(1)                        As Byte
    stuff2(37)                      As Byte
End Type

Public Enum CloseCodes
    CloseTrack = &H1
    CloseSession = &H2
End Enum

Public Enum ReadCDFlags
    SYNC = &H80
    HEADER_CODES = &H60
    USER_DATA = &H10
    EDCECC = &H8
    ERROR_FIELD = &H6
    RAW = &HF8
End Enum

Public Function CDSetCDSpeed(driveid As String, NewReadSpeed As Integer, NewWriteSpeed As Integer, RotCAV As Boolean) As Status
    Dim cdb(11)  As Byte

    cdb(0) = OP_SET_CD_SPEED        ' SET CD SPEED OpCode
    cdb(1) = Abs(RotCAV)            ' CLV/CAV Rotation
    cdb(2) = HiByte(NewReadSpeed)   ' Read Speed MSB
    cdb(3) = LoByte(NewReadSpeed)   ' Read Speed LSB
    cdb(4) = HiByte(NewWriteSpeed)  ' Write Speed MSB
    cdb(5) = LoByte(NewWriteSpeed)  ' Write Speed LSB

    CDSetCDSpeed = scsi.ExecCMD(driveid, cdb, 12, DIR_IN, 0, 0)
End Function

Public Function CDCloseTrackSession(driveid As String, CloseCode As CloseCodes, tracksess As Integer) As Status
    Dim cdb(9)  As Byte

    cdb(0) = OP_CLOSE_TRACK_SESSION ' CLOSE TRACK/SESSION OpCode
    cdb(2) = CloseCode              ' Function
    cdb(4) = HiByte(tracksess)      ' Track/Session MSB
    cdb(5) = LoByte(tracksess)      ' Track/Session LSB

    ' Timeout: 10 Minuten
    CDCloseTrackSession = scsi.ExecCMD(driveid, cdb, 10, DIR_IN, 0, 0, 10 * 60)
End Function

Public Function CDLoad(driveid As String, Optional immed As Boolean) As Status
    Dim cdb(5)  As Byte

    cdb(0) = OP_START_STOP_UNIT     ' START/STOP Unit OpCode
    cdb(1) = Abs(immed)             ' asynchronous processing
    cdb(4) = &H3                    ' Unload Bit/Start Bit

    CDLoad = scsi.ExecCMD(driveid, cdb, 6, DIR_IN, 0, 0)
End Function

Public Function CDUnload(driveid As String, Optional immed As Boolean) As Status
    Dim cdb(5)  As Byte

    cdb(0) = OP_START_STOP_UNIT     ' START/STOP UNIT OpCode
    cdb(1) = Abs(immed)             ' asynchronous processing
    cdb(4) = &H2                    ' Unload Bit

    CDUnload = scsi.ExecCMD(driveid, cdb, 6, DIR_IN, 0, 0)
End Function

Public Function CDSyncCache(driveid As String) As Status
    Dim cdb(9)  As Byte

    cdb(0) = OP_SYNCHRONIZE_CACHE   ' SYNC CACHE OpCode

    CDSyncCache = scsi.ExecCMD(driveid, cdb, 10, DIR_IN, 0, 0)
End Function

Public Function CDWrite10(driveid As String, LBA As Long, sectors As Integer, buffer As Long, bufferlen As Long) As Status
    Dim cdb(9)  As Byte

    cdb(0) = OP_WRITE10             ' WRITE10 OpCode
    cdb(2) = SHR(LBA, 24) And &HFF  ' LBA MSB
    cdb(3) = SHR(LBA, 16) And &HFF
    cdb(4) = SHR(LBA, 8) And &HFF
    cdb(5) = LBA And &HFF           ' LBA LSB
    cdb(7) = HiByte(sectors)        ' Sectors MSB
    cdb(8) = LoByte(sectors)        ' Sectors LSB

    ' Timeout: 10 Minuten
    CDWrite10 = scsi.ExecCMD(driveid, cdb, 10, DIR_OUT, buffer, bufferlen, 10 * 60)
End Function

Public Function CDRead10(driveid As String, LBA As Long, sectors As Integer, buffer As Long, bufferlen As Long) As Status
    Dim cdb(9)  As Byte

    cdb(0) = OP_READ10              ' READ10 OpCode
    cdb(2) = SHR(LBA, 24) And &HFF  ' LBA MSB
    cdb(3) = SHR(LBA, 16) And &HFF
    cdb(4) = SHR(LBA, 8) And &HFF
    cdb(5) = LBA And &HFF           ' MSB LSB
    cdb(7) = HiByte(sectors)        ' sectors MSB
    cdb(8) = LoByte(sectors)        ' sectors LSB

    CDRead10 = scsi.ExecCMD(driveid, cdb, 10, DIR_IN, buffer, bufferlen)
End Function

Public Function CDReadCD(driveid As String, LBA As Long, sectors As Long, buffer As Long, bufferlen As Long, flags As ReadCDFlags) As Status
    Dim cdb(9)  As Byte

    cdb(0) = OP_READ_CD                 ' READ CD OpCode
    cdb(2) = SHR(LBA, 24) And &HFF      ' LBA
    cdb(3) = SHR(LBA, 16) And &HFF
    cdb(4) = SHR(LBA, 8) And &HFF
    cdb(5) = LBA And &HFF
    cdb(6) = SHR(sectors, 16) And &HFF  ' sectors
    cdb(7) = SHR(sectors, 8) And &HFF
    cdb(8) = sectors And &HFF
    cdb(9) = flags                      ' Read flags

    CDReadCD = scsi.ExecCMD(driveid, cdb, 10, DIR_IN, buffer, bufferlen)
End Function

Public Function CDModeSelect10(driveid As String, ByVal buffer As Long, ByVal bufferlen As Integer) As Status
    Dim cdb(9)  As Byte

    cdb(0) = OP_MODE_SELECT10       ' MODE SELECT10 OpCode
    cdb(1) = &H10                   ' PF Bit
    cdb(7) = HiByte(bufferlen)      ' data allocation MSB
    cdb(8) = LoByte(bufferlen)      ' data allocation LSB

    CDModeSelect10 = scsi.ExecCMD(driveid, cdb, 10, DIR_OUT, buffer, bufferlen)
End Function

Public Function CDModeSense10(driveid As String, pc As Byte, buffer As Long, bufferlen As Integer) As Status
    Dim cdb(9)  As Byte

    cdb(0) = OP_MODE_SENSE10        ' MODE SENSE10 OpCode
    cdb(1) = &H8                    ' DBD (Disable Block Descriptors)
    cdb(2) = pc Or &H80             ' &H80 = default values
    cdb(7) = HiByte(bufferlen)      ' data allocation MSB
    cdb(8) = LoByte(bufferlen)      ' data allocation LSB

    CDModeSense10 = scsi.ExecCMD(driveid, cdb, 10, DIR_IN, buffer, bufferlen)
End Function

Public Function CDReadDiscInfo(driveid As String, buffer As discinformation) As Status
    Dim cdb(9)  As Byte

    cdb(0) = OP_READ_DISC_INFORMATION   ' READ DISC INFORMATION OpCode
    cdb(7) = HiByte(Len(buffer))        ' data allocation MSB
    cdb(8) = LoByte(Len(buffer))        ' data allocation LSB

    CDReadDiscInfo = scsi.ExecCMD(driveid, cdb, 10, DIR_IN, VarPtr(buffer), Len(buffer))
End Function

Public Function CDTestUnitReady(driveid As String) As Status
    Dim cdb(5)  As Byte             ' Command Descriptor Block

    cdb(0) = OP_TEST_UNIT_READY

    CDTestUnitReady = scsi.ExecCMD(driveid, cdb, 6, DIR_IN, 0, 0)
End Function

Public Function CDInquiry(driveid As String, buffer As inquiry) As Status
    Dim cdb(5)  As Byte     ' Command Descriptor Block

    cdb(0) = OP_INQUIRY     ' Inquiry OpCode
    cdb(4) = Len(buffer)    ' data allocation

    CDInquiry = scsi.ExecCMD(driveid, cdb, 6, DIR_IN, VarPtr(buffer), Len(buffer))
End Function

Public Function CDReadTOC0(driveid As String, _
                           msf As Boolean, _
                           toc As formated_toc) As Status

    Dim cdb(9)  As Byte

    cdb(0) = OP_READTOC         ' READ TOC OpCode
    cdb(1) = Abs(msf) * 2       ' time format (MSF or LBA)
    cdb(7) = HiByte(Len(toc))   ' data allocation MSB
    cdb(8) = LoByte(Len(toc))   ' data allocation LSB

    CDReadTOC0 = scsi.ExecCMD(driveid, cdb, 10, DIR_IN, VarPtr(toc), Len(toc))
End Function

Public Function IsBitSet(ByVal value As Long, bit As Byte) As Boolean
    IsBitSet = CBool(value And 2 ^ bit)
End Function

Public Function MKWord(ByVal Bh As Byte, Bl As Byte) As Integer
    MKWord = SHL(Bh, 8) Or Bl
End Function

Public Function MKDWord(ByVal Wh As Integer, ByVal Wl As Integer) As Long
    MKDWord = SHL(Wh, 16) Or Wl
End Function

Public Function LoWord(ByVal DWord As Long) As Long
  LoWord = DWord And &HFFFF&
End Function

Public Function HiWord(ByVal DWord As Long) As Long
  HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function LoByte(ByVal Word As Integer) As Byte
  LoByte = Word And &HFF
End Function

Public Function HiByte(ByVal Word As Integer) As Byte
  HiByte = (Word And &HFF00&) \ &H100
End Function

Public Function LoNibble(ByVal Bt As Byte) As Byte
    LoNibble = Bt And &HF
End Function

Public Function HiNibble(ByVal Bt As Byte) As Byte
    HiNibble = (Bt And &HF0) \ &H10
End Function

' >> Operator
' from VB-Accelerator
Public Function SHR(ByVal lThis As Long, ByVal lBits As Long) As Long

    Static Init As Boolean

    If Not Init Then InitShifting: Init = True

    If (lBits <= 0) Then
        SHR = lThis
    ElseIf (lBits > 63) Then
        Exit Function
    ElseIf (lBits > 31) Then
        SHR = 0
    Else
        If (lThis And lngPower2(31)) = lngPower2(31) Then
            SHR = (lThis And &H7FFFFFFF) \ lngPower2(lBits) Or lngPower2(31 - lBits)
        Else
            SHR = lThis \ lngPower2(lBits)
        End If
    End If

End Function

' << Operator
' from VB-Accelerator
Public Function SHL(ByVal lThis As Long, ByVal lBits As Long) As Long

    Static Init As Boolean

    If Not Init Then InitShifting: Init = True

    If (lBits <= 0) Then
        SHL = lThis
    ElseIf (lBits > 63) Then
        Exit Function
    ElseIf (lBits > 31) Then
        SHL = 0
    Else
        If (lThis And lngPower2(31 - lBits)) = lngPower2(31 - lBits) Then
            SHL = (lThis And (lngPower2(31 - lBits) - 1)) * lngPower2(lBits) Or lngPower2(31)
        Else
            SHL = (lThis And (lngPower2(31 - lBits) - 1)) * lngPower2(lBits)
        End If
    End If

End Function

' powers of 2
Private Sub InitShifting()
    Dim i   As Long
    For i = 0 To 30: lngPower2(i) = 2& ^ i: Next
    lngPower2(31) = &H80000000
End Sub

Public Function Dec2Bin(ByVal number As Long) As String
    Dim x As Integer

    If number >= 2 ^ 32 Then Exit Function

    Do
        If (number And 2 ^ x) Then
            Dec2Bin = "1" & Dec2Bin
        Else
            Dec2Bin = "0" & Dec2Bin
        End If
        x = x + 1
    Loop Until 2 ^ x > number
End Function
