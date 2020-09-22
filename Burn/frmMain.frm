VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "ISO Burner"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4875
      Top             =   675
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "ISO images (*.iso)|*.iso"
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   315
      Left            =   1725
      TabIndex        =   8
      Top             =   1200
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdBurn 
      Caption         =   "Write image"
      Height          =   315
      Left            =   1725
      TabIndex        =   7
      Top             =   1575
      Width           =   1515
   End
   Begin VB.CheckBox chkTestMode 
      Caption         =   "Test Mode"
      Height          =   240
      Left            =   300
      TabIndex        =   6
      Top             =   1455
      Width           =   2190
   End
   Begin VB.CheckBox chkFixDisc 
      Caption         =   "close disc"
      Height          =   240
      Left            =   300
      TabIndex        =   5
      Top             =   1215
      Value           =   1  'Aktiviert
      Width           =   2190
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      Top             =   675
      Width           =   465
   End
   Begin VB.TextBox txtISO 
      Height          =   285
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   675
      Width           =   4590
   End
   Begin VB.ComboBox cboDrv 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   150
      Width           =   5865
   End
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Unten ausrichten
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   2070
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "00%"
            TextSave        =   "00%"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkEjectDisc 
      Caption         =   "eject disc"
      Height          =   240
      Left            =   300
      TabIndex        =   9
      Top             =   1680
      Value           =   1  'Aktiviert
      Width           =   2190
   End
   Begin VB.Label lblImage 
      AutoSize        =   -1  'True
      Caption         =   "Image:"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   690
      Width           =   510
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' CD burned, but no files there! What to do?
'
' This problem should only occur when using the SPTI.
' I couldn't find the problem at all, but most likely
' it should be the drive handle's fault.
' We have full access to the device on 2k/XP.
' Anyway it seems that exiting the program should
' be enough to make Windows release the drive
' and display the burned files.
' You can also try to reboot...

Private Declare Sub Sleep Lib "kernel32" ( _
    ByVal ms As Long _
)

Private Sub BurnISO(index As Integer, file As String)
    Const sectors = 31      ' Sectors per write
    Const MAXRETRY = 10     ' retries until abort

    '   drive handle
    Dim driveid     As String
    '   write parameters page
    Dim wpp         As New clsWPP
    '   file handle
    Dim FF          As Integer: FF = FreeFile
    '   read buffer
    Dim buffer()    As Byte
    '   write position
    Dim LBA         As Long
    '   loop counter
    Dim lopcnt      As Integer

    driveid = scsi.DriveHandle(index)

    ' read the write parameters page
    If Not wpp.GetData(driveid) Then
        MsgBox "Couldn't read the Write Parameters Page.", vbExclamation
        Exit Sub
    End If

    wpp.DataBlockType = DB_MODE1_ISO       ' Mode 1 = 2048 Bytes/Sector
    wpp.TrackMode = 4                      ' Data Uninterrupted
    wpp.SessionFormat = SF_CDDA_DATA       ' Session Format
    wpp.Multisession = Abs(Not CBool(chkFixDisc)) * 3 ' Multisession or close disc
    wpp.WriteType = WT_TAO                 ' Track-At-Once
    wpp.TestMode = chkTestMode             ' Test Mode

    wpp.Copy = False
    wpp.FixedPacket = False
    wpp.LinkSizeValid = False

    wpp.AudioPauseLength = 150             ' Pause between tracks - 150 Frames = 2 seconds
    wpp.ApplicationCode = 0
    wpp.LinkSize = 0

    ' set the modified write parameters page
    If Not wpp.SendData(driveid) Then
        MsgBox "Couldn't set the new Write Parameters page.", vbExclamation
        Exit Sub
    End If

    Open file For Binary As #FF

    ReDim buffer(2048& * sectors - 1) As Byte

    ' set the read and write speed
    '   readspeed:  0xFFFF = Maximum
    '   writespeed: 8x (8 * 177 KB/s)
    If CDSetCDSpeed(driveid, &HFFFF, 8 * 177&, False) Then
        Debug.Print "WARNING: Could not set the write speed to 8x!"
    End If

    Do

        If EOF(FF) Then Exit Do

        Get #FF, , buffer

        ' burn the image to disc
        Do While CDWrite10(driveid, LBA, sectors, VarPtr(buffer(0)), UBound(buffer) + 1) <> STATUS_GOOD

            If lopcnt = MAXRETRY Then
                ' unknown error
                ' 10 write errors occured in 10 MS
                ' abort the write
                CDSyncCache driveid
                CDCloseTrackSession driveid, CloseTrack, 1
                CDCloseTrackSession driveid, CloseSession, 1
                If chkEjectDisc Then CDUnload driveid, True
                MsgBox "Schreibfehler (Buffer Under-Run?)", vbExclamation
                Exit Sub
            End If

            ' Write error?
            Select Case scsi.LastSK
                Case KEY_MEDIUM_ERROR:
                    If scsi.LastASC = &HC Then
                        CDSyncCache driveid
                        CDCloseTrackSession driveid, CloseTrack, 1
                        CDCloseTrackSession driveid, CloseSession, 1
                        If chkEjectDisc Then CDUnload driveid, True
                        MsgBox "Schreibfehler (Buffer Under-Run?)", vbExclamation
                        Exit Sub
                    End If
            End Select

            lopcnt = lopcnt + 1
            Sleep 100

        Loop
        lopcnt = 0

        ' set next write position
        LBA = LBA + sectors

        On Error Resume Next
        prg.value = (LBA / (LOF(FF) \ 2048) * 100)
        sbar.Panels(3).Text = Format(prg.value, "00") & "%"
        On Error GoTo 0
        DoEvents

    Loop

    Close #FF

    ' flush the write cache
    If CDSyncCache(driveid) <> STATUS_GOOD Then
        Debug.Print "SYNCHRONIZE CACHE failed."
        Debug.Print "SK: " & Hex$(scsi.LastSK), "ASC: " & Hex$(scsi.LastASC)
    End If

    ' close the track
    If CDCloseTrackSession(driveid, CloseTrack, 1) <> STATUS_GOOD Then
        Debug.Print "Couldn't close the track."
        Debug.Print "SK: " & Hex$(scsi.LastSK), "ASC: " & Hex$(scsi.LastASC)
    End If

    ' close the session
    If CDCloseTrackSession(driveid, CloseSession, 1) <> STATUS_GOOD Then
        Debug.Print "CLOSE TRACK/SESSION failed."
        Debug.Print "SK: " & Hex$(scsi.LastSK), "ASC: " & Hex$(scsi.LastASC)
    End If

    If chkEjectDisc Then CDUnload driveid, True

    MsgBox "Finished!", vbInformation
End Sub

Private Sub CheckTestMode(index As Integer)
    Dim driveid     As String
    Dim page        As page_capabilities

    driveid = scsi.DriveHandle(index)

    If CDModeSense10(driveid, &H2A, VarPtr(page), Len(page)) <> STATUS_GOOD Then
        MsgBox "Konnte Mode Page 2Ah nicht lesen.", vbExclamation
        Exit Sub
    End If

    ' can the drive write to CD-Rs?
    cmdBurn.Enabled = IsBitSet(page.writecaps, 0)
    ' does the drive support Test Mode?
    chkTestMode.Enabled = IsBitSet(page.writecaps, 2)
End Sub

Private Sub AddDrive(index As Integer)
    Dim buffer      As inquiry
    Dim strName     As String
    Dim driveid     As String

    driveid = scsi.DriveHandle(index)

    ' execute INQUIRY
    If CDInquiry(driveid, buffer) <> STATUS_GOOD Then
        MsgBox "INQUIRY failed for device " & index, vbExclamation
        Exit Sub
    End If

    ' device char
    strName = scsi.DriveChar(driveid) & ": "

    ' device name
    strName = strName & StrConv(buffer.vendor, vbUnicode)
    strName = strName & StrConv(buffer.product, vbUnicode)
    strName = strName & StrConv(buffer.revision, vbUnicode)

    ' bus position
    strName = strName & " (" & scsi.HostAdapter(driveid) & ":"
    strName = strName & scsi.TargetID(driveid) & ":"
    strName = strName & scsi.LUN(driveid) & ")"

    ' remove all null chars
    strName = Replace(strName, Chr$(0), "")

    cboDrv.AddItem strName
    cboDrv.ItemData(cboDrv.ListCount - 1) = index
End Sub

Private Sub Init()
    Dim i   As Integer

    i = scsi.DriveCount
    sbar.Panels(2).Text = "Devices: " & i
    sbar.Panels(1).Text = "Interface: " & scsi.Interface

    For i = 1 To i
        AddDrive i
    Next
    cboDrv.ListIndex = 0
End Sub

Private Sub cboDrv_Click()
    CheckTestMode cboDrv.ItemData(cboDrv.ListIndex)
End Sub

Private Sub cmdBrowse_Click()
    On Error Resume Next
    dlg.ShowOpen
    If Err Then Exit Sub
    txtISO.Text = dlg.FileName
End Sub

Private Sub cmdBurn_Click()
    If txtISO.Text = "" Then Exit Sub
    BurnISO cboDrv.ItemData(cboDrv.ListIndex), txtISO.Text
End Sub

Private Sub Form_Load()
    Dim aspi    As New clsASPI
    Dim spti    As New clsSPTI

    Set scsi = spti
    If Not (scsi.Installed And scsi.Initialized) Then

        Set scsi = aspi
        If False = (scsi.Installed And scsi.Initialized) Then
            MsgBox "No working interface found!", vbExclamation
            Unload Me
        End If

    End If

    Init
End Sub
