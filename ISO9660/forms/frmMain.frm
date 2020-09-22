VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMain 
   Caption         =   "Create ISO9660 images"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picInfo 
      BorderStyle     =   0  'Kein
      Height          =   240
      Left            =   3600
      ScaleHeight     =   240
      ScaleWidth      =   4065
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   825
      Width           =   4065
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Add files and directories by dropping them!"
         Height          =   195
         Left            =   900
         TabIndex        =   30
         Top             =   0
         Width           =   3090
      End
   End
   Begin VB.PictureBox picVD 
      Height          =   4665
      Left            =   225
      ScaleHeight     =   4605
      ScaleWidth      =   7380
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   7440
      Begin VB.Frame Frame2 
         Caption         =   "Date"
         Height          =   840
         Left            =   525
         TabIndex        =   24
         Top             =   3525
         Width           =   5865
         Begin VB.CommandButton cmdSetDate 
            Caption         =   "now"
            Height          =   315
            Left            =   4875
            TabIndex        =   28
            Top             =   300
            Width           =   765
         End
         Begin MSComCtl2.DTPicker dateCreation 
            Height          =   315
            Left            =   1650
            TabIndex        =   26
            Top             =   300
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20316161
            CurrentDate     =   38925
         End
         Begin MSComCtl2.DTPicker timeCreation 
            Height          =   315
            Left            =   3300
            TabIndex        =   27
            Top             =   300
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            Format          =   20316162
            CurrentDate     =   38925
         End
         Begin VB.Label lblDescrID 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            Caption         =   "Creation:"
            Height          =   195
            Index           =   5
            Left            =   765
            TabIndex        =   25
            Top             =   345
            Width           =   675
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Descriptor:         "
         Height          =   3015
         Left            =   525
         TabIndex        =   10
         Top             =   375
         Width           =   5865
         Begin VB.TextBox txtVolSetID 
            Height          =   300
            Left            =   1635
            TabIndex        =   23
            Top             =   2475
            Width           =   3915
         End
         Begin VB.TextBox txtVolID 
            Height          =   300
            Left            =   1635
            TabIndex        =   13
            Top             =   600
            Width           =   3915
         End
         Begin VB.TextBox txtSysID 
            Height          =   300
            Left            =   1635
            TabIndex        =   15
            Top             =   975
            Width           =   3915
         End
         Begin VB.TextBox txtAppID 
            Height          =   300
            Left            =   1635
            TabIndex        =   17
            Top             =   1350
            Width           =   3915
         End
         Begin VB.TextBox txtPubID 
            Height          =   300
            Left            =   1635
            TabIndex        =   19
            Top             =   1725
            Width           =   3915
         End
         Begin VB.TextBox txtPrepID 
            Height          =   300
            Left            =   1635
            TabIndex        =   21
            Top             =   2100
            Width           =   3915
         End
         Begin VB.ComboBox cboDescr 
            Height          =   315
            ItemData        =   "frmMain.frx":0442
            Left            =   1050
            List            =   "frmMain.frx":044C
            Style           =   2  'Dropdown-Liste
            TabIndex        =   11
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lblDescrID 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            Caption         =   "Volume Set ID:"
            Height          =   195
            Index           =   6
            Left            =   375
            TabIndex        =   22
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label lblDescrID 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            Caption         =   "Volume ID:"
            Height          =   195
            Index           =   0
            Left            =   645
            TabIndex        =   12
            Top             =   645
            Width           =   780
         End
         Begin VB.Label lblDescrID 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            Caption         =   "System ID:"
            Height          =   195
            Index           =   1
            Left            =   630
            TabIndex        =   14
            Top             =   1020
            Width           =   795
         End
         Begin VB.Label lblDescrID 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            Caption         =   "Application ID:"
            Height          =   195
            Index           =   2
            Left            =   390
            TabIndex        =   16
            Top             =   1395
            Width           =   1050
         End
         Begin VB.Label lblDescrID 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            Caption         =   "Publisher ID:"
            Height          =   195
            Index           =   3
            Left            =   525
            TabIndex        =   18
            Top             =   1770
            Width           =   915
         End
         Begin VB.Label lblDescrID 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            Caption         =   "Data Preparer ID:"
            Height          =   195
            Index           =   4
            Left            =   150
            TabIndex        =   20
            Top             =   2145
            Width           =   1290
         End
      End
   End
   Begin prjISOImageWriter.Splitter spltMain 
      Height          =   3690
      Left            =   2100
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1425
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   6509
      SplitterPos     =   1430
      RatioFromTop    =   0,3
      Child1          =   "tvwDirs"
      Child2          =   "lvwFiles"
      Begin prjISOImageWriter.ucTreeView tvwDirs 
         Height          =   3690
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1430
         _ExtentX        =   2514
         _ExtentY        =   6509
      End
      Begin MSComctlLib.ListView lvwFiles 
         Height          =   3720
         Left            =   1490
         TabIndex        =   8
         Top             =   -15
         Width           =   3415
         _ExtentX        =   6033
         _ExtentY        =   6562
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         _Version        =   393217
         Icons           =   "imgs"
         SmallIcons      =   "imgs"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   5716
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   7425
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "ISO Images (*.iso)|*.iso"
      Flags           =   2
   End
   Begin MSComctlLib.ImageList imgs 
      Left            =   7275
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0461
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09FB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Oben ausrichten
      BackColor       =   &H005B9CBB&
      BorderStyle     =   0  'Kein
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   7890
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7890
      Begin prjISOImageWriter.isButton cmdMenu 
         Height          =   590
         Left            =   225
         TabIndex        =   1
         Top             =   75
         Width           =   680
         _ExtentX        =   1191
         _ExtentY        =   1032
         Icon            =   "frmMain.frx":0F95
         Style           =   4
         IconSize        =   32
         IconAlign       =   0
         iNonThemeStyle  =   4
         ShowFocus       =   -1  'True
         USeCustomColors =   -1  'True
         BackColor       =   6003899
         HighlightColor  =   2271457
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.Label lblHeader2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level 2 + Joliet"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   450
         Width           =   1095
      End
      Begin VB.Line lnDiv2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7875
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line lnDiv 
         BorderColor     =   &H00004080&
         X1              =   0
         X2              =   7875
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ISO9660 Image Writer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   1125
         TabIndex        =   2
         Top             =   75
         Width           =   3615
      End
   End
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5985
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3069
            MinWidth        =   3069
            Text            =   "Size: 0 KB"
            TextSave        =   "Size: 0 KB"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabstrip 
      Height          =   5040
      Left            =   75
      TabIndex        =   6
      Top             =   825
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   8890
      MultiRow        =   -1  'True
      TabFixedWidth   =   2820
      HotTracking     =   -1  'True
      TabMinWidth     =   1352
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filesystem"
            Key             =   "FS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Volume Descriptors"
            Key             =   "VD"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuISOCreate 
         Caption         =   "Create ISO Image..."
      End
      Begin VB.Menu mnuS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrjClear 
         Caption         =   "New project"
      End
      Begin VB.Menu mnuS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMenuDir 
      Caption         =   "Directory"
      Visible         =   0   'False
      Begin VB.Menu mnuDirNew 
         Caption         =   "New directory"
      End
      Begin VB.Menu mnuS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDirRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuDirRen 
         Caption         =   "Rename"
      End
   End
   Begin VB.Menu mnuMenuFiles 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu mnuFilesRem 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuFileRen 
         Caption         =   "Rename"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Drag'n'Drop needs its own format number to identify
' the source of the data
Private Const OLEDragDropFormatLVW  As Integer = 100

Private WithEvents clsISOWrt        As clsISOWriter
Attribute clsISOWrt.VB_VarHelpID = -1

Private Sub cboDescr_Click()
    Select Case cboDescr.ListIndex
        Case 0 ' ISO9660
            txtAppID.MaxLength = 128
            txtAppID.text = clsISOWrt.ApplicationID(False)
            txtPrepID.MaxLength = 128
            txtPrepID.text = clsISOWrt.DataPreparerID(False)
            txtPubID.MaxLength = 128
            txtPubID.text = clsISOWrt.PublisherID(False)
            txtSysID.MaxLength = 32
            txtSysID.text = clsISOWrt.SystemID(False)
            txtVolID.MaxLength = 32
            txtVolID.text = clsISOWrt.VolumeID(False)
            txtVolSetID.MaxLength = 128
            txtVolSetID.text = clsISOWrt.VolumeSetID(False)
        Case 1 ' Joliet
            txtAppID.MaxLength = 64
            txtAppID.text = clsISOWrt.ApplicationID(True)
            txtPrepID.MaxLength = 64
            txtPrepID.text = clsISOWrt.DataPreparerID(True)
            txtPubID.MaxLength = 64
            txtPubID.text = clsISOWrt.PublisherID(True)
            txtSysID.MaxLength = 16
            txtSysID.text = clsISOWrt.SystemID(True)
            txtVolID.MaxLength = 16
            txtVolID.text = clsISOWrt.VolumeID(True)
            txtVolSetID.MaxLength = 64
            txtVolSetID.text = clsISOWrt.VolumeSetID(True)
    End Select
End Sub

Private Sub clsISOWrt_BuildingFilesystem()
    sbar.Panels(1).text = "Building filesystem"
End Sub

Private Sub clsISOWrt_WritingDirectoryRecords()
    sbar.Panels(1).text = "Writing directory records"
End Sub

Private Sub clsISOWrt_WritingFiles( _
    ByVal percent As Long _
)

    sbar.Panels(1).text = "Writing files (" & percent & "%)"
End Sub

Private Sub clsISOWrt_WritingFinished()
    sbar.Panels(1).text = "Ready"
End Sub

Private Sub clsISOWrt_WritingPathTable()
    sbar.Panels(1).text = "Writing path table"
End Sub

Private Sub cmdMenu_Click()
    ' main menu
    PopupMenu mnuFile, _
              vbPopupMenuLeftButton, _
              cmdMenu.Left, _
              cmdMenu.Top + cmdMenu.Height + Screen.TwipsPerPixelY, _
              mnuISOCreate
End Sub

Private Sub cmdSetDate_Click()
    clsISOWrt.VolumeCreation = Now

    dateCreation.Year = Year(clsISOWrt.VolumeCreation)
    dateCreation.Month = Month(clsISOWrt.VolumeCreation)
    dateCreation.Day = Day(clsISOWrt.VolumeCreation)

    timeCreation.Hour = Hour(clsISOWrt.VolumeCreation)
    timeCreation.Minute = Minute(clsISOWrt.VolumeCreation)
    timeCreation.Second = Second(clsISOWrt.VolumeCreation)
End Sub

Private Sub dateCreation_Change()
    With dateCreation
        clsISOWrt.VolumeCreation = .Day & "." & .Month & "." & .Year & " " & _
                                   timeCreation.Hour & ":" & timeCreation.Minute & ":" & timeCreation.Second
    End With
End Sub

Private Sub Form_Load()
    Set clsISOWrt = New clsISOWriter

    With tvwDirs
        .Initialize
        .InitializeImageList
        .AddIcon imgs.ListImages(1).Picture.handle  ' folder icon

        .ItemHeight = 18
        .HasButtons = True
        .HasLines = True
        .HasRootLines = True
        .LabelEdit = True

        .Font.name = "Tahoma"

        .OLEDragMode = drgAutomatic
        .OLEDropMode = drpManual

        ' Root
        .AddNode Key:="\", _
                 text:="root", _
                 Image:=0, _
                 SelectedImage:=0

        .SelectedNode = .GetKeyNode("\")

        .OLEDragInsertStyle = disDropHilite
        .OLEDragAutoExpand = True
    End With

    cboDescr.ListIndex = 0

    cmdSetDate_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    picInfo.Left = Me.ScaleWidth - picInfo.Width

    tabstrip.Width = Me.ScaleWidth
    tabstrip.Height = Me.ScaleHeight - picHeader.Height - sbar.Height
    tabstrip.Top = picHeader.Height
    tabstrip.Left = 0

    spltMain.Top = tabstrip.ClientTop
    spltMain.Left = tabstrip.ClientLeft
    spltMain.Width = tabstrip.ClientWidth
    spltMain.Height = tabstrip.ClientHeight

    picVD.Top = tabstrip.ClientTop
    picVD.Left = tabstrip.ClientLeft
    picVD.Width = tabstrip.ClientWidth
    picVD.Height = tabstrip.ClientHeight
End Sub

Private Sub lvwFiles_AfterLabelEdit( _
    Cancel As Integer, _
    NewString As String _
)

    ' change the name of a file

    Dim clsDir  As clsISODirectory

    ' get the directory from the selected node
    Set clsDir = DirFromSelectedNode()
    If clsDir Is Nothing Then
        Cancel = 1
        Exit Sub
    End If

    ' empty filenames are ilegal
    If Trim$(NewString) = "" Then
        Cancel = 1
        Exit Sub
    End If

    clsDir.Files.File(lvwFiles.SelectedItem.index - 1).name = NewString
End Sub

Private Sub lvwFiles_MouseDown( _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single _
)

    ' options for files
    If Button = vbRightButton Then
        PopupMenu mnuMenuFiles, _
                  vbPopupMenuRightButton, _
                  x + spltMain.CurrSplitterPos + spltMain.SplitterSize + tabstrip.ClientLeft + Screen.TwipsPerPixelX, _
                  y + tabstrip.ClientTop + Screen.TwipsPerPixelY, _
                  mnuFilesRem
    End If
End Sub

Private Sub lvwFiles_OLEDragDrop( _
    Data As MSComctlLib.DataObject, _
    Effect As Long, _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single _
)

    Dim i           As Long
    Dim hNode       As Long
    Dim clsDir      As clsISODirectory
    Dim strFilter() As String
    Dim blnNewDirs  As Boolean

    ReDim strFilter(0) As String
    strFilter(0) = "*"

    ' files/directories were dropped from the Explorer
    If Data.GetFormat(vbCFFiles) Then
        Set clsDir = DirFromSelectedNode()
        If clsDir Is Nothing Then Exit Sub

        For i = 1 To Data.Files.Count
            If DirExists(Data.Files(i)) Then
                With clsDir.AddSubDirectory(GetFilename(Data.Files(i)))
                    ' add local directory to the image
                    .AddLocalDirectory Data.Files(i), strFilter
                End With

                ' Treeview needs to be refreshed
                blnNewDirs = True
            Else
                clsDir.Files.Add Data.Files(i)
            End If
        Next

        If blnNewDirs Then
            ' adding new nodes is faster with redrawing disabled
            tvwDirs.SetRedrawMode False
            ' rebuild the selected node
            ISOBuildTree tvwDirs.SelectedNode, clsDir
            tvwDirs.SetRedrawMode True
        End If

        ShowFilesForDir tvwDirs.SelectedNode
    End If

    sbar.Panels(2).text = "Size: " & FormatFileSize(clsISOWrt.ImageSize)
End Sub

Private Sub ISOBuildTree( _
    ByVal hNode As Long, _
    clsDir As clsISODirectory _
)

    ' build a tree from a directory in the image
    ' (recursive)

    Dim hSubNode    As Long
    Dim i           As Long

    ' first clear all subnodes of the main node
    Do
        hSubNode = tvwDirs.NodeChild(hNode)
        If hSubNode = 0 Then Exit Do
        tvwDirs.DeleteNode hSubNode
    Loop

    ' build subnodes
    For i = 0 To clsDir.SubDirectoryCount - 1
        With clsDir.SubDirectory(i)
            ISOBuildTree tvwDirs.AddNode(hNode, , .FullPath, .name, 0, 0), clsDir.SubDirectory(i)
        End With
    Next
End Sub

' return the directory for the selected node
Private Function DirFromSelectedNode( _
) As clsISODirectory

    With tvwDirs
        Set DirFromSelectedNode = clsISOWrt.DirByPath(.GetNodeKey(.SelectedNode))
    End With
End Function

Private Sub lvwFiles_OLEStartDrag( _
    Data As MSComctlLib.DataObject, _
    AllowedEffects As Long _
)

    Dim btData()    As Byte
    ReDim btData(0) As Byte

    ' only copy nodes, do not move them
    AllowedEffects = vbDropEffectCopy
    Data.SetData btData, OLEDragDropFormatLVW
End Sub

Private Sub mnuDirNew_Click()
    Dim strNewDir   As String
    Dim clsDir      As clsISODirectory
    Dim clsDirNew   As clsISODirectory

    strNewDir = InputBox("New directory's name:")
    If StrPtr(strNewDir) = 0 Then Exit Sub
    If Trim$(strNewDir) = "" Then Exit Sub

    Set clsDir = DirFromSelectedNode()
    If clsDir Is Nothing Then Exit Sub

    Set clsDirNew = clsDir.AddSubDirectory(strNewDir)
    If clsDirNew Is Nothing Then Exit Sub

    With clsDirNew
        tvwDirs.EnsureVisible tvwDirs.AddNode(tvwDirs.SelectedNode, , .FullPath, .name, 0, 0)
    End With
End Sub

Private Sub mnuDirRemove_Click()
    Dim clsDir      As clsISODirectory
    Dim i           As Long

    Set clsDir = DirFromSelectedNode()
    If clsDir Is Nothing Then Exit Sub

    If clsDir.FullPath = "\" Then Exit Sub
    If clsDir.Parent Is Nothing Then Exit Sub

    ' find the index of the directory to remove
    For i = 0 To clsDir.Parent.SubDirectoryCount - 1
        If clsDir Is clsDir.Parent.SubDirectory(i) Then
            clsDir.Parent.RemoveSubDirectory i
            Exit For
        End If
    Next

    ' remove the directory from the treeview
    tvwDirs.DeleteNode tvwDirs.SelectedNode
End Sub

Private Sub mnuDirRen_Click()
    tvwDirs.StartLabelEdit tvwDirs.SelectedNode
End Sub

Private Sub mnuFileRen_Click()
    lvwFiles.StartLabelEdit
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFilesRem_Click()
    Dim i       As Long
    Dim clsDir  As clsISODirectory

    Set clsDir = DirFromSelectedNode()
    If clsDir Is Nothing Then Exit Sub

    ' The index of an item in the Listview should
    ' be equal to (Index - 1) in a files collection.
    ' So don't sort the Listview!

    With lvwFiles.ListItems
        For i = .Count To 1 Step -1
            If .Item(i).Selected Then
                clsDir.Files.Remove i - 1
                .Remove i
            End If
        Next
    End With
End Sub

Private Sub mnuISOCreate_Click()
    With dlg
        .FileName = vbNullString
        .ShowSave
    End With

    If dlg.FileName = vbNullString Then Exit Sub

    If Not clsISOWrt.SaveISO(dlg.FileName) Then
        MsgBox "Failed!", vbExclamation
    End If
End Sub

Private Sub mnuPrjClear_Click()
    Set clsISOWrt = New clsISOWriter

    With tvwDirs
        .Clear

        .AddNode Key:="\", _
                 text:="root", _
                 Image:=0, _
                 SelectedImage:=0

        .SelectedNode = .GetKeyNode("\")
    End With

    lvwFiles.ListItems.Clear
End Sub

Private Sub picHeader_Resize()
    On Error Resume Next

    lnDiv.x1 = 0
    lnDiv.x2 = picHeader.ScaleWidth
    lnDiv2.x1 = lnDiv.x1
    lnDiv2.x2 = lnDiv.x2
End Sub

Private Sub tabstrip_Click()
    Select Case tabstrip.SelectedItem.Key
        Case "FS"
            spltMain.Visible = True
            picVD.Visible = False
        Case "VD"
            spltMain.Visible = False
            picVD.Visible = True
    End Select
End Sub

Private Sub timeCreation_Change()
    With dateCreation
        clsISOWrt.VolumeCreation = .Day & "." & .Month & "." & .Year & " " & _
                                   timeCreation.Hour & ":" & timeCreation.Minute & ":" & timeCreation.Second
    End With
End Sub

Private Sub tvwDirs_AfterLabelEdit( _
    ByVal hNode As Long, _
    Cancel As Integer, _
    NewString As String _
)

    Dim clsDir      As clsISODirectory
    Dim strOldName  As String

    Set clsDir = clsISOWrt.DirByPath(tvwDirs.GetNodeKey(hNode))
    If clsDir Is Nothing Then
        Cancel = 1
        Exit Sub
    End If

    If Trim$(NewString) = "" Then
        Cancel = 1
        Exit Sub
    End If

    clsDir.name = NewString

    ' The key of a node is the full path from the root to the node.
    ' When the node's text gets changed, the key doesn't change.
    ' Consequence is, you can't find the directory in the image
    ' no more. So just rebuild the tree.
    ISOBuildTree tvwDirs.NodeParent(hNode), clsDir.Parent
End Sub

Private Sub tvwDirs_BeforeLabelEdit( _
    ByVal hNode As Long, _
    Cancel As Integer _
)

    ' You can not rename the root
    If hNode = tvwDirs.GetKeyNode("\") Then Cancel = 1
End Sub

Private Sub tvwDirs_MouseDown( _
    Button As Integer, _
    Shift As Integer, _
    x As Long, _
    y As Long _
)

    Dim hNode   As Long

    If Button = vbRightButton Then
        ' make sure the node under the mouse cursor
        ' shows up selected
        hNode = tvwDirs.HitTest(x, y, False)
        If hNode Then tvwDirs.SelectedNode = hNode

        PopupMenu mnuMenuDir, _
                  vbPopupMenuRightButton, _
                  DefaultMenu:=mnuDirNew
    End If
End Sub

Private Function GetFilename( _
    ByVal path As String _
) As String

    GetFilename = Mid$(path, InStrRev(path, "\") + 1)
End Function

Private Function FileExists( _
    ByVal path As String _
) As Boolean

    On Error Resume Next
    FileExists = (GetAttr(path) And (vbDirectory Or vbVolume)) = 0
End Function

Private Function DirExists( _
    ByVal path As String _
) As Boolean

    On Error Resume Next
    DirExists = CBool(GetAttr(path) And vbDirectory)
End Function

Private Sub tvwDirs_MouseUp( _
    Button As Integer, _
    Shift As Integer, _
    x As Long, _
    y As Long _
)

    ' on left button click show files associated with the
    ' node under the mouse cursor
    If Button = 1 Then
        ShowFilesForDir tvwDirs.HitTest(x, y, False)
    End If
End Sub

Private Sub ShowFilesForDir( _
    ByVal hNode As Long _
)

    Dim clsDir  As clsISODirectory
    Dim i       As Long

    ' clsISODirectory by node
    Set clsDir = clsISOWrt.DirByPath(tvwDirs.GetNodeKey(hNode))
    If clsDir Is Nothing Then Exit Sub

    lvwFiles.ListItems.Clear

    ' show file's name and size
    For i = 0 To clsDir.Files.Count - 1
        With lvwFiles.ListItems.Add(text:=clsDir.Files.File(i).name, SmallIcon:=2)
            .SubItems(1) = FormatFileSize(clsDir.Files.File(i).Size)
        End With
    Next
End Sub

Private Sub tvwDirs_OLEDragDrop( _
    Data As DataObject, _
    Effect As Long, _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single _
)

    Dim i           As Long
    Dim hNode       As Long
    Dim hNodeSrc    As Long
    Dim clsDir      As clsISODirectory
    Dim clsDirDst   As clsISODirectory
    Dim strFilter() As String
    Dim blnNewDirs  As Boolean

    ReDim strFilter(0) As String
    strFilter(0) = "*"

    ' data dropped from the listview?
    If Data.GetFormat(OLEDragDropFormatLVW) Then
        ' get the target node
        tvwDirs.OLEGetDropInfo hNode, True
        If hNode = 0 Then Exit Sub

        ' the dropped files have to be in the currently selected directory
        Set clsDir = DirFromSelectedNode()
        If clsDir Is Nothing Then Exit Sub

        ' target directory to move the files to
        Set clsDirDst = clsISOWrt.DirByPath(tvwDirs.GetNodeKey(hNode))
        If clsDirDst Is Nothing Then Exit Sub

        ' target and source are the same, cancel
        If hNode = tvwDirs.SelectedNode Then Exit Sub

        ' move the files to the target directory
        For i = lvwFiles.ListItems.Count To 1 Step -1
            If lvwFiles.ListItems(i).Selected Then
                With clsDir.Files.File(i - 1)
                    clsDirDst.Files.Add .LocalPath, .name
                End With

                lvwFiles.ListItems.Remove i
                clsDir.Files.Remove i - 1
            End If
        Next

        ' refresh Listview
        ShowFilesForDir tvwDirs.SelectedNode

        Exit Sub
    End If

    ' files/directories dropped from the Explorer (or something like that)
    If Data.GetFormat(vbCFFiles) Then
        tvwDirs.OLEGetDropInfo hNode, True
        If hNode = 0 Then Exit Sub

        Set clsDir = clsISOWrt.DirByPath(tvwDirs.GetNodeKey(hNode))
        If clsDir Is Nothing Then Exit Sub

        For i = 1 To Data.Files.Count
            If DirExists(Data.Files(i)) Then
                ' add directories + subdirectories
                With clsDir.AddSubDirectory(GetFilename(Data.Files(i)))
                    .AddLocalDirectory Data.Files(i), strFilter
                End With

                blnNewDirs = True
            Else
                ' must be a file
                clsDir.Files.Add Data.Files(i)
            End If
        Next

        If blnNewDirs Then
            tvwDirs.SetRedrawMode False
            ISOBuildTree hNode, clsDir
            tvwDirs.SetRedrawMode True
        End If

        ShowFilesForDir tvwDirs.SelectedNode

        sbar.Panels(2).text = FormatFileSize(clsISOWrt.ImageSize)

        Exit Sub
    End If

    ' node moved
    If tvwDirs.OLEIsMyFormat(Data) Then
        ' Sourcenode
        tvwDirs.OLEGetDragInfo Data, 0, hNodeSrc
        ' Targetnode
        tvwDirs.OLEGetDropInfo hNode, True

        ' nodes may not move to the the nirvana
        If hNodeSrc = 0 Or hNode = 0 Then Exit Sub

        Set clsDir = clsISOWrt.DirByPath(tvwDirs.GetNodeKey(hNodeSrc))
        If clsDir Is Nothing Then Exit Sub

        Set clsDirDst = clsISOWrt.DirByPath(tvwDirs.GetNodeKey(hNode))
        If clsDirDst Is Nothing Then Exit Sub

        ' source may not be the target
        If clsDir Is clsDirDst Then Exit Sub
        ' source may not be dropped on its parent
        If clsDir.Parent Is clsDirDst Then Exit Sub
        ' source may not be dropped on one of its childs
        If clsISOWrt.DirectoryIsChildOf(clsDir, clsDirDst) Then Exit Sub

        ' find source and remove it from its parent
        For i = 0 To clsDir.Parent.SubDirectoryCount - 1
            If clsDir.Parent.SubDirectory(i) Is clsDir Then
                ' the second parameter is only for cases like this one,
                ' in wich the directory shall be moved!
                ' True causes RemoveSubDirectory to only remove
                ' the directory, but not its subdirectories,
                ' else the moved directory would be empty.
                clsDir.Parent.RemoveSubDirectory i, True
                Exit For
            End If
        Next

        clsDirDst.AddSubDirectoryByRef clsDir

        tvwDirs.DeleteNode hNodeSrc

        ISOBuildTree hNode, clsDirDst

        ShowFilesForDir tvwDirs.SelectedNode
    End If
End Sub

Private Sub tvwDirs_OLEStartDrag( _
    Data As DataObject, _
    AllowedEffects As Long _
)

    Dim clsDir  As clsISODirectory

    Set clsDir = DirFromSelectedNode()
    If clsDir Is Nothing Then Exit Sub

    ' root not movable
    If clsDir.FullPath = "\" Then Exit Sub

    AllowedEffects = vbDropEffectMove
End Sub

Private Sub txtAppID_LostFocus()
    Dim blnJoliet   As Boolean

    blnJoliet = cboDescr.ListIndex = 1

    clsISOWrt.ApplicationID(blnJoliet) = txtAppID.text
    txtAppID.text = clsISOWrt.ApplicationID(blnJoliet)
End Sub

Private Sub txtPrepID_LostFocus()
    Dim blnJoliet   As Boolean

    blnJoliet = cboDescr.ListIndex = 1

    clsISOWrt.DataPreparerID(blnJoliet) = txtPrepID.text
    txtPrepID.text = clsISOWrt.DataPreparerID(blnJoliet)
End Sub

Private Sub txtPubID_LostFocus()
    Dim blnJoliet   As Boolean

    blnJoliet = cboDescr.ListIndex = 1

    clsISOWrt.PublisherID(blnJoliet) = txtPubID.text
    txtPubID.text = clsISOWrt.PublisherID(blnJoliet)
End Sub

Private Sub txtSysID_LostFocus()
    Dim blnJoliet   As Boolean

    blnJoliet = cboDescr.ListIndex = 1

    clsISOWrt.SystemID(blnJoliet) = txtSysID.text
    txtSysID.text = clsISOWrt.SystemID(blnJoliet)
End Sub

Private Sub txtVolID_LostFocus()
    Dim blnJoliet   As Boolean

    blnJoliet = cboDescr.ListIndex = 1

    clsISOWrt.VolumeID(blnJoliet) = txtVolID.text
    txtVolID.text = clsISOWrt.VolumeID(blnJoliet)
End Sub

Private Sub txtVolSetID_LostFocus()
    Dim blnJoliet   As Boolean

    blnJoliet = cboDescr.ListIndex = 1

    clsISOWrt.VolumeSetID(blnJoliet) = txtVolSetID.text
    txtVolSetID.text = clsISOWrt.VolumeSetID(blnJoliet)
End Sub
