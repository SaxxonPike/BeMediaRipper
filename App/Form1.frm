VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BeMedia Ripper iii"
   ClientHeight    =   5895
   ClientLeft      =   150
   ClientTop       =   810
   ClientWidth     =   6600
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6240
      Top             =   3360
   End
   Begin VB.Frame frmSimple 
      BorderStyle     =   0  'None
      Caption         =   "Frame9"
      Height          =   5895
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame14 
         Caption         =   "Rip Volume"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   49
         Top             =   1680
         Width           =   3975
         Begin VB.HScrollBar hsVolumeSimple 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   100
            TabIndex        =   50
            Top             =   240
            Value           =   75
            Width           =   3735
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Content"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   47
         Top             =   4320
         Width           =   2295
         Begin VB.CheckBox chkSimpleRipVids 
            Caption         =   "Videos"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkSimpleRipCharts 
            Caption         =   "Charts"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1080
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkSimpleRipBGM 
            Caption         =   "BGM music"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkSimpleRipKeys 
            Caption         =   "Keysound folders"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Value           =   1  'Checked
            Width           =   2055
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Actions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2520
         TabIndex        =   45
         Top             =   4320
         Width           =   3975
         Begin VB.CommandButton cmdAdvanced 
            Caption         =   "Switch to Advanced Mode"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   3735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Begin"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   3735
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "CD/DVD Drive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4800
         TabIndex        =   42
         Top             =   120
         Width           =   1695
         Begin VB.DriveListBox drvSimple 
            Height          =   315
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Target Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   6375
         Begin VB.CommandButton Command4 
            Caption         =   "Browse..."
            Height          =   255
            Left            =   5160
            TabIndex        =   54
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtSimpleOutput 
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Game Select"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   4575
         Begin VB.ComboBox cmbSimpleGame 
            Height          =   315
            ItemData        =   "Form1.frx":058A
            Left            =   120
            List            =   "Form1.frx":058C
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   240
            Width           =   4335
         End
      End
   End
   Begin VB.Frame frmAdvanced 
      BorderStyle     =   0  'None
      Caption         =   "Frame9"
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame Frame1 
         Caption         =   "Input File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   6375
         Begin VB.TextBox txtInput 
            Height          =   285
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   37
            ToolTipText     =   "Audio source"
            Top             =   240
            Width           =   5055
         End
         Begin VB.CommandButton cmdInputBrowse 
            Caption         =   "Browse..."
            Height          =   255
            Left            =   5280
            TabIndex        =   36
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Output Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   6375
         Begin VB.CommandButton cmdOutputBrowse 
            Caption         =   "Browse..."
            Height          =   255
            Left            =   5280
            TabIndex        =   34
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtOutput 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            ToolTipText     =   "Target folder. It MUST exist already."
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Game Select && Required File (if applicable)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   6375
         Begin VB.ComboBox cmbRipList 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   30
            ToolTipText     =   "Select game type."
            Top             =   240
            Width           =   6135
         End
         Begin VB.TextBox txtEXE 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            OLEDropMode     =   1  'Manual
            TabIndex        =   29
            ToolTipText     =   "SLPM, SLUS, etc."
            Top             =   720
            Width           =   2895
         End
         Begin VB.CommandButton cmdBrowseEXE 
            Caption         =   "Browse..."
            Height          =   255
            Left            =   4560
            TabIndex        =   28
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chkUseEXE 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1320
            TabIndex        =   27
            ToolTipText     =   "Click here to enable/disable the use of the game's executable."
            Top             =   720
            Width           =   225
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Riplist..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   5520
            TabIndex        =   26
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblEXE 
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            ToolTipText     =   "The file required for this game type."
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Rip Selection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   2295
         Begin VB.CheckBox chkKeysounds 
            Caption         =   "Keysounds"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Enable ripping Keysounds. This will create a folder for each set found."
            Top             =   240
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkBGM 
            Caption         =   "BGM"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Enable ripping BGM tracks."
            Top             =   480
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkVideo 
            Caption         =   "Video (VOB)"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Enable ripping standard DVD format videos. These will not be identified by the game's executable file, even if included."
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkRip1 
            Caption         =   "Charts"
            Height          =   255
            Left            =   1320
            TabIndex        =   21
            ToolTipText     =   "Export charts in the appropriate format. (CS, CS2, CS9, etc)"
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "Debug"
            Height          =   255
            Left            =   1320
            TabIndex        =   20
            ToolTipText     =   "Adds misc. info to some folders and creates an info file. (for debug purposes)"
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Progress"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   16
         Top             =   3720
         Width           =   3975
         Begin VB.PictureBox picProgress 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            ScaleHeight     =   2
            ScaleMode       =   0  'User
            ScaleWidth      =   1
            TabIndex        =   17
            Top             =   240
            Width           =   3735
            Begin VB.Label lblProgress 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Width           =   3735
            End
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Actions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2520
         TabIndex        =   14
         Top             =   4920
         Width           =   3975
         Begin VB.CommandButton Command3 
            Caption         =   "Begin"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   3960
         Width           =   2295
         Begin VB.CheckBox Check2 
            Caption         =   "Always stereo output"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Always export WAV files as stereo, even when source is mono"
            Top             =   240
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox chkDDRSilence 
            Caption         =   "DDR: remove silence"
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Remove silence at the beginning of the tracks."
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            Caption         =   "DEBUG: Thru Mode"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Process all code except wave-writing. (for debug purposes)"
            Top             =   480
            Width           =   2055
         End
         Begin VB.CheckBox chkLog 
            Caption         =   "Generate log.txt"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Outputs some different information than ""info"", mostly about offsets. (for debug purposes)"
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtStartOffs 
            Height          =   285
            Left            =   1080
            TabIndex        =   7
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtSkip 
            Height          =   285
            Left            =   1080
            TabIndex        =   6
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Force offset"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Force skip"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2520
         TabIndex        =   1
         Top             =   2760
         Width           =   3975
         Begin VB.HScrollBar hsVolume 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   100
            TabIndex        =   4
            Top             =   240
            Value           =   75
            Width           =   3735
         End
         Begin VB.CheckBox chkUseBGMVol 
            Caption         =   "Use BGM volume setting"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Use the BGM setting in the file. Only applies to IIDX 3rd-8th CS"
            Top             =   510
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkUseBGMBoost 
            Caption         =   "BGM boost"
            Height          =   255
            Left            =   2400
            TabIndex        =   2
            ToolTipText     =   "Boost BGM volume by 20%. Only applies to IIDX 3rd-8th CS"
            Top             =   510
            Value           =   1  'Checked
            Width           =   1455
         End
      End
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuMode 
      Caption         =   "&Mode"
      Begin VB.Menu mnuModeSimple 
         Caption         =   "Simple"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuModeAdvanced 
         Caption         =   "Advanced"
      End
   End
   Begin VB.Menu mnuJobQueue 
      Caption         =   "&Job Queue"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' BeMedia Ripper iii
' by SaxxonPike, 2oo4-2oo7
'  saxxonpike@gmail.com
'
' This program is my child. It grew with me, and I with it.
' Hours of my life has gone into this. And it was never originally planned
' to be what it is now.
'
' (god, this kid must be ugly.)

Private Const ExportRaw1 = False

Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Const IDLE_PRIORITY_CLASS = &H40

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

'===========================================FOLDER CODE
'===========================================FOLDER CODE
'===========================================FOLDER CODE
Private Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lparam As Long
iImage As Long
End Type

'BROWSEINFO.ulFlags values:
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
'===========================================FOLDER CODE
'===========================================FOLDER CODE
'===========================================FOLDER CODE

Dim CLineCode As String
Dim bUnlocked As Boolean
Private OffsetList As New clsOffsetList

Private Type SampleHeaderDJMain
    'Exists As Integer
    'Info0 As Integer
    'Info1 As Integer
    'LowOffset As Integer
    'HiOffset As Integer
    'Info2 As Integer
    'Info3 As Integer
    'Info4 As Integer
    'Info5 As Integer
    'LowEnd As Integer
    'HiEnd As Integer
    'dummy As Integer
    data(0 To 23) As Byte
End Type

Private Type SampleHeaderTwinkle
    IDno As Integer
    SampleRate As Integer
    panning As Byte
    Channels As Byte
    Offs512 As Integer
    Unk4 As Integer
    Unk0 As Integer
    unk1 As Integer
    Unk2 As Byte
    unk3 As Byte
    Leng512 As Integer
End Type

Private Type DJHeaderCustom
    SampleStart As Long
    SampleEnd As Long
    SampleNr As Long
End Type

Private Type SampleHeaderA11
    Ident As Integer
    Unk0(0 To 13) As Byte
End Type

Private Type SampleHeaderB11
    SampleCount As Long
    TotalLength As Long
    Unk0(0 To 7) As Byte
End Type

Private Type SampleInfoA11
    Unk0 As Integer
    unk1 As Byte
    ChanCount As Byte
    Unk2 As Long
    PanLeft As Byte
    PanRight As Byte
    SampleNum As Integer
    volume As Byte
    unk3 As Byte
    Unk4 As Integer
End Type

Private Type SampleInfoB11
    SampOffset As Long
    SampLength As Long
    ChanCount As Integer
    Frequ As Integer
    Unk0 As Long
End Type

Private Type SampleInfo3
    SampleNum As Integer
    Unk0 As Integer
    unk1 As Byte
    vol As Byte
    pan As Byte
    SampType As Byte
    FreqLeft As Long
    FreqRight As Long
    OffsLeft As Long
    OffsRight As Long
    PseudoLeft As Long
    PseudoRight As Long
End Type

Private Type HeaderXA2
    ChannelCount As Long
    ChannelSize As Long
    ChannelLengths(0 To 3) As Long
End Type

Private Type AFSFileName
    fName(0 To 31) As Byte
    Unk0(0 To 15) As Byte
End Type

Private Type AFSFileEntry
    xOffset As Long
    xLength As Long
End Type

Private Type AFSHeader
    AFSident(0 To 3) As Byte
    FileCount As Long
End Type

Private Type PSXSampleHeader
    Unk0 As Long
    unk1(0 To 3) As Byte
    Unk2 As Long
    unk3 As Long
End Type

Private Type PSXSampleInfo
    offs As Long
    Unk0(0 To 2) As Byte
    vol As Byte
    unk1(0 To 2) As Byte
    Unk2 As Byte
    FreqEnc As Long
End Type

'Private Type SampleInfoA9
'End Type

'Private Type SampleInfoB9
'End Type

'Private Type SampleInfo3
'End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private XORType As Long
Private AC2DXDecoder As New cls2DXDecode
Private PSXDecode As New clsPSXDecode
Private bRipping As Boolean
Private EXETableOffset As Long
Private EXETableType As Long
Private EXETitleOffset As Long
Private EXETitleType As Long
Private EXESongOffset As Long
Private EXESongType As Long
Private EXEHasCSFiles As Boolean
Private DataNumberAdjust As Long
Private bRipValidate As Boolean
Private bBassInit As Boolean
Private sMainCaption As String

Private JobMode As Boolean

Private Enum RipList
    AC_DJMain
    AC_Twinkle
    AC_2DX9
    AC_TechnomotionTMM
    CS_beatmania
    CS_2dx03rdStyle
    CS_2dx04thStyle
    CS_2dx05thStyle
    CS_2dx06thStyleA
    CS_2dx06thStyleB
    CS_2dx07thStyleA
    CS_2dx07thStyleB
    CS_2dx07thStyleC
    CS_2dx08thStyleA
    CS_2dx08thStyleB
    CS_2dx08thStyleC
    CS_2dx09thStyle
    CS_2dx10thStyle
    CS_2dx11thStyle
    CS_2dx12thStyleA
    CS_2dx12thStyleB
    CS_2dx12thStyleC
    CS_2dx13thStyleA
    CS_2dx13thStyleB
    CS_2dx13thStyleC
    CS_2dx14thStyleA
    CS_2dx14thStyleB
    CS_2dx14thStyleC
    CS_2dx15thStyleA
    CS_2dx15thStyleB
    CS_2dx15thStyleC
    CS_2dx01stStyleUSA
    CS_2dx01stStyleUSA_JAMPACK
    CS_DDRXBE
    CS_DDRULTRAMIX
    CS_DDRPSX
    CS_DDRCAT
    CS_DDRPS2
    CS_DDRSUPERNOVA
    CS_DDRSUPERNOVA2
    CS_DDRX
    CS_DDRUNIVERSE360
    CS_PopnPSX
    CS_PopnPS2Old
    CS_Popn11
    CS_Taiko7
    MISC_RAWVAG
    MISC_AFS
    MISC_XA2
    MISC_GRANDIA3
    MISC_XG4PS2
    MISC_UNREALTOURNAMENT_PS2
    CS_2dxImage
    CS_2dxImageOld
    CS_2dxCSOld
End Enum

Private Const RM_AC_DJMain = 0
Private Const RM_AC_Twinkle = 1
Private Const RM_AC_2DX9 = 2
Private Const RM_AC_TechnomotionTMM = 3
Private Const RM_CS_beatmania = 4
Private Const RM_CS_2dx03rdStyle = 5
Private Const RM_CS_2dx04thStyle = 6
Private Const RM_CS_2dx05thStyle = 7
Private Const RM_CS_2dx06thStyleA = 8
Private Const RM_CS_2dx06thStyleB = 9
Private Const RM_CS_2dx07thStyleA = 10
Private Const RM_CS_2dx07thStyleB = 11
Private Const RM_CS_2dx07thStyleC = 12
Private Const RM_CS_2dx08thStyleA = 13
Private Const RM_CS_2dx08thStyleB = 14
Private Const RM_CS_2dx08thStyleC = 15
Private Const RM_CS_2dx09thStyle = 16
Private Const RM_CS_2dx10thStyle = 17
Private Const RM_CS_2dx11thStyle = 18
Private Const RM_CS_2dx12thStyleA = 19
Private Const RM_CS_2dx12thStyleB = 20
Private Const RM_CS_2dx12thStyleC = 21
Private Const RM_CS_2dx13thStyleA = 22
Private Const RM_CS_2dx13thStyleB = 23
Private Const RM_CS_2dx13thStyleC = 24
Private Const RM_CS_2dx01stStyleUSA = 25
Private Const RM_CS_2dx01stStyleUSA_JAMPACK = 26
Private Const RM_CS_DDRXBE = 27
Private Const RM_CS_DDRULTRAMIX = 28
Private Const RM_CS_DDRPSX = 29
Private Const RM_CS_DDRCAT = 30
Private Const RM_CS_DDRPS2 = 31
Private Const RM_CS_DDRSUPERNOVA = 32
Private Const RM_CS_DDRSUPERNOVA2 = 33
Private Const RM_CS_DDRUNIVERSE360 = 34
Private Const RM_CS_PopnPSX = 35
Private Const RM_CS_PopnPS2Old = 36
Private Const RM_CS_Popn11 = 37
Private Const RM_CS_Taiko7 = 38
Private Const RM_MISC_RAWVAG = 39
Private Const RM_MISC_AFS = 40
Private Const RM_MISC_XA2 = 41
Private Const RM_MISC_GRANDIA3 = 42
Private Const RM_MISC_XG4PS2 = 43
Private Const RM_MISC_UNREALTOURNAMENT_PS2 = 44
Private Const RM_CS_2dxImage = 45
Private Const RM_CS_2dxImageOld = 46
Private Const RM_CS_2dxCSOld = 47
Private Const RM_CS_2dx14thStyleA = 48
Private Const RM_CS_2dx14thStyleB = 49
Private Const RM_CS_2dx14thStyleC = 50
Private Const RM_CS_DDRX = 51
Private Const RM_CS_2dx15thStyleA = 52
Private Const RM_CS_2dx15thStyleB = 53
Private Const RM_CS_2dx15thStyleC = 54

'==================================================================================
'==================================================================================
'==================================================================================
'==================================================================================
'==================================================================================

Private Sub Rip(RipMode As Long)
    
    Dim AFSFiles() As AFSFileEntry
    Dim AFSFileNames() As AFSFileName
    Dim AFSRealNames() As String
    Dim AFSHead As AFSHeader
    
    Dim PSXSampHeader As PSXSampleHeader
    Dim PSXSampInfo() As PSXSampleInfo
    
    Dim VideoBuffer() As Byte
    Dim InFile As New clsFileStream
    Dim inFile2 As New clsFileStream
    Dim FBuff() As Byte
    Dim FBuff2(0 To 15) As Byte
    Dim CSBuff() As Byte
    Dim ImgBuff() As Long
    Dim ImgBuffB() As Byte
    Dim ImgCount As Long
    Dim fInfo As Long
    Dim bInfo As Boolean
    Dim fIDString As String
    Dim fSize2 As Long
    Dim bCutVideo As Boolean
    Dim FilterString As String
    Dim fInfo2 As Long
    Dim refreshcounts As Long
    Dim Coffs As Double
    Dim FSize As Double
    Dim freq As Long
    Dim RealFreq As Long
    Dim RealSampOffs As Double
    Dim CurrentOffset As Double
    Dim SkipSize As Double
    Dim VAGBlockSize As Double
    Dim HeaderReadSize As Double
    Dim RIPID As Long
    Dim ReadSize As Long
    Dim X As Long
    Dim Y As Long
    Dim z As Long
    Dim IgnoreVid As Boolean
    Dim RipCount As Long
    Dim InvertThing As Long
    Dim InvertVal As Integer
    
    Dim f As Long
    Dim fnameonly As String
    Dim oldripid As Long
    Dim TableSizeLSB As Integer
    Dim bStuffRipped As Boolean
    
    Dim BGMBoost As Double
    
    If chkUseBGMBoost.value = 1 Then
        BGMBoost = 1.2
    Else
        BGMBoost = 1
    End If
    
    'extra formats
    Dim XA2head As HeaderXA2
    
    'ac: twinkle
    Dim TwinkSample(0 To 255) As SampleHeaderTwinkle
    Dim TwinkSampleOffset As Double
    Dim TwinkSampleSize As Long
    
    'ac: djmain
    Dim DJSample(0 To 255) As SampleHeaderDJMain
    Dim DJSampleOffset As Double
    Dim DJSampleOffsetAdd As Long
    Dim DJSampleOffs2 As Double
    Dim DJHead2(0 To 255) As DJHeaderCustom
    Dim DJHeadT As DJHeaderCustom
    
    'IIDX 11 cs
    Dim SampA11() As SampleInfoA11
    Dim SampB11() As SampleInfoB11
    Dim SampA11t As SampleInfoA11
    Dim SampB11t As SampleInfoB11
    Dim SampA11b As SampleInfoA11
    Dim SampB11b As SampleInfoB11
    Dim SampA11Head As SampleHeaderA11
    Dim SampB11Head As SampleHeaderB11
    Dim XOR11 As Long
    Dim bUseTable As Boolean
    
    'IIDX 3 cs
    Dim TableCheck(0 To 7) As Long
    Dim Samp3() As SampleInfo3
    Dim fIndex As Long
    
    'extra
    Dim TextSpinnerPos As Long
    Dim BlankLine As String * 16
    
    ' \\\==================///
    ' ///=== init start ===\\\
    ' \\\==================///
    
    bRipping = True
    
    If mnuModeSimple.Checked = True Then
        Select Case RipMode
            Case RipList.CS_2dx01stStyleUSA
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DATA2.DAT"
            Case RipList.CS_2dx01stStyleUSA_JAMPACK
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BEATM\DATA2.DAT"
            Case RipList.CS_2dx03rdStyle
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_3\BM2DX3.BIN"
            Case RipList.CS_2dx04thStyle
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_4\BM2DX4.BIN"
            Case RipList.CS_2dx05thStyle
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_5\BM2DX5.BIN"
            Case RipList.CS_2dx06thStyleA
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_6\BM2DX6A.BIN"
            Case RipList.CS_2dx06thStyleB
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_6\BM2DX6B.BIN"
            Case RipList.CS_2dx07thStyleA
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_7\BM2DX7A.BIN"
            Case RipList.CS_2dx07thStyleB
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_7\BM2DX7B.BIN"
            Case RipList.CS_2dx07thStyleC
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_7\BM2DX7C.BIN"
            Case RipList.CS_2dx08thStyleA
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_8\BM2DX8A.BIN"
            Case RipList.CS_2dx08thStyleB
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_8\BM2DX8B.BIN"
            Case RipList.CS_2dx08thStyleC
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DX2_8\BM2DX8C.BIN"
            Case RipList.CS_2dx09thStyle
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DATA2.DAT"
            Case RipList.CS_2dx10thStyle
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DATA2.DAT"
            Case RipList.CS_2dx11thStyle
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DATA2.DAT"
            Case RipList.CS_2dx12thStyleA
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX12A.DAT"
            Case RipList.CS_2dx12thStyleB
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX12B.DAT"
            Case RipList.CS_2dx12thStyleC
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX12C.DAT"
            Case RipList.CS_2dx13thStyleA
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX13A.DAT"
            Case RipList.CS_2dx13thStyleB
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX13B.DAT"
            Case RipList.CS_2dx13thStyleC
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX13C.DAT"
            Case RipList.CS_2dx14thStyleA
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX14A.DAT"
            Case RipList.CS_2dx14thStyleB
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX14B.DAT"
            Case RipList.CS_2dx14thStyleC
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX14C.DAT"
            Case RipList.CS_2dx15thStyleA
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX15A.DAT"
            Case RipList.CS_2dx15thStyleB
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX15B.DAT"
            Case RipList.CS_2dx15thStyleC
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\BM2DX15C.DAT"
            'Case RipList.CS_DDRCAT
            Case RipList.CS_DDRPS2
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DATA\FILEDATA.BIN"
            Case RipList.CS_DDRPSX
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\STR.BIN"
            Case RipList.CS_DDRSUPERNOVA, RipList.CS_DDRSUPERNOVA2
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\DATA\MDB_SN1.DAT"
            Case RipList.CS_Popn11
                txtInput.Text = Left$(drvSimple.Drive, 2) + "\SD.DAT"
            Case Else
                MsgBox "Sorry, Simple mode cannot be used for this game type at this time."
                bRipping = False
        End Select
    End If
    
    'create folder if needed
    On Error GoTo 516
    If bRipping Then
        MkDir OutFolder
    End If
516 On Error GoTo 0
    
    fnameonly = GetFileNameOf(txtInput.Text)
    InFile.UseReadBuffer = False
    inFile2.UseReadBuffer = False
    
    If bRipping = True Then
        'parse game executable for names
        If chkUseEXE.value = 1 Then
            OffsetList.SetListFile txtEXE.Text
            OffsetList.ReadFileList EXETableType, EXETableOffset
            OffsetList.ReadInfoList EXESongType, EXESongOffset, DataNumberAdjust, (chkRip1.value = 1)
        End If
        bUseTable = (chkUseEXE.value = 1)
        
        If chkLog.value = 1 Then
            OpenLog OutFolder + "log.txt"
        End If
        
        'open the source file for read access
        InFile.OpenFile txtInput.Text
        InFile.AdvanceOffset = False
        InFile.offset = 0
        FSize = InFile.FileSize
        fInfo = FreeFile
        
        bInfo = (chkInfo.value = 1)
        If bInfo = True And bRipping = True Then
            Open OutFolder + "info.txt" For Output As #fInfo
        End If
        
        Select Case RipMode
            Case RipList.MISC_GRANDIA3
                HeaderReadSize = &H60
                SkipSize = &H800
                If chkUseEXE.value = 0 Then
                    MsgBox "Grandia 3 ripping requires the index file.", vbCritical
                    bRipping = False
                Else
                    bStuffRipped = True
                    f = FreeFile
                    X = 1
                    Open txtEXE.Text For Binary As #f
                    Do While X < LOF(f) And bRipping = True
                        Get #f, X, Y
                        X = X + 4
                        If Y = -1 Then
                            Exit Do
                        End If
                        Y = (Y And &HFFFFF) * &H800
                        InFile.ReadFileBinary FBuff(), 8, Y, True
                        Debug.Print "G3", X \ 4, Hex(Y)
                        DrawProgress X / LOF(f), 0
                        lblProgress.Caption = CStr(X \ 4) + " of " + CStr(LOF(f) \ 4)
                        DoEvents
                        CopyMemory z, FBuff(4), 4
                        PSXDecode.VAGRipSimple InFile, OutFolder + CStr((X \ 4) + 1000) + ".wav", Y + &H800, z, &H800, FBuff(0), , , , , , , True
                    Loop
                    Close #f
                End If
            Case RipList.AC_Twinkle
                HeaderReadSize = &H10
                SkipSize = &H100000
            Case RipList.AC_DJMain
                HeaderReadSize = &H10
                SkipSize = &H1000000
            Case RipList.CS_PopnPSX, RipList.CS_beatmania
                HeaderReadSize = &H10
                SkipSize = &H800
                OffsetList.SetListFile txtInput.Text
                OffsetList.ReadFileList 5, 0
                InFile.AdvanceOffset = False
                OffsetList.CloseListFile
                bUseTable = True
            'Case RipList.CS_PopnPS2Old
            'Case RipList.CS_Popn11
            'Case RipList.CS_Taiko7
            Case RipList.CS_2dx03rdStyle
                HeaderReadSize = &H10
                SkipSize = &H800
            Case RipList.CS_2dx04thStyle
                HeaderReadSize = &H10
                SkipSize = &H800
            Case RipList.CS_2dx05thStyle
                HeaderReadSize = &H10
                SkipSize = &H800
            Case RipList.CS_2dx06thStyleA, RipList.CS_2dx06thStyleB
                HeaderReadSize = &H10
                SkipSize = &H800
            Case RipList.CS_2dx07thStyleA, RipList.CS_2dx07thStyleB, RipList.CS_2dx07thStyleC
                HeaderReadSize = &H10
                SkipSize = &H800
            Case RipList.CS_2dx08thStyleA, RipList.CS_2dx08thStyleB, RipList.CS_2dx08thStyleC
                HeaderReadSize = &H10
                SkipSize = &H800
            Case RipList.CS_2dx09thStyle
                HeaderReadSize = &H20
                SkipSize = &H800
            Case RipList.CS_2dx10thStyle
                HeaderReadSize = &H20
                SkipSize = &H800
            Case RipList.CS_2dx11thStyle 'data2.dat
                HeaderReadSize = &H20
                SkipSize = &H800
            Case RipList.CS_2dx12thStyleA, RipList.CS_2dx12thStyleB, RipList.CS_2dx12thStyleC
                HeaderReadSize = &H20
                SkipSize = &H800
            Case RipList.CS_2dx13thStyleA, RipList.CS_2dx13thStyleB, RipList.CS_2dx13thStyleC
                HeaderReadSize = &H20
                SkipSize = &H800
            Case RipList.CS_2dx14thStyleA, RipList.CS_2dx14thStyleB, RipList.CS_2dx14thStyleC
                HeaderReadSize = &H20
                SkipSize = &H800
                XORType = 1
            Case RipList.CS_2dx15thStyleA, RipList.CS_2dx15thStyleB, RipList.CS_2dx15thStyleC
                HeaderReadSize = &H20
                SkipSize = &H800
                XORType = 1
            Case RipList.CS_2dx01stStyleUSA
                HeaderReadSize = &H20
                SkipSize = &H800
            Case RipList.CS_2dx01stStyleUSA_JAMPACK
                HeaderReadSize = &H20
                SkipSize = &H800
            Case RipList.CS_DDRCAT 'stm#.cat
                VAGBlockSize = &H4000
                HeaderReadSize = &H20
                SkipSize = &H800
                ReadSize = 1
                InFile.offset = &H800
            Case RipList.CS_DDRPS2 'filedata.bin
                SkipSize = &H800
                VAGBlockSize = &H2000
                HeaderReadSize = &H10
                ReadSize = 2
            Case RipList.CS_DDRPSX 'str.bin
                SkipSize = &H800
                VAGBlockSize = &H4000
                HeaderReadSize = &H20
                ReadSize = 1
            Case RipList.CS_DDRSUPERNOVA, RipList.CS_DDRSUPERNOVA2
                SkipSize = &H800
                VAGBlockSize = 16
                HeaderReadSize = &H30
            Case RipList.CS_DDRUNIVERSE360
                SkipSize = &H800
                VAGBlockSize = 16
                HeaderReadSize = &H30
            Case RipList.CS_2dxImage
                SkipSize = &H800
                HeaderReadSize = &H20
            Case RipList.CS_2dxImageOld
                SkipSize = 1
                HeaderReadSize = 5
            Case RipList.CS_2dxCSOld
                SkipSize = 8
                HeaderReadSize = 4
            Case RipList.MISC_XA2
                bStuffRipped = True
                bRipping = False
                InFile.ReadFileObject VarPtr(XA2head), Len(XA2head), 0
                PSXDecode.VAGRipSimple InFile, OutFolder + fnameonly + ".wav", &H800, 44100, XA2head.ChannelSize, XA2head.ChannelCount, False, , , , True
            Case RipList.MISC_XG4PS2
                bStuffRipped = True
                bRipping = False
                PSXDecode.VAGRipSimple InFile, OutFolder + fnameonly + ".wav", 0, 44100, &H800, 2, True
            Case RipList.MISC_RAWVAG
                bStuffRipped = True
                bRipping = False
                With frmRawSettings
                    PSXDecode.VAGRipSimple InFile, OutFolder + fnameonly + ".wav", Val(.Text1(0)), Val(.Text1(3)), Val(.Text1(2)), Val(.Text1(1)), (.Check1.value = 1)
                End With
            Case RipList.MISC_AFS
                bStuffRipped = True
                bRipping = False
                InFile.ReadFileObject VarPtr(AFSHead), 8, 0
                If AFSHead.AFSident(0) = 65 And AFSHead.AFSident(1) = 70 And AFSHead.AFSident(2) = 83 And AFSHead.AFSident(3) = 0 Then
                    ReDim AFSFiles(0 To AFSHead.FileCount) As AFSFileEntry
                    ReDim AFSFileNames(0 To AFSHead.FileCount - 1) As AFSFileName
                    ReDim AFSRealNames(0 To UBound(AFSFileNames)) As String
                    InFile.ReadFileObject VarPtr(AFSFiles(0)), (AFSHead.FileCount + 1) * 8, 8
                    InFile.ReadFileObject VarPtr(AFSFileNames(0)), AFSHead.FileCount * 48, AFSFiles(AFSHead.FileCount).xOffset
                    For X = 0 To AFSHead.FileCount - 1
                        AFSRealNames(X) = LTrim(RTrim(StrConv(AFSFileNames(X).fName(), vbUnicode)))
                        If InStr(AFSRealNames(X), Chr(0)) Then
                            AFSRealNames(X) = Left$(AFSRealNames(X), InStr(AFSRealNames(X), Chr(0)) - 1)
                        End If
                        AFSRealNames(X) = OutFolder + Replace(AFSRealNames(X), "/", "\")
                        MkDir2 AFSRealNames(X), True
                        InFile.ReadFileBinary FBuff(), AFSFiles(X).xLength, AFSFiles(X).xOffset, True
                        f = FreeFile
                        Open AFSRealNames(X) For Binary As #f
                        Put #f, 1, FBuff
                        Close #f
                    Next X
                End If
            Case RipList.AC_2DX9
                bStuffRipped = True
                bRipping = False
                If Not bBassInit Then
                    bBassInit = True
                    If BASS_Init(0, 44100, 0, Me.hWnd, ByVal 0&) = BASSFALSE Then
                        MsgBox "BASS could not be initialized... Bemani PC decoding functions will not work.", vbCritical, "Doh!"
                        BASS_Free
                        bBassInit = False
                    End If
                End If
                If bBassInit Then
                    If InStr(fnameonly, ".") > 0 Then
                        fnameonly = Left$(fnameonly, InStr(fnameonly, ".") - 1)
                    End If
                    If Dir(OutFolder + fnameonly, vbDirectory) = "" Then
                        On Error Resume Next
                        MkDir OutFolder + fnameonly
                        On Error GoTo 0
                    End If
                    OutFolder = OutFolder + fnameonly + "\"
                    AC2DXDecoder.Decrypt InFile
                Else
                    MsgBox "Bass was not initialized properly. Bemani PC conversion cannot continue.", vbCritical, "Doh!"
                End If
            Case RipList.AC_TechnomotionTMM
                ReDim FBuff(0 To &H1400& - 1) As Byte
                ReDim CSBuff(0 To &H1400& - 1) As Byte
                f = FreeFile
                Open OutFolder + "one.raw" For Binary As #f
                z = FreeFile
                Open OutFolder + "two.raw" For Binary As #z
                CurrentOffset = &H1800&
    '            ImgCount = 0
    '            For X = 0 To 7
    '                InFile.ReadFileObject VarPtr(Y), 4, CurrentOffset - &H10
    '                'If Y <> 0 Then
    '                    If ImgCount = 0 Then
    '                        If Y <> 0 Then
    '                            ImgCount = 1 'using this as a temp
    '                        End If
    '                    Else
    '                        If Y = 0 Then
    '                            Exit For
    '                        End If
    '                    End If
    '                'End If
    '                CurrentOffset = CurrentOffset + &HE00
    '            Next X
                'If X < 4 Then
                    X = 1
                    Do While CurrentOffset < (InFile.FileSize + &H2800&)
                        InFile.ReadFileBinary FBuff(), , CurrentOffset
                        InFile.ReadFileBinary CSBuff(), , CurrentOffset + &H1400&
                        Put #f, X, FBuff
                        Put #z, X, CSBuff
                        X = X + &H1400&
                        CurrentOffset = CurrentOffset + &H2800&
                    Loop
                'End If
                Close #f
                Close #z
            Case RipList.MISC_UNREALTOURNAMENT_PS2
                bStuffRipped = True
                bRipping = False
                PSXDecode.VAGRipSimple InFile, OutFolder + fnameonly + ".wav", 0, 22050, 16, 2, , 64
            Case Else
                Debug.Print "NO SKIP SIZE FOR THIS TYPE - DEFAULTING TO 0x800"
                SkipSize = &H800
                HeaderReadSize = 16
        End Select
        
        RIPID = 1001
        fIndex = 0
        f = FreeFile
        
        If bUseTable = True And ((OffsetList.CSCount > 0 And EXESongType = 1) Or EXEHasCSFiles = True) Then
            PrintLog "***** Offset table is being used for CS ripping."
            lblProgress = "Ripping EXE data"
            DoEvents
            If EXESongType = 1 Or EXESongType = 8 Then
                inFile2.OpenFile txtEXE.Text 'executable contains charts internally
            Else
                inFile2.OpenFile txtInput.Text 'data file contains charts
            End If
            If chkRip1.value = 1 Then
            For X = 0 To OffsetList.CSCount - 1
                If Not ((EXESongType = 12) And (OffsetList.CSLink(X) = 0)) Then
                    Debug.Print "CS:", X, OffsetList.CSLink(X), OffsetList.CSOffset(X), OffsetList.CSName(X)
                    If ExportRaw1 Then
                        Y = 0
                        z = 1
                        Do
                            If inFile2.ReadFileObject(VarPtr(z), 4, (Y + OffsetList.CSOffset(X) + &H7FC&)) Then
                                Exit Do
                            End If
                            Y = Y + &H800&
                        Loop Until z = 0
                        If Y > 0 Then
                            inFile2.QuickExtract OutFolder + OffsetList.CSName(X) + ".csR", Y, OffsetList.CSOffset(X)
                        End If
                    
                    ElseIf PSXDecode.DecodeBemani1(inFile2, CSBuff(), (OffsetList.CSOffset(X))) > 0 Then
                        PrintLog CStr(OffsetList.CSOffset(X)) + ": " + OffsetList.CSName(X)
                        If EXETableType = 1 Then
                            FilterString = OffsetList.CSName(X) + ".cs2"
                        Else
                            FilterString = OffsetList.CSName(X) + ".cs"
                        End If
                        If Left$(OffsetList.CSName(X), 1) = " " Then
                            FilterString = "Keys" + CStr(X) + FilterString
                        End If
                        If chkRip1.value = 1 Then
                            Open OutFolder + FilterString For Binary As #f
                            Put #f, 1, CSBuff
                            Close #f
                        End If
                        DrawProgress 0, X / OffsetList.CSCount
                        DoEvents
                    End If
                End If
            Next X
            End If
            inFile2.CloseFile
        End If
        
        If HeaderReadSize > 0 Then
            ReDim FBuff(0 To HeaderReadSize - 1) As Byte
        End If
        
    End If
    
    ' \\\=======================///
    ' ///=== main loop start ===\\\
    ' \\\=======================///
    
    If bUseTable = True Then
        PrintLog "***** Offset table is being used for audio/video ripping."
    End If
    
    If txtStartOffs.Text <> "" Then
        InFile.offset = Val(txtStartOffs.Text)
    End If
    If txtSkip.Text <> "" Then
        SkipSize = Val(txtSkip.Text)
    End If
    
    Do
        If bRipping = False Then Exit Do
        bStuffRipped = True
        refreshcounts = refreshcounts + 1
        If refreshcounts = 1500 Then
            TextSpinnerPos = TextSpinnerPos + 1
            lblProgress = "  Searching..."
            Select Case TextSpinnerPos
                Case 0, 4
                    lblProgress = lblProgress + " "
                    TextSpinnerPos = 0
                Case 1
                    lblProgress = lblProgress + "  "
                Case 2
                    lblProgress = lblProgress + "   "
                Case 3
                    lblProgress = lblProgress + "  "
            End Select
            refreshcounts = 0
            DrawProgress InFile.offset / FSize, 0
            DoEvents
        End If
        
        If bUseTable = True Then
            If OffsetList.ListLength(fIndex) = 0 Then
                Exit Do
            End If
            
            If fIndex >= OffsetList.ListCount Then
                Exit Do
            End If
            
            InFile.offset = OffsetList.ListOffset(fIndex)
            RIPID = 1000 + fIndex
            PrintLog CStr(InFile.offset) + ": " + CStr(RIPID) + " " + OffsetList.ListName(fIndex)
            
            'Debug.Print "SCAN:", InFile.Offset
            
            fIndex = fIndex + 1
        End If
        
        InFile.ReadFileBinary FBuff()
        
        If bUseTable = True Then
            If FBuff(6) = &HBB& And FBuff(0) = 0 Then
                InFile.offset = InFile.offset + 2048
                InFile.ReadFileBinary FBuff()
            End If
        End If
        
        'VIDEO
        If FBuff(0) = 0 And FBuff(1) = 0 And FBuff(2) = 1 And FBuff(3) = &HB9 Then
            IgnoreVid = False
        End If
        
        If FBuff(0) = 0 And FBuff(1) = 0 And FBuff(2) = 1 And FBuff(3) = &HBA And IgnoreVid = False Then
            If InFile.offset > 0 Then
                InFile.ReadFileBinary FBuff2(), , InFile.offset - 16
                InFile.offset = InFile.offset + 16
            End If
            If Compare16(FBuff2(), BlankLine) Or InFile.offset = 0 Then
                IgnoreVid = True
                ReDim VideoBuffer(1 To &H4000&) As Byte
                Debug.Print "VOB_Video: rip"; InFile.offset
                If chkVideo.value = 1 And frmRipList.IsSelected(RIPID) = True Then
                    RipCount = RipCount + 1
                    CurrentOffset = InFile.offset
                    InFile.AdvanceOffset = True
                    X = FreeFile
                    Debug.Print "VOB:"; InFile.offset
                    lblProgress = "Extracting Video " + CStr(RIPID)
                    fIDString = CStr(RIPID)
                    If bUseTable = True Then
                        fIDString = fIDString + " " + OffsetList.ListName(RIPID - 1000) + " vid"
                    End If
                    Open OutFolder + fIDString + ".vob" For Binary As #X
                    DoEvents
                    freq = -1191116800
                    Do
                        InFile.ReadFileBinary VideoBuffer()
                        If VideoBuffer(1) = 0 And VideoBuffer(2) = 0 And VideoBuffer(3) = 1 And VideoBuffer(4) = &HB9 Then
                            Put #X, , freq
                        Else
                            Put #X, , VideoBuffer
                        End If
                    Loop Until (VideoBuffer(1) = 0 And VideoBuffer(2) = 0 And VideoBuffer(3) = 1 And VideoBuffer(4) = &HB9)
                    Close #X
                    IgnoreVid = False
                    DoEvents
                    InFile.AdvanceOffset = False
                    InFile.offset = (InFile.offset - &H4000) - SkipSize
                End If
                
                RIPID = RIPID + 1
            End If
        End If
        
        ' CS simfiles
        Select Case RipMode
            Case RipList.CS_2dx03rdStyle, _
                 RipList.CS_2dx04thStyle, _
                 RipList.CS_2dx05thStyle, _
                 RipList.CS_2dx06thStyleA, RipList.CS_2dx06thStyleB, _
                 RipList.CS_2dx07thStyleA, RipList.CS_2dx07thStyleB, RipList.CS_2dx07thStyleC, _
                 RipList.CS_2dx08thStyleA, RipList.CS_2dx08thStyleB, RipList.CS_2dx08thStyleC, _
                 RipList.CS_2dx09thStyle, _
                 RipList.CS_2dx10thStyle, _
                 RipList.CS_2dx11thStyle, _
                 RipList.CS_2dx12thStyleA, RipList.CS_2dx12thStyleB, RipList.CS_2dx12thStyleC, _
                 RipList.CS_2dx13thStyleA, RipList.CS_2dx13thStyleB, RipList.CS_2dx13thStyleC, _
                 RipList.CS_2dx14thStyleA, RipList.CS_2dx14thStyleB, RipList.CS_2dx14thStyleC, _
                 RipList.CS_2dxCSOld, RipList.CS_2dx01stStyleUSA, RipList.CS_2dx01stStyleUSA_JAMPACK
                 
                If FBuff(0) = &H64 And FBuff(1) = 8 And FBuff(2) = 0 And FBuff(3) = &H80 And bUnlocked = True And (EXETableType = 0) Then
                    Debug.Print "CSfile: rip"; InFile.offset
                    DoEvents
                    If chkRip1.value = 1 Then
                        lblProgress = "Ripping CS " + CStr(RIPID)
                        PSXDecode.DecodeBemani1 InFile, CSBuff()
                        f = FreeFile
                        If CSBuff(UBound(CSBuff) - 5) <> 0 Then
                            'old format
                            Open OutFolder + "[b" + CStr(FindBPM2(CSBuff())) + "] " + CStr(RIPID) + " (n" + CStr(FindNoteCount2(CSBuff())) + ") [" + BME(FindHiKey2(CSBuff())) + "].CS2" For Binary As #f
                        Else
                            'new format
                            Open OutFolder + "[b" + CStr(FindBPM(CSBuff())) + "] " + CStr(RIPID) + " (n" + CStr(FindNoteCount(CSBuff())) + ") [" + BME(FindHiKey(CSBuff())) + "].CS" For Binary As #f
                        End If
                        Put #f, 1, CSBuff
                        Close #f
                        RipCount = RipCount + 1
                    End If
                    RIPID = RIPID + 1
                End If
            
            Case Else
        End Select
        
        
        ' ---------------------------=== MAIN RIP CODE ===---------------------------
        ' ---------------------------=== MAIN RIP CODE ===---------------------------
        ' ---------------------------=== MAIN RIP CODE ===---------------------------
        
        
        Select Case RipMode
            Case RipList.AC_DJMain
                
                CurrentOffset = InFile.offset
                InFile.AdvanceOffset = False
                
                InFile.ReadFileObject VarPtr(DJSample(0)), 22
                If DJSample(0).data(0) = &H38 Then
                    On Local Error Resume Next
                    If (chkKeysounds.value = 1) Or (chkRip1.value = 1) Then
                        MkDir OutFolder + CStr(RIPID)
                    End If
                    On Local Error GoTo 0
                    For X = LBound(DJSample) To UBound(DJSample)
                        InFile.ReadFileObject VarPtr(DJSample(X)), 22, CurrentOffset + (X * 22)
                    Next X
                    InFile.ReadFileBinary FBuff(), 4, CurrentOffset + &H800&, True
                    If FBuff(0) = 0 And FBuff(1) = 0 And FBuff(2) = 0 And FBuff(3) = 0 Then
                        DJSampleOffs2 = &H800&
                    Else
                        DJSampleOffs2 = &H2000&
                    End If
                    
                    For X = LBound(DJSample) To UBound(DJSample)
                        With DJSample(X)
                            If .data(0) = &H38& Then
                                DJSampleOffset = DWBytes(.data(7), .data(6), .data(9), 0) + CurrentOffset + DJSampleOffs2
                                DJSampleOffsetAdd = (DWBytes(.data(16), .data(19), .data(18), 0)) ' - DJSampleOffset
                                If .data(16) = &HA And .data(19) = &HA And .data(18) = &HA Then
                                    DJSampleOffsetAdd = 0
                                End If
                                DJSampleOffset = DJSampleOffset
                                DJSampleOffsetAdd = DJSampleOffsetAdd + CurrentOffset + DJSampleOffs2
                                
                                If (chkKeysounds.value = 1) Then
                                    
                                    If .data(5) = &H7F Then
                                        Exit For
                                    End If
                                    Debug.Print BME((X * 2) + 1), DJSampleOffset, .data(5)
                                    PSXDecode.DJMainToWav InFile, OutFolder + CStr(RIPID) + "\" + BME((X * 2) + 1) + ".wav", DJSampleOffset + 0, 0, 44100, 1, 0, , , False
                                    
                                    If .data(14) = &H7F Then
                                        Exit For
                                    End If
                                    Debug.Print BME((X * 2) + 2), DJSampleOffsetAdd, .data(14)
                                    PSXDecode.DJMainToWav InFile, OutFolder + CStr(RIPID) + "\" + BME((X * 2) + 2) + ".wav", DJSampleOffsetAdd + 0, 0, 44100, 1, 0, , , False
                                    
                                End If
                                
                            End If
                            If (.data(4) = &HA And .data(5) = &HA) Or (.data(14) = &H7F) Then
                                Exit For
                            End If
                        End With
                    Next X
                    RIPID = RIPID + 1
                End If
                
                X = 0
                InFile.offset = CurrentOffset ' + SkipSize
                InFile.AdvanceOffset = False
                DoEvents
                
            Case RipList.AC_Twinkle
                
                CurrentOffset = InFile.offset
                InFile.AdvanceOffset = True
                'Check for data
                InFile.ReadFileObject VarPtr(TwinkSample(0)), 18
                If TwinkSample(0).unk3 = 7 And TwinkSample(0).Unk2 = 0 And TwinkSample(0).Leng512 <> 0 Then
                    
                    On Local Error Resume Next
                    If (chkKeysounds.value = 1) Or (chkRip1.value = 1) Then
                        MkDir OutFolder + CStr(RIPID)
                    End If
                    On Local Error GoTo 0
                    
                    InFile.ReadFileObject VarPtr(TwinkSample(0)), 4608, CurrentOffset
                    
                    f = FreeFile
                    
                    For X = 0 To 3
                        InFile.ReadFileBinary FBuff(), &H4000&, CurrentOffset + &H2000& + (X * &H4000&), True
                        If FBuff(3) <> 0 And (chkRip1.value = 1) Then
                            Open OutFolder + CStr(RIPID) + "\" + "@" + CStr(X) + ".cs2" For Binary As #f
                            Put #f, 1, FBuff
                            Close #f
                        End If
                    Next X
                    
                    Debug.Print RIPID, CurrentOffset
                    For X = 0 To 254
                        If TwinkSample(X).Leng512 > 0 And TwinkSample(X).SampleRate <> 2570 Then
                            TwinkSampleOffset = TwinkSample(X).Offs512
                            If TwinkSampleOffset < 0 Then
                                TwinkSampleOffset = TwinkSampleOffset + 65536
                            End If
                            freq = TwinkSample(X).SampleRate
                            If freq < 0 Then
                                freq = freq + 65536
                            End If
                            TwinkSampleOffset = CDbl((TwinkSampleOffset * 512) + &H100000) + CurrentOffset
                            TwinkSampleSize = Int(CLng(TwinkSample(X).Leng512) * 1445.1 * (freq / 44100))
                            If TwinkSample(X).Channels > &H40 Then
                                TwinkSample(X).Channels = 1
                            Else
                                TwinkSample(X).Channels = (&H40 - TwinkSample(X).Channels)
                            End If
                            If chkKeysounds.value = 1 Then
                                If TwinkSample(X).panning = 0 Or TwinkSample(X).panning > 127 Then
                                    TwinkSample(X).panning = 64
                                End If
                                PSXDecode.TwinkleToWAV InFile, OutFolder + CStr(RIPID) + "\" + BME(X + 1) + ".wav", TwinkSampleOffset, CDbl(TwinkSampleSize), freq, TwinkSample(X).Channels, TwinkSample(X).panning, (TwinkSample(X).Unk2 <> 0)
                            End If
                        End If
                    Next X
                    If TwinkSample(0).Leng512 = TwinkSample(1).Leng512 Then
                        PSXDecode.CombineWaves OutFolder + CStr(RIPID) + "\" + "01.wav", OutFolder + CStr(RIPID) + "\" + "02.wav", 0 '154
                        If Dir(OutFolder + CStr(RIPID) + "\" + "02.wav") <> "" Then
                            Kill OutFolder + CStr(RIPID) + "\" + "02.wav"
                        End If
                    End If
                    RIPID = RIPID + 1
                    CurrentOffset = CurrentOffset + &H1700000
                End If
                
                X = 0
                InFile.offset = CurrentOffset ' + SkipSize
                InFile.AdvanceOffset = False
                DoEvents
                
            Case RipList.CS_PopnPSX, RipList.CS_beatmania
                
                CurrentOffset = InFile.offset
                
                'Charts
                InFile.ReadFileObject VarPtr(Y), 4, (CurrentOffset + OffsetList.ListLength(fIndex - 1)) - 4
                If Y = &H7FFF& Then
                    If chkRip1.value = 1 Then
                        ReDim CSBuff(0 To OffsetList.ListLength(fIndex - 1) - 1) As Byte
                        InFile.ReadFileBinary CSBuff(), , CurrentOffset
                        f = FreeFile
                        If RipMode = RipList.CS_PopnPSX Then
                            Open OutFolder + CStr(RIPID) + ".cs9" For Binary As #f
                        Else
                            Open OutFolder + CStr(RIPID) + ".cs5" For Binary As #f
                        End If
                        Put #f, 1, CSBuff
                        Close #f
                    End If
                    RIPID = RIPID + 1
                
                'Keysounds
                ElseIf FBuff(0) = 0 And FBuff(1) = 0 And FBuff(2) = 8 And FBuff(4) = 0 And FBuff(14) = 0 And FBuff(15) = 0 Then
                    InFile.ReadFileObject VarPtr(PSXSampHeader), 16, CurrentOffset
                    'reverse endian on length
                    TableSizeLSB = MakeWord(PSXSampHeader.unk1(3), PSXSampHeader.unk1(2))
                    ReDim PSXSampInfo(1 To (TableSizeLSB \ 16)) As PSXSampleInfo
                    For X = 1 To (TableSizeLSB \ 16)
                        InFile.ReadFileObject VarPtr(PSXSampInfo(X).offs), 16, CurrentOffset + (X * 16)
                    Next X
                    fIDString = " [" + CStr(CurrentOffset) + "]"
                    If chkKeysounds.value = 1 Then
                        On Local Error Resume Next
                        If (Check1.value = 0) Then
                            MkDir OutFolder + CStr(RIPID) + fIDString
                        End If
                        On Local Error GoTo 0
                        RipCount = RipCount + 1
                        For X = 1 To (TableSizeLSB \ 16)
                            'Debug.Print x, PSXSampInfo(x).Offs - 4096 + 48 + CurrentOffset + PSXSampHeader.TableSizeLSB
                            If RipMode = RipList.CS_PopnPSX Then
                                PSXDecode.VAGRipSimple InFile, OutFolder + CStr(RIPID) + fIDString + "\" + BME(X + 1) + ".wav", PSXSampInfo(X).offs - 4096 + 48 + CurrentOffset + TableSizeLSB, 18900, 16, 1, , , , PSXSampInfo(X).vol
                            Else
                                PSXDecode.VAGRipSimple InFile, OutFolder + CStr(RIPID) + fIDString + "\" + BME(X + 1) + ".wav", PSXSampInfo(X).offs - 4096 + 48 + CurrentOffset + TableSizeLSB, 37800, 16, 1, , , , PSXSampInfo(X).vol
                            End If
                            Debug.Print RIPID, X, PSXSampInfo(X).FreqEnc
                        Next X
                    End If
                    InFile.offset = CurrentOffset + &H800&
                    RIPID = RIPID + 1
                End If
                
            Case RipList.CS_2dx03rdStyle, _
                 RipList.CS_2dx04thStyle, _
                 RipList.CS_2dx05thStyle, _
                 RipList.CS_2dx06thStyleA, RipList.CS_2dx06thStyleB, _
                 RipList.CS_2dx07thStyleA, RipList.CS_2dx07thStyleB, RipList.CS_2dx07thStyleC, _
                 RipList.CS_2dx08thStyleA, RipList.CS_2dx08thStyleB, RipList.CS_2dx08thStyleC
                
                CurrentOffset = InFile.offset
                fIDString = ""
                
                
                    
                
                If Compare16(FBuff(), Chr(0) + Chr(2) + String(14, Chr(0))) Then
                    InFile.ReadFileBinary FBuff2(), , CurrentOffset - 16
                    If Compare16(FBuff2(), BlankLine) Then
                        InFile.ReadFileBinary FBuff2(), , CurrentOffset + &H800
                        If Compare16(FBuff2(), Chr(0) + Chr(2) + String(14, Chr(0))) Then
                            InFile.ReadFileBinary FBuff2(), , CurrentOffset - &H800
                            If bUseTable = True Then
                                fIDString = " " + OffsetList.ListName(RIPID - 1000)
                            End If
                            lblProgress = "Extracting BGM " + CStr(RIPID)
                            Debug.Print "CS_2dx08thStyle: BGM"; RIPID, InFile.offset
                            If FBuff2(5) > 170 Then
                                FBuff2(5) = 170
                            End If
                            DoEvents
                            If chkBGM.value = 1 And frmRipList.IsSelected(RIPID) = True Then
                                RipCount = RipCount + 1
                                If chkUseBGMVol.value = 1 Then
                                    fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + fIDString + ".wav", CurrentOffset, 48000, &H800, 2, , 256, , CByte((CDbl(FBuff2(5)) / 100) * 128 * BGMBoost))
                                Else
                                    fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + fIDString + ".wav", CurrentOffset, 48000, &H800, 2, , 256)
                                End If
                            End If
                            RIPID = RIPID + 1
                            InFile.offset = CurrentOffset + &H1000&
                        End If
                    End If
                End If
                
                'Keys
                InFile.ReadFileObject VarPtr(TableCheck(0)), 32, CurrentOffset
                If TableCheck(0) > 0 And TableCheck(1) > 0 And TableCheck(2) >= 0 And TableCheck(2) < TableCheck(1) And TableCheck(1) < TableCheck(0) Then
                    If (TableCheck(0) - TableCheck(2)) - TableCheck(1) = 16384 And TableCheck(3) = 0 Then
                        If TableCheck(4) = 0 And TableCheck(5) = 0 And TableCheck(6) = 0 And TableCheck(7) = 0 Then
                            Debug.Print "CS_2dx08thStyle: keys"; RIPID, CurrentOffset
                            lblProgress = "Extracting Keysound Set " + CStr(RIPID)
                            DoEvents
                            ReDim Samp3(1 To 511) As SampleInfo3
                            InFile.ReadFileObject VarPtr(Samp3(1)), &H3FE0&, CurrentOffset + &H20&
                            freq = 0
                            For X = 1 To 511
                                If Samp3(X).SampleNum <> 0 Then
                                    If X > 1 Then
                                        If Samp3(X - 1).SampleNum <> 255 And Samp3(X).SampleNum = 256 Then
                                            freq = X - 1
                                            Exit For
                                        End If
                                    End If
                                    freq = X
                                End If
                            Next X
                            If bUseTable = True Then
                                fIDString = " " + OffsetList.ListName(RIPID - 1000)
                            End If
                            
                            Do While Right$(fIDString, 1) = "."
                                fIDString = RTrim$(Left$(fIDString, Len(fIDString) - 1))
                            Loop

                            If chkKeysounds.value = 1 And frmRipList.IsSelected(RIPID) = True Then
                                On Local Error Resume Next
                                If (Check1.value = 0) Then
                                    MkDir OutFolder + CStr(RIPID) + fIDString
                                End If
                                On Local Error GoTo 0
                                RipCount = RipCount + 1
                            End If
                            For X = 1 To freq
                                If Samp3(X).FreqLeft > 0 And (Samp3(X).OffsLeft > 0 Or Samp3(X).PseudoLeft > 0) Then
                                    With Samp3(X)
                                        .SampType = (.SampType And &HF)
                                        If chkKeysounds.value = 1 And frmRipList.IsSelected(RIPID) = True Then
                                            If .SampType = 2 Then
                                                fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + fIDString + "\" + BME(.SampleNum) + ".wav", (.OffsLeft + CurrentOffset) - 61456, .FreqLeft, 16, 1, , , .pan, .vol)
                                            ElseIf .SampType = 3 Then
                                                If X > 1 Then
                                                    If Not (Samp3(X - 1).SampleNum <> 255 And .SampleNum = 256) Then
                                                        fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + fIDString + "\" + BME(.SampleNum) + ".wav", TableCheck(1) + CurrentOffset + .OffsLeft + 16400, .FreqLeft, 16, 1, , , .pan, .vol)
                                                    End If
                                                Else
                                                    fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + fIDString + "\" + BME(.SampleNum) + ".wav", TableCheck(1) + CurrentOffset + .OffsLeft + 16400, .FreqLeft, 16, 1, , , .pan, .vol)
                                                End If
                                            ElseIf .SampType = 4 Then
                                                fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + fIDString + "\" + BME(.SampleNum) + ".wav", (.OffsLeft + CurrentOffset) - 61456, .FreqLeft, (.OffsRight - .OffsLeft), 2, , , .pan, .vol)
                                            End If
                                        End If
                                    End With
                                    DrawProgress 0, X / freq
                                    If bRipping = False Then
                                        Exit For
                                    End If
                                End If
                            Next X
                            InFile.offset = CurrentOffset + &H4000
                            RIPID = RIPID + 1
                            picProgress.Cls
                        End If
                    End If
                End If
                
            Case RipList.CS_2dx09thStyle, _
                 RipList.CS_2dx10thStyle, _
                 RipList.CS_2dx11thStyle, RipList.CS_2dx01stStyleUSA, RipList.CS_2dx01stStyleUSA_JAMPACK, _
                 RipList.CS_2dx12thStyleA, RipList.CS_2dx12thStyleB, RipList.CS_2dx12thStyleC, _
                 RipList.CS_2dx13thStyleA, RipList.CS_2dx13thStyleB, RipList.CS_2dx13thStyleC, _
                 RipList.CS_2dx14thStyleA, RipList.CS_2dx14thStyleB, RipList.CS_2dx14thStyleC, _
                 RipList.CS_2dx15thStyleA, RipList.CS_2dx15thStyleB, RipList.CS_2dx15thStyleC, _
                 RipList.CS_Popn11
                
                InFile.AdvanceOffset = False
                fIDString = ""
                
                
                
                'BGM ---
                If FBuff(0) = 1 And FBuff(1) = 0 And FBuff(2) = &H64 And FBuff(3) = 8 Then
                    CurrentOffset = InFile.offset
                    InFile.ReadFileObject VarPtr(freq), 4, CurrentOffset + &H18&
                    lblProgress = "Extracting BGM " + CStr(RIPID)
                    Debug.Print "CS_2dx11thStyle: BGM"; RIPID, InFile.offset
                    If bUseTable = True Then
                        fIDString = " " + OffsetList.ListName(RIPID - 1000)
                    End If
                    If chkBGM.value = 1 And frmRipList.IsSelected(RIPID) = True Then
                        XOR11 = 0
                        Select Case RipMode
                            Case RipList.CS_2dx14thStyleA, RipList.CS_2dx14thStyleB, RipList.CS_2dx14thStyleC
                                InFile.ReadFileObject VarPtr(XOR11), 4, CurrentOffset + &H800&
                                XORType = 1
                            Case RipList.CS_2dx15thStyleA, RipList.CS_2dx15thStyleB, RipList.CS_2dx15thStyleC
                                InFile.ReadFileObject VarPtr(XOR11), 4, CurrentOffset + &H800&
                                XORType = 1
                        End Select
                        fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + fIDString + ".wav", CurrentOffset + &H800&, freq, &H10, 2, , 512, , , , XOR11, , XORType)
                        RipCount = RipCount + 1
                    End If
                    RIPID = RIPID + 1
                    InFile.offset = CurrentOffset + &H1000&
                End If
                
                'KEYS ---
                If (FBuff(0) = &H77 And FBuff(1) = &H70 And FBuff(3) = 0) Or _
                    (FBuff(0) = &H65 And FBuff(1) = &H66 And FBuff(2) = 1 And FBuff(3) = 0) Or _
                    (FBuff(0) = &HCF And FBuff(1) = &H2 And FBuff(2) = 0 And FBuff(3) = 0) Then
                
                    X = FBuff(0)
                    CurrentOffset = InFile.offset
                    InFile.ReadFileObject VarPtr(SampA11Head), Len(SampA11Head), CurrentOffset
                    InFile.ReadFileObject VarPtr(SampB11Head), Len(SampB11Head), CurrentOffset + &H8000&
                    If bInfo Then
                        fIDString = " [" + Hex(SampB11Head.TotalLength + SampB11Head.SampleCount) + "]"
                    End If
                    If bUseTable = True Then
                        fIDString = " " + OffsetList.ListName(RIPID - 1000)
                    End If
                    Do While Right$(fIDString, 1) = "."
                        fIDString = RTrim$(Left$(fIDString, Len(fIDString) - 1))
                    Loop
                    
                    If SampB11Head.SampleCount > 0 Then
                        lblProgress = "Extracting Keysound Set " + CStr(RIPID)
                        Debug.Print "CS_2dx11thStyle: keys"; RIPID, CurrentOffset
                        ReDim SampA11(1 To 2047) As SampleInfoA11
                        ReDim SampB11(1 To SampB11Head.SampleCount) As SampleInfoB11
                        
                        ReDim FBuff(0 To (2047 * 16) - 1) As Byte
                        
                        InFile.ReadFileObject VarPtr(SampA11(1).Unk0), 32752, CurrentOffset + &H10&
                        InFile.ReadFileObject VarPtr(SampB11(1).SampOffset), SampB11Head.SampleCount * 16, CurrentOffset + &H8010&
                        
                        InFile.ReadFileObject VarPtr(fSize2), 4, CurrentOffset + &H8000&
                        
                        FBuff(0) = X
                        If FBuff(0) = &H77 Then 'red XOR encoding only
                            InFile.ReadFileObject VarPtr(XOR11), 4, CurrentOffset + &H8010& + (SampB11Head.SampleCount * 16)
                        End If
                        If chkKeysounds.value = 1 And frmRipList.IsSelected(RIPID) = True Then
                            On Error Resume Next
                            MkDir OutFolder + CStr(RIPID) + fIDString
                            If bInfo Then
                                fInfo2 = FreeFile
                                Open OutFolder + CStr(RIPID) + fIDString + "\keysound.log" For Output As #fInfo2
                                Print #fInfo2, "Num", "Frequency", "Guessed Freq"
                            End If
                            On Error GoTo 0
                            RipCount = RipCount + 1
                        End If
                        freq = 2
                        If SampA11(2).SampleNum <> SampA11(1).SampleNum Or SampA11(1).volume <> 0 Then
                            freq = 1
                        End If
                        For X = freq To 2047
                            
                            If SampA11(X).ChanCount > 0 Then
                                
                                RealSampOffs = SampB11(SampA11(X).SampleNum + 1).SampOffset + &H8020& + (SampB11Head.SampleCount * 16) + CurrentOffset
                                If chkKeysounds.value = 1 And frmRipList.IsSelected(RIPID) = True Then
                                    RealFreq = ConvertFrequency(((SampB11(SampA11(X).SampleNum + 1).Frequ)))
                                    If bInfo Then
                                        Print #fInfo2, X, SampB11(SampA11(X).SampleNum + 1).Frequ, RealFreq
                                    End If
                        
                        'DECRYPTION SETUP
                        Select Case RipMode
                            Case RipList.CS_2dx14thStyleA, RipList.CS_2dx14thStyleB, RipList.CS_2dx14thStyleC
                                XORType = 2
                            Case RipList.CS_2dx15thStyleA, RipList.CS_2dx15thStyleB, RipList.CS_2dx15thStyleC
                                XORType = 2
                        End Select
                        
                                    fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + fIDString + "\" + BME(X + 0) + ".wav", RealSampOffs, RealFreq, SampB11(SampA11(X).SampleNum + 1).SampLength, SampB11(SampA11(X).SampleNum + 1).ChanCount, , , ((SampA11(X).PanLeft + SampA11(X).PanRight) \ 2), SampA11(X).volume, , XOR11, , XORType)
                                End If
                                DrawProgress 0, X / SampB11Head.SampleCount
                                If bRipping = False Then
                                    Exit For
                                End If
                            End If
                        Next X
                        If bInfo Then
                            Close #fInfo2
                        End If
                    End If
                    InFile.offset = CurrentOffset + &H8000&
                    RIPID = RIPID + 1
                    picProgress.Cls
                    
                ElseIf RipMode = RipList.CS_Popn11 Then
                
                'POP'N 11+ CHARTS
                
                    If FBuff(0) = &H20 And FBuff(1) = 0 And FBuff(2) = 0 And FBuff(3) = 0 Then
                        CurrentOffset = InFile.offset
                        InFile.ReadFileObject VarPtr(TableCheck(0)), 4, CurrentOffset + 32
                        If TableCheck(0) = 0 Then
                            InFile.ReadFileObject VarPtr(TableCheck(0)), 32, CurrentOffset
                            Y = 0
                            For X = 0 To 7
                                If TableCheck(X) > Y Then
                                    Y = TableCheck(X)
                                End If
                            Next X
                            Y = Y - 32
                            InFile.QuickExtract OutFolder + CStr(RIPID) + ".cs9", Y, CurrentOffset + 32
                            RIPID = RIPID + 1
                            CurrentOffset = CurrentOffset + ((Y \ &H800) * &H800) + &H800
                        End If
                    End If
                End If
                
            Case RipList.CS_DDRCAT
                If Compare16(FBuff(), BlankLine) Then
                    If FBuff(&H10) <> 0 And (FBuff(&H11) And &HF0) = 0 Then
                        Debug.Print "CS_DDRCAT: rip"; InFile.offset
                        lblProgress = "DDR song " + CStr(RIPID)
                        If chkBGM.value = 1 Then
                            fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + ".wav", InFile.offset, 44100, VAGBlockSize, 2, (chkDDRSilence.value = 1), ReadSize, , , True)
                            RipCount = RipCount + 1
                        End If
                        RIPID = RIPID + 1
                        InFile.offset = InFile.offset + VAGBlockSize
                    End If
                End If
            Case RipList.CS_DDRPS2
                If FBuff(0) = &H53 And FBuff(1) = &H76 And FBuff(2) = &H61 And FBuff(3) = &H67 Then
                    CopyMemory freq, FBuff(8), 4
                    Debug.Print "CS_DDRPS2: rip"; InFile.offset
                    lblProgress = "DDR song " + CStr(RIPID)
                    If chkBGM.value = 1 Then
                        fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + ".wav", InFile.offset + &H800, freq, VAGBlockSize, FBuff(12), (chkDDRSilence.value = 1), ReadSize)
                        RipCount = RipCount + 1
                    End If
                    RIPID = RIPID + 1
                    InFile.offset = InFile.offset + VAGBlockSize
                End If
            Case RipList.CS_DDRPSX
                If Compare16(FBuff(), BlankLine) Then
                    If (FBuff(&H11) And 4) Then
                        Debug.Print "CS_DDRPSX: rip"; InFile.offset
                        lblProgress = "DDR song " + CStr(RIPID)
                        If chkBGM.value = 1 Then
                            fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + ".wav", InFile.offset, 44100, VAGBlockSize + 0, 2, (chkDDRSilence.value = 1), ReadSize)
                            RipCount = RipCount + 1
                        End If
                        RIPID = RIPID + 1
                        InFile.offset = InFile.offset + VAGBlockSize
                    End If
                End If
            Case RipList.CS_DDRSUPERNOVA, RipList.CS_DDRSUPERNOVA2
                If FBuff(0) = 1 And FBuff(1) = 0 And FBuff(2) = &H64 And FBuff(3) = 8 Then
                    CurrentOffset = InFile.offset
                    InFile.ReadFileObject VarPtr(freq), 4, CurrentOffset + &H18&
                    lblProgress = "SuperNova Song " + CStr(RIPID)
                    Debug.Print "CS_SuperNova: BGM"; RIPID, InFile.offset
                    If chkBGM.value = 1 Then
                        fSize2 = PSXDecode.VAGRipSimple(InFile, OutFolder + CStr(RIPID) + ".wav", CurrentOffset + &H800&, freq, &H10, 2, , 256)
                        RipCount = RipCount + 1
                    End If
                    RIPID = RIPID + 1
                    InFile.offset = CurrentOffset + &H1000&
                End If
            Case RipList.CS_2dxImage
                If InFile.offset >= 16 Then
                    InFile.ReadFileObject FBuff2(), 16, CurrentOffset - 16
                End If
                If Compare16(FBuff2(), BlankLine) Then
                    InFile.ReadFileObject VarPtr(ImgCount), 4, CurrentOffset
                    If ImgCount > 0 And ImgCount < 65536 Then
                        ReDim ImgBuff(0 To ImgCount - 1) As Long
                        InFile.ReadFileObject VarPtr(ImgBuff(0)), ImgCount * 4, CurrentOffset + 16
                        For X = 0 To ImgCount - 1
                            If ImgBuff(X) <> 0 Then
                                PSXDecode.DecodeBemani1 InFile, ImgBuffB(), CurrentOffset + ImgBuff(X)
                                Open OutFolder + CStr(RIPID) + "_" + BME(X + 0) + ".tim" For Binary As #f
                                Put #f, 1, ImgBuffB
                                Close #f
                            End If
                        Next X
                    End If
                End If
            Case RipList.CS_2dxImageOld
                If (FBuff(0) And 15) = 4 And FBuff(1) = 16 And FBuff(2) = 0 And FBuff(3) = 128 Then
                    If (FBuff(4) And 15) = 8 Or (FBuff(4) And 15) = 9 Then
                        f = FreeFile
                        CurrentOffset = InFile.offset
                        'PSXDecode.DecodeBemani1 InFile, ImgBuffB(), currentoffset
                        'Open OutFolder + CStr(RIPID) + "_" + BME(x + 0) + ".tim" For Binary As #f
                        'Put #f, 1, ImgBuffB
                        'Close #f
                        RIPID = RIPID + 1
                    End If
                End If
        End Select
        
        If oldripid <> RIPID Then
            DoEvents
            oldripid = RIPID
        End If
        
        InFile.SeekFileRel SkipSize
        
    
    Loop While InFile.BytesLeft >= SkipSize And bRipping = True
    bRipping = False
    lblProgress.Caption = ""
    picProgress.Cls
    InFile.CloseFile
    Command3.Caption = "Begin"
    If JobMode = False And bStuffRipped = True Then
        MsgBox "Ripping Complete!"
    End If
    hsVolume.Enabled = True
    Frame4.Enabled = True
    Frame7.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    If bInfo Then
        Close #fInfo
    End If
    CloseLog
End Sub

'==================================================================================
'==================================================================================
'==================================================================================
'==================================================================================
'==================================================================================



Private Sub Check1_Click()
    PSXDecode.ThruMode = (Check1.value = 1)
End Sub

Private Sub Check2_Click()
    PSXDecode.AutoMono = (Check2.value = 0)
End Sub

Private Sub chkSimpleRipBGM_Click()
    chkBGM.value = chkSimpleRipBGM.value
End Sub

Private Sub chkSimpleRipCharts_Click()
    chkRip1.value = chkSimpleRipCharts.value
End Sub

Private Sub chkSimpleRipKeys_Click()
    chkKeysounds.value = chkSimpleRipKeys.value
End Sub

Private Sub chkSimpleRipVids_Click()
    chkVideo.value = chkSimpleRipVids.value
End Sub

Private Sub chkUseBGMVol_Click()
    chkUseBGMBoost.Enabled = chkUseBGMVol.value
End Sub

Private Sub chkUseEXE_Click()
    SetRequiredFile
    Command1.Enabled = (chkUseEXE.value = 1)
End Sub

Private Sub cmbRipList_Click()
    SetRequiredFile
    If cmbRipList.ListIndex = RipList.MISC_RAWVAG Then
        frmRawSettings.Show
    Else
        frmRawSettings.Hide
    End If
    bRipValidate = True
End Sub

Private Sub cmbSimpleGame_Click()
    cmbRipList.ListIndex = cmbSimpleGame.ListIndex
End Sub

Private Sub cmdAdvanced_Click()
    mnuModeAdvanced_Click
End Sub

Private Sub cmdBrowseEXE_Click()
    CMD.ShowOpen
    txtEXE.Text = CMD.filename
End Sub

Private Sub cmdInputBrowse_Click()
    CMD.ShowOpen
    txtInput.Text = CMD.filename
End Sub

Private Sub cmdOutputBrowse_Click()
    Dim bInf As BROWSEINFO
    Dim RetVal As Long
    Dim PathID As Long
    Dim RetPath As String
    Dim offset As Integer
    
    bInf.hOwner = Me.hWnd
    bInf.pidlRoot = 0&
    bInf.lpszTitle = "Select a folder to extract files to:"
    bInf.ulFlags = BIF_RETURNONLYFSDIRS
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        offset = InStr(RetPath, Chr$(0))
        txtOutput.Text = Left$(RetPath, offset - 1)
    End If
End Sub

Private Sub Command1_Click()
    Dim f As Long
    Dim X As Long
    Dim cTest As Long
    If txtEXE.Text = "" Then
        Exit Sub
    End If
    If chkUseEXE.value = 0 Then
        MsgBox "Please make sure the checkbox next to the additional file entry is CHECKED.", vbCritical, "Oops."
        Exit Sub
    End If
    If bRipValidate Then
        f = FreeFile
        Open txtEXE.Text For Binary As #f
        Get #f, 1, cTest
        If cTest <> &H464C457F Then
            MsgBox "A list could not be created because this file is not a valid ELF executable.", vbCritical, "Oops."
            Close #f
            Exit Sub
        End If
        Close #f
        OffsetList.SetListFile txtEXE.Text
        OffsetList.ReadFileList EXETableType, EXETableOffset
        OffsetList.ReadInfoList EXESongType, EXESongOffset, DataNumberAdjust, (chkRip1.value = 1)
        frmRipList.List1.Clear
        For X = 0 To OffsetList.ListCount
            frmRipList.List1.AddItem CStr(X + 1000) + " " + OffsetList.ListName(X)
        Next X
        For X = frmRipList.List1.ListCount - 1 To 0 Step -1
            frmRipList.List1.Selected(X) = True
        Next X
        bRipValidate = False
    End If
    frmRipList.Visible = True
End Sub

Private Sub Command2_Click()
    SetRequiredFile
    If chkUseEXE.Enabled = True Then
        txtEXE.Text = Left$(drvSimple.Drive, 2) + "\" + lblEXE.Caption
        chkUseEXE.value = 1
    End If
    Command3_Click
End Sub

Private Sub Command3_Click()
    If Not bRipping Then
        If cmbRipList.ListIndex < 0 Then
            MsgBox "You must first select a game type from the drop down list."
        Else
            If chkUseEXE.value <> 0 And txtEXE.Text = "" Then
                MsgBox "You need to provide the required file to use this feature.", , "Warning"
                Exit Sub
            End If
            If txtInput.Text = "" And mnuModeSimple.Checked = False Then
                MsgBox "Please select the Input file to process.", , "Warning"
                Exit Sub
            End If
            OutFolder = Trim$(txtOutput.Text)
            If OutFolder = "" Then
                OutFolder = txtInput.Text
                If InStr(OutFolder, "\") > 0 Then
                    OutFolder = Left$(OutFolder, InStrRev(OutFolder, "\"))
                End If
            End If
            OutFolder = Replace(OutFolder, "/", "\")
            If Right$(OutFolder, 1) <> "\" Then
                OutFolder = OutFolder + "\"
            End If
            Command3.Caption = "Stop"
            hsVolume.Enabled = False
            Frame4.Enabled = False
            Frame7.Enabled = False
            Frame1.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
            CalculateFreqTable
'            SetRequiredFile
            Rip cmbRipList.ListIndex
        End If
    Else
        If MsgBox("Really end conversion?" + vbCrLf + "Note: you may need to wait for the current conversion to end.", vbYesNo, "Confirmation") = vbYes Then
            PSXDecode.CancelOperation
            Command3.Caption = "Begin"
            bRipping = False
            Frame1.Enabled = True
            Frame2.Enabled = True
            Frame3.Enabled = True
            Frame4.Enabled = True
            Frame7.Enabled = True
            hsVolume.Enabled = True
        End If
    End If
End Sub

Sub DoUnlock()
    bUnlocked = True
    chkRip1.Visible = True
    chkInfo.Visible = True
End Sub

Private Sub Command4_Click()
    cmdOutputBrowse_Click
    txtSimpleOutput.Text = txtOutput.Text
End Sub

Private Sub Form_Load()
    Me.Visible = True
    DoEvents
    DoUnlock
    DoEvents
    SetPriorityClass GetCurrentProcess, IDLE_PRIORITY_CLASS
    CalculateFreqTable
    Call hsVolume_Change
    Me.Caption = Me.Caption + " (" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + ")"
    sMainCaption = Me.Caption
    Me.Show
    PSXDecode.AutoMono = False
    CreateRipList
    'fill Simple Game list
    Dim X As Long
    For X = 0 To cmbRipList.ListCount - 1
        cmbSimpleGame.AddItem cmbRipList.List(X)
        cmbSimpleGame.ItemData(X) = cmbRipList.ItemData(X)
    Next X
    ResetCaption
End Sub

Private Sub CreateRipList()
    With cmbRipList
        .AddItem "AC-DJMain hardware based systems (RAW format ONLY! use CHDMAN)"
        .AddItem "AC-Twinkle hardware based systems (hard drive image)"
        .AddItem "AC-Bemani PC .2DX file"
        .AddItem "AC-Technomotion .TMM file"
        .AddItem "CS-beatmania PSX (BMDATA.PAK)"
        .AddItem "CS-beatmaniaIIDX 3rd style: BM2DX3.BIN"
        .AddItem "CS-beatmaniaIIDX 4th style: BM2DX4.BIN"
        .AddItem "CS-beatmaniaIIDX 5th style: BM2DX5.BIN"
        .AddItem "CS-beatmaniaIIDX 6th style A: BM2DX6A.BIN"
        .AddItem "CS-beatmaniaIIDX 6th style B: BM2DX6B.BIN"
        .AddItem "CS-beatmaniaIIDX 7th style A: BM2DX7A.BIN"
        .AddItem "CS-beatmaniaIIDX 7th style B: BM2DX7B.BIN"
        .AddItem "CS-beatmaniaIIDX 7th style C: BM2DX7C.BIN"
        .AddItem "CS-beatmaniaIIDX 8th style A: BM2DX8A.BIN"
        .AddItem "CS-beatmaniaIIDX 8th style B: BM2DX8B.BIN"
        .AddItem "CS-beatmaniaIIDX 8th style C: BM2DX8C.BIN"
        .AddItem "CS-beatmaniaIIDX 9th style: DATA2.BIN"
        .AddItem "CS-beatmaniaIIDX 10th style: DATA2.BIN"
        .AddItem "CS-beatmaniaIIDX 11th style RED: DATA2.BIN"
        .AddItem "CS-beatmaniaIIDX 12th style HAPPYSKY A: BM2DX12A.DAT"
        .AddItem "CS-beatmaniaIIDX 12th style HAPPYSKY B: BM2DX12B.DAT"
        .AddItem "CS-beatmaniaIIDX 12th style HAPPYSKY C: BM2DX12C.DAT"
        .AddItem "CS-beatmaniaIIDX 13th style DistorteD A: BM2DX13A.DAT"
        .AddItem "CS-beatmaniaIIDX 13th style DistorteD B: BM2DX13B.DAT"
        .AddItem "CS-beatmaniaIIDX 13th style DistorteD C: BM2DX13C.DAT"
        .AddItem "CS-beatmaniaIIDX 14th style GOLD A: BM2DX14A.DAT"
        .AddItem "CS-beatmaniaIIDX 14th style GOLD B: BM2DX14B.DAT"
        .AddItem "CS-beatmaniaIIDX 14th style GOLD C: BM2DX14C.DAT"
        .AddItem "CS-beatmaniaIIDX 15th style DJ TROOPERS A: BM2DX15A.DAT"
        .AddItem "CS-beatmaniaIIDX 15th style DJ TROOPERS B: BM2DX15B.DAT"
        .AddItem "CS-beatmaniaIIDX 15th style DJ TROOPERS C: BM2DX15C.DAT"
        .AddItem "CS-beatmania USA: DATA2.DAT"
        .AddItem "CS-beatmania USA (JAMPACK): DATA2.DAT"
        .AddItem "CS-Dance Dance Revolution ULTRAMIX (*.XBE)"
        .AddItem "CS-Dance Dance Revolution ULTRAMIX Package (*.HBN)"
        .AddItem "CS-Dance Dance Revolution (PSX STR.BIN)"
        .AddItem "CS-Dance Dance Revolution (PSX *.CAT)"
        .AddItem "CS-Dance Dance Revolution (PS2 FILEDATA.BIN)"
        .AddItem "CS-Dance Dance Revolution SUPERNOVA (MDB_SN1.DAT)"
        .AddItem "CS-Dance Dance Revolution SUPERNOVA 2 (MDB_SN1.DAT)"
        .AddItem "CS-Dance Dance Revolution X (MDB_X1.DAT)"
        .AddItem "CS-Dance Dance Revolution UNIVERSE (xbox360 ISO image) " + "[incomplete]"
        .AddItem "CS-Pop'n Music PSX (POPDATA.PAK)"
        .AddItem "CS-Pop'n Music PS2 early ()"
        .AddItem "CS-Pop'n Music PS2 11+ ()"
        .AddItem "CS-Taiko No Tatsujin 7 ()"
        .AddItem "Generic Audio: Raw VAG (see options)"
        .AddItem "Generic Audio: " + "[incomplete]" + " AFS audio bank"
        .AddItem "Extreme-G 3: XA2 (XG3/PS2)"
        .AddItem "Grandia 3: GR3_STR.STZ"
        .AddItem "XGRA: XG4PS2 (XGRA/PS2)"
        .AddItem "Unreal Tournament (PS2 *.VGM)"
        'If bUnlocked Then
        '    .AddItem "[incomplete]" + " CS-beatmaniaIIDX" + " T" + "I" + "M" + " " + "(7th-8th " + "BM2DX" + "#A)"
        '    .AddItem "[incomplete]" + " CS-beatmaniaIIDX" + " T" + "I" + "M" + " (3rd-6th " + "SLPM)"
        '    .AddItem "[" + "adva" + "nced" + "]" + " CS-beatmaniaIIDX" + " C" + "S" + " (3rd-6th " + "SLPM)"
        'End If
    End With
End Sub

Sub SetRequiredFile()
    Dim rf As String
    rf = ""
    EXEHasCSFiles = False
    EXETitleType = 0
    EXETitleOffset = 0
    EXETableOffset = -1
    EXETableType = -1
    EXESongOffset = -1
    EXESongType = -1
    DataNumberAdjust = 0
    Select Case cmbRipList.ListIndex
        Case RipList.CS_2dx03rdStyle
            rf = "SLPM_650.06"
            EXETableOffset = 1334480
            EXESongOffset = 491464
            EXETableType = 4
            EXESongType = 7
        Case RipList.CS_2dx04thStyle
            rf = "SLPM_650.26"
            EXETableOffset = 1274960
            EXESongOffset = 572568
            EXETableType = 4
            EXESongType = 10
        Case RipList.CS_2dx05thStyle
            rf = "SLPM_650.49"
            EXETableOffset = 1587160
            EXESongOffset = 714016
            EXETableType = 1
            EXESongType = 1
        Case RipList.CS_2dx06thStyleA
            rf = "SLPM_651.56"
            EXETableOffset = 1572952
            EXESongOffset = 1607096
            EXETableType = 1
            EXESongType = 8
            EXEHasCSFiles = True
        Case RipList.CS_2dx06thStyleB
            rf = "SLPM_651.56"
            EXETableOffset = 1578472
            EXESongOffset = 1607096
            EXETableType = 1
            EXESongType = 8
        Case RipList.CS_2dx07thStyleB
            rf = "SLPM_655.93"
            EXETableOffset = 1799264
            EXETableType = 2
            EXESongOffset = 1841904
            EXESongType = 2
        Case RipList.CS_2dx07thStyleC
            rf = "SLPM_655.93"
            EXETableOffset = 1808944
            EXETableType = 2
            EXESongOffset = 1841904
            EXESongType = 2
            EXEHasCSFiles = True
        Case RipList.CS_2dx08thStyleB
            rf = "SLPM_657.68"
            EXETableOffset = 1681728
            EXETableType = 2
            EXESongOffset = 1720416
            EXESongType = 3
        Case RipList.CS_2dx08thStyleC
            rf = "SLPM_657.68"
            EXETableOffset = 1683584
            EXETableType = 2
            EXESongOffset = 1720416
            EXESongType = 3
            EXEHasCSFiles = True
        Case RipList.CS_2dx09thStyle
            rf = "SLPM_659.46"
            EXETableOffset = 774704
            EXETableType = 3
            EXESongOffset = 791808
            EXESongType = 4
            EXEHasCSFiles = True
        Case RipList.CS_2dx10thStyle
            rf = "SLPM_661.80"
            EXETableOffset = 842896
            EXETableType = 3
            EXESongOffset = 1096416
            EXESongType = 4
            EXEHasCSFiles = True
        Case RipList.CS_2dx11thStyle
            rf = "SLPM_664.26"
            EXETableOffset = 975936
            EXETableType = 3
            EXESongOffset = 1843696
            EXESongType = 6
            EXEHasCSFiles = True
        Case RipList.CS_2dx12thStyleB
            rf = "SLPM_666.21"
            EXETableOffset = 1069888
            EXETableType = 3
            EXESongOffset = 1138448
            EXESongType = 9
            EXEHasCSFiles = True
            DataNumberAdjust = 28
        Case RipList.CS_2dx12thStyleC
            rf = "SLPM_666.21"
            EXETableOffset = 1075776
            EXETableType = 3
            EXESongOffset = 1138448
            EXESongType = 9
            DataNumberAdjust = 764
        Case RipList.CS_2dx13thStyleB 'DD A-file starts at 1124864 and DATA1 at 1145968 both format 3
            rf = "SLPM_668.28"
            EXETableOffset = 1125056
            EXETableType = 3
            EXESongOffset = 1266656
            EXESongType = 11
            DataNumberAdjust = 24
            EXEHasCSFiles = True
        Case RipList.CS_2dx13thStyleC
            rf = "SLPM_668.28"
            EXETableOffset = 1131456
            EXETableType = 3
            EXESongOffset = 1266656
            EXESongType = 11
            DataNumberAdjust = 824
            EXEHasCSFiles = True
        Case RipList.CS_2dx14thStyleA
            'rf = "SLPM_669.95"
            'EXETableOffset = 1158496
            'EXETableType = 4
            'DataNumberAdjust = 0
        Case RipList.CS_2dx14thStyleB
            rf = "SLPM_669.95"
            EXETableOffset = 1158496
            EXETableType = 4
            EXESongOffset = 1401600
            EXESongType = 12
            DataNumberAdjust = 0
            EXEHasCSFiles = True
        Case RipList.CS_2dx14thStyleC
            rf = "SLPM_669.95"
            EXETableOffset = 1168972
            EXETableType = 4
            EXESongOffset = 1401600
            EXESongType = 12
            DataNumberAdjust = 873
            EXEHasCSFiles = True
        Case RipList.CS_2dx15thStyleA
            rf = "SLPM_551.17"
            EXETableOffset = 1261600
            EXETableType = 4
            EXESongOffset = 1506912
            EXESongType = 13
            DataNumberAdjust = 0
            'EXEHasCSFiles = True
        Case RipList.CS_2dx15thStyleB
            rf = "SLPM_551.17"
            EXETableOffset = 1262176    'A is 1261600
            EXETableType = 4
            EXESongOffset = 1506912
            EXESongType = 13
            DataNumberAdjust = 48
            EXEHasCSFiles = True
        Case RipList.CS_2dx15thStyleC
            rf = "SLPM_551.17"
            EXETableOffset = 1272352
            EXETableType = 4
            EXESongOffset = 1506912
            EXESongType = 13
            DataNumberAdjust = 896
            EXEHasCSFiles = True
        Case RipList.CS_2dx01stStyleUSA
            rf = "SLUS_212.39"
            EXETableOffset = 763664
            EXETableType = 3
            EXESongOffset = 783632
            EXESongType = 5
            EXEHasCSFiles = True
        Case RipList.CS_2dx01stStyleUSA_JAMPACK
            rf = "BEATM\BEATM.ELF"
            EXETableOffset = 771552
            EXETableType = 3
            EXESongOffset = 787088
            EXESongType = 5
            EXEHasCSFiles = True
        Case RipList.CS_DDRSUPERNOVA
            rf = ""
        Case RipList.CS_DDRSUPERNOVA2
            rf = "SLUS_216.08"
            EXESongOffset = 2038944
            EXESongType = 100
        Case RipList.CS_DDRX
            rf = "SLUS_217.67"
            EXESongOffset = 2626512
            EXESongType = 101
        Case RipList.CS_PopnPSX
            rf = ""
            EXETableOffset = 0
            EXETableType = 5
        Case RipList.CS_beatmania
            rf = ""
            EXETableOffset = 0
            EXETableType = 5
        Case RipList.MISC_GRANDIA3
            rf = "GR3_STR.IDX"
    End Select
    txtEXE.Enabled = (rf <> "")
    chkUseEXE.Enabled = txtEXE.Enabled
    If chkUseEXE.Enabled = False Then
        chkUseEXE.value = 0
    End If
    'If chkUseEXE.Value = 0 Then
    '    rf = ""
    'End If
    lblEXE.Caption = rf
End Sub

Private Function Compare16(in16() As Byte, str16 As String) As Boolean
    Dim tstrn As String
    tstrn = StrConv(in16(), vbUnicode)
    If Left$(str16, 16) = Left$(tstrn, 16) Then
        Compare16 = True
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If bRipping Then
        If MsgBox("There is a ripping operation in progress!" + vbCrLf + "Really quit?", vbYesNo, "Confirmation") = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bRipping = False
    PSXDecode.CancelOperation
    If bBassInit Then
        BASS_Free
    End If
    DoEvents
    End
End Sub

Private Sub hsVolume_Change()
    Frame8.Caption = "Conversion Volume: " + CStr((hsVolume.value / hsVolume.max) * 100) + "%"
    PSXDecode.ConvertVolume = (hsVolume.value / hsVolume.max) * 100
    AC2DXDecoder.SetVolume hsVolume.value / hsVolume.max
End Sub

Private Sub hsVolume_Scroll()
    hsVolume_Change
End Sub

Private Sub hsVolumeSimple_Change()
    hsVolume.value = hsVolumeSimple.value
End Sub

Private Sub hsVolumeSimple_Scroll()
    hsVolume.value = hsVolumeSimple.value
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuJobQueue_Click()
    frmJobs.Show
End Sub

Private Sub mnuModeAdvanced_Click()
    If mnuModeAdvanced.Checked = False Then
        mnuModeSimple.Checked = False
        mnuModeAdvanced.Checked = True
        frmSimple.Visible = False
        frmAdvanced.Visible = True
        ResetCaption
    End If
End Sub

Private Sub ResetCaption()
    Me.Caption = sMainCaption
    If mnuModeAdvanced.Checked = True Then
        Me.Caption = Me.Caption + " [advanced mode]"
    Else
        Me.Caption = Me.Caption + " [simple mode]"
    End If
End Sub

Private Sub mnuModeSimple_Click()
    If mnuModeSimple.Checked = False Then
        mnuModeSimple.Checked = True
        mnuModeAdvanced.Checked = False
        frmSimple.Visible = True
        frmAdvanced.Visible = False
        ResetCaption
    End If
End Sub

Private Sub Timer1_Timer()
    If Me.Visible = False Then
        End
    End If
    Command2.Caption = Command3.Caption
    Command2.Enabled = Command3.Enabled
    Frame14.Caption = Frame8.Caption
End Sub

Private Sub txtEXE_Change()
    If Not bRipValidate Then
        bRipValidate = True
        frmRipList.List1.Clear
    End If
End Sub

Private Sub txtEXE_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtEXE.Text = data.Files(1)
    bRipValidate = True
End Sub

Private Sub txtInput_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInput = data.Files(1)
End Sub

Private Sub DrawProgress(ByVal prog As Single, ByVal prog2 As Single)
    If prog > 1 Then prog = 1
    If prog2 > 1 Then prog2 = 1
    picProgress.Line (0, 0)-(prog, 1), vbBlue, BF
    picProgress.Line (0, 1)-(prog2, 2), vbRed, BF
End Sub

Private Function BME(inval As Integer) As String
    Dim BMEString As String
    BMEString = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    BME = Mid$(BMEString, (inval \ 36) + 1, 1) + Mid$(BMEString, (inval Mod 36) + 1, 1)
End Function

Private Function GetFileNameOf(inPath As String) As String
    Dim X As String
    If InStr(inPath, "\") = 0 Then
        GetFileNameOf = inPath
        Exit Function
    End If
    X = Mid$(inPath, InStrRev(inPath, "\") + 1)
    GetFileNameOf = X
End Function

Public Function SimulateJob()
    JobMode = True
    Command3_Click
    JobMode = False
End Function

Private Sub txtSimpleOutput_Change()
    txtOutput.Text = txtSimpleOutput.Text
End Sub

