VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BeMedia Player"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Visual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2880
      TabIndex        =   28
      Top             =   120
      Width           =   2535
      Begin VB.PictureBox picVis 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   8
         TabIndex        =   29
         Top             =   240
         Width           =   2295
         Begin VB.Timer tmrVis 
            Interval        =   1
            Left            =   240
            Top             =   240
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   5160
      Width           =   5295
   End
   Begin VB.Frame Frame5 
      Caption         =   "Volume"
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
      TabIndex        =   13
      Top             =   1080
      Width           =   2655
      Begin VB.HScrollBar hsVolume 
         Height          =   255
         LargeChange     =   20
         Left            =   120
         Max             =   100
         TabIndex        =   14
         Top             =   240
         Value           =   80
         Width           =   2415
      End
   End
   Begin VB.CheckBox chkDiskRecord 
      Caption         =   "Record to Disk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   5295
      Begin VB.TextBox txtRecord 
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Text            =   "output.wav"
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "wav"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "(Uses default recording device)"
         Height          =   225
         Left            =   2640
         TabIndex        =   12
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   5295
      Begin VB.CommandButton cmdFramePreset 
         Caption         =   "Troopers"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdFramePreset 
         Caption         =   "GOLD"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   25
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdFramePreset 
         Caption         =   "9th-DD"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkInverse 
         Caption         =   "rate"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         ToolTipText     =   "If checked this will be the multiplier. If unchecked this will be the frame rate."
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtFrame 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "PRESETS (machine type)"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Frame mult (blank=default)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      Begin VB.Timer tmrInfo 
         Interval        =   1
         Left            =   120
         Top             =   240
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         Caption         =   "playframe"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "playtime"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   5295
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmPlayer.frx":0000
         Left            =   3960
         List            =   "frmPlayer.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   990
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   18
         Top             =   600
         Width           =   4335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmPlayer.frx":0053
         Left            =   840
         List            =   "frmPlayer.frx":0075
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   990
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "file type"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Chart#"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "soundfile"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "chart"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ChanLevels(0 To 7) As Single
Private Decoder As New cls2DXDecode
Private SoundFile As New clsFileStream
Private bUseSoundFile As Boolean
Private bRecording As Boolean
Private hRecord As Long
Private hMixer As Long
Private xVolume As Long
Private sLastSoundFile As String

Private Sub cmdFramePreset_Click(index As Integer)
    Select Case index
        Case 0 '9th-DD
            chkInverse.value = 0
            txtFrame = "59.94"
        Case 1 'GOLD
            chkInverse.value = 0
            txtFrame = "60.04"
        Case 2 'Troopers
            chkInverse.value = 1
            txtFrame = "1"
    End Select
End Sub

Private Sub Command1_Click()
    ProcessClick
End Sub

Sub ProcessClick()
    Dim x As Long
    Dim yfile As String
    Dim bLoadNew As Boolean
    If InStr(Text1.Text, "\") > 0 Then
        yfile = Left$(Text1.Text, InStrRev(Text1.Text, "\"))
    End If
    If Command1.Caption = "Play" Then
        If Combo1.ListIndex = 0 Then
            MsgBox "Select a chart type.", vbInformation, "BeMedia Player"
            Exit Sub
        End If
        If sLastSoundFile <> Text2 Or Text2 = "" Then
            BASS_Free
            If bUseSoundFile Then
                SoundFile.CloseFile
                bUseSoundFile = False
                Decoder.Reset
                For x = 1 To 1295
                    SampleHandles(x) = 0
                Next x
            End If
            DoEvents
            If BASS_Init(1, 44100, 0, Me.hWnd, ByVal 0&) = BASSFALSE Then
                MsgBox "BASS23 could not be loaded.", vbCritical, "Doh."
                BASS_Free
                Exit Sub
            End If
            sLastSoundFile = Text2
            bLoadNew = True
        Else
            bLoadNew = False
        End If
        'hMixer = BASS_Mixer_StreamCreate(44100, 2, BASS_MIXER_NONSTOP)
        BASS_SetConfig BASS_CONFIG_GVOL_SAMPLE, xVolume
        BASS_Start
        If txtFrame.Text = "" Then
            txtFrame.Text = "0"
        End If
        If txtFrame.Text = "0" And chkInverse.value = 1 Then
            txtFrame.Text = "59.94"
        End If
        If chkInverse.value = 1 Then
            CS.LoadFile Text1.Text, Combo1.ListIndex, Combo3.ListIndex, 1000 / CDbl(txtFrame.Text)
        Else
            CS.LoadFile Text1.Text, Combo1.ListIndex, Combo3.ListIndex, CDbl(txtFrame.Text)
        End If
        If CS.GetSimFileData(VarPtr(NextEvent), 3) = False Then
            MsgBox "The simfile data could not be loaded.", vbCritical, "Doh."
            BASS_Free
            Exit Sub
        End If
        CS.GetSimFileData VarPtr(NextEvent), 1
        Command1.Enabled = False
        Command1.Caption = "Loading..."
        If Text2.Text = "" Then
            For x = 1 To 1295
                If CS.KeysoundFileName(x) <> "" Then
                    SampleHandles(x) = BASS_SampleLoad(BASSFALSE, yfile + CS.KeysoundFileName(x), 0, 0, 1, BASS_SAMPLE_OVER_POS)
                    If SampleHandles(x) = 0 Then
                        SampleHandles(x) = BASS_SampleLoad(BASSFALSE, yfile + Left$(CS.KeysoundFileName(x), 2) + ".ogg", 0, 0, 1, BASS_SAMPLE_OVER_POS)
                    End If
                End If
            Next x
        Else
            If bLoadNew Then
                bUseSoundFile = True
                SoundFile.OpenFile Text2.Text, True, False
                Decoder.Decrypt SoundFile
                For x = 1 To 1295
                    SampleHandles(x) = Decoder.SampleHandle(x)
                Next x
            End If
        End If
        bRecording = (chkDiskRecord.value = 1)
        If bRecording Then
            If BASS_RecordInit(-1) = BASSTRUE Then
                hRecord = BASS_RecordStart(44100, 2, 0, AddressOf RECORDPROC1, 0)
                StartRecordFile txtRecord
            End If
        End If
        CurrentEvent = 1
        bPlaying = True
        Command1.Caption = "Stop"
        Command1.Enabled = True
        Sleep 1
        Sleep 0
        DoEvents
        PlayTimer.Init
    Else
        bPlaying = False
        For x = 1 To 1295
            BASS_SampleStop SampleHandles(x)
        Next x
        Command1.Caption = "Play"
        CurrentMS = 0
        CurrentFrame = 0
        BASS_Pause
        chkDiskRecord.value = 0
        If bRecording Then
            BASS_RecordFree
            StopRecordFile
        End If
    End If
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
    hsVolume_Change
    If App.Major = 0 Then
        Me.Caption = Me.Caption + " [beta " + CStr(App.Revision) + "]"
    End If
    If BASS_RecordInit(-1) = BASSFALSE Then
        MsgBox "Could not init recording system."
    End If
    BASS_RecordFree
    Combo3.ListIndex = 0
    Decoder.EnablePlayer True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BASS_Free
    bRunning = False
End Sub

Private Sub hsVolume_Change()
    Frame5.Caption = "Volume: " + CStr(hsVolume.value) + "%"
    xVolume = hsVolume.value
    If bPlaying Then
        BASS_SetConfig BASS_CONFIG_GVOL_SAMPLE, xVolume
    End If
End Sub

Private Sub hsVolume_Scroll()
    hsVolume_Change
End Sub

Private Sub picVis_Click()
    tmrVis.Enabled = (Not tmrVis.Enabled)
    If tmrVis.Enabled Then
        Frame6.Caption = "Visual"
    Else
        Frame6.Caption = "Visual (disabled)"
        picVis.Cls
    End If
End Sub

Private Sub Text1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Text1.Text = data.Files(1)
End Sub

Private Sub Text2_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Text2.Text = data.Files(1)
End Sub

Private Sub tmrInfo_Timer()
    Dim xSec As String
    If CurrentMS >= 0 Then
        xSec = CStr(((CurrentMS \ 1000) Mod 60))
        If Len(xSec) = 1 Then
            xSec = "0" + xSec
        End If
        lblTime = "Time: " + CStr(CurrentMS \ 60000) + ":" + xSec
        lblFrame = "Frame: " + CStr(CurrentFrame) + " (" + CStr(CS.FrameTiming) + ")"
    Else
        lblTime = "- Warming Up -"
        lblFrame = CStr(Int(CurrentMS))
    End If
End Sub

Private Sub tmrVis_Timer()
    Dim x As Long
    picVis.Cls
    For x = LBound(ChanLevels) To UBound(ChanLevels)
        Select Case x
            Case 0, 2, 4, 6
                picVis.Line (x, 1 - ChanLevels(x))-(x + 1, 1), &HAAAAAA, BF
            Case 1, 3, 5
                picVis.Line (x, 1 - ChanLevels(x))-(x + 1, 1), vbBlue, BF
            Case 7
                picVis.Line (x, 1 - ChanLevels(x))-(x + 1, 1), vbRed, BF
        End Select
        ChanLevels(x) = ChanLevels(x) * 0.8
    Next x
    picVis.Refresh
End Sub
Public Sub PlayChan(ByVal x As Long)
    If x >= 64 Or x < 0 Then
        Exit Sub
    End If
    x = (x Mod 32)
    If x >= LBound(ChanLevels) And x <= UBound(ChanLevels) Then
        ChanLevels(x) = 1
    End If
End Sub

