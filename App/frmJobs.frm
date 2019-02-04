VERSION 5.00
Begin VB.Form frmJobs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Job Queue"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5670
   Icon            =   "frmJobs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Description"
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5655
      Begin VB.Label lblDescription 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdBeginQueue 
      Caption         =   "Begin Queue"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemoveJob 
      Caption         =   "Remove Job"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddJob 
      Caption         =   "Add Job"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ListBox lstJobs 
      Height          =   2595
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
   End
End
Attribute VB_Name = "frmJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type JobDescription
    InFile As String
    OutFolder As String
    GameType As Long
    bUseExe As Boolean
    ExeFile As String
    bKeysounds As Boolean
    bBGM As Boolean
    bVideo As Boolean
    bCS As Boolean
    bInfo As Boolean
    bAlwaysStereo As Boolean
    bDDRSilence As Boolean
    bDebugThru As Boolean
    RipVolume As Long
    RawOffset As String
    RawChannels As String
    RawBlockSize As String
    RawFrequency As String
End Type

Dim Jobs() As JobDescription

Private Sub cmdAddJob_Click()
    With Jobs(UBound(Jobs))
        .bAlwaysStereo = (frmMain.Check2.value = 1)
        .bBGM = (frmMain.chkBGM.value = 1)
        .bCS = (frmMain.chkRip1.value = 1)
        .bDDRSilence = (frmMain.chkDDRSilence.value = 1)
        .bDebugThru = (frmMain.Check1.value = 1)
        .bInfo = (frmMain.chkInfo.value = 1)
        .bKeysounds = (frmMain.chkKeysounds.value = 1)
        .bUseExe = (frmMain.chkUseEXE.value = 1)
        .bVideo = (frmMain.chkVideo.value = 1)
        .ExeFile = frmMain.txtEXE.Text
        .GameType = frmMain.cmbRipList.ListIndex
        .InFile = frmMain.txtInput.Text
        .OutFolder = frmMain.txtOutput.Text
        .RawBlockSize = frmRawSettings.Text1(2)
        .RawChannels = frmRawSettings.Text1(1)
        .RawFrequency = frmRawSettings.Text1(3)
        .RawOffset = frmRawSettings.Text1(0)
        .RipVolume = frmMain.hsVolume.value
        lstJobs.AddItem .InFile
    End With
    ReDim Preserve Jobs(0 To UBound(Jobs) + 1) As JobDescription
End Sub

Private Sub cmdBeginQueue_Click()
    Dim X As Long
    Dim Y As Long
    X = lstJobs.ListCount - 1
    lstJobs.Enabled = False
    For Y = 0 To X
        With Jobs(Y)
            lblDescription.Caption = "Now Processing: " + CStr(Y) + vbCrLf + .InFile
            frmMain.Check2.value = Abs(.bAlwaysStereo)
            frmMain.chkBGM.value = Abs(.bBGM)
            frmMain.chkRip1.value = Abs(.bCS)
            frmMain.chkDDRSilence.value = Abs(.bDDRSilence)
            frmMain.Check1.value = Abs(.bDebugThru)
            frmMain.chkInfo.value = Abs(.bInfo)
            frmMain.chkKeysounds.value = Abs(.bKeysounds)
            frmMain.chkUseEXE.value = Abs(.bUseExe)
            frmMain.chkVideo.value = Abs(.bVideo)
            frmMain.txtEXE.Text = .ExeFile
            frmMain.cmbRipList.ListIndex = .GameType
            frmMain.txtInput.Text = .InFile
            frmMain.txtOutput.Text = .OutFolder
            frmMain.hsVolume.value = .RipVolume
            frmRawSettings.Text1(2) = .RawBlockSize
            frmRawSettings.Text1(1) = .RawChannels
            frmRawSettings.Text1(3) = .RawFrequency
            frmRawSettings.Text1(0) = .RawOffset
            frmMain.SimulateJob
        End With
    Next Y
    lstJobs.Enabled = True
    If MsgBox("Job list complete!" + vbCrLf + "Clear the list?", vbYesNo, "Success") = vbYes Then
        lstJobs.Clear
        ReDim Jobs(0) As JobDescription
    End If
End Sub

Private Sub cmdRemoveJob_Click()
    Dim X As Long
    If lstJobs.ListIndex > -1 Then
        For X = lstJobs.ListIndex To lstJobs.ListCount - 2
            Jobs(X) = Jobs(X + 1)
        Next X
        lstJobs.RemoveItem lstJobs.ListIndex
        ReDim Preserve Jobs(0 To UBound(Jobs)) As JobDescription
    End If
End Sub

Private Sub Form_Load()
    ReDim Jobs(0) As JobDescription
    lstJobs.Clear
    Me.Height = 4455
End Sub

Private Sub lstJobs_Click()
    If lstJobs.ListIndex > -1 Then
        With Jobs(lstJobs.ListIndex)
            lblDescription = "Output Folder: " + .OutFolder + vbCrLf + "Game Type: " + frmMain.cmbRipList.List(.GameType)
        End With
    End If
End Sub

Private Sub lstJobs_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim xx As Long
    Dim f As Long
    Dim s As String
    If frmMain.cmbRipList.ListIndex = -1 Then
        MsgBox "You must specify the game type (and executable file, if you choose) before adding a job."
        Exit Sub
    End If
    If frmMain.txtOutput = "" Then
        MsgBox "You must specify the target folder before adding a job."
        Exit Sub
    End If
    If data.Files.Count = 1 And Right$(UCase$(data.Files(1)), 4) = ".TXT" Then
        f = FreeFile
        Open data.Files(1) For Input As #f
        Do While Not EOF(f)
            Line Input #f, s
            frmMain.txtInput.Text = s
            cmdAddJob_Click
        Loop
    Else
        For xx = 1 To data.Files.Count
            frmMain.txtInput.Text = data.Files(xx)
            cmdAddJob_Click
        Next xx
        If data.Files.Count > 1 Then
            MsgBox CStr(data.Files.Count) + " files added.", vbInformation, "Information"
        End If
    End If
End Sub
