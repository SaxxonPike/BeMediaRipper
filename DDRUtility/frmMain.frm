VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DDR Utility"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Quick File conversion (steps, .pBav, .tim)"
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
      TabIndex        =   21
      Top             =   120
      Width           =   4455
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Simply drag and drop your file(s) onto this text..."
         Height          =   255
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   22
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Output Folder (must exist)"
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
      TabIndex        =   17
      Top             =   1920
      Width           =   4455
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "A folder where extracted data will go. This folder MUST exist first!"
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   4455
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4215
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
      Height          =   1575
      Left            =   2280
      TabIndex        =   12
      Top             =   2640
      Width           =   2295
      Begin VB.CheckBox chkConvertUnref 
         Caption         =   "and convert (unref)"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         ToolTipText     =   $"frmMain.frx":014A
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkSolo 
         Caption         =   "SOLO step mode"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Check this if you are using files from DDR Solo games."
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox chkDelTemps 
         Caption         =   "Delete temp files"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Delete source files after conversion, only if it was successful."
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkAllFiles 
         Caption         =   "then check everything"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "Check for non-referenced files (note that they won't always be complete)"
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox chkFileList 
         Caption         =   "AC: check file list first"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "ALWAYS leave this checked unless you experience problems"
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Types"
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
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
      Begin VB.CheckBox chkConvertBMP 
         Caption         =   "-> BMP"
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         ToolTipText     =   "Convert all TIM images to BMP"
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Data"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Extra data of unknown format"
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkConvertSM 
         Caption         =   "-> SM"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "Convert all step formats to StepMania's SM"
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkConvertWAV 
         Caption         =   "-> WAV"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         ToolTipText     =   "Convert all VAG audio to WAV"
         Top             =   600
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkSteps 
         Caption         =   "Steps"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Step Data (step, SSQ)"
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "Sound"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "VAG-format audio"
         Top             =   600
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkTIM 
         Caption         =   ".TIM"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   ".TIM playstation images"
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source Files (drag && drop enabled)"
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
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      Begin VB.TextBox txtInFile 
         Height          =   285
         Index           =   1
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         ToolTipText     =   "CARD.DAT if applicable"
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtInFile 
         Height          =   285
         Index           =   0
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         ToolTipText     =   "Put the GAME.DAT here. A 16MB file that looks similar to ""GN895JAA.DAT"" works too."
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "CARD"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   630
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "GAME"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
'
' DDR Utility
'  Copyright © SaxxonPike 2oo8
'  for the sole use of recovering "the steps that time forgot"
'
' DANCE DANCE REVOLUTION and BEMANI are trademarks of Konami Corporation.
' Konami is a registered trademark of Konami Corporation.
'
' ************************************************************************
' "DDR Utility" is free software: you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation, either version 3 of the License, or
'  (at your option) any later version.
'
' "DDR Utility" is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
'  along with this program.  If not, see <http://www.gnu.org/licenses/>.
' ************************************************************************

Private Const GAMETYPE_UNKNOWN = 0
Private Const GAMETYPE_DDRAC_OLD = 1
Private Const GAMETYPE_DDRAC_NEW = 2
Private Const GAMETYPE_DDRCS_OLD = 3
Private Const GAMETYPE_DDRCS_NEW = 4
Private Const GAMETYPE_DDRCS_IMAGEDAT = 5

Private Const FORMAT_NOTRIPPED = -1
Private Const FORMAT_UNKNOWN = 0
Private Const FORMAT_TIM = 1                'playstation image
Private Const FORMAT_PBAV = 2               'audio (old announcers)
Private Const FORMAT_STEP = 3               'step format used in 2nd-4th mix
Private Const FORMAT_SSQ = 4                'modern DDR step format
Private Const FORMAT_STEP2 = 5              'old DDR 1st mix step format
Private Const FORMAT_TIM2 = 6               'new graphics format
Private Const FORMAT_ANNOUNCER = 7          'announcer set

Private bMarkUnref As Boolean

Private Sub chkFileList_Click()
    If chkFileList.value = 0 Then
        chkAllFiles.Enabled = False
        chkAllFiles.value = 1
    Else
        chkAllFiles.Enabled = True
    End If
    chkConvertUnref.Enabled = chkAllFiles.Enabled
End Sub

Private Sub cmdExecute_Click()
    DoRip
End Sub

Private Sub Form_Load()
    'add version info to window title
    Me.Caption = Me.Caption + " (rev." + CStr(App.Revision) + ")"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim StepDecoder As New clsSSQ
    Dim AudioDecoder As New clsDeVAG
    Dim ImageDecoder As New clsTIM
    Dim x1 As Long
    cmdExecute.Visible = False
    For x1 = 1 To Data.Files.Count
        Label3 = Data.Files(x1)
        DoEvents
        If InStr(Data.Files(x1), ".") > 0 Then
            Select Case UCase$(Mid$(Data.Files(x1), InStrRev(Data.Files(x1), ".") + 1))
                Case "SSQ"
                    StepDecoder.ConvertSSQ Data.Files(x1), Data.Files(x1) + ".sm"
                Case "STEP"
                    StepDecoder.ConvertStep1 Data.Files(x1), Data.Files(x1) + ".sm", (chkSolo.value = 1)
                Case "STEP2"
                    StepDecoder.ConvertStep2 Data.Files(x1), Data.Files(x1) + ".sm"
                Case "PBAV"
                    AudioDecoder.VAGtoWAV Data.Files(x1), Data.Files(x1) + ".wav", 3120, -1, 44100
                Case "TIM"
                    ImageDecoder.TIMtoBMP Data.Files(x1), Data.Files(x1) + Chr$(1) + ".bmp"
                Case "TIM2", "TM2"
                    
                Case "SOUNDS"
                    AudioDecoder.ExtractSoundPack Data.Files(x1), Data.Files(x1), 0, -1
            End Select
        End If
    Next x1
    cmdExecute.Visible = True
End Sub

Private Sub txtInFile_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    txtInFile(Index) = Data.Files(1)
End Sub

Private Sub DoRip()
    'loop counters
    Dim x As Long
    Dim y As Long
    Dim z As Long
    'ripping
    Dim rf As Long
    Dim bu As Long
    Dim lFileTable(0 To 3) As Long
    Dim lReadOffset As Long
    Dim lRipOffset As Long
    Dim lParms(0 To 7) As Long
    Dim lGameType As Long
    Dim sOutFolder As String
    'BlocksOccupied is used to determine which blocks of arcade data
    'haven't been refereced by the file table - this is handy for
    'speeding up the search for "lost" data and allows us to mark
    'data that's been ripped outside the table as "unreferenced"
    Dim lBlocksOccupied() As Long
    
    'file verification ---
    If txtInFile(0) = "" Then
        MsgBox "Please select a GAME file to use."
        Exit Sub
    End If
    
    Dim inFile As New clsFileStream
    Dim InFile2 As New clsFileStream
    
    sOutFolder = txtOutput
    sOutFolder = Replace$(sOutFolder, "/", "\")
    If Right$(sOutFolder, 1) <> "\" Then
        sOutFolder = sOutFolder + "\"
    End If
    
    If inFile.OpenFile(txtInFile(0)) Then
        MsgBox "There was a problem opening the GAME file."
        inFile.CloseFile
        Exit Sub
    End If
    If inFile.FileSize < &H1000000 Then
        MsgBox "The GAME file is too small to use."
        inFile.CloseFile
        Exit Sub
    End If
    If inFile.FileSize = &H1000000 Then
        'arcade files
        inFile.ReadFileObject VarPtr(lParms(0)), 4, &HFE4000
        inFile.ReadFileObject VarPtr(lParms(1)), 4, &H24
        If lParms(1) = &H582D5350 Then
            If lParms(0) < 0 Then
                lGameType = GAMETYPE_DDRAC_OLD
                lReadOffset = &H100000
                ReDim lBlocksOccupied(0 To (inFile.FileSize \ &H800&) - 1) As Long
                bMarkUnref = True
            Else
                lGameType = GAMETYPE_DDRAC_NEW
                lReadOffset = &HFE4000
                If txtInFile(1) = "" Then
                    If MsgBox("No CARD file is specified. This means some data may be missing." + vbCrLf + "Continue anyway? (ignore this message for SOLO)", vbYesNo) = vbNo Then
                        inFile.CloseFile
                        Exit Sub
                    End If
                ElseIf InFile2.OpenFile(txtInFile(1)) Then
                    MsgBox "There was a problem opening the CARD file."
                    inFile.CloseFile
                    InFile2.CloseFile
                    Exit Sub
                End If
                ReDim lBlocksOccupied(0 To ((inFile.FileSize + InFile2.FileSize) \ &H800&) - 1) As Long
                bMarkUnref = True
            End If
        End If
    End If
    Debug.Print "GAMETYPE:", lGameType
    cmdExecute.Visible = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame5.Enabled = False
    
    ' ripping procedure ---
    Select Case lGameType
        Case GAMETYPE_DDRAC_OLD, GAMETYPE_DDRAC_NEW
            If chkFileList.value = 1 Then
                Do
                    inFile.ReadFileObject VarPtr(lFileTable(0)), 16, lReadOffset
                    If lFileTable(0) = -1 Or (lFileTable(0) = 0 And lFileTable(1) = 0 And lFileTable(2) = 0 And lFileTable(3) = 0) Then
                        'end of list
                        Exit Do
                    End If
                    lReadOffset = lReadOffset + 16
                    lRipOffset = (lFileTable(1) And &HFFFF&) * &H800&
                    If (lFileTable(1) And &HFF0000) <> 0 Then
                        lRipOffset = lRipOffset + inFile.FileSize
                    End If
                    If lFileTable(3) < (&H8000000) Then '>128mb = bad
                        For x = (lRipOffset \ &H800&) To ((lRipOffset + lFileTable(3)) \ &H800&)
                            If x >= 0 And x < UBound(lBlocksOccupied) Then
                                lBlocksOccupied(x) = 1
                            End If
                        Next x
                        If lRipOffset >= inFile.FileSize Then
                            'card data
                            lRipOffset = lRipOffset - inFile.FileSize
                            rf = RipFile(sOutFolder + "Card" + FormatNumberString(lRipOffset, 8), InFile2, lRipOffset, lFileTable(3), (lFileTable(2) = 1), False, bu)
                        Else
                            'game data
                            rf = RipFile(sOutFolder + "Game" + FormatNumberString(lRipOffset, 8), inFile, lRipOffset, lFileTable(3), (lFileTable(2) = 1), False, bu)
                        End If
                    End If
                    Label3 = CStr(rf) + "/" + CStr(lRipOffset)
                    Debug.Print "rip FT:", lRipOffset
                    DoEvents
                Loop While lReadOffset < inFile.FileSize
            Else
                'at least eliminate compressed material before searching for
                'uncompressed material when not using the file table
                For z = 0 To UBound(lBlocksOccupied)
                    z = (&H1000000 \ &H800&) + 15
                    lRipOffset = z * &H800&
                    If lRipOffset >= inFile.FileSize Then
                        lRipOffset = lRipOffset - inFile.FileSize
                        rf = RipFile(sOutFolder + "Card" + FormatNumberString(lRipOffset, 8), InFile2, lRipOffset, InFile2.FileSize - lRipOffset, True, True, bu)
                    Else
                        rf = RipFile(sOutFolder + "Game" + FormatNumberString(lRipOffset, 8), inFile, lRipOffset, inFile.FileSize - lRipOffset, True, True, bu)
                    End If
                    If rf <> FORMAT_NOTRIPPED Then
                        lBlocksOccupied(z) = 1
                    End If
                    Label3 = "rCo:" + CStr(lRipOffset)
                    DoEvents
                Next z
            End If
            If chkAllFiles.value = 1 Then
                For z = 0 To UBound(lBlocksOccupied)
                    If lBlocksOccupied(z) = 0 Then
                        'found an unused block, find the last unused block in the chain
                        For y = z To UBound(lBlocksOccupied)
                            If lBlocksOccupied(y) = 0 Then
                                x = y
                            Else
                                Exit For
                            End If
                        Next y
                        Debug.Print "BLOCKS", z, "TO", x
                        If z <= 18170 And x >= 18170 Then
                            x = x
                        End If
                        
                        'start ripping stuff all the way there
                        Debug.Print "rip RU:", lRipOffset
                        For y = z To x
                            lRipOffset = y * &H800&
                            If lRipOffset >= inFile.FileSize Then
                                lRipOffset = lRipOffset - inFile.FileSize
                                rf = RipFile(sOutFolder + "Card" + FormatNumberString(lRipOffset, 8), InFile2, lRipOffset, (((x - z) + 1) * &H800&), False, True, bu)
                            Else
                                rf = RipFile(sOutFolder + "Game" + FormatNumberString(lRipOffset, 8), inFile, lRipOffset, (((x - z) + 1) * &H800&), False, True, bu)
                            End If
                            Label3 = "rU:" + CStr(lRipOffset)
                            DoEvents
                        Next y
                        z = x
                    End If
                Next z
            End If
    End Select
    InFile2.CloseFile
    inFile.CloseFile
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame5.Enabled = True
    cmdExecute.Visible = True
End Sub

Private Function FormatNumberString(lNumber As Long, lCount As Long) As String
    FormatNumberString = CStr(lNumber)
    Do While Len(FormatNumberString) < lCount
        FormatNumberString = "0" + FormatNumberString
    Loop
End Function

'returns file type ripped
Private Function RipFile(ByVal sOutputFile As String, ByRef xFileStream As clsFileStream, lOffset As Long, ByVal lLength As Long, bForceCompressed As Boolean, bUnreferenced As Boolean, ByRef retBytesUsed As Long) As Long
    Dim AudioDecoder As New clsDeVAG
    Dim StepDecoder As New clsSSQ
    Dim ImageDecoder As New clsTIM
    Dim BemaniDecompressor As New clsDecompress1
    Dim bFData() As Byte
    Dim bDoRip As Boolean
    Dim f As Long
    Dim l As Long
    Dim bu As Long
    Dim x As Long
    Dim bFormatFound As Boolean
    Dim bValidCompressed As Boolean
    f = FreeFile
    RipFile = FORMAT_NOTRIPPED
    'read file data
    If bUnreferenced Then
        If bMarkUnref Then
            sOutputFile = sOutputFile + " (unref)"
        End If
        l = BemaniDecompressor.DecompressBemani1(xFileStream, bFData(), lOffset, bu)
        If l = 0 Then
            ReDim bFData(0) As Byte
        End If
        retBytesUsed = bu
    Else
        If bForceCompressed Then
            If BemaniDecompressor.DecompressBemani1(xFileStream, bFData(), lOffset, bu) = 0 Then
                Exit Function
            End If
            retBytesUsed = bu
        Else
            If xFileStream.ReadFileBinary(bFData(), lLength, lOffset, True) Then
                Exit Function
            End If
            retBytesUsed = lLength
        End If
    End If
    GoSub 1
    
    'for unreferenced files, try both compressed and uncompressed
    If bUnreferenced And (bFormatFound = False) Then
        If bForceCompressed = True Then
            'we're only doing a fast-seek on bForceCompressed + bUnreferenced
            If xFileStream.ReadFileBinary(bFData(), 2048, lOffset, True) Then
                Exit Function
            End If
            GoSub 1
            If Not bFormatFound Then
                Exit Function
            Else
                If xFileStream.ReadFileBinary(bFData(), lLength, lOffset, True) Then
                    Exit Function
                End If
                f = f
            End If
        Else
            If xFileStream.ReadFileBinary(bFData(), lLength, lOffset, True) Then
                Exit Function
            End If
            retBytesUsed = lLength
            GoSub 1
        End If
    End If
    
    'if we still haven't found the format, fall back to "dat"
    If Not bFormatFound Then
        sOutputFile = sOutputFile + ".dat"
        bDoRip = (chkData.value = 1)
        RipFile = FORMAT_UNKNOWN
        If bUnreferenced Then
            bDoRip = False
        End If
    End If
    
    If bDoRip Then
        'empty file first
        Open sOutputFile For Output As #f
        Close #f
        'then write it again
        Open sOutputFile For Binary As #f
        Put #f, 1, bFData
        Close #f
        If (Not bUnreferenced) Or (chkConvertUnref.value = 1) Then
            Select Case RipFile
                Case FORMAT_PBAV
                    If (chkConvertWAV.value = 1) Then
                        Label3 = "[DeVAG] .pbav"
                        DoEvents
                        If Not AudioDecoder.VAGtoWAV(sOutputFile, sOutputFile + ".wav", 3120, lLength, 44100) Then
                            If (chkDelTemps.value = 1) Then
                                Kill sOutputFile
                            End If
                        End If
                    End If
                Case FORMAT_ANNOUNCER
                    If (chkConvertWAV.value = 1) Then
                        Label3 = "[DeVAG] .sounds"
                        DoEvents
                        If Not AudioDecoder.ExtractSoundPack(sOutputFile, sOutputFile, 0, -1) Then
                            If (chkDelTemps.value = 1) Then
                                Kill sOutputFile
                            End If
                        End If
                    End If
                Case FORMAT_STEP2
                    If (chkConvertSM.value = 1) Then
                        Label3 = "[Step2SM] .step2"
                        DoEvents
                        If StepDecoder.ConvertStep2(sOutputFile, sOutputFile + ".sm") Then
                            If (chkDelTemps.value = 1) Then
                                Kill sOutputFile
                            End If
                        End If
                    End If
                Case FORMAT_SSQ
                    If (chkConvertSM.value = 1) Then
                        Label3 = "[Step2SM] .ssq"
                        DoEvents
                        If StepDecoder.ConvertSSQ(sOutputFile, sOutputFile + ".sm") Then
                            If (chkDelTemps.value = 1) Then
                                Kill sOutputFile
                            End If
                        End If
                    End If
                Case FORMAT_STEP
                    If (chkConvertSM.value = 1) Then
                        Label3 = "[Step2SM] .step"
                        DoEvents
                        If StepDecoder.ConvertStep1(sOutputFile, sOutputFile + ".sm", (chkSolo.value = 1)) Then
                            If (chkDelTemps.value = 1) Then
                                Kill sOutputFile
                            End If
                        End If
                    End If
                Case FORMAT_TIM
                    If (chkConvertBMP.value = 1) Then
                        Label3 = "[TIM2BMP] .tim"
                        DoEvents
                        If (Not ImageDecoder.TIMtoBMP(sOutputFile, sOutputFile + Chr$(1) + ".bmp")) Then
                            If (chkDelTemps.value = 1) Then
                                Kill sOutputFile
                            End If
                        End If
                    End If
            End Select
        End If
    End If
    Exit Function
    
    'determine format ---
1   If UBound(bFData) < 15 Then
        'don't bother with files under 16 bytes
        Return
    End If
    l = UBound(bFData) + 1
    
    'TIM image
    If bFData(0) = 16 And bFData(1) = 0 And bFData(2) = 0 And bFData(3) = 0 And (bFData(4) = 2 Or bFData(4) = 3 Or bFData(4) = 8 Or bFData(4) = 9) And bFData(5) = 0 Then
        bFormatFound = True
        sOutputFile = sOutputFile + ".tim"
        RipFile = FORMAT_TIM
        bDoRip = (chkTIM.value = 1)
        If bUnreferenced Or (Not bForceCompressed) Then
            'get the image's size from its header
            x = bFData(8) + (bFData(9) * 256&) + 8
            If bFData(4) <> 2 And bFData(4) <> 3 Then '16-bit and 24-bit color has no table
                x = x + (bFData(x) + (bFData(x + 1) * 256&) + (bFData(x + 2) * 65536))
            End If
            lLength = x
            ReDim Preserve bFData(0 To lLength - 1) As Byte
        End If
    
    'pBAV sound
    ElseIf bFData(0) = &H70 And bFData(1) = &H42 And bFData(2) = &H41 And bFData(3) = &H56 Then
        bFormatFound = True
        sOutputFile = sOutputFile + ".pbav"
        RipFile = FORMAT_PBAV
        bDoRip = (chkSound.value = 1)
        If bUnreferenced Or (Not bForceCompressed) Then
            x = bFData(12) + (bFData(13) * &H100&) + (bFData(14) * &H10000) + (bFData(15) * &H1000000)
            lLength = x
            ReDim Preserve bFData(0 To lLength - 1) As Byte
        End If
        
    'TGCD graphics collection
    ElseIf bFData(0) = &H54 And bFData(1) = &H47 And bFData(2) = &H43 And bFData(3) = &H44 Then
        
    
    'old DDR stepfile
    ElseIf (bFData(0) <> 0 Or bFData(1) <> 0) And bFData(2) = 0 And bFData(3) = 0 And bFData(l - 1) = 0 And bFData(l - 2) = 0 And bFData(l - 3) = 0 And bFData(l - 4) = 0 And bFData(l - 5) = 255 And bFData(l - 6) = 255 And bFData(l - 7) = 255 And bFData(l - 8) = 255 Then
        x = bFData(0) + (bFData(1) * 256&)
        If x < 2000 Then
            If bFData(x + 12) = 255 And bFData(x + 13) = 255 And bFData(x + 14) = 255 And bFData(x + 15) = 255 Then
                bFormatFound = True
                sOutputFile = sOutputFile + ".step"
                RipFile = FORMAT_STEP
                bDoRip = (chkSteps.value = 1)
            End If
        End If
        
    'new DDR stepfile
    ElseIf bFData(0) <> 0 And ((bFData(0) Mod 4) = 0) And bFData(2) = 0 And bFData(3) = 0 And bFData(4) = 1 And bFData(5) = 0 And (bFData(6) <> 0 Or bFData(7) <> 0) And bFData(8) <> 0 Then
        bFormatFound = True
        sOutputFile = sOutputFile + ".ssq"
        RipFile = FORMAT_SSQ
        bDoRip = (chkSteps.value = 1)
        
    'really, really old DDR stepfile
    ElseIf (bFData(0) Mod 4 = 0) And bFData(0) = bFData(4) And bFData(1) = bFData(5) And bFData(2) = bFData(6) And bFData(3) = bFData(7) And (bFData(0) <> 0 Or bFData(1) <> 0) And bFData(2) = 0 And bFData(3) = 0 And ((l Mod 4) = 0) Then
        bFormatFound = True
        sOutputFile = sOutputFile + ".step2"
        RipFile = FORMAT_STEP2
        bDoRip = (chkSteps.value = 1)
    
    End If
    
    If lOffset = 25952256 Then
        x = x
    End If
    
    'more complex stuff...
    If Not bFormatFound Then
        'announcer set
        If bFData(0) > 0 And bFData(1) = 0 And bFData(2) = 0 And bFData(3) = 0 Then
            Debug.Print lOffset
            x = (bFData(0) * 24) + 24
            x = (x + (CLng(bFData(12)) + (CLng(bFData(13)) * 256))) - 32
            If x < UBound(bFData) And (bForceCompressed = False) Then
                If bFData(x) = 0 And bFData(x + 1) = 7 And bFData(x + 2) = 119 Then
                    bFormatFound = True
                    sOutputFile = sOutputFile + ".sounds"
                    RipFile = FORMAT_ANNOUNCER
                    bDoRip = (chkSound.value = 1)
                End If
            End If
        End If
    End If
    
    Return
    
End Function

Private Function Round800h(iVal As Long) As Long
    Round800h = ((iVal + &H7FF&) \ &H800&) * &H800&
End Function
