VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCSFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================
' clsCSFile
' SaxxonPike 2oo6-2oo9
'
' interpreter class for note charts for Bemani's key-based games
' * can also load/save BME
'
' last updated January 14, 2009
'================================================================

Private Const BMEString = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const BadFileName = "?*/\""<>|"

Private Const InternalType_Measure = 1
Private Const InternalType_BPM = 2
Private Const InternalType_BPM03 = 3
Private Const InternalType_Note = 4
Private Const InternalType_NoteChange = 5
Private Const InternalType_EndSong = 6
Private Const InternalType_Metronome = 7
Private Const InternalType_Bad = 255

Private Const FileMode_Unknown = 0
Private Const FileMode_IIDXAC = 1
Private Const FileMode_IIDXCS = 2
Private Const FileMode_IIDXCSmulti = 3
Private Const FileMode_IIDXCS2 = 4
Private Const FileMode_IIDXCS5 = 5
Private Const FileMode_POPN = 6
Private Const FileMode_DrumMania = 7
Private Const FileMode_EZ2DJ = 8
Private Const FileMode_BMS = 9
Private Const FileMode_IIDXCS3 = 10

Private Const PopnMode_Unknown = 0
Private Const PopnMode_4 = 1
Private Const PopnMode_6 = 2
Private Const PopnMode_8 = 3

Private Const KeyNumbering_BME = 0
Private Const KeyNumbering_XBME = 1
Private Const KeyNumbering_BMS = 2
Private Const KeyNumbering_Decimal = 3

Private Const SyncTolerance = 2

Private Type Format4
    Offs As Integer
    Var1 As Byte
    Var2 As Byte
End Type
Private Type Format8
    Offs As Long
    Var1 As Byte
    Var2 As Byte
    Var3 As Integer
End Type
Private Type FormatPopn4
    Offs As Integer
    Var1 As Byte
    Var2 As Byte
End Type
Private Type FormatPopn6
    Offs As Integer
    Var1 As Byte
    xUnk0 As Byte
    Var2 As Byte
    Var3 As Byte
End Type
Private Type FormatPopn8
    Offs As Long
    Var1 As Byte
    Var2 As Byte
    Var3 As Integer
End Type
Private Type FormatPopnAC15
    Offs As Long
    Flag0 As Byte
    Var1 As Byte
    Var2 As Byte
    Var3 As Byte
End Type

Private Type FormatInternal
    OffsetMS As Double
    OffsetMetric As Double
    value As Double
    OffsetInternal As Double
    Lane As Double
    xType As Double
    OffsetBMSMetric As Double
    temp0 As Long
    temp1 As Integer
    xEnabled As Integer
End Type

Private Type FormatBMS
    Measure As Long
    Lane As Long
    Va As Double
End Type

Dim MeasureLengths() As Double
Dim MeasureSizes() As Double
Dim SimFileData() As FormatInternal
Dim KeysoundUsed() As Boolean
Dim KeysoundFile() As String
Dim bLoaded As Boolean
Private notecount(0 To 2) As Long
Private SongLength As Long
Private xBPM As Double
Private BPMlist() As Double
Private CurrentMode As Long
Private PlayerCount As Long
Private WorkingFile As String
Private xFrameTiming As Double
Private xCSSize As Long

Public Function GetSimFileData(dataPtr As Long, SimNumber As Long) As Boolean
    If SimNumber <= UBound(SimFileData) Then
        CopyMemory ByVal dataPtr, SimFileData(SimNumber), LenB(SimFileData(SimNumber))
        GetSimFileData = True
    End If
End Function

Public Function IsLoaded() As Boolean
    IsLoaded = bLoaded
End Function

Public Function TotalNotes() As Long
    TotalNotes = notecount(0) + notecount(1)
End Function

Public Function TotalSongLength() As Long
    TotalSongLength = SongLength
End Function

Public Function MainBPM() As Long
    MainBPM = xBPM
End Function

Public Sub WriteBMS(ByVal outfile As String, ByVal bWriteInfo As Boolean, ByVal bWriteTag As Boolean, ByVal Title As String, ByVal Artist As String, ByVal Genre As String, ByVal Difficulty As String, ByVal AddTags As String, ByVal KeysoundPrefix As String, ByVal KeyNumbering As Long, ByVal KeyQuant As Long, ByVal SetTitleToFileName As Boolean, ByVal bLineReduce As Boolean)
    
    Dim f As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim a As String
    Dim b As String
    Dim ret As String
    Dim maxmeasure As Long
    Dim currentmeasure As Long
    Dim linehead As String
    
    If Not bLoaded Then
        Exit Sub
    End If
    
    ' ===========================================================================
    '  HEADER
    ' ===========================================================================
    
    If SetTitleToFileName Then
        
        'auto get title
        a = outfile
        b = Mid$(a, InStrRev(a, "\") + 1)
        If InStr(b, " [") > 0 Then
            a = Mid$(outfile, InStrRev(outfile, " ["))
        Else
            a = "."
        End If
        a = Left$(a, InStrRev(a, ".") - 1)
        a = Title + a
        a = Replace(a, " [7key]", "")
        a = Replace(a, " [Normal]", "")
        Title = a
        
        'auto get difficulty
        If InStrRev(WorkingFile, "] (") > 0 Then
            a = WorkingFile
            a = Mid$(a, InStrRev(a, "] ("))
            a = Mid$(a, 4)
            a = Left$(a, InStr(a, ")") - 1)
            Difficulty = a
        End If
        
    End If
    
    outfile = FixFileName(outfile)
    
    f = FreeFile
    On Local Error GoTo 66
    Open outfile For Output As #f
    On Local Error GoTo 0
    Print #f, "; >>> Bemani2BMS v" + CStr(App.Major) + "." + CStr(App.Minor) + " release " + CStr(App.Revision)
    If bWriteTag Then
        Print #f, "; File created "; Date$; " // "; Time$
    End If
    If bWriteInfo Then
        Print #f, "; ------------------------------------------------------------------------"
        Print #f, "; info ... ... ... ... ... ... ... ... ... ... ... ... ... ... ... ... ..."
        Print #f, "; ------------------------------------------------------------------------"
        Print #f, ";          TITLE: "; Title
        Print #f, ";         ARTIST: "; Artist
        Print #f, ";          GENRE: "; Genre
        Print #f, ";"
        Print #f, ";  1P Note count:"; notecount(0)
        Print #f, ";  2P Note count:"; notecount(1)
        Print #f, "; BGM Note count:"; notecount(2)
        Print #f, ";"
        Print #f, ";    Song Length:"; SongLength; "seconds"
        Print #f, "; ------------------------------------------------------------------------"
        Print #f, "; ... ... ... ... ... ... ... ... ... ... ... ... ... ... ... ... ... info"
        Print #f, "; ------------------------------------------------------------------------"
    End If
    Print #f, ""
    Print #f, "#TITLE "; Title
    Print #f, "#ARTIST "; Artist
    Print #f, "#GENRE "; Genre
    If Len(Trim(Difficulty)) > 0 Then
        Print #f, "#PLAYLEVEL "; Difficulty
    End If
    Print #f, "#BPM "; CStr(xBPM)
    If CurrentMode <> FileMode_POPN Then
        Print #f, "#PLAYER ";
        Select Case PlayerCount
            Case 1
                Print #f, "1"
            Case 2
                Print #f, "3"
            Case Else
                Print #f, "0"
        End Select
    Else
        Print #f, "#PLAYER 3"
    End If
    If AddTags <> "" Then
        Print #f, AddTags
    End If
    For x = 1 To 1295
        If KeysoundUsed(x) Then
            Print #f, "#WAV"; BME(x); " "; KeysoundPrefix;
            Select Case KeyNumbering
                Case KeyNumbering_BME
                    Print #f, BME(x);
                Case KeyNumbering_XBME
                    Print #f, "00"; BME(x);
                Case KeyNumbering_BMS
                    Print #f, BMS(x);
                Case KeyNumbering_Decimal
                    Print #f, CStr(x);
            End Select
            Print #f, ".wav"
        End If
    Next x
    
    maxmeasure = Int(SimFileData(UBound(SimFileData)).OffsetBMSMetric + 1)
    
    y = 0
    For x = LBound(SimFileData) To UBound(SimFileData)
        With SimFileData(x)
            If .xType = InternalType_BPM And Dec2(.temp1) > y Then
                Print #f, "#BPM"; Dec2(.temp1); " "; CStr(.value)
                y = .temp1
            End If
        End With
    Next x
    
    For x = 0 To maxmeasure
        currentmeasure = x + 1
        If MeasureLengths(x) > 0 Then
            Print #f, "#" + String(3 - Len(CStr(currentmeasure - 1)), "0") + CStr(currentmeasure - 1) + "02:" + CStr(MeasureLengths(x))
        End If
    Next x
    
    ' ===========================================================================
    '  BODY
    ' ===========================================================================
    
    Select Case CurrentMode
        
        Case FileMode_IIDXAC, FileMode_IIDXCS, FileMode_IIDXCS2, FileMode_IIDXCSmulti, _
            FileMode_POPN, FileMode_IIDXCS3
    
            For x = 0 To maxmeasure
                currentmeasure = x + 1
                For y = 1 To 29
                    ret = ""
                    linehead = "#"
                    linehead = linehead + String(3 - Len(CStr(currentmeasure - 1)), "0") + CStr(currentmeasure - 1)
                    linehead = linehead + Dec2(y + 0) + ":"
                    Select Case y
                    'BGM ----
                        Case 1
                            Do Until ret = String(KeyQuant * 2, "0")
                                ret = BMSLine(InternalType_Note, x, 64, KeyQuant, KeyNumbering_BME)
                                If ret <> String(KeyQuant * 2, "0") Then
                                    Print #f, linehead; ret
                                End If
                            Loop
                    'BPM ----
                        Case 3
                            ret = BMSLine(InternalType_BPM03, x, 0, KeyQuant, KeyNumbering_BMS)
                        Case 8
                            ret = BMSLine(InternalType_BPM, x, 0, KeyQuant, KeyNumbering_Decimal)
                    End Select
                    
                    Select Case CurrentMode
                        Case FileMode_IIDXAC, FileMode_IIDXCS, FileMode_IIDXCS2, FileMode_IIDXCSmulti, FileMode_IIDXCS3
                            Select Case y
                            '1P ----
                                Case 11
                                    ret = BMSLine(InternalType_Note, x, 0, KeyQuant, KeyNumbering_BME)
                                Case 12
                                    ret = BMSLine(InternalType_Note, x, 1, KeyQuant, KeyNumbering_BME)
                                Case 13
                                    ret = BMSLine(InternalType_Note, x, 2, KeyQuant, KeyNumbering_BME)
                                Case 14
                                    ret = BMSLine(InternalType_Note, x, 3, KeyQuant, KeyNumbering_BME)
                                Case 15
                                    ret = BMSLine(InternalType_Note, x, 4, KeyQuant, KeyNumbering_BME)
                                Case 18
                                    ret = BMSLine(InternalType_Note, x, 5, KeyQuant, KeyNumbering_BME)
                                Case 19
                                    ret = BMSLine(InternalType_Note, x, 6, KeyQuant, KeyNumbering_BME)
                                Case 16
                                    ret = BMSLine(InternalType_Note, x, 7, KeyQuant, KeyNumbering_BME)
                                Case 17
                                    ret = BMSLine(InternalType_Note, x, 8, KeyQuant, KeyNumbering_BME)
                            '2P ----
                                Case 21
                                    ret = BMSLine(InternalType_Note, x, 32, KeyQuant, KeyNumbering_BME)
                                Case 22
                                    ret = BMSLine(InternalType_Note, x, 33, KeyQuant, KeyNumbering_BME)
                                Case 23
                                    ret = BMSLine(InternalType_Note, x, 34, KeyQuant, KeyNumbering_BME)
                                Case 24
                                    ret = BMSLine(InternalType_Note, x, 35, KeyQuant, KeyNumbering_BME)
                                Case 25
                                    ret = BMSLine(InternalType_Note, x, 36, KeyQuant, KeyNumbering_BME)
                                Case 28
                                    ret = BMSLine(InternalType_Note, x, 37, KeyQuant, KeyNumbering_BME)
                                Case 29
                                    ret = BMSLine(InternalType_Note, x, 38, KeyQuant, KeyNumbering_BME)
                                Case 26
                                    ret = BMSLine(InternalType_Note, x, 39, KeyQuant, KeyNumbering_BME)
                                Case 27
                                    ret = BMSLine(InternalType_Note, x, 40, KeyQuant, KeyNumbering_BME)
                            'extra ----
                            End Select
                        Case FileMode_POPN
                            Select Case y
                            '1P ----
                                Case 11
                                    ret = BMSLine(InternalType_Note, x, 0, KeyQuant, KeyNumbering_BME)
                                Case 12
                                    ret = BMSLine(InternalType_Note, x, 1, KeyQuant, KeyNumbering_BME)
                                Case 13
                                    ret = BMSLine(InternalType_Note, x, 2, KeyQuant, KeyNumbering_BME)
                                Case 14
                                    ret = BMSLine(InternalType_Note, x, 3, KeyQuant, KeyNumbering_BME)
                                Case 15
                                    ret = BMSLine(InternalType_Note, x, 4, KeyQuant, KeyNumbering_BME)
                                Case 22
                                    ret = BMSLine(InternalType_Note, x, 5, KeyQuant, KeyNumbering_BME)
                                Case 23
                                    ret = BMSLine(InternalType_Note, x, 6, KeyQuant, KeyNumbering_BME)
                                Case 24
                                    ret = BMSLine(InternalType_Note, x, 7, KeyQuant, KeyNumbering_BME)
                                Case 25
                                    ret = BMSLine(InternalType_Note, x, 8, KeyQuant, KeyNumbering_BME)
                            End Select
                    End Select
                    If bLineReduce Then
                        If ret <> "" And ret <> String(KeyQuant * 2, "0") Then
                            ret = ReduceLine(ret)
                            If ret <> "" Then
                                Print #f, linehead; ret
                            End If
                        End If
                    Else
                        If ret <> String(KeyQuant * 2, "0") And ret <> "" Then
                            Print #f, linehead; ret
                        End If
                    End If
                Next y
            Next x
    End Select
    
    Close #f
    Exit Sub
    
66  Close #f
    On Local Error GoTo 67
    a = UCase$(outfile)
    For x = InStrRev(outfile, "\") + 1 To Len(outfile)
        y = (Asc(Mid$(a, x, 1)))
        If y < 32 Or y > 126 Then
            Mid$(outfile, x, 1) = "_"
        End If
    Next x
    
    a = ""
    Resume

67  Close #f
    On Local Error GoTo 0
End Sub

Public Sub LoadFile(ByVal inFile As String, ByVal FileMode As Long, ByVal SubMode As Long, Optional ByVal TimeOverride As Double = 0, Optional ByVal SyncMode As Long = -1)
    
    'defines
    Dim TimingConstant As Double
    Dim f As Long
    Dim verifysig As Long
    Dim tlong As Long
    Dim tlong2 As Long
    Dim sortflag As Boolean
    Dim currentmeasure As Long
    Dim F4() As Format4
    Dim F8() As Format8
    Dim FP4() As FormatPopn4
    Dim FP6() As FormatPopn6
    Dim FP8() As FormatPopn8
    Dim FPAC15() As FormatPopnAC15
    Dim TempInternal As FormatInternal
    Dim tF4 As Format4
    Dim tF8 As Format8
    Dim tFP8 As FormatPopn8
    Dim FBMS() As FormatBMS
    ReDim MeasureLengths(0 To 999) As Double
    ReDim MeasureSizes(0 To 999) As Double
    Dim LastMeasure As Long
    Dim LastCheck As Long
    Dim NML As Double
    Dim dp As Long
    Dim useformat As Long
    Dim foffs As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim a As Long
    Dim bEndFlag As Boolean
    Dim PopnMode As Long
    Dim DefaultNote(0 To 99) As Double
    Dim tNote As FormatInternal
    Dim currentoffset As Double
    Dim CurrentMetric As Double
    Dim CreateNote As Boolean
    ReDim SimFileData(0 To 2) As FormatInternal
    Dim SimFileDataCount As Long
    ReDim KeysoundUsed(1 To 1295) As Boolean
    ReDim KeysoundFile(1 To 1295) As String
    ReDim BPMlist(0 To 100) As Double
    ReDim BPMlist03(0 To 1295) As Double
    Dim BPMCount As Long
    Dim BPMCount03 As Long
    Dim CurrentBPM As Double
    Dim LastMeasureMetric As Double
    Dim LastRealMeasure As Long
    Dim bPopnAC15 As Boolean
    Dim MeasureCount As Long
    Dim xBMSLine As String
    
    notecount(0) = 0
    notecount(1) = 0
    notecount(2) = 0
    bLoaded = False
    BPMCount = 0
    PlayerCount = 1
    xCSSize = 0
    
    CurrentMode = FileMode_Unknown
    
    'better have a good mode
    If FileMode = FileMode_Unknown Then
        Debug.Print "Convertfile: Can't convert from unknown type."
        Exit Sub
    End If
    
    'if the input doesn't exist, cancel
    If Dir(inFile) = "" Then
        Debug.Print "Convertfile: Input file doesn't exist."
        Exit Sub
    End If
    
    ReDim FBMS(0) As FormatBMS
    WorkingFile = inFile
    
    f = FreeFile
    If FileMode <> FileMode_BMS Then
        Open inFile For Binary As #f
    Else
        Open inFile For Input As #f
    End If
    
    'verify stage
    Select Case FileMode
        Case FileMode_IIDXAC
            y = 0
            For x = 0 To 11
                Get #f, (x * 8) + 1, verifysig
                If (verifysig >= &H60 And verifysig < LOF(f)) Then
                    y = 1
                    Exit For
                End If
            Next x
            If y = 0 Then
                Debug.Print "Convertfile: IIDXAC verification failed."
                Close #f
                Exit Sub
            End If
        Case FileMode_IIDXCS
            Get #f, 1, verifysig
            If verifysig <> 8 Then
                Debug.Print "Convertfile: IIDXCS verification failed."
                Close #f
                Exit Sub
            End If
        Case FileMode_IIDXCS2
            Get #f, 1, verifysig
            If verifysig <> 8 Then
                Get #f, LOF(1) - 3, verifysig
                If verifysig <> &H7FFF& And LOF(1) <> &H4000& Then
                    Debug.Print "Convertfile: IIDXCS2 verification failed."
                    Close #f
                    Exit Sub
                End If
                ReDim F4(0 To (LOF(f) \ 4) - 1) As Format4
                ReDim F8(0 To UBound(F4)) As Format8
                Get #f, 5, F4
            Else
                ReDim F4(0 To (LOF(f) \ 4) - 2) As Format4
                ReDim F8(0 To UBound(F4)) As Format8
                Get #f, 9, F4
            End If
            For x = 0 To UBound(F8)
                F8(x).Offs = F4(x).Offs
                F8(x).Var1 = F4(x).Var1 And &HF
                F8(x).Var2 = (F4(x).Var1 And &HF0) \ &H10
                F8(x).Var3 = F4(x).Var2
            Next x
        Case FileMode_IIDXCS3
            Get #f, 1, verifysig
            If verifysig <> 8 Then
                Debug.Print "Convertfile: IIDXCS3 verification failed."
                Close #f
                Exit Sub
            End If
            
        Case FileMode_POPN
            'autodetect mode: we must find out the width of our fields
            PopnMode = PopnMode_Unknown
            
            Get #f, 1, verifysig
            
            If verifysig = 32 And SubMode < 6 Then
                TimingConstant = 1
                Get #f, 1 + (SubMode * 4), foffs
                Get #f, 5 + (SubMode * 4), tlong2
                If tlong2 > 0 Then
                    PopnMode = PopnMode_8       'newer 8-wide style (CS9)
                End If
                tlong2 = tlong2 - foffs
                ReDim FP8(0 To (tlong2 \ 8) - 1) As FormatPopn8
                Get #f, foffs + 1, FP8
            Else
                If SubMode = 0 Then
                    Get #f, 1, tlong
                    Get #f, 5, tlong2
                    If tlong = 1 And tlong2 = 1 Then
                        Get #f, 17, foffs
                        Get #f, 21, tlong2
                        Get #f, (foffs + tlong2) - 3, verifysig
                        If (verifysig And &HFFFF&) = &H7FFF& Then
                            PopnMode = PopnMode_4
                            ReDim FP4(0 To (tlong2 \ 4) - 1) As FormatPopn4
                            ReDim FP8(0 To (tlong2 \ 4) - 1) As FormatPopn8
                            Get #f, foffs + 1, FP4
                        Else
                            PopnMode = PopnMode_6
                            ReDim FP6(0 To (tlong2 \ 6) - 1) As FormatPopn6
                            ReDim FP8(0 To (tlong2 \ 6) - 1) As FormatPopn8
                            Get #f, foffs + 1, FP6
                        End If
                        
                        TimingConstant = (50 / 3)
                        
                    Else
                        Get #f, LOF(f) - 3, tlong
                        If tlong = 32767 Then
                            PopnMode = PopnMode_4
                            tlong2 = (LOF(f) \ 4)
                            ReDim FP4(0 To (tlong2 \ 4) - 1) As FormatPopn4
                            ReDim FP8(0 To (tlong2 \ 4) - 1) As FormatPopn8
                            Get #f, 1, FP4
                            TimingConstant = 50 / 3
                        ElseIf (tlong And &H7FFFFF00) = &H600& Then
                            'arcade
                            PopnMode = PopnMode_8
                            tlong2 = LOF(f)
                            ReDim FP8(0 To (tlong2 \ 8) - 1) As FormatPopn8
                            ReDim FPAC15(0 To (tlong2 \ 8) - 1) As FormatPopnAC15
                            Get #f, 1, FPAC15
                            For x = LBound(FP8) To UBound(FP8)
                                With FPAC15(x)
                                    FP8(x).Offs = .Offs
                                    FP8(x).Var2 = .Var1
                                    FP8(x).Var1 = .Var3 \ 16
                                    FP8(x).Var3 = .Var2
                                    If FP8(x).Var2 = 1 Then
                                        FP8(x).Var1 = FP8(x).Var3
                                        FP8(x).Var3 = 0
                                    ElseIf FP8(x).Var2 = 10 Then
                                        FP8(x).Var2 = 0
                                        If FP8(x).Offs = 0 Then
                                            FP8(x).Var2 = 255
                                        End If
                                    ElseIf FP8(x).Var2 = 2 Or FP8(x).Var2 = 7 Then
                                        FP8(x).Var3 = FP8(x).Var3 + ((CLng(.Var3) And 15) * 256)
                                    ElseIf FP8(x).Var2 = 0 Then
                                        FP8(x).Var2 = 255
                                    End If
                                End With
                            Next x
                            TimingConstant = 1
                            bPopnAC15 = True
                        End If
                    End If
                End If
                        
            End If
            
            If PopnMode = PopnMode_Unknown Then
                Debug.Print "ConvertFile: Pop'n mode could not be determined."
                Close #f
                Exit Sub
            End If
            
        Case FileMode_DrumMania
            
            Get #f, 1, verifysig
            verifysig = verifysig And &HFFFF&
            Get #f, (verifysig * 40) + 1 + 16, tlong
            If tlong <> &H31515344 Then
                Debug.Print "ConvertFile: Could not recognize DrumMania simfile."
                Close #f
                Exit Sub
            End If
            If SubMode > (verifysig - 1) Then
                Close #f
                Exit Sub
            End If
            Get #f, (SubMode * 40) + 48 + 1, tlong      'offset
            Get #f, (SubMode * 40) + 48 + 5, tlong2     'length
            tlong2 = tlong2 - 16
            tlong = tlong + (verifysig * 40) + 1 + 16
            ReDim F8(0 To (tlong2 \ 8)) As Format8
            Get #f, tlong + 16, F8
    End Select
    
    'pre-conditioning stage
    Select Case FileMode
        Case FileMode_IIDXAC
            TimingConstant = (1 / 59.94) * 1000
            Get #f, 1 + (SubMode * 8), tlong
            Get #f, 5 + (SubMode * 8), tlong2
            If tlong = 0 Or tlong2 = 0 Then
                Debug.Print "ConvertFile: AC mode doesn't exist in the file."
                Close #f
                Exit Sub
            End If
            ReDim F8(0 To (tlong2 \ 8)) As Format8
            Get #f, 1 + tlong, F8
            For x = LBound(F8) To UBound(F8) - 1
                With F8(x)
                    'remove duplicate measures (14key stuff)
                    If .Var1 = &HC And .Var2 <> 0 Then
                        .Var1 = 255
                    End If
                    'remove measures at offset 0
                    If .Offs = 0 And .Var1 = &HC Then
                        .Var1 = 255
                    End If
                    'remove bad notes
                    If (.Var1 = 0 Or .Var1 = 1) And .Var3 <> 0 Then
                        .Var1 = 255
                    End If
                    If (.Var1 = 0 Or .Var1 = 1) And .Offs = 0 Then
                        .Var1 = 255
                    End If
                End With
            Next x
        Case FileMode_IIDXCS, FileMode_IIDXCS3
            If FileMode = FileMode_IIDXCS3 Then
                x = 5 + (&H10000 * SubMode)
                y = &H10000
            Else
                x = 5
                y = LOF(f)
            End If
            Get #f, x, tlong
            If tlong = 0 Then
                Debug.Print "Convertfile: Bad CS time constant."
                Close #f
                Exit Sub
            End If
            TimingConstant = (tlong / 1000)
            ReDim F8(0 To (y \ 8) - 1) As Format8
            Get #f, x + 4, F8
            For x = LBound(F8) To UBound(F8)
                With F8(x)
                    'transfer lanes
                    If (.Var3 = 0 And .Var2 <> 0) Then
                        .Var3 = .Var2
                        .Var2 = 0
                    End If
                    If (.Var1 And &HF) = 4 Then
                        .Var3 = .Var3 + ((CLng(.Var1) And &HF0&) * &H10&)
                        .Var1 = &H4
                    End If
                    'convert to AC .1 format
                    .Var2 = .Var1 \ &H10&
                    .Var1 = .Var1 And &HF&
                    'remove duplicate measures (14key stuff)
                    If .Var1 = &HC And .Var2 <> 0 Then
                        .Var1 = 255
                    End If
                    'remove measures at offset 0
                    If .Offs = 0 And .Var1 = &HC Then
                        .Var1 = 255
                    End If
                    'remove bad notes
                    If (.Var1 = 0 Or .Var1 = 1) And .Var3 <> 0 Then
                        .Var1 = 255
                    End If
                    If .Offs = &H7FFF And FileMode = FileMode_IIDXCS3 Then
                        For y = x To UBound(F8)
                            F8(y).Offs = &H7FFF
                            F8(y).Var1 = 0
                            F8(y).Var2 = 0
                            F8(y).Var3 = 0
                        Next y
                    End If
                End With
            Next x
            FileMode = FileMode_IIDXCS
        Case FileMode_IIDXCS2
            TimingConstant = 16.722
            If verifysig = 8 Then
                Get #f, 5, tlong
                TimingConstant = (tlong / 1000)
            End If
            FileMode = FileMode_IIDXCS
            For x = LBound(F8) To UBound(F8)
                With F8(x)
                    'remove duplicate measures (14key stuff)
                    If .Var1 = &HC And .Var2 <> 0 Then
                        .Var1 = 255
                    End If
                    'remove measures at offset 0
                    If .Offs = 0 And .Var1 = &HC Then
                        .Var1 = 255
                    End If
                    'remove bad notes
                    If (.Var1 = 0 Or .Var1 = 1) And .Var3 <> 0 Then
                        .Var1 = 255
                    End If
                End With
            Next x
        Case FileMode_POPN
            Select Case PopnMode 'convert all old pop'n formats to new 8-wide
                Case PopnMode_4
                    For x = LBound(FP4) To UBound(FP4)
                        With FP8(x)
                            .Offs = FP4(x).Offs
                            .Var1 = (FP4(x).Var1 And &HF0&) \ &H10&
                            .Var2 = FP4(x).Var1 And &HF&
                            .Var3 = FP4(x).Var2
                            If .Var1 = 12 Then
                                .Var1 = 0
                                .Var2 = 10
                            End If
                        End With
                    Next x
                    FP8(0).Offs = 0
                    FP8(0).Var1 = 0
                    FP8(0).Var2 = 3
                    FP8(0).Var3 = 0
                Case PopnMode_6
                    For x = LBound(FP6) To UBound(FP6)
                        With FP8(x)
                            .Offs = FP6(x).Offs
                            .Var1 = FP6(x).Var3
                            .Var2 = FP6(x).Var1
                            .Var3 = FP6(x).Var2
                        End With
                    Next x
                    FP8(1).Offs = 0
                    FP8(1).Var1 = 0
                    FP8(1).Var2 = 250
                    FP8(1).Var3 = 0
            End Select
            
            For x = LBound(FP8) To UBound(FP8)
                With FP8(x)
                    If .Var1 = 0 And .Var2 = 0 And .Var3 <> 0 Then
                        .Var2 = 255
                    End If
                    'swap the actions so Var1 is the action, Var2 is the lane
                    tlong = .Var2
                    .Var2 = .Var1
                    .Var1 = tlong
                End With
                If PopnMode = PopnMode_4 Then
                    If FP8(x).Var1 = 6 Then
                        FP8(x).Var1 = 8
                    ElseIf FP8(x).Var1 = 2 Then
                        FP8(x).Var1 = 4
                    ElseIf FP8(x).Var1 = 1 Then
                        FP8(x).Var1 = 2
                    ElseIf FP8(x).Var1 = 0 Then
                        FP8(x).Var1 = 1
                    ElseIf FP8(x).Var1 = 5 Then
                        FP8(x).Var1 = 11
                    ElseIf FP8(x).Var1 = 4 Then
                        FP8(x).Var1 = 6
                    End If
                ElseIf PopnMode = PopnMode_6 Then
                    If FP8(x).Var1 = 3 Then
                        FP8(x).Var1 = 2
                    ElseIf FP8(x).Var1 = 250 Then
                        FP8(x).Var1 = 3
                    ElseIf FP8(x).Var1 = 16 Then
                        FP8(x).Var1 = 11
                    ElseIf FP8(x).Var1 = 12 Then
                        FP8(x).Var1 = 10
                    End If
                End If
            Next x
        
        Case FileMode_DrumMania
            'events
            '    offs-offset(units unknown)
            '    var1-action
            '    var2-volume
            '    var3-value(2b)
            ' 0=hihat
            ' 1=snare
            ' 2=bassdrum
            ' 3=hitom
            ' 4=lowtom
            ' 5=cymbal
            
            ' 7=measure
            ' 8=metronome
            ' A=end
        
        Case FileMode_EZ2DJ
        
        
        
        Case FileMode_BMS
            'BMS loader added for BeMedia Player:
            ' 02 August 2007
            y = 1
            For x = 0 To 999
                MeasureLengths(x) = 1
            Next x
            ReDim SimFileData(0) As FormatInternal
            LastMeasureMetric = 0
            
            Do While Not EOF(f)
                Line Input #f, xBMSLine
                xBMSLine = UCase$(xBMSLine)
                If Left$(xBMSLine, 1) = "#" And Mid$(xBMSLine, 7, 1) = ":" Then
                    x = Val(Mid$(xBMSLine, 5, 2))
                    z = Val(Mid$(xBMSLine, 2, 3))
                    LastMeasureMetric = 0
                    If x <> 2 Then
                        For a = 0 To z - 1
                            LastMeasureMetric = LastMeasureMetric + MeasureLengths(a)
                        Next a
                        If x > MeasureCount Then
                            MeasureCount = x
                        End If
                        xBMSLine = Mid$(xBMSLine, 8)
                        tlong = Len(xBMSLine) \ 2
                        For a = 1 To Len(xBMSLine) Step 2
                            If Mid$(xBMSLine, a, 2) <> "00" Then
                                ReDim Preserve SimFileData(0 To y) As FormatInternal
                                With SimFileData(y)
                                    Select Case x
                                        Case 11: .Lane = 0
                                        Case 12: .Lane = 1
                                        Case 13: .Lane = 2
                                        Case 14: .Lane = 3
                                        Case 15: .Lane = 4
                                        Case 18: .Lane = 5
                                        Case 19: .Lane = 6
                                        Case 16: .Lane = 7
                                        Case 21: .Lane = 32
                                        Case 22: .Lane = 33
                                        Case 23: .Lane = 34
                                        Case 24: .Lane = 35
                                        Case 25: .Lane = 36
                                        Case 28: .Lane = 37
                                        Case 29: .Lane = 38
                                        Case 26: .Lane = 39
                                        Case 1: .Lane = 64
                                    End Select
                                    If x <> 3 And x <> 8 Then
                                        .xType = InternalType_Note
                                        .value = BMEToDec(Mid$(xBMSLine, a, 2)) 'Val(Mid$(xBMSLine, a, 2))
                                    ElseIf x = 3 Then
                                        .xType = InternalType_BPM03
                                        .value = CLng("&H" + Mid$(xBMSLine, a, 2))
                                    ElseIf x = 8 Then
                                        .xType = InternalType_BPM
                                        .value = BPMlist(Val(Mid$(xBMSLine, a, 2)))
                                        .temp1 = Val(Mid$(xBMSLine, a, 2))
                                    End If
                                    .OffsetBMSMetric = LastMeasureMetric + ((((a - 1) \ 2) / tlong) * MeasureLengths(z))
                                End With
                                y = y + 1
                            End If
                        Next a
                    Else
                        MeasureLengths(z) = Val(Mid$(xBMSLine, 8))
                    End If
                ElseIf Left$(xBMSLine, 5) = "#BPM " Then
                    CurrentBPM = Val(Mid$(xBMSLine, 6))
                ElseIf Left$(xBMSLine, 4) = "#BPM" Then
                    BPMlist(Val(Mid$(xBMSLine, 5, 2))) = Val(Mid$(xBMSLine, 8))
                ElseIf Left$(xBMSLine, 4) = "#WAV" Then
                    KeysoundFile(BMEToDec(Mid$(xBMSLine, 5, 2))) = Mid$(xBMSLine, 8)
                End If
            Loop
            
            ' sort by metric
            Do
                sortflag = False
                For x = 0 To y - 2
                    If SimFileData(x).OffsetBMSMetric > SimFileData(x + 1).OffsetBMSMetric Then
                        TempInternal = SimFileData(x)
                        SimFileData(x) = SimFileData(x + 1)
                        SimFileData(x + 1) = TempInternal
                        sortflag = True
                    End If
                Next x
            Loop While sortflag = True
            
            ' develop offset by metric
            NML = GetNML(CurrentBPM, 1)
            TimingConstant = 1
            LastMeasureMetric = 0
            LastCheck = 0
            For x = 0 To y - 1
                With SimFileData(x)
                    .OffsetMS = ((.OffsetBMSMetric - LastMeasureMetric) * NML) + LastCheck
                    .OffsetInternal = .OffsetMS
                    .OffsetMetric = .OffsetBMSMetric
                    .xEnabled = 0
                    Select Case .xType
                        Case InternalType_BPM, InternalType_BPM03
                            CurrentBPM = .value
                            LastCheck = .OffsetMS
                            LastMeasureMetric = .OffsetBMSMetric
                            NML = GetNML(CurrentBPM, 1)
                        Case InternalType_Note
                            KeysoundUsed(.value) = True
                    End Select
                End With
            Next x
            
            ReDim Preserve SimFileData(0 To y) As FormatInternal
            With SimFileData(y)
                .xType = InternalType_EndSong
                .OffsetMS = SimFileData(y - 1).OffsetMS
                .OffsetMetric = SimFileData(y - 1).OffsetMetric
                .OffsetInternal = .OffsetMS
            End With
            
    End Select
    
    'do some fixing on really odd occasions like "Stop Violence!" which
    'remove some vital information on some notes
    LastCheck = 0
    Select Case FileMode
        Case FileMode_IIDXCS
            For x = LBound(F8) To UBound(F8)
                If (F8(x).Offs And &HFF00&) > LastCheck Then
                    LastCheck = F8(x).Offs And &HFF00&
                End If
                If LastCheck > 0 And (F8(x).Offs And &HFF00&) = 0 Then
                    F8(x).Offs = F8(x).Offs + (LastCheck + &H100&)
                End If
            Next x
        Case Else
    End Select
    
    'override timing if needed
    If TimeOverride > 0 Then
        TimingConstant = (1000 / TimeOverride)
    End If
    xFrameTiming = TimingConstant
    
    'pre-sort stage
    Select Case FileMode
        Case FileMode_IIDXCS, FileMode_IIDXAC
            sortflag = True
            Do Until sortflag = False
                sortflag = False
                For x = LBound(F8) + 1 To UBound(F8)
                    If F8(x).Offs < F8(x - 1).Offs Then
                        tF8 = F8(x)
                        F8(x) = F8(x - 1)
                        F8(x - 1) = tF8
                        sortflag = True
                    End If
                Next x
            Loop
            currentmeasure = 0
            For x = LBound(F8) To UBound(F8)
                If F8(x).Var1 = 4 Then
                    If F8(x).Var2 = 0 Then F8(x).Var2 = 1
                    If NML > 0 Then
                        MeasureSizes(currentmeasure) = MeasureSizes(currentmeasure) + (F8(x).Offs - LastMeasure)
                        MeasureLengths(currentmeasure) = MeasureLengths(currentmeasure) + ((F8(x).Offs - LastMeasure) / NML)
                    Else
                        xBPM = F8(x).Var3 / F8(x).Var2
                    End If
                    NML = GetNML(F8(x).Var3 / F8(x).Var2, TimingConstant)
                    LastMeasure = F8(x).Offs
                End If
                If F8(x).Var1 = &HC Or F8(x).Var1 = 6 Then
                    If (F8(x).Offs - LastRealMeasure) > 0 And NML > 0 Then
                        MeasureSizes(currentmeasure) = MeasureSizes(currentmeasure) + (F8(x).Offs - LastMeasure)
                        MeasureLengths(currentmeasure) = MeasureLengths(currentmeasure) + ((F8(x).Offs - LastMeasure) / NML)
                        LastRealMeasure = F8(x).Offs
                        LastMeasure = LastRealMeasure
                        LastCheck = F8(x).Offs
                        currentmeasure = currentmeasure + 1
                    End If
                End If
            Next x
            
            
            
            
        Case FileMode_POPN
            sortflag = True
            Do Until sortflag = False
                sortflag = False
                For x = LBound(FP8) + 1 To UBound(FP8)
                    If FP8(x).Offs < FP8(x - 1).Offs Then
                        tFP8 = FP8(x)
                        FP8(x) = FP8(x - 1)
                        FP8(x - 1) = tFP8
                        sortflag = True
                    End If
                Next x
            Loop
            currentmeasure = 0
            For x = LBound(FP8) To UBound(FP8)
                If FP8(x).Var1 = 4 Then
                    If NML > 0 Then
                        MeasureSizes(currentmeasure) = MeasureSizes(currentmeasure) + (FP8(x).Offs - LastMeasure)
                        MeasureLengths(currentmeasure) = MeasureLengths(currentmeasure) + ((FP8(x).Offs - LastMeasure) / NML)
                    Else
                        xBPM = FP8(x).Var3
                    End If
                    NML = GetNML(FP8(x).Var3 + 0, TimingConstant)
                    LastMeasure = FP8(x).Offs
                End If
                If ((FP8(x).Var1 = &HA) Or (bPopnAC15 = True And FP8(x).Var1 = 0)) And NML > 0 Then
                    MeasureSizes(currentmeasure) = MeasureSizes(currentmeasure) + (FP8(x).Offs - LastMeasure)
                    MeasureLengths(currentmeasure) = MeasureLengths(currentmeasure) + ((FP8(x).Offs - LastMeasure) / NML)
                    LastMeasure = FP8(x).Offs
                    LastRealMeasure = LastMeasure
                    LastCheck = FP8(x).Offs
                    currentmeasure = currentmeasure + 1
                End If
            Next x
    End Select
    
    'internal conversion stage
    currentoffset = 0
    CurrentMetric = 0
    currentmeasure = 0
    LastMeasure = 0
    NML = 0
    
    Select Case FileMode
        Case FileMode_IIDXAC, FileMode_IIDXCS
            For x = LBound(F8) To UBound(F8)
                With F8(x)
                
                    If bEndFlag Then
                        CreateNote = False
                        Exit For
                    End If
                    
                    If NML > 0 Then
                        If .Var1 = 7 And .Var3 = 1 Then
                            .Offs = .Offs + 1
                        End If
                        tNote.OffsetMetric = ((.Offs - currentoffset) / NML) + CurrentMetric
                        tNote.OffsetMS = .Offs
                        tNote.OffsetMS = tNote.OffsetMS * TimingConstant
                        tNote.OffsetInternal = .Offs
                        If MeasureSizes(currentmeasure) > 0 Then
                            tNote.OffsetBMSMetric = ((tNote.OffsetMetric - LastMeasureMetric) / MeasureLengths(currentmeasure)) + currentmeasure
                        Else
                            
                        End If
                    End If
                    tNote.value = 0
                    tNote.xType = InternalType_Bad
                    tNote.Lane = 0
                    
                    CreateNote = True
                    
                    Select Case .Var1
                        Case 0   '1P note
                            tNote.Lane = .Var2
                            tNote.xType = 0
                            If tNote.Lane = 8 Then
                                tNote.value = DefaultNote(7) 'free scratch zone
                            Else
                                tNote.value = DefaultNote(tNote.Lane)
                            End If
                            tNote.xType = InternalType_Note
                            notecount(0) = notecount(0) + 1
                        Case 1   '2P note
                            tNote.Lane = .Var2 + 32
                            tNote.value = DefaultNote(tNote.Lane)
                            tNote.xType = InternalType_Note
                            notecount(1) = notecount(1) + 1
                            PlayerCount = 2
                        Case 2   '1P change
                            tNote.Lane = .Var2
                            tNote.value = .Var3
                            If tNote.value = 0 And FileMode = FileMode_IIDXCS And tNote.OffsetMS > 0 Then
                                tNote.value = 8
                            End If
                            tNote.xType = InternalType_NoteChange
                            DefaultNote(tNote.Lane) = tNote.value
                            If tNote.value > 0 And tNote.value <= 1295 Then
                                KeysoundUsed(tNote.value) = True
                            End If
                        Case 3   '2P change
                            tNote.Lane = .Var2 + 32
                            tNote.value = .Var3
                            tNote.xType = InternalType_NoteChange
                            DefaultNote(tNote.Lane) = tNote.value
                            If tNote.value > 0 And tNote.value <= 1295 Then
                                KeysoundUsed(tNote.value) = True
                            End If
                        Case 4   'BPM
                            tNote.value = .Var3 / .Var2
                            tNote.xType = InternalType_BPM
                            If tNote.value > 0 And tNote.value <> CurrentBPM Then
                                z = 0
                                If tNote.value = Int(tNote.value) And tNote.value < 256 Then
                                    tNote.xType = InternalType_BPM03
                                Else
                                    For y = 1 To BPMCount
                                        If BPMlist(y) = tNote.value Then
                                            z = y
                                            Exit For
                                        End If
                                    Next y
                                    If z = 0 Then
                                        BPMCount = BPMCount + 1
                                        BPMlist(BPMCount) = tNote.value
                                        z = BPMCount
                                    End If
                                End If
                                NML = GetNML(tNote.value, TimingConstant)
                                currentoffset = .Offs
                                CurrentMetric = tNote.OffsetMetric
                                CurrentBPM = tNote.value
                                If tNote.xType = InternalType_BPM Then
                                    tNote.temp1 = z
                                End If
                            Else
                                CreateNote = False
                            End If
                        Case 5   'measure size
                            CreateNote = False
                        Case 6   'end of song
                            tNote.xType = InternalType_EndSong
                            SongLength = Int(tNote.OffsetMS / 1000) + 1
                            bEndFlag = True
                        Case 7   'BGM note
                            tNote.xType = InternalType_Note
                            tNote.Lane = 64
                            tNote.value = .Var3
                            notecount(2) = notecount(2) + 1
                            If tNote.value = 1 And FileMode <> FileMode_IIDXAC Then
                                If KeysoundUsed(1) = True Then
                                    CreateNote = False
                                End If
                            End If
                            If tNote.value = 0 And FileMode = FileMode_IIDXCS Then
                                tNote.value = 8
                            End If
                            If tNote.value > 0 And tNote.value <= 1295 Then
                                KeysoundUsed(tNote.value) = True
                            End If
                        Case 8   'judgement adjustment
                            CreateNote = False
                        Case 12  'measure
                            tNote.xType = InternalType_Measure
                            currentmeasure = currentmeasure + 1
                            CurrentMetric = tNote.OffsetMetric
                            currentoffset = .Offs
                            LastMeasure = .Offs
                            LastMeasureMetric = tNote.OffsetMetric
                        Case 16  'note count
                            CreateNote = False
                        Case 255 'bad
                            CreateNote = False
                        Case Else 'wtf?
                            CreateNote = False
                            Debug.Print "UNKNOWN EVENT:", Hex(.Var1)
                    End Select
                End With
                If CreateNote Then
                    tNote.xEnabled = -1
                    SimFileDataCount = SimFileDataCount + 1
                    ReDim Preserve SimFileData(0 To SimFileDataCount) As FormatInternal
                    SimFileData(SimFileDataCount) = tNote
                End If
            Next x
        
        
        
        
        Case FileMode_POPN
        
            For x = LBound(FP8) To UBound(FP8)
                With FP8(x)
                
                    If NML > 0 Then
                        tNote.OffsetMetric = ((.Offs - currentoffset) / NML) + CurrentMetric
                        tNote.OffsetMS = .Offs * TimingConstant
                        tNote.OffsetInternal = .Offs
                        Debug.Print .Offs, tNote.OffsetMS
                        If MeasureSizes(currentmeasure) > 0 Then
                            tNote.OffsetBMSMetric = ((tNote.OffsetMetric - LastMeasureMetric) / MeasureLengths(currentmeasure)) + currentmeasure
                        Else
                            
                        End If
                    End If
                    tNote.value = 0
                    tNote.xType = InternalType_Bad
                    tNote.Lane = 0
                    
                    CreateNote = True
                    
                    Select Case .Var1
                        Case 1   '1P note
                            tNote.Lane = .Var2
                            tNote.xType = 0
                            tNote.value = DefaultNote(tNote.Lane)
                            tNote.xType = InternalType_Note
                            notecount(0) = notecount(0) + 1
                        Case 2   '1P change
                            tNote.Lane = .Var2
                            tNote.value = .Var3
                            tNote.xType = InternalType_NoteChange
                            DefaultNote(tNote.Lane) = tNote.value
                            If tNote.value > 0 And tNote.value <= 1295 Then
                                KeysoundUsed(tNote.value) = True
                            End If
                        Case 3   'bgm start
                            tNote.xType = InternalType_Note
                            tNote.Lane = 64
                            tNote.value = 1295
                            KeysoundUsed(1295) = True
                        Case 4   'BPM
                            tNote.value = .Var3 ' / .Var2
                            tNote.xType = InternalType_BPM

'                           OLD METHOD

'                            NML = GetNML(tNote.Value, TimingConstant)
'                            currentoffset = .Offs
'                            CurrentMetric = tNote.OffsetMetric
'                            If tNote.Value = CurrentBPM Then
'                                CreateNote = False
'                            Else
'                                BPMCount = BPMCount + 1
'                                BPMlist(BPMCount) = tNote.Value
'                                CurrentBPM = tNote.Value
'                            End If
'                            If BPMCount > 1 Then
'                                tNote.temp1 = BPMCount
'                            End If
                            
                            If tNote.value > 0 And tNote.value <> CurrentBPM Then
                                z = 0
                                If tNote.value = Int(tNote.value) And tNote.value < 256 Then
                                    tNote.xType = InternalType_BPM03
                                Else
                                    For y = 1 To BPMCount
                                        If BPMlist(y) = tNote.value Then
                                            z = BPMlist(y)
                                            Exit For
                                        End If
                                    Next y
                                    If z = 0 Then
                                        BPMCount = BPMCount + 1
                                        BPMlist(BPMCount) = tNote.value
                                        z = BPMCount
                                    End If
                                End If
                                NML = GetNML(tNote.value, TimingConstant)
                                currentoffset = .Offs
                                CurrentMetric = tNote.OffsetMetric
                                CurrentBPM = tNote.value
                                If tNote.xType = InternalType_BPM Then
                                    tNote.temp1 = z
                                End If
                            Else
                                CreateNote = False
                            End If
                        Case 5   'measure size
                            CreateNote = False
                        Case 6   'end of song
                            tNote.xType = InternalType_EndSong
                            SongLength = Int(tNote.OffsetMS / 1000) + 1
                        Case 7   'BGM note
                            tNote.xType = InternalType_Note
                            tNote.Lane = 64
                            tNote.value = .Var3
                            notecount(2) = notecount(2) + 1
                            If tNote.value > 0 And tNote.value <= 1295 Then
                                KeysoundUsed(tNote.value) = True
                            End If
                        Case 8   'judgement adjustment
                            CreateNote = False
                        Case 9   'metronome?
                            CreateNote = False
                        Case 10, 0 'measure
                            If (bPopnAC15 = True) Or .Var1 = 10 Then
                                tNote.xType = InternalType_Measure
                                'Debug.Print "------------------------------------------------"
                                currentmeasure = currentmeasure + 1
                                CurrentMetric = tNote.OffsetMetric
                                currentoffset = .Offs
                                LastMeasure = .Offs
                                LastMeasureMetric = tNote.OffsetMetric
                            Else
                                CreateNote = False
                            End If
                        Case 11 'metronome
                            CreateNote = False
                        Case 255 'bad
                            CreateNote = False
                        Case Else 'wtf?
                            CreateNote = False
                            Debug.Print "UNKNOWN EVENT:", Hex(.Var1)
                    End Select
                End With
                If CreateNote Then
                    tNote.xEnabled = -1
                    SimFileDataCount = SimFileDataCount + 1
                    ReDim Preserve SimFileData(0 To SimFileDataCount) As FormatInternal
                    SimFileData(SimFileDataCount) = tNote
                End If
            Next x
        
            
        
    End Select
    xCSSize = LOF(f)
    Close #f
    
    'post-sorting 1 (types)
    sortflag = True
    Do Until sortflag = False
        sortflag = False
        For x = LBound(SimFileData) + 1 To UBound(SimFileData)
            If SimFileData(x).xType < SimFileData(x - 1).xType Then
                TempInternal = SimFileData(x)
                SimFileData(x) = SimFileData(x - 1)
                SimFileData(x - 1) = TempInternal
                sortflag = True
            End If
        Next x
    Loop
    
    'post-sorting 2 (offsets)
    sortflag = True
    Do Until sortflag = False
        sortflag = False
        For x = LBound(SimFileData) + 1 To UBound(SimFileData)
            If SimFileData(x).OffsetMS < SimFileData(x - 1).OffsetMS Then
                TempInternal = SimFileData(x)
                SimFileData(x) = SimFileData(x - 1)
                SimFileData(x - 1) = TempInternal
                sortflag = True
            End If
        Next x
    Loop
    
    bLoaded = (NML > 0)
    CurrentMode = FileMode
    
    Close #f
    
End Sub

Private Function GetNML(InBPM As Double, TimingConstant As Double) As Double
    GetNML = (((60 / InBPM) * 4) * (1000 / TimingConstant))
End Function

Private Function BME(ByVal inval As Integer) As String
    Dim res As String
    res = Mid$(BMEString, (inval \ Len(BMEString)) + 1, 1)
    res = res + Mid$(BMEString, (inval Mod Len(BMEString)) + 1, 1)
    BME = res
End Function

Private Function BMS(ByVal inval As Integer) As String
    Dim res As String
    inval = inval And 255
    res = Hex(inval)
    If Len(res) = 1 Then
        res = "0" + res
    End If
    BMS = res
End Function

Private Function Dec2(ByVal inval As Integer) As String
    Dim res As String
    res = CStr(inval Mod 100)
    If Len(res) = 1 Then
        res = "0" + res
    End If
    Dec2 = res
End Function

Private Function BMSLine(ByVal intype As Long, ByVal inmeasure As Long, ByVal inlane As Long, ByVal quant As Long, ByVal numbering As Long) As String
    Dim res As String
    Dim x As Long
    Dim cm As Long
    Dim p As Double
    Dim z As Long
    cm = 0
    res = String(quant * 2, "0")
    For x = LBound(SimFileData) To UBound(SimFileData)
        With SimFileData(x)
            If Int(.OffsetBMSMetric) = inmeasure Then
                If .Lane = inlane Or inlane = -1 Then
                    If .xType = intype Or intype = -1 Then
                        If .xEnabled <> 0 Then
                            p = .OffsetBMSMetric - Int(.OffsetBMSMetric)
                            p = p * quant
                            p = Round(p)
                            p = p * 2
                            p = p + 1
                            If p > Len(res) Then
                                 p = Len(res) - 1
                            End If
                            If Mid$(res, p, 2) = "00" Then
                                Select Case numbering
                                    Case KeyNumbering_BME
                                        Mid$(res, p, 2) = BME(.value)
                                    Case KeyNumbering_BMS
                                        Mid$(res, p, 2) = BMS(.value)
                                    Case KeyNumbering_Decimal
                                        Mid$(res, p, 2) = Dec2(.temp1)
                                        If .temp1 <> 0 Then
                                            x = x
                                        End If
                                End Select
                                .xEnabled = 0
                            End If
                        End If
                    End If
                End If
            End If
            If Int(.OffsetBMSMetric) > inmeasure Then
                Exit For
            End If
        End With
    Next x
    BMSLine = res
End Function

Function ReduceLine(ByVal xInLine As String) As String
    Dim x As Long
    Dim y As Long
    Dim z As String
    Dim a As String
    Dim b As String
    Dim f As Boolean
    
    x = (Len(xInLine) \ 4)
    If xInLine = "" Then
        Exit Function
    End If
    
    Do While x >= 2
        If Len(xInLine) Mod (x * 2) = 0 Then
            b = String((x - 1) * 2, "0")
            For y = 1 To Len(xInLine) - 1 Step (x * 2)
                f = True
                z = Mid$(xInLine, y + 2, (x - 1) * 2)
                If z <> b Then
                    f = False
                    Exit For
                End If
            Next y
            If f = True Then
                For y = 1 To Len(xInLine) - 1 Step (x * 2)
                    a = a + Mid$(xInLine, y, 2)
                Next y
                x = (Len(xInLine) \ 4)
                xInLine = a
                a = ""
            Else
                x = x - 1
            End If
        Else
            x = x - 1
        End If
    Loop
    If xInLine = "00" Then
        xInLine = ""
    End If
    ReduceLine = xInLine
End Function

Public Function FrameTiming() As Double
    FrameTiming = xFrameTiming
End Function
'
Public Function KeysoundFileName(KeyNumber As Long) As String
    If KeyNumber > 1295 Or KeyNumber < 0 Then
        Exit Function
    End If
    If KeysoundUsed(KeyNumber) Then
        If KeysoundFile(KeyNumber) = "" Then
            KeysoundFileName = BME(KeyNumber) + ".wav"
        Else
            KeysoundFileName = KeysoundFile(KeyNumber)
        End If
    End If
End Function

Public Function BMEToDec(inBME As String) As Long
    BMEToDec = ((InStr(BMEString, Left$(inBME, 1)) - 1) * Len(BMEString)) + (InStr(BMEString, Mid$(inBME, 2, 1)) - 1)
End Function

Public Function FixFileName(inFileName As String) As String
    Dim x As Long
    Dim y As Long
    Dim s As Long
    s = InStrRev(inFileName, "\") + 1
    FixFileName = inFileName
    For x = s To Len(inFileName)
        For y = 1 To Len(BadFileName)
            If Mid$(FixFileName, x, 1) = Mid$(BadFileName, y, 1) Then
                Mid$(FixFileName, x, 1) = "_"
                Exit For
            End If
        Next y
    Next x
End Function

Public Function CSSize() As Long
    CSSize = xCSSize
End Function
