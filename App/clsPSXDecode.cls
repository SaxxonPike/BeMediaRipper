VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPSXDecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ========================================================
' Playstation VAG/XA decoder wrapper
' Plus BM RAW and IIDX RAW conversion
'  SaxxonPike 2006-2008
' ========================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub DecodeVAG Lib "BeMedia2.dll" (in16 As Any, out28 As Any, ByRef p1 As Long, ByRef p2 As Long)

Private Const SilenceConstant = 64
Private Const ExportRaw = False
Private Const ExportConverted = True

Private bBusy As Boolean
Private lConVol As Long
Private xThruMode As Boolean
Private xAutoMono As Boolean

Public Property Get AutoMono() As Boolean
    AutoMono = xAutoMono
End Property
Public Property Let AutoMono(i As Boolean)
    xAutoMono = i
End Property
Public Property Get ConvertVolume() As Long
    ConvertVolume = lConVol
End Property
Public Property Let ConvertVolume(i As Long)
    lConVol = i
End Property
Public Property Let ThruMode(i As Boolean)
    xThruMode = i
End Property

Public Function DJMainToWav(ByRef inFS As clsFileStream, ByVal OutFile As String, ByVal offs As Double, ByVal Leng As Double, ByVal freq As Long, ByVal Channels As Long, ByVal BitStyle As Long, Optional ByVal volume As Byte = &H7F, Optional ByVal panning As Byte = &H40, Optional ByVal bStereo As Boolean = False) As Long
    Dim X As Long
    Dim Y As Long
    Dim bConverted As Boolean

    Dim Coffs As Double
    Dim COffs2 As Double
    Dim BGJump As Double
    Dim CJump As Double
    Dim CheckBytes(0 To 31) As Byte
    Dim Unswapped(0 To 31) As Byte
    Dim HeaderBytes() As Byte
    Dim RawWave() As Byte
    Dim RawWave2() As Byte
    Dim ExpWave() As Integer
    Dim ExpWave2() As Integer
    Dim CombWave() As Integer
    Dim IDNum As Long
    Dim keynum As Long
    Dim Kv1 As Long
    Dim Kv2 As Long
    Dim KeyOffs As Double
    Dim WriteWave As Boolean
    Dim WaveHead() As Byte
    Dim DoSoundReadRefresh As Boolean
    Dim SongEnd As Boolean
    Dim z As Byte
    Dim f1 As Long
    Dim f2 As Long
    
    
    WaveHead = LoadResData(101, "CUSTOM")
    WaveHead(32) = 4
    WaveHead(22) = 2
    
    Dim NeedsByteSwap As Boolean
    Dim AudioBits As Byte

    KeyOffs = offs
    COffs2 = offs
    inFS.ReadFileBinary Unswapped(), , COffs2
    For Y = 0 To 31 Step 2
        CheckBytes(Y + 1) = Unswapped(Y)
        CheckBytes(Y) = Unswapped(Y + 1)
    Next Y
    z = 0
    For X = 0 To &H67FFFF
        z = z + 1
        If z = 8 Then
            inFS.ReadFileObject VarPtr(Unswapped(16)), 16, COffs2 + 16
            For Y = 16 To 31 Step 2
                CheckBytes(Y + 1) = Unswapped(Y)
                CheckBytes(Y) = Unswapped(Y + 1)
            Next Y
            z = 0
        End If
        
        If CheckBytes(0) = &H80 And CheckBytes(1) = &H80 And CheckBytes(2) = &H80 And CheckBytes(3) = &H80 Then
            '4-bit signed
            If (COffs2 - KeyOffs) > 0 Then
                inFS.ReadFileBinary RawWave(), (COffs2 - KeyOffs), KeyOffs, True
            Else
                ReDim RawWave(0) As Byte
            End If
            WriteWave = True
            COffs2 = COffs2 + 4
            AudioBits = 8
            NeedsByteSwap = True
        ElseIf (Unswapped(0) = &H88 And Unswapped(1) = &H88 And Unswapped(2) = &H88 And Unswapped(3) = &H88) Or _
        CheckBytes(0) = &H88 And CheckBytes(1) = &H88 And CheckBytes(2) = &H88 And CheckBytes(3) = &H88 Then
            '4-bit signed 2
            If (COffs2 - KeyOffs) > 0 Then
                inFS.ReadFileBinary RawWave(), (COffs2 - KeyOffs), KeyOffs, True
            Else
                ReDim RawWave(0) As Byte
            End If
            WriteWave = True
            COffs2 = COffs2 + 4
            AudioBits = 8
            NeedsByteSwap = True
            freq = 22050
        ElseIf CheckBytes(0) = &H80 And CheckBytes(1) = &H80 And CheckBytes(2) = &H80 And CheckBytes(3) = &H80 Then
            '8-bit signed
            If (COffs2 - KeyOffs) > 0 Then
                inFS.ReadFileBinary RawWave(), (COffs2 - KeyOffs), KeyOffs, True
            Else
                ReDim RawWave(0) As Byte
            End If
            WriteWave = True
            COffs2 = COffs2 + 8
            AudioBits = 8
            NeedsByteSwap = True
        ElseIf CheckBytes(0) = &H80 And CheckBytes(1) = 0 And CheckBytes(2) = &H80 And CheckBytes(3) = 0 And CheckBytes(4) = &H80 And CheckBytes(5) = 0 And CheckBytes(6) = &H80 And CheckBytes(7) = 0 And CheckBytes(8) = &H80 And CheckBytes(9) = 0 And CheckBytes(10) = &H80 And CheckBytes(11) = 0 And CheckBytes(12) = &H80 And CheckBytes(13) = 0 And CheckBytes(14) = &H80 And CheckBytes(15) = 0 Then
            '16-bit intel
            If (COffs2 - KeyOffs) > 0 Then
                inFS.ReadFileBinary RawWave(), COffs2 - KeyOffs, KeyOffs, True
            Else
                ReDim RawWave(0) As Byte
            End If
            WriteWave = True
            COffs2 = COffs2 + 16
            AudioBits = 16
        ElseIf CheckBytes(1) = &H80 And CheckBytes(0) = 0 And CheckBytes(3) = &H80 And CheckBytes(2) = 0 And CheckBytes(5) = &H80 And CheckBytes(4) = 0 And CheckBytes(7) = &H80 And CheckBytes(6) = 0 And CheckBytes(9) = &H80 And CheckBytes(8) = 0 And CheckBytes(11) = &H80 And CheckBytes(10) = 0 And CheckBytes(13) = &H80 And CheckBytes(12) = 0 And CheckBytes(15) = &H80 And CheckBytes(14) = 0 Then
            '16-bit motorola (reversed)
            If (COffs2 - KeyOffs) > 0 Then
                inFS.ReadFileBinary RawWave(), COffs2 - KeyOffs, KeyOffs, True
            Else
                ReDim RawWave(0) As Byte
            End If
            WriteWave = True
            COffs2 = COffs2 + 16
            NeedsByteSwap = True
            AudioBits = 16
        End If
        If WriteWave = True Then
            SongEnd = True
            DoSoundReadRefresh = True
            If UBound(RawWave) > 512 Then
                If NeedsByteSwap Then
                    ByteSwap RawWave()
                    'ByteSwap RawWave2()
                    NeedsByteSwap = False
                End If
                    f1 = FreeFile
                    Open OutFile For Binary As #f1
                    f2 = FreeFile
                    Open OutFile + ".raw" For Binary As #f2
                    Put #f2, 1, RawWave
                    Close #f2
                    UpSample AudioBits + 0, RawWave(), ExpWave()
                    'UpSample AudioBits + 0, RawWave2(), ExpWave2()
                    ExpWave2() = ExpWave()
                    CombineStreams ExpWave(), ExpWave2(), CombWave()
                    Put #f1, 1, WaveHead
                    Put #f1, , CombWave
                    Kv1 = LOF(f1) - (UBound(WaveHead) + 1)
                    Kv2 = LOF(f1) - 8
                    Put #f1, 5, Kv2
                    Put #f1, UBound(WaveHead) - 2, Kv1
                    Kv1 = freq * 4
                    Put #f1, 29, Kv1
                    Kv1 = freq
                    Put #f1, 25, Kv1
                    Close #f1
                    Exit For
            ElseIf UBound(RawWave) > 0 Then
                Exit For
            End If
        Else
            COffs2 = COffs2 + 2
            If COffs2 Mod 32768 = 0 Then
                'Debug.Print COffs2
                DoEvents
            End If
        End If
        
        For Y = 0 To 29
            CheckBytes(Y) = CheckBytes(Y + 2)
            Unswapped(Y) = Unswapped(Y + 2)
        Next Y
    
    Next X

    'Set OutWave = Nothing
    bBusy = False
End Function

Private Sub CombineStreams(ByRef inL() As Integer, ByRef inR() As Integer, ByRef outLR() As Integer)
    Dim sSize As Long
    Dim X As Long
    sSize = (UBound(inL) + 1)
    ReDim outLR(0 To (sSize * 2) - 1) As Integer
    For X = 0 To sSize - 2
        outLR((X * 2) + 0) = inL(X)
        outLR((X * 2) + 1) = inR(X)
    Next X
End Sub

Public Function TwinkleToWAV(ByRef inFS As clsFileStream, ByVal OutFile As String, ByVal offs As Double, ByVal Leng As Double, ByVal freq As Long, Optional ByVal volume As Byte = &H7F, Optional ByVal panning As Byte = &H40, Optional ByVal bStereo As Boolean = False) As Long
    bBusy = True
    Dim iData() As Integer
    Dim wData() As Integer
    Dim outWave As New clsWAVFile
    Dim X As Long
    Dim f As Long
    Dim a As Long
    Dim b As Long
    Dim VolAmp As Double
    Dim VolChan(0 To 1) As Double
    Dim opf As Long
    Dim StereoSize As Long
    Dim ChannelCount As Long
    Dim SampleSize As Long
        
    VolChan(0) = 2 - (panning / &H40&)
    VolChan(1) = panning / &H40&
    VolAmp = volume / &H3F
    If VolAmp > 1 Then VolAmp = 1
    VolAmp = (VolAmp * (lConVol / 100)) / 2
    'If VolChan(0) > 1 Then VolChan(0) = 1
    'If VolChan(1) > 1 Then VolChan(1) = 1
    'If VolChan(0) < 0.02 Then VolChan(0) = 0
    'If VolChan(1) < 0.02 Then VolChan(1) = 0
    
    'make it properly divisible
    If bStereo Then
        Leng = (Leng \ 4) * 4
    End If
    
    ReDim iData(0 To (Leng \ 2) - 1) As Integer
    inFS.ReadFileObject VarPtr(iData(0)), (UBound(iData) + 1) * 2, offs
    
    If bStereo Then
        StereoSize = UBound(iData)
        ChannelCount = 2
        SampleSize = UBound(iData)
    Else
        StereoSize = (UBound(iData) * 2) + 1
        ChannelCount = 1
        SampleSize = UBound(iData)
    End If
    
    ReDim wData(0 To StereoSize) As Integer
    'convert data
    For X = 0 To UBound(iData)
        If iData(X) < 0 Then
            iData(X) = -(iData(X) And &H7FFF&)
        End If
    Next X
    'apply stereo effects
    For X = 0 To SampleSize Step ChannelCount
        a = iData(X) * VolChan(0) * VolAmp
        If Not bStereo Then
            b = iData(X) * VolChan(1) * VolAmp
        Else
            b = iData(X + 1) * VolChan(1) * VolAmp
        End If
        If a > 32767 Then a = 32767
        If a < -32768 Then a = -32768
        If b > 32767 Then b = 32767
        If b < -32768 Then b = -32768
        If Not bStereo Then
            wData((X * 2) + 0) = a
            wData((X * 2) + 1) = b
        Else
            wData(X) = a
            wData(X + 1) = b
        End If
        If Not bBusy Then
            Exit Function
        End If
    Next X
    
    DoEvents
    f = FreeFile
    Open OutFile For Binary As #f
    Put #f, &H28 + 1, f
    Put #f, , wData
    Close #f
    outWave.Channels = 2
    outWave.Frequency = freq
    outWave.WriteHeader OutFile, ((Leng \ 2) * 2)
    Set outWave = Nothing
    bBusy = False
End Function

Public Function VAGRipSimple(ByRef inFS As clsFileStream, ByVal OutFile As String, ByVal offs As Double, ByVal freq As Long, ByVal Blocksize As Long, ByVal ChannelCount As Long, Optional ByVal RemoveLeadingSilence As Boolean = False, Optional ByVal SimultaneousReads As Long = 1, Optional ByVal panning As Byte = &H40, Optional ByVal volume As Byte = &H7F, Optional ByVal EndOnBlankLine As Boolean = False, Optional ByVal XORValue As Long = 0, Optional ByVal EndOnBlankLine2 As Boolean = False, Optional ByVal XORType As Long = 0) As Long
    '--------------------------------------------------
    'xThruMode = False
    If xThruMode = True Then
        Exit Function
    End If
    
    Dim chan As Long
    Dim VolAmp As Double
    Dim VolChan(0 To 1) As Double
    Dim MakeStereo As Boolean
    Dim Y As Long
    Dim X As Long
    Dim f As Long
    Dim z As Long
    Dim iloc As Long
    Dim BufferSize As Long
    Dim decodeoffset As Long
    Dim decodedx As Long
    Dim totaldecoded As Long
    Dim OldOffset As Double
    Dim OldAdvanceSetting As Boolean
    Dim Decoded() As Integer
    Dim DecodeBuffer() As Integer
    Dim InputBuffer() As Byte
    Dim InvalidFlag As Boolean
    Dim Filter0() As Long
    Dim Filter1() As Long
    Dim XORvar As Long
    Dim EndFlag As Boolean
    Dim outWave As New clsWAVFile
    Dim TempBufOffs As Long
    Dim InvertInt As Integer
    Dim InvertThing As Integer
    Dim InvertInt2 As Integer
    Dim TempConv As Long
    Dim FilterString As String
    Dim OldBufferSize As Long
    Dim breakcounter As Long
    Dim xflag As Boolean
    '--------------------------------------------------
    bBusy = True
    OldOffset = inFS.offset
    inFS.offset = offs
    OldAdvanceSetting = inFS.AdvanceOffset
    inFS.AdvanceOffset = True
    OldBufferSize = inFS.BufferSize
    inFS.BufferSize = 512
    'If Volume > &H7F Then
    '    Volume = &H7F
    'End If
    If panning > &H7F Then
        panning = &H40
    End If
    '--------------------------------------------------
    'If Panning <> &H40 And ChannelCount = 1 Then
    '    VolAmp = Panning
    '    VolAmp = (((&H40 - Abs(VolAmp - &H40)) / &H40) + 1) * Volume
    'End If
    '--------------------------------------------------
    VolChan(0) = 2 - (panning / &H40&)
    VolChan(1) = panning / &H40&
    'If VolChan(0) > 1 Then VolChan(0) = 1
    'If VolChan(1) > 1 Then VolChan(1) = 1
    VolAmp = volume / &H7F
    'If VolAmp > 1 Then VolAmp = 1
    VolAmp = VolAmp * (lConVol / 100)
    '--------------------------------------------------
    BufferSize = (ChannelCount * Blocksize) * SimultaneousReads
    If panning <> &H40 And ChannelCount = 1 Then
        MakeStereo = True
        ChannelCount = 2
    End If
    If AutoMono = False Then
        If ChannelCount = 1 Then
            MakeStereo = True
            ChannelCount = 2
        End If
    End If
    '--------------------------------------------------
    If ChannelCount = 0 Then
        Exit Function
    End If
    
    ReDim DecodeBuffer(0 To (ChannelCount * 28) - 1) As Integer
    ReDim Filter0(1 To ChannelCount) As Long
    ReDim Filter1(1 To ChannelCount) As Long
    ReDim Decoded(0 To (28 * ChannelCount) - 1) As Integer
    ReDim InputBuffer(0 To BufferSize - 1) As Byte
    '--------------------------------------------------
    FilterString = "*?/><|"
    For X = 1 To Len(FilterString)
        OutFile = Replace(OutFile, Mid(FilterString, X, 1), "_")
    Next X
    '--------------------------------------------------
    If Dir(OutFile) <> "" Then
        Kill OutFile
    End If
    '--------------------------------------------------
    f = FreeFile
    If ExportRaw Then
        If Dir(OutFile + ".buf") <> "" Then
            Kill OutFile + ".buf"
        End If
        Open OutFile + ".buf" For Binary As (f + 1)
    End If
    '--------------------------------------------------
    On Local Error GoTo 71
    If ExportConverted Then
        Open OutFile For Binary As #f
        Put #f, &H28 + 1, f
    End If
    On Local Error GoTo 0
    
    Do While (Not EndFlag) And bBusy And (Not InvalidFlag)
        If inFS.ReadFileBinary(InputBuffer()) = True Then
            InvalidFlag = True
        End If
    '--------------------------------------------------
        If XORValue <> 0 Then
            For iloc = 0 To (UBound(InputBuffer) \ 16) * 16 Step 16
                '------------------------------------------------------
                'decode: part 2
                '------------------------------------------------------
                Select Case XORType
                    '------------------------------------------------------
                    'RED to DistorteD
                    '------------------------------------------------------
                    Case 0
                        If InputBuffer(iloc + 1) = 4 Then
                            XORValue = XORValue And &HFFFF00FF
                        End If
                        CopyMemory XORvar, InputBuffer(iloc), 4
                        XORvar = XORvar Xor XORValue
                        CopyMemory InputBuffer(iloc), XORvar, 4
                        InvertThing = InputBuffer(iloc + 2)
                        If InvertThing <> 0 Then
                            If InvertThing And &H10 Then
                                InvertThing = &H120 - InvertThing
                            Else
                                InvertThing = &H100 - InvertThing
                            End If
                            If InvertThing < 0 Then InvertThing = InvertThing + &H100
                            If InvertThing > 255 Then InvertThing = InvertThing - &H100
                            InputBuffer(iloc + 2) = InvertThing
                        End If
                        InvertThing = InputBuffer(iloc + 3)
                        If InvertThing <> 0 And InputBuffer(iloc + 1) <= 1 Then
                            InvertInt = &H100
                            If InvertThing And &H8 Then
                                InvertInt = InvertInt + &H10
                            End If
                            If InvertThing And &H40 Then
                                InvertInt = InvertInt - &H80
                            End If
                            InvertThing = InvertInt - InvertThing
                            If InvertThing < 0 Then InvertThing = InvertThing + &H100
                            If InvertThing > 255 Then InvertThing = InvertThing - &H100
                        End If
                        InputBuffer(iloc + 3) = InvertThing
                    '------------------------------------------------------
                    'GOLD BGM
                    '------------------------------------------------------
                    Case 1
                        CopyMemory XORvar, InputBuffer(iloc), 4
                        XORvar = XORvar Xor XORValue
                        CopyMemory InputBuffer(iloc), XORvar, 4
                        InvertThing = InputBuffer(iloc + 2)
                        If InvertThing <> 0 Then
                            If InvertThing And &H2 Then
                                InvertThing = InvertThing Xor &HFC&
                            ElseIf InvertThing And &H4 Then
                                InvertThing = InvertThing Xor &HF8&
                            ElseIf InvertThing And &H8 Then
                                InvertThing = InvertThing Xor &HF0&
                            ElseIf InvertThing And &H10 Then
                                InvertThing = InvertThing Xor &HE0&
                            ElseIf InvertThing And &H20 Then
                                InvertThing = InvertThing Xor &HC0&
                            ElseIf InvertThing And &H40 Then
                                InvertThing = InvertThing Xor &H80&
                            End If
                            InputBuffer(iloc + 2) = InvertThing
                        End If
                    '------------------------------------------------------
                    'GOLD Keysounds
                    '------------------------------------------------------
                    Case 2
                        CopyMemory XORvar, InputBuffer(iloc), 4
                        XORvar = XORvar Xor XORValue
                        CopyMemory InputBuffer(iloc), XORvar, 4
                        InvertThing = InputBuffer(iloc + 2)
                        InvertInt = InvertThing + &H100&
                        If InvertThing <> 0 Then
                            'Debug.Print Hex(InvertInt), ;
                            If (InvertThing And &H4&) Then InvertInt = InvertInt - &H8&
                            If (InvertThing And &H8&) Then InvertInt = InvertInt - &H10&
                            If (InvertThing And &H10&) Then InvertInt = InvertInt - &H20&
                            'If (InvertThing And &H20&) Then InvertInt = InvertInt - &H40&
                            If (InvertThing And &H40&) Then InvertInt = InvertInt - &H80&
                            'Debug.Print Hex(InvertInt)
                            InvertInt = (InvertInt And &HFF&)
                            'Debug.Print Hex(InputBuffer(iloc + 2)), Hex(InvertInt)
                            InputBuffer(iloc + 2) = InvertInt
                        End If
                End Select
            Next iloc
        End If
        
        
    '--------------------------------------------------
        If ExportRaw Then
            Put (f + 1), TempBufOffs + 1, InputBuffer
            TempBufOffs = TempBufOffs + UBound(InputBuffer) + 1
        End If
    '--------------------------------------------------
        
        
        For z = 0 To SimultaneousReads - 1
            For decodeoffset = 0 To (Blocksize \ 16) - 1
                For chan = 1 To ChannelCount
                    If chan = 1 Or (Not MakeStereo) Then
                        If MakeStereo = False Then
                            iloc = ((chan - 1) * Blocksize) + (decodeoffset * 16) + (z * Blocksize * ChannelCount)
                        Else
                            iloc = ((chan - 1) * Blocksize) + (decodeoffset * 16) + (z * Blocksize)
                        End If
                        If iloc > UBound(InputBuffer) Then
                            Exit For
                        End If
                        If Not EndFlag Then
                            'DecodeVAG InputBuffer(iloc), DecodeBuffer((chan - 1) * 28), Filter0(chan), Filter1(chan)
                            DecodeVAG InputBuffer(iloc), DecodeBuffer(GetDecodeChan(chan - 1, ChannelCount) * 28), Filter0(chan), Filter1(chan)
                        Else
                            For X = 0 To 27
                                DecodeBuffer(GetDecodeChan(chan - 1, ChannelCount) * 28) = 0
                            Next X
                        End If
                        If totaldecoded > 0 Then
                            If EndOnBlankLine Then
                                EndFlag = True
                                If InputBuffer(iloc) <> 0 Then
                                    For X = 0 To 15
                                        If InputBuffer(iloc + X) <> 0 Then
                                            EndFlag = False
                                            Exit For
                                        End If
                                    Next X
                                End If
                            End If
                        End If
                        If ((InputBuffer(iloc + 1) And 1) = 1) Then 'And EndOnBlankLine = False Then
                            EndFlag = True
                            Exit For
                        End If
                    Else
                        For X = 0 To 27
                            DecodeBuffer((GetDecodeChan(chan - 1, ChannelCount) * 28) + X) = DecodeBuffer(X)
                        Next X
                    End If
                Next chan
                '---------------------------
                'postprocessing (volume/pan)
                '---------------------------
                For chan = 1 To ChannelCount
                    For X = 0 To 27
                        If EndFlag = True Then
                            DecodeBuffer((GetDecodeChan(chan - 1, ChannelCount) * 28) + X) = 0
                        Else
                            If DecodeBuffer((GetDecodeChan(chan - 1, ChannelCount) * 28) + X) <> 0 Then
                                TempConv = CLng(DecodeBuffer((GetDecodeChan(chan - 1, ChannelCount) * 28) + X)) * ((VolChan((chan - 1) Mod 2)) * VolAmp)
                                If TempConv < -32768 Then TempConv = -32768
                                If TempConv > 32767 Then TempConv = 32767
                                DecodeBuffer((GetDecodeChan(chan - 1, ChannelCount) * 28) + X) = TempConv
                            End If
                        End If
                    Next X
                Next chan
    '--------------------------------------------------
                For Y = 0 To 27
                    For chan = 1 To ChannelCount
                        Decoded((Y * ChannelCount) + (chan - 1)) = DecodeBuffer(((chan - 1) * 28) + Y)
                        If RemoveLeadingSilence Then
                            iloc = Decoded((Y * ChannelCount) + (chan - 1))
                            If iloc > SilenceConstant Or iloc < -SilenceConstant Then
                                RemoveLeadingSilence = False
                            End If
                        End If
                    Next chan
                Next Y
                If RemoveLeadingSilence = False And EndFlag = False Then
                    If ExportConverted Then
                        Put #f, totaldecoded + 1, Decoded
                    End If
                    totaldecoded = totaldecoded + (ChannelCount * 56)
                End If
                If EndFlag = True Then
                    Exit For
                End If
            Next decodeoffset
        Next z
        If EndOnBlankLine2 Then 'this one uses the end of the data instead of the beginning
            EndFlag = True
            'If InputBuffer(UBound(InputBuffer) - 15) <> 0 Then
                For X = UBound(InputBuffer) - 15 To UBound(InputBuffer)
                    If InputBuffer(X) <> 0 Then
                        EndFlag = False
                        Exit For
                    End If
                Next X
            'End If
            If (InputBuffer(0) = 1 Or InputBuffer(0) = 2) And InputBuffer(1) = 0 And InputBuffer(2) = 0 And InputBuffer(3) = 0 And InputBuffer(4) <> 0 Then
                EndFlag = True
            End If
        End If
        breakcounter = breakcounter + (((ChannelCount * 56) * (Blocksize \ 16)) * ChannelCount * SimultaneousReads)
        If breakcounter > &H400000 Then
            DoEvents
            breakcounter = breakcounter - &H400000
        End If
        'DoEvents
    '------------------------------------------------        DoEvents
    Loop
    '--------------------------------------------------
    If ExportConverted Then
        Close #f
    End If
    If ExportRaw Then
        Close (f + 1)
    End If
    '--------------------------------------------------
    
    outWave.Channels = ChannelCount
    outWave.Frequency = freq
    If ExportConverted Then
        outWave.WriteHeader OutFile, totaldecoded
    End If
    If totaldecoded = 0 Then
        Kill OutFile
    End If
    
72     Set outWave = Nothing
    inFS.offset = OldOffset
    inFS.AdvanceOffset = OldAdvanceSetting
    bBusy = False
    bRipping = False
    inFS.BufferSize = OldBufferSize
    VAGRipSimple = totaldecoded
    
    Exit Function
    
71     Debug.Print "VAGRipSimple Error: could not create " + OutFile
    Err.Clear
    On Local Error GoTo 0
    Close #f
    GoTo 72
End Function

Public Sub CancelOperation()
    bBusy = False
End Sub

Public Function DecodeBemani1(ByRef inFS As clsFileStream, ByRef outdata() As Byte, Optional ByVal offs As Double = -1) As Long
    
    On Error GoTo 616
    
    'internal
    Dim OldOffset As Double
    Dim BytesDecoded As Long
    Dim EndFlag As Boolean
    Dim OldAdvanceSetting As Boolean
    Dim Decoded() As Byte
    Dim DecodeOffs As Double
    Dim f As Long
    
    'decoder
    Dim i As Long
    Dim flags As Long
    Dim j As Long
    Dim c As Byte
    Dim t As Byte
    Dim src As Long
    Dim dst As Long
    Dim xloop As Boolean
    
    'code
    OldAdvanceSetting = inFS.AdvanceOffset
    OldOffset = inFS.offset
    DecodeOffs = OldOffset
    inFS.AdvanceOffset = False
    If offs <> -1 Then
        inFS.offset = offs
        DecodeOffs = offs
    End If
    
    Do
        Do
            flags = ShiftRight(flags, 1)
            If (flags And &H100&) = 0 Then
                flags = inFS.ReadFileByteFast(DecodeOffs) Or &HFF00&
                DecodeOffs = DecodeOffs + 1
            End If
            c = inFS.ReadFileByteFast(DecodeOffs)
            If (flags And 1&) = 0 Then
                AddByte Decoded(), c
                DecodeOffs = DecodeOffs + 1
                Exit Do
            End If
            If (c And &H80&) = 0 Then
                i = inFS.ReadFileByteFast(DecodeOffs + 1)
                DecodeOffs = DecodeOffs + 2
                i = i Or ShiftLeft(c And 3, 8)
                j = ShiftRight(c + 0, 2) + 2
                xloop = True
            End If
            If Not xloop Then
                DecodeOffs = DecodeOffs + 1
                If (c And &H40&) = 0 Then
                    i = (c And 15) + 1
                    j = (ShiftRight(c + 0, 4)) + 1 - 8
                    xloop = True
                End If
            End If
            If xloop Then
                xloop = False
                Do
                    If (UBound(Decoded) + 1) - i > 0 And (UBound(Decoded) + 1) - i < UBound(Decoded) Then
                        AddByte Decoded(), Decoded((UBound(Decoded) + 1) - i)
                    ElseIf (UBound(Decoded) + 1) - i = UBound(Decoded) Then
                        AddByte Decoded(), Decoded(UBound(Decoded))
                    Else
                        AddByte Decoded(), 0
                    End If
                    j = j - 1
                Loop While j >= 0
                Exit Do
            End If
            If c = 255 Then
                EndFlag = True
                Exit Do
            End If
            j = c - &HC0& + 7
            Do
                AddByte Decoded(), inFS.ReadFileByteFast(DecodeOffs)
                DecodeOffs = DecodeOffs + 1
                j = j - 1
            Loop While j >= 0
        Loop
    Loop While EndFlag = False
    BytesDecoded = UBound(Decoded) + 1
    
    ReDim outdata(LBound(Decoded) To UBound(Decoded))
    CopyMemory outdata(LBound(Decoded)), Decoded(LBound(Decoded)), BytesDecoded
    
617 inFS.offset = OldOffset
    inFS.AdvanceOffset = OldAdvanceSetting
    DecodeBemani1 = BytesDecoded
    Exit Function
    
616 Debug.Print "clsPSXDecode.DecodeBemani1: ERR DecOffs:" + CStr(DecodeOffs)
    BytesDecoded = 0
    On Error GoTo 0
End Function

Private Sub AddByte(ByRef inArr() As Byte, ByVal va As Byte)
    On Local Error GoTo 1
    inArr(0) = inArr(0)
    On Local Error GoTo 0
    ReDim Preserve inArr(LBound(inArr) To UBound(inArr) + 1) As Byte
    inArr(UBound(inArr)) = va
    Exit Sub
1   ReDim inArr(0) As Byte
    inArr(0) = va
    On Local Error GoTo 0
End Sub

Private Function ShiftRight(ByVal value As Long, ByVal ShiftCount As Long) As Long
' from VBSPEED
' by Jost Schwider, jost@schwider.de, 20011001
  If value And &H80000000 Then
    Select Case ShiftCount
    Case 0&:  ShiftRight = value
    Case 1&:  ShiftRight = &H40000000 Or (value And &H7FFFFFFF) \ &H2&
    Case 2&:  ShiftRight = &H20000000 Or (value And &H7FFFFFFF) \ &H4&
    Case 3&:  ShiftRight = &H10000000 Or (value And &H7FFFFFFF) \ &H8&
    Case 4&:  ShiftRight = &H8000000 Or (value And &H7FFFFFFF) \ &H10&
    Case 5&:  ShiftRight = &H4000000 Or (value And &H7FFFFFFF) \ &H20&
    Case 6&:  ShiftRight = &H2000000 Or (value And &H7FFFFFFF) \ &H40&
    Case 7&:  ShiftRight = &H1000000 Or (value And &H7FFFFFFF) \ &H80&
    Case 8&:  ShiftRight = &H800000 Or (value And &H7FFFFFFF) \ &H100&
    Case 9&:  ShiftRight = &H400000 Or (value And &H7FFFFFFF) \ &H200&
    Case 10&: ShiftRight = &H200000 Or (value And &H7FFFFFFF) \ &H400&
    Case 11&: ShiftRight = &H100000 Or (value And &H7FFFFFFF) \ &H800&
    Case 12&: ShiftRight = &H80000 Or (value And &H7FFFFFFF) \ &H1000&
    Case 13&: ShiftRight = &H40000 Or (value And &H7FFFFFFF) \ &H2000&
    Case 14&: ShiftRight = &H20000 Or (value And &H7FFFFFFF) \ &H4000&
    Case 15&: ShiftRight = &H10000 Or (value And &H7FFFFFFF) \ &H8000&
    Case 16&: ShiftRight = &H8000& Or (value And &H7FFFFFFF) \ &H10000
    Case 17&: ShiftRight = &H4000& Or (value And &H7FFFFFFF) \ &H20000
    Case 18&: ShiftRight = &H2000& Or (value And &H7FFFFFFF) \ &H40000
    Case 19&: ShiftRight = &H1000& Or (value And &H7FFFFFFF) \ &H80000
    Case 20&: ShiftRight = &H800& Or (value And &H7FFFFFFF) \ &H100000
    Case 21&: ShiftRight = &H400& Or (value And &H7FFFFFFF) \ &H200000
    Case 22&: ShiftRight = &H200& Or (value And &H7FFFFFFF) \ &H400000
    Case 23&: ShiftRight = &H100& Or (value And &H7FFFFFFF) \ &H800000
    Case 24&: ShiftRight = &H80& Or (value And &H7FFFFFFF) \ &H1000000
    Case 25&: ShiftRight = &H40& Or (value And &H7FFFFFFF) \ &H2000000
    Case 26&: ShiftRight = &H20& Or (value And &H7FFFFFFF) \ &H4000000
    Case 27&: ShiftRight = &H10& Or (value And &H7FFFFFFF) \ &H8000000
    Case 28&: ShiftRight = &H8& Or (value And &H7FFFFFFF) \ &H10000000
    Case 29&: ShiftRight = &H4& Or (value And &H7FFFFFFF) \ &H20000000
    Case 30&: ShiftRight = &H2& Or (value And &H7FFFFFFF) \ &H40000000
    Case 31&: ShiftRight = &H1&
    End Select
  Else
    Select Case ShiftCount
    Case 0&:  ShiftRight = value
    Case 1&:  ShiftRight = value \ &H2&
    Case 2&:  ShiftRight = value \ &H4&
    Case 3&:  ShiftRight = value \ &H8&
    Case 4&:  ShiftRight = value \ &H10&
    Case 5&:  ShiftRight = value \ &H20&
    Case 6&:  ShiftRight = value \ &H40&
    Case 7&:  ShiftRight = value \ &H80&
    Case 8&:  ShiftRight = value \ &H100&
    Case 9&:  ShiftRight = value \ &H200&
    Case 10&: ShiftRight = value \ &H400&
    Case 11&: ShiftRight = value \ &H800&
    Case 12&: ShiftRight = value \ &H1000&
    Case 13&: ShiftRight = value \ &H2000&
    Case 14&: ShiftRight = value \ &H4000&
    Case 15&: ShiftRight = value \ &H8000&
    Case 16&: ShiftRight = value \ &H10000
    Case 17&: ShiftRight = value \ &H20000
    Case 18&: ShiftRight = value \ &H40000
    Case 19&: ShiftRight = value \ &H80000
    Case 20&: ShiftRight = value \ &H100000
    Case 21&: ShiftRight = value \ &H200000
    Case 22&: ShiftRight = value \ &H400000
    Case 23&: ShiftRight = value \ &H800000
    Case 24&: ShiftRight = value \ &H1000000
    Case 25&: ShiftRight = value \ &H2000000
    Case 26&: ShiftRight = value \ &H4000000
    Case 27&: ShiftRight = value \ &H8000000
    Case 28&: ShiftRight = value \ &H10000000
    Case 29&: ShiftRight = value \ &H20000000
    Case 30&: ShiftRight = value \ &H40000000
    Case 31&: ShiftRight = &H0&
    End Select
  End If
End Function

Private Function ShiftLeft(ByVal value As Long, ByVal ShiftCount As Long) As Long
' from VBSPEED
' by Jost Schwider, jost@schwider.de, 20011001
  Select Case ShiftCount
  Case 0&
    ShiftLeft = value
  Case 1&
    If value And &H40000000 Then
      ShiftLeft = (value And &H3FFFFFFF) * &H2& Or &H80000000
    Else
      ShiftLeft = (value And &H3FFFFFFF) * &H2&
    End If
  Case 2&
    If value And &H20000000 Then
      ShiftLeft = (value And &H1FFFFFFF) * &H4& Or &H80000000
    Else
      ShiftLeft = (value And &H1FFFFFFF) * &H4&
    End If
  Case 3&
    If value And &H10000000 Then
      ShiftLeft = (value And &HFFFFFFF) * &H8& Or &H80000000
    Else
      ShiftLeft = (value And &HFFFFFFF) * &H8&
    End If
  Case 4&
    If value And &H8000000 Then
      ShiftLeft = (value And &H7FFFFFF) * &H10& Or &H80000000
    Else
      ShiftLeft = (value And &H7FFFFFF) * &H10&
    End If
  Case 5&
    If value And &H4000000 Then
      ShiftLeft = (value And &H3FFFFFF) * &H20& Or &H80000000
    Else
      ShiftLeft = (value And &H3FFFFFF) * &H20&
    End If
  Case 6&
    If value And &H2000000 Then
      ShiftLeft = (value And &H1FFFFFF) * &H40& Or &H80000000
    Else
      ShiftLeft = (value And &H1FFFFFF) * &H40&
    End If
  Case 7&
    If value And &H1000000 Then
      ShiftLeft = (value And &HFFFFFF) * &H80& Or &H80000000
    Else
      ShiftLeft = (value And &HFFFFFF) * &H80&
    End If
  Case 8&
    If value And &H800000 Then
      ShiftLeft = (value And &H7FFFFF) * &H100& Or &H80000000
    Else
      ShiftLeft = (value And &H7FFFFF) * &H100&
    End If
  Case 9&
    If value And &H400000 Then
      ShiftLeft = (value And &H3FFFFF) * &H200& Or &H80000000
    Else
      ShiftLeft = (value And &H3FFFFF) * &H200&
    End If
  Case 10&
    If value And &H200000 Then
      ShiftLeft = (value And &H1FFFFF) * &H400& Or &H80000000
    Else
      ShiftLeft = (value And &H1FFFFF) * &H400&
    End If
  Case 11&
    If value And &H100000 Then
      ShiftLeft = (value And &HFFFFF) * &H800& Or &H80000000
    Else
      ShiftLeft = (value And &HFFFFF) * &H800&
    End If
  Case 12&
    If value And &H80000 Then
      ShiftLeft = (value And &H7FFFF) * &H1000& Or &H80000000
    Else
      ShiftLeft = (value And &H7FFFF) * &H1000&
    End If
  Case 13&
    If value And &H40000 Then
      ShiftLeft = (value And &H3FFFF) * &H2000& Or &H80000000
    Else
      ShiftLeft = (value And &H3FFFF) * &H2000&
    End If
  Case 14&
    If value And &H20000 Then
      ShiftLeft = (value And &H1FFFF) * &H4000& Or &H80000000
    Else
      ShiftLeft = (value And &H1FFFF) * &H4000&
    End If
  Case 15&
    If value And &H10000 Then
      ShiftLeft = (value And &HFFFF&) * &H8000& Or &H80000000
    Else
      ShiftLeft = (value And &HFFFF&) * &H8000&
    End If
  Case 16&
    If value And &H8000& Then
      ShiftLeft = (value And &H7FFF&) * &H10000 Or &H80000000
    Else
      ShiftLeft = (value And &H7FFF&) * &H10000
    End If
  Case 17&
    If value And &H4000& Then
      ShiftLeft = (value And &H3FFF&) * &H20000 Or &H80000000
    Else
      ShiftLeft = (value And &H3FFF&) * &H20000
    End If
  Case 18&
    If value And &H2000& Then
      ShiftLeft = (value And &H1FFF&) * &H40000 Or &H80000000
    Else
      ShiftLeft = (value And &H1FFF&) * &H40000
    End If
  Case 19&
    If value And &H1000& Then
      ShiftLeft = (value And &HFFF&) * &H80000 Or &H80000000
    Else
      ShiftLeft = (value And &HFFF&) * &H80000
    End If
  Case 20&
    If value And &H800& Then
      ShiftLeft = (value And &H7FF&) * &H100000 Or &H80000000
    Else
      ShiftLeft = (value And &H7FF&) * &H100000
    End If
  Case 21&
    If value And &H400& Then
      ShiftLeft = (value And &H3FF&) * &H200000 Or &H80000000
    Else
      ShiftLeft = (value And &H3FF&) * &H200000
    End If
  Case 22&
    If value And &H200& Then
      ShiftLeft = (value And &H1FF&) * &H400000 Or &H80000000
    Else
      ShiftLeft = (value And &H1FF&) * &H400000
    End If
  Case 23&
    If value And &H100& Then
      ShiftLeft = (value And &HFF&) * &H800000 Or &H80000000
    Else
      ShiftLeft = (value And &HFF&) * &H800000
    End If
  Case 24&
    If value And &H80& Then
      ShiftLeft = (value And &H7F&) * &H1000000 Or &H80000000
    Else
      ShiftLeft = (value And &H7F&) * &H1000000
    End If
  Case 25&
    If value And &H40& Then
      ShiftLeft = (value And &H3F&) * &H2000000 Or &H80000000
    Else
      ShiftLeft = (value And &H3F&) * &H2000000
    End If
  Case 26&
    If value And &H20& Then
      ShiftLeft = (value And &H1F&) * &H4000000 Or &H80000000
    Else
      ShiftLeft = (value And &H1F&) * &H4000000
    End If
  Case 27&
    If value And &H10& Then
      ShiftLeft = (value And &HF&) * &H8000000 Or &H80000000
    Else
      ShiftLeft = (value And &HF&) * &H8000000
    End If
  Case 28&
    If value And &H8& Then
      ShiftLeft = (value And &H7&) * &H10000000 Or &H80000000
    Else
      ShiftLeft = (value And &H7&) * &H10000000
    End If
  Case 29&
    If value And &H4& Then
      ShiftLeft = (value And &H3&) * &H20000000 Or &H80000000
    Else
      ShiftLeft = (value And &H3&) * &H20000000
    End If
  Case 30&
    If value And &H2& Then
      ShiftLeft = (value And &H1&) * &H40000000 Or &H80000000
    Else
      ShiftLeft = (value And &H1&) * &H40000000
    End If
  Case 31&
    If value And &H1& Then
      ShiftLeft = &H80000000
    Else
      ShiftLeft = &H0&
    End If
  End Select
End Function

Private Function GetDecodeChan(inChan As Long, totalChan As Long)
    Select Case totalChan
        Case 4
            Select Case inChan
                Case 0: GetDecodeChan = 2
                Case 1: GetDecodeChan = 3
                Case 2: GetDecodeChan = 0
                Case 3: GetDecodeChan = 1
            End Select
        Case Else
            GetDecodeChan = inChan
    End Select
End Function

Private Function Match16(bytes16() As Byte, bst As String) As Boolean
    Match16 = (Get16(bytes16()) = bst)
End Function

Private Function Get16(bytes16() As Byte) As String
    Dim X As Long
    Dim s As String
    For X = 0 To 15
        s = Hex(bytes16(X))
        If Len(s) < 2 Then
            s = "0" + s
        End If
        Get16 = Get16 + s
    Next X
End Function

Sub ByteSwap(inArr() As Byte)
    Dim X As Long
    Dim Y As Byte
    For X = 0 To ((UBound(inArr) \ 2) * 2) - 1
        Y = inArr(X)
        inArr(X) = inArr(X + 1)
        inArr(X + 1) = Y
    Next X
End Sub

Function CheckEmpty(inArr() As Byte) As Boolean
    Dim X As Long
    CheckEmpty = True
    For X = LBound(inArr) To UBound(inArr)
        If inArr(X) <> 0 And inArr(X) <> 255 Then
            CheckEmpty = False
            Exit For
        End If
    Next X
End Function

Sub UpSample(inBits As Long, inArr() As Byte, ByRef outArr() As Integer)
    Dim X As Long
    Dim Y As Long
    Dim pv As Integer
    Dim nv As Integer
    If inBits = 4 Then
        ReDim outArr(0 To (UBound(inArr) * 2) + 1) As Integer
        For X = 0 To UBound(inArr)
            Y = ((inArr(X) \ &H10) * &H1000&)
            If Y > 32767 Then Y = Y - 65536
            Y = Y \ 2
            outArr((X * 2) + 0) = Y
            Y = ((inArr(X) And 15) * &H1000&)
            If Y > 32767 Then Y = Y - 65536
            Y = Y \ 2
            outArr((X * 2) + 1) = Y
        Next X
    ElseIf inBits = 8 Then
        ReDim outArr(0 To UBound(inArr)) As Integer
        For X = 0 To UBound(inArr)
            Y = (inArr(X))
            Y = Y * &H101
            If Y > 32767 Then Y = Y - 65536
            nv = Y
            Y = ((Y + pv) \ 2)
            pv = nv
            outArr(X) = Y
        Next X
        outArr(0) = 0
        outArr(1) = 0
        outArr(2) = 0
        outArr(3) = 0
        outArr(UBound(outArr) - 3) = 0
        outArr(UBound(outArr) - 2) = 0
        outArr(UBound(outArr) - 1) = 0
        outArr(UBound(outArr)) = 0
    ElseIf inBits = 16 Then
        ReDim outArr(0 To UBound(inArr) \ 2) As Integer
        CopyMemory outArr(0), inArr(0), UBound(inArr) + 1
    End If
    'volume reduction
    'For x = LBound(outArr) To UBound(outArr)
    '    outArr(x) = outArr(x) * 0.8
    'Next x
End Sub

Function BMEString(inval As Long) As String
    BMEString = Mid(xBMEString, (inval \ 36) + 1, 1) + Mid(xBMEString, (inval Mod 36) + 1, 1)
End Function

Sub ByteSwap16(inArr() As Integer)
    Dim X As Long
    Dim Y As Integer
    Debug.Print "Byteswap ", UBound(inArr) + 1
    For X = LBound(inArr) To ((UBound(inArr) \ 2) * 2) - 1 Step 2
        Y = inArr(X)
        inArr(X) = inArr(X + 1)
        inArr(X + 1) = Y
    Next X
End Sub

Sub CombineWaves(inFile1 As String, inFile2 As String, stereoOffset As Long)
    '1 is LEFT, 2 is RIGHT
    Dim inArr1() As Integer
    Dim inArr2() As Integer
    Dim f1 As Long
    Dim f2 As Long
    If Dir(inFile1) = "" Or Dir(inFile2) = "" Then
        Exit Sub
    End If
    f1 = FreeFile
    Open inFile1 For Binary As #f1
    f2 = FreeFile
    Open inFile2 For Binary As #f2
    Debug.Print LOF(f1)
    ReDim inArr1(0 To (LOF(f1) - 44) \ 2) As Integer
    ReDim inArr2(0 To UBound(inArr1)) As Integer
    Get #f1, 45, inArr1
    Get #f2, 45, inArr2
    Close #f2
    Dim X As Long
    For X = 1 To (UBound(inArr1) - 1) - stereoOffset Step 2
        inArr1(X) = inArr2(X + stereoOffset)
    Next X
    Put #f1, 45, inArr1
    Close #f1
End Sub
