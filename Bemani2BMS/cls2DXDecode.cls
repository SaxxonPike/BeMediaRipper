VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls2DXDecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'2dx decoder
' * based off "crack2dx" code by Tau
' * uses Flamed's ACM code
'
'saxxonpike 2oo7

Private Const KeysoundAdjust = 5
Private Const OnlyExtractDecrypted = False  'don't separate keysounds
Private Const ShortExtract = False          'only extract first 64 bytes for examination

Private Const BMEString = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const keys_09 = "971E24A09A00102B91EFD77ACD11AFAF8D265DBBE0C61B2B"
Private Const keys_10 = "2D865662D7FDCAA4B32460262481DBC257B1746FA7529921"
Private Const keys_11 = "EDF09C90441A5A03AB07C199232432C75F32A597AD980F8F"
Private Const keys_16 = "28222854633F0E426F454E506753617C0446003B132B456A"

Private Type WaveHead
    sRIFF As Long
    sRIFFsize As Long
    sWAVE As Long
    sfmt As Long
    sfmtsize As Long
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    sdata As Long
    sdatasize As Long
End Type

Private Type Header2DX9
    ArchiveName(0 To 15) As Byte
    DataStart As Long
    DataCount As Long
    Unk0(0 To 11) As Long
End Type

Private Type Entry2DX9
    ID2DX9(0 To 3) As Byte
    TagSize As Long
    TotalSize As Long
    flags(0 To 5) As Integer '2=pan, 3=volume (this is odd, lower is louder)
End Type

'Private WAVDecode As New FL_DecoderWAV

Private Keys(0 To 2, 0 To 7) As Byte
Private xVolume As Double

Private Type ByteBlock8
    Bytes(0 To 7) As Byte
End Type
Private Type CompressedHeader
    Ident2DX(0 To 3) As Byte
    FSize As Long
End Type

Public Sub SetVolume(newvol As Double)
    xVolume = newvol
End Sub

Private Function FillKeys(xVersion As Long) As Boolean
    Dim fString As String
    Dim X As Long
    Dim Y As Long
    FillKeys = True
    Select Case xVersion
        Case 9
            fString = keys_09
        Case 10
            fString = keys_10
        Case 11, 12
            fString = keys_11
        Case 16
            fString = keys_16
        Case Else
            fString = String$(48, 0)
            FillKeys = False
    End Select
    For X = 0 To 23
        Keys(X \ 8, X Mod 8) = Val("&H" + Mid$(fString, (X * 2) + 1, 2))
    Next X
End Function

Private Sub DecryptCommon(ByRef xBlock As ByteBlock8)
    Dim a As Byte
    Dim b As Byte
    Dim c As Byte
    Dim d As Byte
    Dim e As Byte
    Dim f As Byte
    Dim g As Byte
    Dim h As Byte
    Dim i As Byte
    a = (CLng(xBlock.Bytes(0)) * 63) Mod 256
    b = (CLng(xBlock.Bytes(3)) + a) Mod 256
    c = (CLng(xBlock.Bytes(1)) * 17) Mod 256
    d = (CLng(xBlock.Bytes(2)) + c) Mod 256
    e = (CLng(d) + b) Mod 256
    f = (CLng(xBlock.Bytes(3)) * e) Mod 256
    g = (CLng(f) + b + 51) Mod 256
    h = b Xor d
    i = g Xor e
    xBlock.Bytes(4) = xBlock.Bytes(4) Xor h
    xBlock.Bytes(5) = xBlock.Bytes(5) Xor d
    xBlock.Bytes(6) = xBlock.Bytes(6) Xor i
    xBlock.Bytes(7) = xBlock.Bytes(7) Xor g
End Sub

'void decrypt_normal(const byte *key1, const byte *key2, const byte *key3)
'{
'    byte tmp;
'    int i;
'
'    for (i = 0 ; i < 8 ; i++) block[i] ^= key1[i];
'    decrypt_common();
'
'    for (i = 0 ; i < 4 ; i++) {
'        tmp = block[i];
'        block[i] = block[i + 4];
'        block[i + 4] = tmp;
'    }
'
'    for (i = 0 ; i < 8 ; i++) block[i] ^= key2[i];
'    decrypt_common();
'
'    for (i = 0 ; i < 8 ; i++) block[i] ^= key3[i];
'}

Private Sub DecryptNormal(ByRef xBlock As ByteBlock8)
    Dim tmp As Byte
    Dim i As Long
    
    For i = 0 To 7
        xBlock.Bytes(i) = xBlock.Bytes(i) Xor Keys(0, i)
    Next i
    DecryptCommon xBlock
    
    For i = 0 To 3
        tmp = xBlock.Bytes(i)
        xBlock.Bytes(i) = xBlock.Bytes(i + 4)
        xBlock.Bytes(i + 4) = tmp
    Next i
    
    For i = 0 To 7
        xBlock.Bytes(i) = xBlock.Bytes(i) Xor Keys(1, i)
    Next i
    DecryptCommon xBlock
    
    For i = 0 To 7
        xBlock.Bytes(i) = xBlock.Bytes(i) Xor Keys(2, i)
    Next i
    
End Sub

Private Sub DecryptBlock(xBlock As ByteBlock8, gamemode As Long)
    Dim tmp As Byte
    Dim i As Long
    
    Select Case gamemode
    
        Case 9, 10, 11, 16
            DecryptNormal xBlock
            
        Case 12
            For i = 0 To 7
                xBlock.Bytes(i) = xBlock.Bytes(i) Xor Keys(0, i)
            Next i
            DecryptCommon xBlock
            For i = 0 To 3
                tmp = xBlock.Bytes(i)
                xBlock.Bytes(i) = xBlock.Bytes(i + 4)
                xBlock.Bytes(i + 4) = tmp
            Next i
            
    End Select
End Sub

Public Function Decrypt(ByRef InFile As clsFileStream) As Boolean
    Dim curBlock As ByteBlock8
    Dim prevBlock As ByteBlock8
    Dim xBlock As ByteBlock8
    Dim xInfo As CompressedHeader
    Dim X As Long
    Dim i As Long
    Dim f As Long
    Dim offs As Double
    Dim decpos As Long
    Dim offssize As Double
    Dim gamemode As Long
    Dim xIdentString As String
    Dim sdata(0 To 7) As Byte
    Dim fData() As Byte
    
    InFile.ReadFileBinary sdata()
    
    CopyMemory xInfo, sdata(0), 8
    xIdentString = Chr$(sdata(0)) + Chr$(sdata(1)) + Chr$(sdata(2)) + Chr$(sdata(3))
    ReDim fData(0 To InFile.FileSize) As Byte
    
    Select Case xIdentString
        Case "%eNc"
            gamemode = 9
        Case "%e10"
            gamemode = 10
        Case "%e11"
            gamemode = 11
        Case "%e12", "%hid"
            gamemode = 12
        Case "%iO0"
            gamemode = 16
        Case Else
            InFile.ReadFileObject VarPtr(X), 4, 16
            If X > 0 Then
                InFile.ReadFileObject VarPtr(i), 4, X
                If i = &H39584432 Then
                    gamemode = -1
                    InFile.ReadFileBinary fData(), , 0
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
    End Select
    
    If gamemode > -1 Then
        If FillKeys(gamemode) = False Then
            Exit Function
        End If
    
        'begin conversion :)
        
        offssize = InFile.FileSize
        offs = 8
        Do While offs < offssize
            InFile.ReadFileBinary xBlock.Bytes(), , offs
            curBlock = xBlock
            DecryptBlock xBlock, gamemode
            For i = 0 To 7
                xBlock.Bytes(i) = xBlock.Bytes(i) Xor prevBlock.Bytes(i)
                fData(decpos) = xBlock.Bytes(i)
                decpos = decpos + 1
            Next i
            prevBlock = curBlock
            offs = offs + 8
            If (ShortExtract) And (offs > 64) Then
                Exit Do
                ReDim Preserve fData(0 To 63) As Byte
            End If
        Loop
    End If
    
    If OnlyExtractDecrypted Then
        f = FreeFile
        Open OutFolder + "extracted.bin" For Binary As #f
        Put #f, 1, fData()
        Close #f
    Else
        ExtractSounds fData()
        Decrypt = True
    End If
    
End Function

Private Function BME(ByVal inval As Integer) As String
    Dim res As String
    res = Mid$(BMEString, (inval \ Len(BMEString)) + 1, 1)
    res = res + Mid$(BMEString, (inval Mod Len(BMEString)) + 1, 1)
    BME = res
End Function

Private Sub ExtractSounds(ByRef InFile() As Byte)
    Dim KeyList() As Long
    Dim KeyHeader As Entry2DX9
    Dim FileHeader As Header2DX9
    'Dim WaveHeader As WAVEFORMATEX
    Dim Decoded() As Integer
    Dim Processed() As Integer
    Dim Isolated() As Byte
    Dim VolAmp(0 To 1) As Double
    Dim Proc2(0 To 1) As Long
    Dim MasterAmp As Double
    Dim hStream As Long
    Dim hStreamSize As Long
    Dim X As Long
    Dim Y As Long
    Dim f As Long
    Dim fName As String
    Dim BassChannelInfo As BASS_CHANNELINFO
    Dim OutWaveHeader As WaveHead
    
    With OutWaveHeader
        .sRIFF = &H46464952     ' "RIFF"
        .sWAVE = &H45564157     ' "WAVE"
        .sfmt = &H20746D66      ' "fmt "
        .sdata = &H61746164     ' "data"
        .nChannels = 2
        .wBitsPerSample = 16
        .wFormatTag = 1
        .nBlockAlign = 4
        .sfmtsize = 16
    End With
    
    'load filelist
    CopyMemory FileHeader, InFile(0), 72
    ReDim KeyList(0 To FileHeader.DataCount - 1) As Long
    CopyMemory KeyList(0), InFile(72), ((UBound(KeyList) + 1) * 4)
    
    For X = LBound(KeyList) To UBound(KeyList)
        Y = KeyList(X) + 24
        CopyMemory KeyHeader, InFile(KeyList(X)), 24
        fName = OutFolder + BME(X + 1) + ".wav"
        If InFile(KeyList(X)) = 50 Then
            ReDim Isolated(0 To KeyHeader.TotalSize - 1) As Byte
            CopyMemory Isolated(0), InFile(Y), KeyHeader.TotalSize ' in order for this to work,
            Y = VarPtr(Isolated(0))                                ' BASS likes the data isolated (wtf)
            hStream = BASS_StreamCreateFile(BASSTRUE, ByVal Y, 0, KeyHeader.TotalSize, BASS_STREAM_DECODE)
            If hStream <> 0 Then
                hStreamSize = BASS_ChannelGetLength(hStream)
                ReDim Decoded(0 To (hStreamSize \ 2) - 1) As Integer
                BASS_ChannelGetData hStream, Decoded(0), hStreamSize
                BASS_ChannelGetInfo hStream, BassChannelInfo
                'convert mono to stereo
                If BassChannelInfo.chans = 1 Then
                    ReDim Processed(0 To hStreamSize - 1) As Integer
                    For Y = 0 To UBound(Decoded)
                        Processed(Y * 2) = Decoded(Y)
                        Processed((Y * 2) + 1) = Decoded(Y)
                    Next Y
                Else
                    Processed() = Decoded()
                End If
                'adjust volumes for panning & global
                Debug.Print fName, "v"; KeyHeader.flags(3), "p"; KeyHeader.flags(2)
                KeyHeader.flags(3) = KeyHeader.flags(3) - KeysoundAdjust
                If KeyHeader.flags(3) < 1 Then
                    KeyHeader.flags(3) = 1
                End If
                MasterAmp = (1 - ((KeyHeader.flags(3) - 1) / 40))
                MasterAmp = MasterAmp * xVolume
                VolAmp(0) = (KeyHeader.flags(2) - 1) / 63
                VolAmp(1) = 2 - VolAmp(0)
                If VolAmp(0) > 1 Then VolAmp(0) = 1
                If VolAmp(1) > 1 Then VolAmp(1) = 1
                VolAmp(0) = VolAmp(0) * MasterAmp
                VolAmp(1) = VolAmp(1) * MasterAmp
                If VolAmp(0) <> 1 Or VolAmp(1) <> 1 Then
                    For Y = 0 To UBound(Processed) Step 2
                        Proc2(0) = CLng(Processed(Y)) * VolAmp(0)
                        Proc2(1) = CLng(Processed(Y + 1)) * VolAmp(1)
                        If Proc2(0) < -32768 Then Proc2(0) = -32768
                        If Proc2(0) > 32767 Then Proc2(0) = 32767
                        If Proc2(1) < -32768 Then Proc2(1) = -32768
                        If Proc2(1) > 32767 Then Proc2(1) = 32767
                        Processed(Y) = Proc2(0)
                        Processed(Y + 1) = Proc2(1)
                    Next Y
                End If
                'output
                f = FreeFile
                If Dir(fName) <> "" Then
                    Kill fName
                End If
                On Error GoTo 73
                Open fName For Binary As #f
                Put #f, 45, Processed
                With OutWaveHeader
                    .nSamplesPerSec = BassChannelInfo.freq
                    .nAvgBytesPerSec = (.nSamplesPerSec * .nChannels * (.wBitsPerSample \ 8))
                    .sRIFFsize = LOF(f) - 8
                    .sdatasize = LOF(f) - 44
                End With
                Put #f, 1, OutWaveHeader
73              Close #f
            End If
            On Error GoTo 0
            BASS_StreamFree hStream
            hStream = 0
        End If
    Next X
    
End Sub

Private Sub Class_Terminate()
    BASS_Free
End Sub
