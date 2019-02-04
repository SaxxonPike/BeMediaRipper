VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWAVFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' WAV file class
' saxxonpike 2006

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Type WAVHeader
    RIFFstring As Long
    RIFFsize As Long
    WAVE As Long
    fmtstring As Long
    fmtsize As Long
    AudioFormat As Integer
    Channels As Integer
    SampleRate As Long
    ByteRate As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    Datastring As Long
    datasize As Long
End Type
Private WaveHeader As WAVHeader
Private WaveData() As Integer

Public Property Get Frequency() As Long
    Frequency = WaveHeader.SampleRate
End Property
Public Property Let Frequency(v As Long)
    WaveHeader.SampleRate = v
    WaveHeader.ByteRate = v * WaveHeader.BlockAlign
    'Debug.Print WaveHeader.ByteRate
End Property
Public Property Get Channels() As Long
    Channels = WaveHeader.Channels
End Property
Public Property Let Channels(v As Long)
    WaveHeader.Channels = v
    WaveHeader.BlockAlign = WaveHeader.BitsPerSample * WaveHeader.Channels / 8
End Property





Public Sub ClearData()
    ReDim WaveData(0) As Integer
    datasize = 0
End Sub

Private Sub Class_Initialize()
    ClearData
    With WaveHeader
        .AudioFormat = 1
        .BitsPerSample = 16
        .BlockAlign = 4
        .ByteRate = 176400
        .Channels = 2
        .datasize = 0
        .Datastring = &H61746164
        .fmtsize = &H10
        .fmtstring = &H20746D66
        .RIFFsize = ((UBound(WaveData) * 2) + Len(WaveHeader)) - 8
        .RIFFstring = &H46464952
        .SampleRate = 44100
        .WAVE = &H45564157
    End With
End Sub

Public Sub AddRawData(rawdata() As Integer)
    Dim xdatasize As Long
    Dim mdatasize As Long
    xdatasize = (UBound(rawdata) - LBound(rawdata)) + 1
    mdatasize = UBound(WaveData)
    ReDim Preserve WaveData(0 To UBound(WaveData) + xdatasize)
    CopyMemory WaveData(mdatasize), rawdata(LBound(rawdata)), xdatasize * 2
    WaveHeader.RIFFsize = WaveHeader.RIFFsize + (2 * xdatasize)
    WaveHeader.datasize = WaveHeader.datasize + (2 * xdatasize)
End Sub

Public Sub WriteHeader(sFileName As String, WaveDataSize As Long)
    Dim f As Long
    Dim NewHeader As WAVHeader
    NewHeader = WaveHeader
    f = FreeFile
    Open sFileName For Binary As #f
    With NewHeader
        .datasize = LOF(f) - Len(NewHeader)
        .RIFFsize = LOF(f) - 8
    End With
    Put #f, 1, NewHeader
    Close #f
End Sub

Public Sub SaveWave(sFileName As String)
    Dim f As Long
    If Dir(sFileName) <> "" Then
        Kill sFileName
    End If
    f = FreeFile
    Open sFileName For Binary As #f
    Put #f, 1, WaveHeader
    Put #f, , WaveData
    Close #f
End Sub

Public Sub SaveWave2(sFileName As String, ByRef wave16() As Integer)
    Dim f As Long
    Dim z As Long
    z = ((UBound(wave16) - LBound(wave16)) * 2) + 2
    If Dir(sFileName) <> "" Then
        Kill sFileName
    End If
    WaveHeader.RIFFsize = WaveHeader.RIFFsize + z
    WaveHeader.datasize = z
    f = FreeFile
    Open sFileName For Binary As #f
    Put #f, 1, WaveHeader
    Put #f, , wave16
    Close #f
    WaveHeader.RIFFsize = WaveHeader.RIFFsize - z
End Sub

