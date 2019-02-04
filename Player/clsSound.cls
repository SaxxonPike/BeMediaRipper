VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' multimedia functions
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Private Const SND_ASYNC = &H1
Private Const SND_LOOP = &H8
Private Const SND_NODEFAULT = &H2
Private Const SND_SYNC = &H0
Private Const SND_NOSTOP = &H10
Private Const SND_MEMORY = &H4
Private RetValue As Long

Private Const xVersion = 1

Private Type MBufferType
    RIFFHead As Long
    RIFFLength As Long
    WAVETag As Long
    fmtTag As Long
    fmtSize As Long
    WaveType As Integer
    WaveChannels As Integer
    WaveFreq As Long
    WaveAvgBPS As Long
    BytesPerSample As Integer
    SampleBits As Integer
    dataTag As Long
    dataLength As Long
    WaveData(0 To 7055) As Byte
End Type

Private Buffer As MBufferType
Private xBufferSize As Long
Private xBufferPtr As Long
Private BufferTimer As Currency

Private Sub Class_Initialize()
    With Buffer
        .BytesPerSample = 4
        .dataLength = 0
        .dataTag = &H61746164
        .fmtSize = 16
        .fmtTag = &H20746D66
        .RIFFHead = &H46464952
        .RIFFLength = 36
        .SampleBits = 16
        .WaveAvgBPS = 176400
        .WaveChannels = 2
        .WaveFreq = 44100
        .WAVETag = &H45564157
        .WaveType = 1
    End With
    xBufferSize = UBound(Buffer.WaveData) + 1
    RecreateBuffer
End Sub

Public Property Get Version() As Long
    Version = xVersion
End Property

Public Sub Init()
    'init code here
End Sub

Public Property Get BufferSize() As Long
    BufferSize = xBufferSize
End Property

'Public Property Let BufferSize(newsize As Long)
'    xBufferSize = newsize
'    RecreateBuffer
'End Property

Private Sub RecreateBuffer()
    'ReDim Buffer.WaveData(0 To (xBufferSize * Buffer.BytesPerSample) - 1) As Byte
    Buffer.dataLength = UBound(Buffer.WaveData) + 1
    Buffer.RIFFLength = 36 + Buffer.dataLength
    xBufferPtr = VarPtr(Buffer)
    BufferTimer = ((UBound(Buffer.WaveData) + 1) / Buffer.WaveAvgBPS) * 1000
End Sub

Public Property Get BufferPtr() As Long
    BufferPtr = xBufferPtr
End Property

Public Property Get BufferDelay() As Currency
    BufferDelay = BufferTimer
End Property

Public Sub Play()
    RetValue = sndPlaySound(Buffer, SND_MEMORY Or SND_NODEFAULT Or SND_ASYNC) 'memory + async
End Sub

Public Sub Randomize()
    Dim x As Long
    For x = 0 To UBound(Buffer.WaveData)
        Buffer.WaveData(x) = Int(Rnd * 256)
    Next x
    'Open "h:\wavdump.wav" For Binary As #1
    'Put #1, 1, Buffer
    'Close #1
End Sub
