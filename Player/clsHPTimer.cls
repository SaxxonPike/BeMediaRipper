VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHPTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
'Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
'Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private PerformanceFreq As Currency
Private LastTick As Currency
Private CurrentTick As Currency

Private Sub Class_Initialize()
    QueryPerformanceFrequency PerformanceFreq
    PerformanceFreq = PerformanceFreq / 1000
End Sub

Public Sub Init()
    QueryPerformanceCounter LastTick
End Sub

Public Property Get TimeElapsed() As Currency
    QueryPerformanceCounter CurrentTick
    TimeElapsed = ((CurrentTick - LastTick) / PerformanceFreq)
End Property

Public Sub Tick()
    QueryPerformanceCounter LastTick
End Sub

Public Sub TickInc(IncVal As Currency)
    LastTick = LastTick + (IncVal * PerformanceFreq)
End Sub
