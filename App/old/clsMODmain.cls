VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMODmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' saxxonpike's amiga mod loader
' 2oo7

Private Type ModHeader
    xName(0 To 19) As Byte
    xSampleInfo(1 To 31) As SampleInfo
    xSongLength As Byte
    xRepeat As Byte
    xOrders(0 To 127) As Byte
    xID(0 To 3) As Byte
End Type

Private Type ModSampleInfo
    xName(0 To 21) As Byte
    xSize As Integer
    xFinetune As Byte
    xVolume As Byte
    xRepeat As Integer
    xRepeatLength As Integer
End Type

Private Type ModNote
    xData(0 To 3) As Byte
End Type

Private Type ModPattern
    xNotes(0 To 255) As Byte
End Type
