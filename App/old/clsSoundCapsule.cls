VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoundCapsule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsSoundCapsule
' for BEMANI games
'
' SaxxonPike 2oo7

Private Const BMEString = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Private Type HeaderPopn15
    ArchiveName As String * 16
    DataStart As Long
    DataCount As Long
    Unk0(0 To 11) As Long
End Type

Private Type Entry2DX9
    ID2DX9 As String * 4
    TagSize As Long
    TotalSize As Long
    Flags(0 To 5) As Integer
        'flags:
        '2-volume?
End Type

Private Type CapsuleInfo
    KeyCount As Long
    KeyOffsets() As Long
End Type

Private Capsule As CapsuleInfo
Private f As Long

Public Sub OpenPopn15Capsule(inFile As String)
    Dim PNMInfo As HeaderPopn15
    f = FreeFile
    Open inFile For Binary As #f
    Get #f, 1, PNMInfo
    Capsule.KeyCount = PNMInfo.DataCount
    ReDim Capsule.KeyOffsets(0 To PNMInfo.DataCount - 1) As Long
    Get #f, 73, Capsule.KeyOffsets
End Sub

Public Sub CloseCapsule()
    Close #f
End Sub

Public Sub ExtractKeysound(inNumber As Long, inTarget As String)
    Dim TEntry As Entry2DX9
    Dim WaveData() As Byte
    Dim f2 As Long
    If inNumber > UBound(Capsule.KeyOffsets) Or inNumber < 0 Then
        Exit Sub
    End If
    Get #f, Capsule.KeyOffsets(inNumber) + 1, TEntry
    If TEntry.ID2DX9 <> "2DX9" Then
        Exit Sub
    End If
    With TEntry
        ReDim WaveData(0 To TEntry.TotalSize - 1) As Byte
        Get #f, Capsule.KeyOffsets(inNumber) + 1 + TEntry.TagSize, WaveData
        f2 = FreeFile
        Open inTarget + BME(inNumber + 1) + ".wav" For Binary As #f2
        Put #f2, 1, WaveData
        Close #f2
    End With
End Sub

Public Sub ExtractAll(inTarget As String)
    Dim x As Long
    For x = 0 To Capsule.KeyCount - 1
        ExtractKeysound x, inTarget
    Next x
    MsgBox "Conversion Complete."
End Sub

Public Function KeysoundCount() As Long
    KeysoundCount = Capsule.KeyCount
End Function

Private Function BME(ByVal inval As Integer) As String
    Dim res As String
    res = Mid$(BMEString, (inval \ Len(BMEString)) + 1, 1)
    res = res + Mid$(BMEString, (inval Mod Len(BMEString)) + 1, 1)
    BME = res
End Function

