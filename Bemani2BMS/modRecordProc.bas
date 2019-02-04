Attribute VB_Name = "modRecordProc"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private f As Long
Private o As Long

Function RECORDPROC1(ByVal handle As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Long
    'CALLBACK FUNCTION !!!

    ' Recording callback function.
    ' handle : The recording handle
    ' buffer : Buffer containing the recorded samples
    ' length : Number of bytes
    ' user   : The 'user' parameter value given when calling BASS_RecordStart
    ' RETURN : BASSTRUE = continue recording, BASSFALSE = stop
    
    Dim b() As Byte
    
    If f <> 0 Then
        On Local Error Resume Next
        
        ReDim b(0 To length - 1) As Byte
        CopyMemory b(0), ByVal buffer, length
        Put #f, o, b
        o = o + length
        RECORDPROC1 = BASSTRUE
    End If

End Function

Sub StartRecordFile(sFile As String)
    On Local Error GoTo 2
    If f <> 0 Then
        Close #f
    End If
    f = FreeFile
    Open sFile For Output As #f
    Close #f
    Open sFile For Binary As #f
    o = 45
    Exit Sub
2   Close #f
    f = 0
End Sub

Sub StopRecordFile()
    Dim s As String
    Dim l As Long
    Dim b(0 To 43) As Byte
    If f = 0 Then
        Exit Sub
    End If
    b(16) = 16
    b(20) = 1
    b(22) = 2
    b(32) = 4
    b(34) = 16
    Put #f, 1, b
    s = "RIFF"
    Put #f, 1, s
    s = "WAVEfmt "
    Put #f, 9, s
    s = "data"
    Put #f, 37, s
    l = LOF(f) - 8
    Put #f, 5, l
    l = LOF(f) - 44
    Put #f, 41, l
    l = 44100
    Put #f, 25, l
    l = 44100 * 4
    Put #f, 29, l
    Close #f
    f = 0
End Sub
