Attribute VB_Name = "modBeMedia"
Private FreqTable(0 To 32767) As Double
Private LogFileNumber As Long
Public OutFolder As String

Public Function CalculateFreqTable()
    Dim FreqCount As Long
    Dim FreqCurr As Long
    Dim FreqConst(0 To 20, 0 To 1) As Long
    Dim CurSlope As Double
    Dim CurAmt As Long
    Dim FreqOffs As Long
    Dim FreqMult As Double
    Dim x As Long
    Dim Y As Long
    
    FreqCount = 15
    FreqConst(Y, 0) = 44100:            FreqConst(Y, 1) = 15940:        Y = Y + 1
    FreqConst(Y, 0) = 40000:            FreqConst(Y, 1) = 16491:        Y = Y + 1
    FreqConst(Y, 0) = 36000:            FreqConst(Y, 1) = 16642:        Y = Y + 1
    FreqConst(Y, 0) = 37000:            FreqConst(Y, 1) = 16703:        Y = Y + 1
    FreqConst(Y, 0) = 37000:            FreqConst(Y, 1) = 16708:        Y = Y + 1
    FreqConst(Y, 0) = 37818:            FreqConst(Y, 1) = 16762:        Y = Y + 1
    FreqConst(Y, 0) = 36000:            FreqConst(Y, 1) = 16899:        Y = Y + 1
    FreqConst(Y, 0) = 35002:            FreqConst(Y, 1) = 16964:        Y = Y + 1
    FreqConst(Y, 0) = 32000:            FreqConst(Y, 1) = 17533:        Y = Y + 1
    FreqConst(Y, 0) = 30000:            FreqConst(Y, 1) = 17774:        Y = Y + 1
    FreqConst(Y, 0) = 28000:            FreqConst(Y, 1) = 18005:        Y = Y + 1
    FreqConst(Y, 0) = 24000:            FreqConst(Y, 1) = 18432:        Y = Y + 1
    FreqConst(Y, 0) = 22050:            FreqConst(Y, 1) = 19007:        Y = Y + 1
    FreqConst(Y, 0) = 22050:            FreqConst(Y, 1) = 19012:        Y = Y + 1
    FreqConst(Y, 0) = 11025:            FreqConst(Y, 1) = 19714:        Y = Y + 1
    
    
    FreqCurr = 0
    CurSlope = (FreqConst(1, 1) - FreqConst(0, 1))
    FreqMult = (FreqConst(1, 0) - FreqConst(0, 0))
    CurAmt = FreqConst(0, 0)
    FreqOffs = FreqConst(0, 1)
    
    For x = 0 To 32767
        If FreqCurr < (FreqCount - 2) Then
            If x = FreqConst(FreqCurr + 1, 1) Then
                FreqCurr = FreqCurr + 1
                CurSlope = (FreqConst(FreqCurr + 1, 1) - FreqConst(FreqCurr, 1))
                FreqMult = (FreqConst(FreqCurr + 1, 0) - FreqConst(FreqCurr, 0))
                FreqOffs = FreqConst(FreqCurr, 1)
            End If
        End If
        FreqTable(x) = (((x - FreqConst(FreqCurr, 1)) / CurSlope) * FreqMult) + FreqConst(FreqCurr, 0)
    Next x
End Function

Public Function ConvertFrequency(InFreq As Long)
    If InFreq > 0 And InFreq <= 32767 Then
        ConvertFrequency = Round(FreqTable(InFreq), 0)
        If Int(FreqTable(InFreq)) <> FreqTable(InFreq) Then
            Debug.Print InFreq, FreqTable(InFreq)
            InFreq = InFreq
        End If
    Else
        ConvertFrequency = 44100
    End If
    'End If
End Function

Public Function FindBPM(inArr() As Byte) As Long
    If inArr(0) <> 8 Then
        Exit Function
    End If
    Dim Coffs As Long
    Coffs = 8
    Do
        If inArr(Coffs + 0) = 0 And inArr(Coffs + 1) = 0 Then
            If (inArr(Coffs + 4) And &HF) = 4 Then
                FindBPM = inArr(Coffs + 5)
                Exit Function
            End If
        End If
        Coffs = Coffs + 8
    Loop While Coffs < UBound(inArr) - 4
    x = x
End Function

Public Function FindNoteCount(inArr() As Byte) As Long
    If inArr(0) <> 8 Then
        Exit Function
    End If
    Dim Coffs As Long
    Dim fnc As Long
    Coffs = 8
    fnc = 0
    Do
        If inArr(Coffs + 6) = 0 And inArr(Coffs + 7) = 0 Then
            If (inArr(Coffs + 4) And &HF) <= 1 Then
                If inArr(Coffs) <> 0 Or inArr(Coffs + 1) <> 0 Or inArr(Coffs + 2) <> 0 Or inArr(Coffs + 3) <> 0 Then
                    fnc = fnc + 1
                End If
            End If
        End If
        Coffs = Coffs + 8
    Loop While Coffs < UBound(inArr) - 4
    x = x
    FindNoteCount = fnc
End Function

Public Function FindHiKey(inArr() As Byte) As Long
    If inArr(0) <> 8 Then
        Exit Function
    End If
    Dim Coffs As Long
    Dim fnc As Long
    Dim keynum As Long
    Coffs = 8
    fnc = 0
    Do
        If inArr(Coffs + 6) <> 0 Or inArr(Coffs + 7) <> 0 Then
            If (inArr(Coffs + 4) And &HF) = 2 Or (inArr(Coffs + 4) And &HF) = 3 Or (inArr(Coffs + 4) And &HF) = 7 Then
                keynum = CLng(inArr(Coffs + 7)) * 256
                keynum = keynum + CLng(inArr(Coffs + 6))
                If keynum > fnc Then
                    fnc = keynum
                End If
            End If
        End If
        Coffs = Coffs + 8
    Loop While Coffs < UBound(inArr) - 4
    x = x
    FindHiKey = fnc
End Function

Public Function FindBPM2(inArr() As Byte) As Long
    If inArr(0) <> 8 Then
        Exit Function
    End If
    Dim Coffs As Long
    Coffs = 8
    Do
        If inArr(Coffs + 0) = 0 And inArr(Coffs + 1) = 0 Then
            If (inArr(Coffs + 2) And &HF) = 4 Then
                FindBPM2 = inArr(Coffs + 3)
                Exit Function
            End If
        End If
        Coffs = Coffs + 4
    Loop While Coffs < UBound(inArr) - 4
    x = x
End Function

Public Function FindNoteCount2(inArr() As Byte) As Long
    If inArr(0) <> 8 Then
        Exit Function
    End If
    Dim Coffs As Long
    Dim fnc As Long
    Coffs = 8
    fnc = 0
    Do
        If inArr(Coffs + 3) = 0 Then
            If (inArr(Coffs + 2) And &HF) <= 1 Then
                If inArr(Coffs) <> 0 Or inArr(Coffs + 1) <> 0 Then
                    fnc = fnc + 1
                End If
            End If
        End If
        Coffs = Coffs + 4
    Loop While Coffs < UBound(inArr) - 4
    x = x
    FindNoteCount2 = fnc
End Function

Public Function FindHiKey2(inArr() As Byte) As Long
    If inArr(0) <> 8 Then
        Exit Function
    End If
    Dim Coffs As Long
    Dim fnc As Long
    Dim keynum As Long
    Coffs = 8
    fnc = 0
    Do
        If inArr(Coffs + 3) <> 0 Then
            If (inArr(Coffs + 2) And &HF) = 2 Or (inArr(Coffs + 2) And &HF) = 3 Or (inArr(Coffs + 2) And &HF) = 7 Then
                keynum = inArr(Coffs + 3)
                If keynum > fnc Then
                    fnc = keynum
                End If
            End If
        End If
        Coffs = Coffs + 4
    Loop While Coffs < UBound(inArr) - 4
    x = x
    FindHiKey2 = fnc
End Function

Public Sub MkDir2(ByVal fName As String, Optional bIsFileName As Boolean = False)
    Dim x As Long
    Dim c As Long
    If Not bIsFileName Then
        fName = fName + "\"
    End If
    If InStr(fName, Chr(0)) > 0 Then
        fName = Left(fName, InStr(fName, Chr(0)) - 1)
    End If
    For x = 1 To Len(fName)
        If Mid(fName, x, 1) = "\" Then
            If c > 0 Or Mid(fName, 2, 1) <> ":" Then
                On Error Resume Next
                MkDir Left(fName, x - 1)
                On Error GoTo 0
            End If
            c = c + 1
        End If
    Next x
End Sub

Public Function AppPath() As String
    AppPath = App.Path
    If Right(AppPath, 1) = "\" Then
        AppPath = Left(AppPath, Len(AppPath) - 1)
    End If
End Function

Public Function MakeWord(LoByte As Byte, HiByte As Byte) As Integer
  If HiByte And &H80 Then
    MakeWord = ((HiByte * &H100&) Or LoByte) Or &HFFFF0000
  Else
    MakeWord = (HiByte * &H100) Or LoByte
  End If
End Function

Public Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeDWord = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function

Public Function DWBytes(byte0 As Byte, byte1 As Byte, byte2 As Byte, hibyte3 As Byte)
    DWBytes = MakeDWord(MakeWord(byte0, byte1), MakeWord(byte2, hibyte3))
End Function

Public Sub PrintLog(lText As String)
    If LogFileNumber = 0 Then
        Exit Sub
    End If
    Print #LogFileNumber, lText
End Sub

Public Sub OpenLog(fName As String)
    LogFileNumber = FreeFile
    Open fName For Output As #LogFileNumber
End Sub

Public Sub CloseLog()
    Close #LogFileNumber
    LogFileNumber = 0
End Sub
