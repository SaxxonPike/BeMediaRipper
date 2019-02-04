VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
'
' Step Conversion class
'  Copyright © SaxxonPike 2oo8
'
' ************************************************************************
' "DDR Utility" is free software: you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation, either version 3 of the License, or
'  (at your option) any later version.
'
' "DDR Utility" is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
'  along with this program.  If not, see <http://www.gnu.org/licenses/>.
' ************************************************************************

Private Const bLog = False

Private Const MODE_SINGLE = &H100&
Private Const MODE_COUPLE = &H200&
Private Const MODE_DOUBLE = &H300&
Private Const MODE_SOLO = &H400&
Private Const MODE_SOLO3STEP = &H500&
Private Const MODE_MASK = &H7F00&
Private Const DIFFICULTY_BEGINNER = &H1&
Private Const DIFFICULTY_LIGHT = &H2&
Private Const DIFFICULTY_STANDARD = &H4&
Private Const DIFFICULTY_HEAVY = &H8&
Private Const DIFFICULTY_CHALLENGE = &H10&
Private Const DIFFICULTY_MASK = &H7F&

Private xUNote(0 To 20000) As UniversalNote
Private xUBPM() As UniversalBPM
Private xNoteCount As Long
Private xBPMCount As Long

Private Type SSQHead
    xSize As Long
    xParm1 As Integer
    xParm2 As Integer
    xParm3 As Integer
    xParm4 As Integer
End Type
Private Type StepForm1Note
    xOffset As Long
    xStep As Long
End Type
Private Type StepForm1BPM
    xOffset As Long
    xDuration As Long
End Type
Private Type StepForm2BPM
    xOffset As Long
    xStepSize As Long
    xUnknown As Long
    xStepStart As Long
    xStepEnd As Long
End Type
Private Type UniversalNote
    xOffset192 As Long
    xPlayer As Long
    xChart As Long
    xStepLeft As Long
    xStepDown As Long
    xStepUp As Long
    xStepRight As Long
    xStepUL As Long
    xStepUR As Long
    xFreezeLeft As Long
    xFreezeDown As Long
    xFreezeUp As Long
    xFreezeRight As Long
    xFreezeUL As Long
    xFreezeUR As Long
    xEndFreezeLeft As Long
    xEndFreezeDown As Long
    xEndFreezeUp As Long
    xEndFreezeRight As Long
    xEndFreezeUL As Long
    xEndFreezeUR As Long
End Type
Private Type UniversalBPM 'stops also
    xOffset192 As Long
    xFlags As Long '+1=is a stop, not bpm
    xBPM As Double
End Type
Private Type FreezeNotice
    xOffset192 As Long
    xStepNumber As Long
End Type
Private bFreezeNotice As Boolean
Private xFreezeNotices() As FreezeNotice


Public Function ConvertStep1(inFile As String, outFile As String, Optional bSolo As Boolean = False) As Boolean
    Dim f1 As Long 'file handle input
    Dim l As Long  'data length
    Dim m As Long  'mode (as defined by the class)
    Dim t As Long  'type (as defined in the file)
    Dim o As Long  'file offset
    Dim x As Long  'loop counter
    Dim y As Long  'loop counter
    Dim g As Long  '#gap
    Dim s As Long  'step count
    Dim ep As String
    Dim BPMInfo() As StepForm1BPM
    Dim NoteInfo() As StepForm1Note
    
    WipeNotes
    Debug.Print inFile
    
    f1 = FreeFile: Open inFile For Binary As #f1
    o = 1
    m = 0
    s = -1
    xNoteCount = 0
    
    Do While o < LOF(f1)
        Get #f1, o, l
        If l = 0 Then Exit Do
        Get #f1, o + 4, t
        Select Case m
            Case 0 'BPMs/stops
                If l >= 20 Then
                    g = (l \ 8) - 1
                    ReDim BPMInfo(0 To g) As StepForm1BPM
                    g = UBound(BPMInfo) - 1
                    ReDim xUBPM(0 To g) As UniversalBPM
                    Get #f1, o + 4, BPMInfo
                    g = (BPMInfo(0).xDuration / 75) * 1000
                    For x = 1 To UBound(BPMInfo)
                        With xUBPM(x - 1)
                            .xOffset192 = (((BPMInfo(x - 1).xOffset) / &H1000&) * 192)
                            If x = 1 Then
                                .xOffset192 = 0
                            End If
                            If BPMInfo(x).xOffset <> BPMInfo(x - 1).xOffset Then
                                'bpm change
                                .xBPM = (((BPMInfo(x).xOffset - BPMInfo(x - 1).xOffset) / &H400&)) / _
                                    (((BPMInfo(x).xDuration - BPMInfo(x - 1).xDuration) / 75) / 60)
                            Else
                                'stop
                                .xBPM = ((BPMInfo(x).xDuration - BPMInfo(x - 1).xDuration) / 75)
                                If .xBPM > 0 Then
                                    .xFlags = 1
                                Else
                                    .xBPM = 0
                                End If
                            End If
                        End With
                    Next x
                End If
                m = 1
            Case 1 'steps
                If l > 16 Or l < 0 Then
                    If l > LOF(f1) Or l < 0 Then
                        Exit Do
                    End If
                    ReDim NoteInfo(0 To (l \ 8) - 2) As StepForm1Note
                    Get #f1, o + 4, t 'chart#
                    Get #f1, o + 8, NoteInfo
                    If Not bSolo Then
                        Select Case t 'all other games
                            Case 0: t = MODE_SINGLE + DIFFICULTY_LIGHT
                            Case 1: t = MODE_COUPLE + DIFFICULTY_LIGHT
                            Case 2: t = MODE_DOUBLE + DIFFICULTY_LIGHT
                            Case 256: t = MODE_SINGLE + DIFFICULTY_STANDARD
                            Case 257: t = MODE_COUPLE + DIFFICULTY_STANDARD
                            Case 258: t = MODE_DOUBLE + DIFFICULTY_STANDARD
                            Case 512: t = MODE_SINGLE + DIFFICULTY_HEAVY
                            Case 513: t = MODE_COUPLE + DIFFICULTY_HEAVY
                            Case 514: t = MODE_DOUBLE + DIFFICULTY_HEAVY
                        End Select
                    Else
                        Select Case t 'DDR SOLO only
                            Case 0: t = MODE_SOLO + DIFFICULTY_LIGHT
                            Case 1: t = MODE_SINGLE + DIFFICULTY_LIGHT
                            Case 2: t = MODE_SOLO3STEP + DIFFICULTY_LIGHT
                            Case 256: t = MODE_SOLO + DIFFICULTY_STANDARD
                            Case 257: t = MODE_SINGLE + DIFFICULTY_STANDARD
                            Case 258: t = MODE_SOLO3STEP + DIFFICULTY_STANDARD
                            Case 512: t = MODE_SOLO + DIFFICULTY_HEAVY
                            Case 513: t = MODE_SINGLE + DIFFICULTY_HEAVY
                            Case 514: t = MODE_SOLO3STEP + DIFFICULTY_HEAVY
                        End Select
                    End If
                    For x = 0 To UBound(NoteInfo)
                        If NoteInfo(x).xStep <> -1 Then
                            s = s + 2
                            xNoteCount = s
                            With xUNote(s - 1)
                                .xPlayer = 1
                                .xOffset192 = (((NoteInfo(x).xOffset) / &H1000&) * 192)
                                .xStepLeft = CLng(CBool(NoteInfo(x).xStep And &H1&))
                                .xStepDown = CLng(CBool(NoteInfo(x).xStep And &H10&))
                                .xStepUp = CLng(CBool(NoteInfo(x).xStep And &H100&))
                                .xStepRight = CLng(CBool(NoteInfo(x).xStep And &H1000&))
                                .xChart = t
                            End With
                            If Not bSolo Then
                                With xUNote(s)
                                    .xPlayer = 2
                                    .xOffset192 = (((NoteInfo(x).xOffset) / &H1000&) * 192)
                                    .xStepLeft = CLng(CBool(NoteInfo(x).xStep And &H10000))
                                    .xStepDown = CLng(CBool(NoteInfo(x).xStep And &H100000))
                                    .xStepUp = CLng(CBool(NoteInfo(x).xStep And &H1000000))
                                    .xStepRight = CLng(CBool(NoteInfo(x).xStep And &H10000000))
                                    .xChart = t
                                End With
                            Else
                                With xUNote(s)
                                    .xPlayer = 1
                                    .xOffset192 = (((NoteInfo(x).xOffset) / &H1000&) * 192)
                                    .xStepUL = CLng(CBool(NoteInfo(x).xStep And &H10000))
                                    .xStepUR = CLng(CBool(NoteInfo(x).xStep And &H1000000))
                                    .xChart = t
                                End With
                            End If
                        End If
                    Next x
                End If
        End Select
        o = o + l
    Loop
    Close #f1
    
    SortNotes
    WriteSM outFile, , , g
    DoEvents
    ConvertStep1 = True
End Function

Private Sub SortNotes()
    Dim x As Long
    Dim y As Long
    Dim tUNote As UniversalNote
    Dim e As Boolean
    Do While Not e
        e = True
        For x = 1 To xNoteCount
            If xUNote(x).xOffset192 < xUNote(x - 1).xOffset192 Then
                tUNote = xUNote(x)
                xUNote(x) = xUNote(x - 1)
                xUNote(x - 1) = tUNote
                e = False
            End If
        Next x
    Loop
End Sub

Private Sub WipeNotes()
    Dim EmptyNote As UniversalNote
    Dim x As Long
    For x = 0 To 20000
        xUNote(x) = EmptyNote
    Next x
    bFreezeNotice = False
    ReDim xFreezeNotices(0) As FreezeNotice
End Sub

Public Function ConvertSSQ(ByVal inFile As String, ByVal outFile As String) As Boolean
    Dim SSQh As SSQHead
    Dim e As Boolean
    Dim m As Long
    Dim f As Long
    Dim o As Long
    Dim l As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim TimeMode As Long
    Dim fr As Boolean
    Dim LastStep As Long
    Dim oT As String
    Dim LogPos As Long
    Dim SSQd1() As Long 'for offsets
    Dim SSQd2() As Long 'for BPM/stops
    Dim SSQd3() As Byte 'for steps
    Dim SSQd4() As Integer 'for unknown type 2
    Dim SSQd5() As Long 'for unknown type 2 offsets
    Dim d1 As Long
    Dim d2 As Long
    Dim fnc As Long 'freeze notice counter
    On Local Error Resume Next
    WipeNotes
    s = -1
    f = FreeFile
    Open inFile For Binary As #f
    m = 0
    o = 1
    LogPos = 1
    Do
        Get #f, o, SSQh
        If SSQh.xSize = 0 Then
            Exit Do
        End If
        Select Case SSQh.xParm1
            Case 1 'bpm changes
                TimeMode = SSQh.xParm2
                If SSQh.xParm3 > 1 Then
                    ReDim SSQd1(1 To SSQh.xParm3) As Long
                    ReDim SSQd2(1 To SSQh.xParm3) As Long
                    Get #f, , SSQd1
                    Get #f, , SSQd2
                    'fix BPMs
                    For x = 2 To SSQh.xParm3
                        If (SSQd1(x) < SSQd1(x - 1)) And (SSQd2(x) >= SSQd2(x - 1)) Then
                            SSQd1(x) = SSQd1(x - 1)
                            SSQd2(x) = SSQd2(x - 1)
                        End If
                    Next x
                    ReDim xUBPM(0 To SSQh.xParm3 - 2) As UniversalBPM
                    For x = 2 To SSQh.xParm3
                        With xUBPM(x - 2)
                            .xOffset192 = ((SSQd1(x - 1) / &H1000&) * 192)
                            If x = 2 Then
                                .xOffset192 = 0
                            End If
                            If SSQd1(x) <> SSQd1(x - 1) Then
                                'bpm change
                                If SSQd2(x) < &H100 And (SSQh.xParm3 = 2) Then
                                    oT = "*** USE DISPLAY BPM ***"
                                    .xBPM = 123.456789
                                Else
                                    .xBPM = (((SSQd1(x) - SSQd1(x - 1)) / &H400&)) / _
                                        (((SSQd2(x) - SSQd2(x - 1)) / TimeMode) / 60)
                                End If
                            Else
                                'stop
                                .xBPM = ((SSQd2(x) - SSQd2(x - 1)) / TimeMode)
                                If .xBPM > 0 Then
                                    .xFlags = 1
                                Else
                                    .xBPM = 0
                                End If
                            End If
                        End With
                    Next x
                End If
            Case 2 'unknown
                If SSQh.xParm3 > 0 Then
                    ReDim SSQd4(1 To SSQh.xParm3) As Integer
                    ReDim SSQd5(1 To SSQh.xParm3) As Long
                    Get #f, , SSQd5
                    Get #f, , SSQd4
                End If
            Case 3 'notes
                If SSQh.xParm3 > 0 Then
                    ReDim SSQd1(1 To SSQh.xParm3) As Long
                    ReDim SSQd3(1 To SSQh.xParm3) As Byte
                    Get #f, , SSQd1
                    Get #f, , SSQd3
                    '*** some logging ***
                    If bLog Then
                        f2 = FreeFile
                        Open "h:\log.dat" For Binary As #f2
                        LogPos = LOF(f2) + 1
                        Put #f2, LogPos, inFile
                        LogPos = LogPos + 32
                        Put #f2, LogPos - 2, SSQh.xParm2
                        For x = 1 To SSQh.xParm3
                            Put #f2, LogPos, x
                            Put #f2, LogPos + 4, SSQd3(x)
                            Put #f2, LogPos + 12, SSQd1(x)
                            LogPos = LogPos + 16
                        Next x
                        Close #f2
                        Open "h:\log2.dat" For Binary As #f2
                        LogPos = LOF(f2) + 1
                        Put #f2, LogPos, inFile
                        LogPos = LogPos + 32
                        Put #f2, LogPos - 2, SSQh.xParm2
                        Close #f2
                    End If
                    '***
                    If SSQh.xParm3 > 16 Then
                        Select Case SSQh.xParm2
                            Case &H114: t = MODE_SINGLE + DIFFICULTY_LIGHT
                            Case &H118: t = MODE_DOUBLE + DIFFICULTY_LIGHT
                            Case &H214: t = MODE_SINGLE + DIFFICULTY_STANDARD
                            Case &H218: t = MODE_DOUBLE + DIFFICULTY_STANDARD
                            Case &H314: t = MODE_SINGLE + DIFFICULTY_HEAVY
                            Case &H318: t = MODE_DOUBLE + DIFFICULTY_HEAVY
                            Case &H414: t = MODE_SINGLE + DIFFICULTY_BEGINNER
                            Case &H418: t = MODE_DOUBLE + DIFFICULTY_BEGINNER
                            Case &H614: t = MODE_SINGLE + DIFFICULTY_CHALLENGE
                            Case &H618: t = MODE_DOUBLE + DIFFICULTY_CHALLENGE
                            Case Else
                                x = x
                                Debug.Print "MODE:", Hex(SSQh.xParm2)
                        End Select
                        For x = 1 To SSQh.xParm3
                            If SSQd3(x) <> 0 Then
                                s = s + 2
                                xNoteCount = s
                                With xUNote(s - 1)
                                    .xPlayer = 1
                                    .xOffset192 = ((SSQd1(x) / &H1000&) * 192) + 192
                                    .xStepLeft = CLng(CBool(SSQd3(x) And 1))
                                    .xStepDown = CLng(CBool(SSQd3(x) And 2))
                                    .xStepUp = CLng(CBool(SSQd3(x) And 4))
                                    .xStepRight = CLng(CBool(SSQd3(x) And 8))
                                    .xChart = t
                                End With
                                With xUNote(s)
                                    .xPlayer = 2
                                    .xOffset192 = ((SSQd1(x) / &H1000&) * 192) + 192
                                    .xStepLeft = CLng(CBool(SSQd3(x) And 16))
                                    .xStepDown = CLng(CBool(SSQd3(x) And 32))
                                    .xStepUp = CLng(CBool(SSQd3(x) And 64))
                                    .xStepRight = CLng(CBool(SSQd3(x) And 128))
                                    .xChart = t
                                End With
                                LastStep = s
                                z = x
                            Else 'freeze
                            
                            
                    '********
                    'NOTE ***
                    
                    ' This freeze arrow detection method is incomplete and is not
                    ' guaranteed to work properly under all circumstances. Its behavior
                    ' is not accurate when dealing with two steps that happen at
                    ' the same time with ONLY one of them being a freeze.   -Saxx
                    
                    '********
                    '********
                    
                            
                                fr = False
                                s = s + 2
                                xNoteCount = s
                                y = 255
                                
                                If SSQd3(z) <> 1 And SSQd3(z) <> 2 And SSQd3(z) <> 4 And SSQd3(z) <> 8 And SSQd3(z) <> 16 And SSQd3(z) <> 32 And SSQd3(z) <> 64 And SSQd3(z) <> 128 Then
                                    'read-ahead for jumps to find out which arrow
                                    'should receive the freeze - at least this will
                                    'help us eliminate "impossible" freezes
                                    y = SSQd3(z)
                                    For a = 0 To 7
                                        b = 2 ^ a
                                        If CBool(y And b) Then
                                            For c = x + 1 To SSQh.xParm3
                                                If CBool(SSQd3(c) And b) And (SSQd1(c) <= SSQd1(x)) And (SSQd1(c) > SSQd1(z)) Then
                                                    'blocked arrow
                                                    y = y Xor b
                                                    Exit For
                                                End If
                                            Next c
                                        End If
                                    Next a
                                    If y <> 1 And y <> 2 And y <> 4 And y <> 8 And y <> 16 And y <> 32 And y <> 64 And y <> 128 Then
                                        'at this point we have no way to determine the
                                        'freeze step for sure
                                        If SSQd3(z + 1) = 0 And SSQd3(z + 2) <> 0 Then 'verify that it is in fact a one-freeze and not 2
                                            If bLog Then
                                                ' more loggage
                                                f2 = FreeFile
                                                Open "h:\log2.dat" For Binary As #f2
                                                LogPos = LOF(f2) + 1
                                                Put #f2, LogPos, "1f  "
                                                Put #f2, LogPos + 4, y
                                                Put #f2, LogPos + 8, SSQd1(z)
                                                Put #f2, LogPos + 12, (SSQd1(z + 1) - SSQd1(z))
                                                Close #f2
                                            End If
                                            bFreezeNotice = True
                                            If fnc = 0 Then
                                                e = True
                                            ElseIf xFreezeNotices(fnc).xOffset192 <> (((SSQd1(z) / &H1000&) * 192) + 192) Then
                                                e = True
                                            End If
                                            If e Then
                                                fnc = fnc + 1
                                                ReDim Preserve xFreezeNotices(0 To fnc) As FreezeNotice
                                                xFreezeNotices(fnc).xOffset192 = ((SSQd1(z) / &H1000&) * 192) + 192
                                                xFreezeNotices(fnc).xStepNumber = t
                                                e = False
                                            End If
                                            'by default, just take the leftmost arrow
                                            If CBool(y And 1) Then y = 1
                                            If CBool(y And 2) Then y = 2
                                            If CBool(y And 4) Then y = 4
                                            If CBool(y And 8) Then y = 8
                                            If CBool(y And 16) Then y = 16
                                            If CBool(y And 32) Then y = 32
                                            If CBool(y And 64) Then y = 64
                                            If CBool(y And 128) Then y = 128
                                            SSQd3(z) = SSQd3(z) Xor y
                                        End If
                                    End If
                                End If
                                
                                With xUNote(s - 1)
                                    'freeze search
                                    .xPlayer = 1
                                    .xOffset192 = ((SSQd1(x) / &H1000&) * 192) + 192
                                    If (xUNote(LastStep - 1).xStepLeft <> 0) And CBool(y And 1) Then
                                        .xEndFreezeLeft = xUNote(LastStep - 1).xStepLeft
                                        xUNote(LastStep - 1).xStepLeft = 0
                                        xUNote(LastStep - 1).xFreezeLeft = -1
                                        fr = True
                                    ElseIf (xUNote(LastStep - 1).xStepDown <> 0) And CBool(y And 2) Then
                                        .xEndFreezeDown = xUNote(LastStep - 1).xStepDown
                                        xUNote(LastStep - 1).xStepDown = 0
                                        xUNote(LastStep - 1).xFreezeDown = -1
                                        fr = True
                                    ElseIf (xUNote(LastStep - 1).xStepUp <> 0) And CBool(y And 4) Then
                                        .xEndFreezeUp = xUNote(LastStep - 1).xStepUp
                                        xUNote(LastStep - 1).xStepUp = 0
                                        xUNote(LastStep - 1).xFreezeUp = -1
                                        fr = True
                                    ElseIf (xUNote(LastStep - 1).xStepRight <> 0) And CBool(y And 8) Then
                                        .xEndFreezeRight = xUNote(LastStep - 1).xStepRight
                                        xUNote(LastStep - 1).xStepRight = 0
                                        xUNote(LastStep - 1).xFreezeRight = -1
                                        fr = True
                                    End If
                                    .xChart = t
                                End With
                                If Not fr Then
                                    'there wasn't a freeze on the 1p side, so let's search 2p
                                    With xUNote(s)
                                        .xPlayer = 2
                                        .xOffset192 = ((SSQd1(x) / &H1000&) * 192) + 192
                                        If (xUNote(LastStep).xStepLeft <> 0) And CBool(y And 16) Then
                                            .xEndFreezeLeft = xUNote(LastStep).xStepLeft
                                            xUNote(LastStep).xStepLeft = 0
                                            xUNote(LastStep).xFreezeLeft = -1
                                        ElseIf (xUNote(LastStep).xStepDown <> 0) And CBool(y And 32) Then
                                            .xEndFreezeDown = xUNote(LastStep).xStepDown
                                            xUNote(LastStep).xStepDown = 0
                                            xUNote(LastStep).xFreezeDown = -1
                                        ElseIf (xUNote(LastStep).xStepUp <> 0) And CBool(y And 64) Then
                                            .xEndFreezeUp = xUNote(LastStep).xStepUp
                                            xUNote(LastStep).xStepUp = 0
                                            xUNote(LastStep).xFreezeUp = -1
                                        ElseIf (xUNote(LastStep).xStepRight <> 0) And CBool(y And 128) Then
                                            .xEndFreezeRight = xUNote(LastStep).xStepRight
                                            xUNote(LastStep).xStepRight = 0
                                            xUNote(LastStep).xFreezeRight = -1
                                        End If
                                        .xChart = t
                                    End With
                                End If
                            End If
                        Next x
                    End If
                End If
            Case 4 'bga changes
                'todo
        End Select
        o = o + SSQh.xSize
    Loop
    Close #f
    Close #f2
    
    'sort and write
    WriteSM outFile, oT
    ConvertSSQ = (xNoteCount > 0)
    
End Function

Private Sub WriteSM(outFile As String, Optional outTitle As String = "", Optional outArtist As String = "", Optional outGap As Long = 0, Optional outMusic As String = "")
    Dim f As Long
    Dim x As Long
    Dim fnc As Long
    Dim y As Long
    Dim z As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim m As Long
    Dim t As Long
    Dim p As Long
    Dim MeasureDivisor As Long
    Dim mm As Double
    Dim e As Boolean
    Dim en As Boolean
    Dim nc As Long
    Dim s As String
    Dim n As String
    Dim tn As Long
    If xNoteCount < 0 Then
        Exit Sub
    End If
    f = FreeFile
    Open outFile For Output As #f
    fnc = 1
    
    'a little information
    Print #f, "// DDRu rev"; CStr(App.Revision)
    
    'this will only be needed until freeze notes are fixed
    If bFreezeNotice Then
        Print #f, "// "
        Print #f, "// NOTE: Measures with inaccurate freezes have been marked with a comment like this:"
        Print #f, "// INACCURATE FREEZE"
        Print #f, "// "
    End If
    
    'title card
    Print #f, "#TITLE:" + outTitle + ";"
    Print #f, "#ARTIST:" + outArtist + ";"
    Print #f, "#OFFSET:" + CStr((-outGap) / 1000) + ";"
    Print #f, "#MUSIC:" + outMusic + ";"
    
    'bpms
    Print #f, "#BPMS:";
    a = UBound(xUBPM)
    For x = 0 To a
        With xUBPM(x)
            If (.xFlags And 1) = 0 And .xBPM > 0 Then
                p = (.xOffset192 / 48) ' - 4
                'If p < 0 Then
                '    p = 0
                'End If
                e = True
                .xBPM = Round(.xBPM, 2)
                If x > 0 Then
                    If xUBPM(x - 1).xBPM = .xBPM Then
                        e = False
                    End If
                End If
                If e Then
                    Print #f, CStr(p) + "=" + CStr(.xBPM);
                    If x < a Then
                        Print #f, ",";
                    End If
                End If
            End If
        End With
    Next x
    Print #f, ";"
    
    'stops
    Print #f, "#STOPS:";
    For x = 1 To a
        With xUBPM(x)
            If (.xFlags And 1) = 1 And .xBPM > 0 Then
                p = (.xOffset192 / 48) ' - 4
                If p < 0 Then p = 0
                Print #f, CStr(p) + "=" + CStr(.xBPM);
                If x < a Then
                    Print #f, ",";
                End If
            End If
        End With
    Next x
    Print #f, ";"
    
    'write steps
    tn = xNoteCount
    Do
        Select Case m
            Case 0: t = MODE_SINGLE + DIFFICULTY_LIGHT
            Case 1: t = MODE_SINGLE + DIFFICULTY_STANDARD
            Case 2: t = MODE_SINGLE + DIFFICULTY_HEAVY
            Case 3: t = MODE_DOUBLE + DIFFICULTY_LIGHT
            Case 4: t = MODE_DOUBLE + DIFFICULTY_STANDARD
            Case 5: t = MODE_DOUBLE + DIFFICULTY_HEAVY
            Case 6: t = MODE_COUPLE + DIFFICULTY_LIGHT
            Case 7: t = MODE_COUPLE + DIFFICULTY_STANDARD
            Case 8: t = MODE_COUPLE + DIFFICULTY_HEAVY
            Case 9: t = MODE_SINGLE + DIFFICULTY_BEGINNER
            Case 10: t = MODE_DOUBLE + DIFFICULTY_BEGINNER
            Case 11: t = MODE_COUPLE + DIFFICULTY_BEGINNER
            Case 12: t = MODE_SINGLE + DIFFICULTY_CHALLENGE
            Case 13: t = MODE_DOUBLE + DIFFICULTY_CHALLENGE
            Case 14: t = MODE_COUPLE + DIFFICULTY_CHALLENGE
            Case 15: t = MODE_SOLO + DIFFICULTY_BEGINNER
            Case 16: t = MODE_SOLO + DIFFICULTY_LIGHT
            Case 17: t = MODE_SOLO + DIFFICULTY_STANDARD
            Case 18: t = MODE_SOLO + DIFFICULTY_HEAVY
            Case 19: t = MODE_SOLO + DIFFICULTY_CHALLENGE
            Case 20: t = MODE_SOLO3STEP + DIFFICULTY_BEGINNER
            Case 21: t = MODE_SOLO3STEP + DIFFICULTY_LIGHT
            Case 22: t = MODE_SOLO3STEP + DIFFICULTY_STANDARD
            Case 23: t = MODE_SOLO3STEP + DIFFICULTY_HEAVY
            Case 24: t = MODE_SOLO3STEP + DIFFICULTY_CHALLENGE
            Case 25: Exit Do
        End Select
        Select Case (t And MODE_MASK)
            Case MODE_SINGLE: s = "dance-single": z = 4: p = 1
            Case MODE_DOUBLE: s = "dance-double": z = 4: p = 2
            Case MODE_COUPLE: s = "dance-couple": z = 4: p = 2
            Case MODE_SOLO: s = "dance-solo": z = 6: p = 1
            Case MODE_SOLO3STEP: s = "dance-3panel": z = 3: p = 1
        End Select
        s = "     " + s + ":" + vbCrLf
        s = s + "     " + ":" + vbCrLf + "     "
        Select Case (t And DIFFICULTY_MASK)
            Case DIFFICULTY_LIGHT: s = s + "Easy"
            Case DIFFICULTY_STANDARD: s = s + "Medium"
            Case DIFFICULTY_HEAVY: s = s + "Hard"
            Case DIFFICULTY_BEGINNER: s = s + "Beginner"
            Case DIFFICULTY_CHALLENGE: s = s + "Challenge"
        End Select
        s = s + ":" + vbCrLf
        s = s + "     " + "1" + ":" + vbCrLf
        s = s + "     " + "1.000,1.000,1.000,1.000,1.000" + ":"
        e = False
        nc = 0
        For x = 0 To tn
            If xUNote(x).xChart = t Then
                If xUNote(x).xPlayer <= p Then
                    If (xUNote(x).xStepDown Or xUNote(x).xStepUp Or xUNote(x).xStepRight Or xUNote(x).xStepLeft Or xUNote(x).xStepUL Or xUNote(x).xStepUR) Or _
                        (xUNote(x).xFreezeDown Or xUNote(x).xFreezeUp Or xUNote(x).xFreezeRight Or xUNote(x).xFreezeLeft Or xUNote(x).xFreezeUL Or xUNote(x).xFreezeUR) _
                    Then
                        e = True
                        nc = nc + 1
                    End If
                End If
            End If
        Next x
        If e = True Then
            
            Print #f, "// ----------------------- Step count: "; CStr(nc)
            Print #f, "#NOTES:"
            Print #f, s;
            y = 1
            a = 0
            Print #f, "   // measure " + CStr(y)
            Do
                MeasureDivisor = 48 'minimum measure size: 4 beats - 192/4=48
                mm = 1
                s = String$(z * 192 * p, 48)
                e = True
                
                
                For x = a To tn
                    With xUNote(x)
                        If .xChart = t Then
                            If .xPlayer <= p Then
                                If .xOffset192 >= (y * 192) And .xOffset192 < ((y + 1) * 192) Then
                                    en = False
                                    If z = 4 Then '4-panel setup
                                        If .xStepLeft Then Mid$(s, (1 + ((z * p * (.xOffset192 - (y * 192)) * mm))) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xStepDown Then Mid$(s, (2 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xStepUp Then Mid$(s, (3 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xStepRight Then Mid$(s, (4 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xFreezeLeft Then Mid$(s, (1 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "2": en = True
                                        If .xFreezeDown Then Mid$(s, (2 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "2": en = True
                                        If .xFreezeUp Then Mid$(s, (3 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "2": en = True
                                        If .xFreezeRight Then Mid$(s, (4 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "2": en = True
                                        If .xEndFreezeLeft Then Mid$(s, (1 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "3": en = True
                                        If .xEndFreezeDown Then Mid$(s, (2 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "3": en = True
                                        If .xEndFreezeUp Then Mid$(s, (3 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "3": en = True
                                        If .xEndFreezeRight Then Mid$(s, (4 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "3": en = True
                                    ElseIf z = 6 Then '6-panel setup
                                        If .xStepLeft Then Mid$(s, (1 + ((z * p * (.xOffset192 - (y * 192)) * mm))) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xStepUL Then Mid$(s, (2 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xStepDown Then Mid$(s, (3 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xStepUp Then Mid$(s, (4 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xStepUR Then Mid$(s, (5 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xStepRight Then Mid$(s, (6 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                    ElseIf z = 3 Then '3-panel setup
                                        If .xStepUL Then Mid$(s, (1 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xStepDown Then Mid$(s, (2 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                        If .xStepUR Then Mid$(s, (3 + (z * p * (.xOffset192 - (y * 192)) * mm)) + ((.xPlayer - 1) * z)) = "1": en = True
                                    End If
                                    If en Then
                                        'this code finds the smallest time unit that can be used
                                        'to appropriately represent the steps - saves a lot of
                                        'space in the SM
                                        Do While ((.xOffset192 - (y * 192)) Mod MeasureDivisor <> 0) And (MeasureDivisor > 1)
                                            MeasureDivisor = MeasureDivisor \ 2
                                        Loop
                                    End If
                                End If
                            End If
                        End If
                        If .xOffset192 > (y * 192) Then
                            e = False
                        End If
                    End With
                Next x
                
                y = y + 1
                For b = 1 To Len(s) Step ((z * p) * MeasureDivisor)
                    Print #f, Mid$(s, b, (z * p))
                Next b
                If e = False Then
                    Print #f, ","; "  // measure "; CStr(y); " ";
                    If bFreezeNotice Then
                        For c = 1 To UBound(xFreezeNotices)
                            If (xFreezeNotices(c).xOffset192 < ((y + 1) * 192)) And (xFreezeNotices(c).xOffset192 >= ((y + 0) * 192)) Then
                                If xFreezeNotices(c).xStepNumber = t Then
                                    Print #f, "// INACCURATE FREEZE";
                                    Exit For
                                End If
                            End If
                        Next c
                    End If
                    Print #f, ""
                Else
                    Print #f, ";"
                    Exit Do
                End If
            Loop
        End If
        m = m + 1
    Loop
    Close #f
End Sub

Private Sub Class_Initialize()
    xNoteCount = -1
    DoEvents
End Sub

Public Function ConvertStep2(inFile As String, outFile As String)
    Dim f As Long
    Dim o As Long
    Dim x As Long
    Dim y As Long
    Dim p As Long
    Dim h As Long
    Dim StepPosition(1 To 2) As Long
    Dim s2d() As StepForm2BPM
    Dim s2() As Byte
    Dim PlayerCount As Long
    
    PlayerCount = 1
    xNoteCount = -1
    f = FreeFile
    
    'load header data
    Open inFile For Binary As #f
    o = 1
    Get #f, o, h
    ReDim s2d(0 To ((h - 4) \ &H14) - 1) As StepForm2BPM
    Get #f, o + 4, s2d
    o = o + h
    
    'old DDR step files don't contain BPM... default to 120
    ReDim xUBPM(0) As UniversalBPM
    xUBPM(0).xBPM = 120
    xUBPM(0).xFlags = 0
    xUBPM(0).xOffset192 = 0
    
    'load notes
    y = LOF(f) - h
    y = (y \ PlayerCount) - 1
    ReDim s2(0 To y) As Byte
    Get #f, o, s2
    Close #f
    
    'determine whether it's single or double
    p = 1
    PlayerCount = 1
    For y = 1 To UBound(s2d)
        If s2d(y).xStepStart = s2d(y - 1).xStepStart And s2d(y).xStepEnd = s2d(y - 1).xStepEnd And s2d(y).xStepStart > 0 And (s2d(y).xStepStart <> s2d(UBound(s2d)).xStepStart) Then
            PlayerCount = 2
            If (y Mod 2) = 0 Then
                p = 2
            End If
            Exit For
        End If
    Next y
    
    'each byte in a STEP2 file is 1/64th
    For y = 0 To UBound(s2d)
        o = s2d(y).xOffset - h
        For x = 0 To (s2d(y).xStepSize * 4) - 1
            If s2(x + o) <> 0 Then
                xNoteCount = xNoteCount + 1
                With xUNote(xNoteCount)
                    If PlayerCount = 1 Then
                        .xChart = MODE_SINGLE + DIFFICULTY_LIGHT
                    ElseIf PlayerCount = 2 Then
                        .xChart = MODE_DOUBLE + DIFFICULTY_LIGHT
                    End If
                    .xOffset192 = StepPosition(p) * 3 'because 192\64=3
                    .xPlayer = p
                    .xStepLeft = CBool(s2(x + o) And 1)
                    .xStepRight = CBool(s2(x + o) And 2)
                    .xStepUp = CBool(s2(x + o) And 4)
                    .xStepDown = CBool(s2(x + o) And 8)
                End With
            End If
            StepPosition(p) = StepPosition(p) + 1
        Next x
        If y = 0 Then
            If p = 2 Then
                StepPosition(1) = StepPosition(2)
            Else
                StepPosition(2) = StepPosition(1)
            End If
        End If
        If PlayerCount = 2 Then
            If p = 1 Then
                p = 2
            Else
                p = 1
            End If
        End If
    Next y
    
    'sort and write
    SortNotes
    WriteSM outFile
    
End Function

Private Function TweakFreeze(xOffset As Long, xStep As Byte) As Byte
    
End Function
