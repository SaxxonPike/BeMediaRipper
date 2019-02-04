VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOffsetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Offset List functions for all games
' -> files and song names
'
' SaxxonPike 2006-2007

Private Type DDRXInfo '152b
    SongName As String * 6
    songfileID2 As Integer
    songfileID As Long
    songlistID As Long
    previewfileID As Long
    unk1a As Long
    unk1b As Long
    unk2a As Integer
    unk2b As Integer
    unk3 As Long
    stepinfo(0 To 11) As Integer
    extradata(0 To 103) As Byte
End Type

Private Type DDRSNInfo '168b
    SongName As String * 6
    unk0a As Byte
    unk0b As Byte
    songlistID As Integer
    songlistID2 As Integer
    songfileID As Long
    previewfileID As Long
    unk1a As Long
    unk1b As Long
    unk1c As Long
    unk2a As Integer
    unk2b As Integer
    unk3a As Long
    unk3b As Long
    unk3c As Long
    unk3d As Long
    unk4a As Long
    unk4b As Long
    unk4c As Long
    stepinfo(0 To 7) As Integer
    extradata(0 To 43) As Integer
End Type

Private Type OffsetOld12
    FOffs800 As Long
    FLength As Long
    Unk0 As Long
End Type

Private Type OffsetNew8
    FOffs800 As Long
    FLength As Long
End Type

Private Type InfoNew280
    SongName As String * 64
    Unk0(0 To 5) As Integer
    Difficulty(0 To 5) As Byte
    FolderNumber As Integer
    unk1(0 To 7) As Byte
    VideoNumbers(0 To 3) As Integer
    Unk2(0 To 83) As Byte
    KeyOffsets(0 To 7) As Long
    SetNumbers(0 To 15) As Integer
    ExtraInfo(0 To 15) As Integer
End Type

Private Type InfoNew284
    SongName As String * 64
    Unk0(0 To 5) As Integer
    Difficulty(0 To 5) As Byte
    FolderNumber As Integer
    unk1(0 To 7) As Byte
    VideoNumbers(0 To 3) As Integer
    Unk2(0 To 79) As Byte
    KeyOffsets(0 To 9) As Long
    SetNumbers(0 To 15) As Integer
    ExtraInfo(0 To 15) As Integer
End Type

Private Type InfoNew308
    SongName As String * 64
    Unk0(0 To 4) As Integer
    Difficulty(0 To 9) As Byte
    FolderNumber As Integer
    unk1(0 To 105) As Byte
    KeyOffsets(0 To 9) As Long
    SetNumbers(0 To 15) As Integer
    ExtraInfo(0 To 21) As Integer
End Type

Private Type InfoNew324
    SongName As String * 64
    Unk0(0 To 4) As Integer
    Difficulty(0 To 7) As Byte
    unk1(0 To 7) As Byte
    VideoNumbers(0 To 3) As Integer
    Unk2(0 To 137) As Byte
    KeyOffsets(0 To 7) As Long
    SetNumbers(0 To 15) As Integer
    ExtraInfo(0 To 11) As Integer
End Type

Private Type InfoNew320
    SongName As String * 64
    Unk0(0 To 5) As Integer
    Difficulty(0 To 7) As Byte
    unk1(0 To 7) As Byte
    VideoNumbers(0 To 3) As Integer
    Unk2(0 To 155) As Byte
    KeyOffsets(0 To 7) As Long
    SetNumbers(0 To 15) As Integer
End Type

Private Type InfoNew372US
    SongName As String * 64
    Unk0(0 To 2) As Integer
    MinBPM As Integer
    MaxBPM As Integer
    Difficulty(0 To 7) As Byte
    unk1(0 To 9) As Byte
    VideoNumbers(0 To 3) As Integer
    Unk2(0 To 203) As Byte
    KeyOffsets(0 To 7) As Long
    SetNumbers(0 To 17) As Integer
End Type

Private Type InfoNew364
    SongName As String * 64
    Unk0(0 To 4) As Integer
    Difficulty(0 To 6) As Byte
    unk1(0 To 6) As Byte
    VideoNumbers(0 To 3) As Integer
    Unk2(0 To 203) As Byte
    KeyOffsets(0 To 7) As Long
    SetNumbers(0 To 15) As Integer
End Type

Private Type OffsetOld16
    FOffs800 As Long
    Unk0 As Long
    FLength As Long
    unk1 As Long
End Type

Private Type InfoOld164
    ShortNameOffset As Long
    LongNameOffset As Long
    Difficulty(0 To 5) As Byte
    Unk0(0 To 2) As Integer
    unk1(0 To 17) As Long
    KeyOffsets(0 To 7) As Long
    SetNumbers(0 To 19) As Integer
End Type

Private Type InfoOld160
    ShortNameOffset As Long
    LongNameOffset As Long
    unk3(0 To 1) As Long
    Difficulty(0 To 5) As Byte
    Unk0(0 To 2) As Integer
    unk1(0 To 14) As Long
    KeyOffsets(0 To 7) As Long
    SetNumbers(0 To 19) As Integer
End Type

Private Type InfoOld1606th
    ShortNameOffset As Long
    LongNameOffset As Long
    unk3(0 To 1) As Long
    Difficulty(0 To 5) As Byte
    Unk0(0 To 2) As Integer
    unk1(0 To 14) As Long
    KeyOffsets(0 To 7) As Long
    SetNumbers(0 To 19) As Integer
End Type

Private Type InfoOld156
    ShortNameOffset As Long
    LongNameOffset As Long
    Difficulty(0 To 5) As Byte
    Unk0(0 To 6) As Integer
    unk1(0 To 13) As Long
    KeyOffsets(0 To 7) As Long
    SetNumbers(0 To 19) As Integer
End Type

Private Type InfoOld124 '3rd
    ShortNameOffset As Long
    LongNameOffset As Long
    Unk0 As Long
    Difficulty(0 To 3) As Byte
    unk1(0 To 16) As Long
    KeyOffsets(0 To 5) As Long
    SetNumbers(0 To 7) As Integer
End Type

Private Type InfoOld144 '4th
    ShortNameOffset As Long
    LongNameOffset As Long
    Difficulty(0 To 3) As Byte
    Unk0(0 To 1) As Integer
    unk1(0 To 17) As Long
    KeyOffsets(0 To 5) As Long
    SetNumbers(0 To 15) As Integer
End Type

Private Type DDRInfo20
    SongName As String * 8
    FileDataID As Long
    Info1 As Long
    Info2 As Long
End Type

'Private Type InfoOld160
    
'End Type

Private Type OffsetUniversal
    FLength As Double
    FOffs As Double
    fLink As Long
    bRipped As Boolean
    SongName As String
End Type

Private bUseList As Boolean
Private xList() As OffsetUniversal
Private yList() As OffsetUniversal
Private ff As Long
Private xListFile As String

Public Property Get ListIsLoaded() As Boolean
    ListIsLoaded = bUseList
End Property

Public Sub ReadFileList(InfoType As Long, offs As Long)
    Dim OO16 As OffsetOld16
    Dim ON8 As OffsetNew8
    Dim OO12 As OffsetOld12
    Dim TDbl As Double
    Dim Coffs As Long
    Dim ZeroCount As Long
    Coffs = offs + 1
    ReDim xList(0) As OffsetUniversal
    ReDim yList(0) As OffsetUniversal
    Select Case InfoType
        Case 1
            Do
                Get #ff, Coffs, OO16
                Coffs = Coffs + 16
                If OO16.FLength = 0 And OO16.FOffs800 = 0 Then
                    Exit Do
                End If
                If UBound(xList) > 0 Then
                    TDbl = OO16.FOffs800
                    TDbl = TDbl * &H800
                    If TDbl < (xList(UBound(xList) - 1).FOffs) Then
                        Exit Do
                    End If
                End If
                With xList(UBound(xList))
                    .FLength = OO16.FLength
                    .FOffs = OO16.FOffs800
                    .FOffs = .FOffs * 2048
                    If UBound(xList) > 0 Then
                        If (xList(UBound(xList) - 1).FLength = xList(UBound(xList)).FLength) And (OO16.Unk0 = 0 And OO16.unk1 = 0) Then
                            '.FLength = 0
                        End If
                    End If
                    .SongName = ""
                End With
                ReDim Preserve xList(0 To UBound(xList) + 1) As OffsetUniversal
            Loop
        Case 2
            Do
                Get #ff, Coffs, OO16
                Coffs = Coffs + 16
                If OO16.FLength = 0 And OO16.FOffs800 = 0 Then
                    Exit Do
                End If
                If UBound(xList) > 0 Then
                    If OO16.FOffs800 < (xList(UBound(xList) - 1).FOffs \ 2048) Then
                        Exit Do
                    End If
                End If
                With xList(UBound(xList))
                    .FLength = OO16.FLength
                    .FOffs = OO16.FOffs800
                    .FOffs = .FOffs * 2048
                    If UBound(xList) > 0 Then
                        If (xList(UBound(xList) - 1).FLength = xList(UBound(xList)).FLength) And (OO16.Unk0 = 0 And OO16.unk1 = 0) Then
                            .FLength = 0
                        End If
                    End If
                    .SongName = ""
                End With
                ReDim Preserve xList(0 To UBound(xList) + 1) As OffsetUniversal
            Loop
        Case 3
            Do
                Get #ff, Coffs, ON8
                Coffs = Coffs + 8
                If ON8.FLength = 0 Then
                    Exit Do
                End If
                'If UBound(xList) > 0 Then
                '    If ON8.FOffs800 < (xList(UBound(xList) - 1).FOffs \ 2048) Then
                '        Exit Do
                '    End If
                'End If
                
                If ON8.FLength > 0 And (UBound(xList) = 0 Or ON8.FOffs800 > 0) Then
                With xList(UBound(xList))
                    .FLength = ON8.FLength
                    .FOffs = ON8.FOffs800
                    .FOffs = .FOffs * 2048
                    .FLength = .FLength * 2048
                    .SongName = ""
                    Debug.Print UBound(xList), "O:"; CStr(.FOffs), "L:"; CStr(.FLength), ;
                    If UBound(xList) > 0 Then
                        Debug.Print "E:"; CStr(xList(UBound(xList) - 1).FOffs + xList(UBound(xList) - 1).FLength)
                    Else
                        Debug.Print
                    End If
                End With
                ReDim Preserve xList(0 To UBound(xList) + 1) As OffsetUniversal
                Else
                    Debug.Print "============= END OF LIST ============="
                    Exit Do
                End If
            Loop
        Case 4 '4th style
            OO12.FOffs800 = 0
            Do
                'ZeroCount = OO12.FOffs800
                Get #ff, Coffs, OO12
                Coffs = Coffs + 12
                If OO12.FLength = 0 Then ' Or ZeroCount > OO12.FOffs800 Then
                    Exit Do
                End If
                With xList(UBound(xList))
                    .FLength = OO12.FLength
                    .FOffs = OO12.FOffs800
                    .FOffs = .FOffs * 2048
                    .SongName = ""
                End With
                ReDim Preserve xList(0 To UBound(xList) + 1) As OffsetUniversal
            Loop
        Case 5 'popn PSX
            ZeroCount = 1001
            Do
                Get #ff, Coffs, ON8
                Coffs = Coffs + 8
                If ON8.FLength = 0 Then
                    Exit Do
                End If
                'If UBound(xList) > 0 Then
                '    If ON8.FOffs800 < (xList(UBound(xList) - 1).FOffs \ 2048) Then
                '        Exit Do
                '    End If
                'End If
                
                If ON8.FLength > 0 And (UBound(xList) = 0 Or ON8.FOffs800 > 0) Then
                With xList(UBound(xList))
                    .FLength = ON8.FLength
                    .FOffs = ON8.FOffs800
                    .FOffs = .FOffs * 2048
                    .SongName = "untitled" + CStr(ZeroCount)
                End With
                ZeroCount = ZeroCount + 1
                ReDim Preserve xList(0 To UBound(xList) + 1) As OffsetUniversal
                Else
                    Exit Do
                End If
                bUseList = True
            Loop
        Case Else
    End Select
End Sub

Public Property Get ListCount() As Long
    ListCount = UBound(xList)
End Property

Public Sub ReadInfoList(InfoType As Long, offs As Long, NumberAdjust As Long, bAutoRipCharts As Boolean)
    Dim IO164 As InfoOld164
    Dim IO160 As InfoOld160
    Dim IO156 As InfoOld156
    Dim IN364 As InfoNew364
    Dim IN372 As InfoNew372US
    Dim IN324 As InfoNew324
    Dim IN320 As InfoNew320
    Dim IN280 As InfoNew280
    Dim IN284 As InfoNew284
    Dim IN308 As InfoNew308
    Dim IO144 As InfoOld144
    Dim IO124 As InfoOld124
    Dim DDRI20 As DDRInfo20
    Dim DDRSN168 As DDRSNInfo
    Dim DDRX152 As DDRXInfo
    
    Dim ChartExtract As New clsPSXDecode
    Dim ExtractBuffer() As Byte
    Dim ExtractFileObj As New clsFileStream
    Dim Coffs As Long
    Dim Csong As String
    Dim lTag As Long
    Dim X As Long
    Dim Y As Long
    Dim DSong As String
    lTag = 8192
    Coffs = offs + 1
    'MsgBox Len(IO144)
    'MsgBox Len(IO124)
    
    Select Case InfoType
        Case 1 '5th
            Do
                Get #ff, Coffs, IO164
                Coffs = Coffs + 164
                With IO164
                    If (Abs(.LongNameOffset - .ShortNameOffset) Mod 8) <> 0 Then
                        Exit Do
                    End If
                    If .LongNameOffset <> 0 Then
                        Csong = ApplyFilter(LoadTitle(.LongNameOffset - &HFF000))
                    End If
                    If .Difficulty(0) > 0 Then
                        For X = 0 To 19
                            If .SetNumbers(X) > 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        For X = 0 To 7
                            If .KeyOffsets(X) > 0 Then
                                For Y = 0 To UBound(yList)
                                    If yList(Y).FOffs = (.KeyOffsets(X) - &HFF000) Then
                                        Exit For
                                    End If
                                    If Y = UBound(yList) Then
                                        Select Case X
                                            Case 0: yList(UBound(yList)).SongName = Csong + " [7key] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                            Case 1: yList(UBound(yList)).SongName = Csong + " [14key] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                            Case 2: yList(UBound(yList)).SongName = Csong + " [Another7] (10) b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                            Case 3: yList(UBound(yList)).SongName = Csong + " [Another14] (10) b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                            Case 4: yList(UBound(yList)).SongName = Csong + " [Light7] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                            Case 5: yList(UBound(yList)).SongName = Csong + " [Light14] (" + CStr(.Difficulty(3)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                            Case 6: yList(UBound(yList)).SongName = Csong + " [Beginner7] (" + CStr(.Difficulty(4)) + ") b" + CStr(.SetNumbers(14)) + " k" + CStr(.SetNumbers(12))
                                            Case 7: yList(UBound(yList)).SongName = Csong + " [Beginner14] (" + CStr(.Difficulty(5)) + ") b" + CStr(.SetNumbers(15)) + " k" + CStr(.SetNumbers(13))
                                        End Select
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " b-1", "")
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " k-1", "")
                                        Debug.Print yList(UBound(yList)).SongName
                                        yList(UBound(yList)).FOffs = .KeyOffsets(X) - &HFF000
                                        ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                        Exit For
                                    End If
                                Next Y
                            End If
                        Next X
                    End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
        Case 2, 8 '7th
            Do
                Get #ff, Coffs, IO160
                Coffs = Coffs + 160
                With IO160
                    If (Abs(.LongNameOffset - .ShortNameOffset) Mod 8) <> 0 Then
                        Exit Do
                    End If
                    If .LongNameOffset <> 0 Then
                        If InfoType = 2 Then
                            Csong = ApplyFilter(LoadTitle(((.LongNameOffset - &H100000) + &H80)))
                        Else
                            Csong = ApplyFilter(LoadTitle(((.LongNameOffset - &HFF000))))
                        End If
                    End If
                    If .Difficulty(0) > 0 Then
                        For X = 0 To 19
                            If .SetNumbers(X) > 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                        Debug.Print .SetNumbers(X), Csong
                                    Else
                                        X = X
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        For X = 0 To 7
                            If .KeyOffsets(X) <> 0 Then
                                For Y = 0 To UBound(yList)
                                    If Y = UBound(yList) Then
                                        Select Case X
                                            Case 0: yList(UBound(yList)).SongName = Csong + " [7key] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                            Case 1: yList(UBound(yList)).SongName = Csong + " [14key] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                            Case 2: yList(UBound(yList)).SongName = Csong + " [Another7] (10) b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                            Case 3: yList(UBound(yList)).SongName = Csong + " [Another14] (10) b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                            Case 4: yList(UBound(yList)).SongName = Csong + " [Light7] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                            Case 5: yList(UBound(yList)).SongName = Csong + " [Light14] (" + CStr(.Difficulty(3)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                            Case 6: yList(UBound(yList)).SongName = Csong + " [Beginner7] (" + CStr(.Difficulty(4)) + ") b" + CStr(.SetNumbers(14)) + " k" + CStr(.SetNumbers(12))
                                            Case 7: yList(UBound(yList)).SongName = Csong + " [Beginner14] (" + CStr(.Difficulty(5)) + ") b" + CStr(.SetNumbers(15)) + " k" + CStr(.SetNumbers(13))
                                        End Select
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " b-1", "")
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " k-1", "")
                                        Debug.Print yList(UBound(yList)).SongName
                                        If Left(yList(UBound(yList)).SongName, 8) = "dissolve" Then
                                            X = X
                                        End If
                                        If InfoType = 2 Then
                                            If (.KeyOffsets(X) And &HFFFF&) < UBound(xList) Then
                                                yList(UBound(yList)).FOffs = xList((.KeyOffsets(X) And &HFFFF&)).FOffs '.KeyOffsets(X) - &HFF000
                                                ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                            End If
                                        Else
                                            yList(UBound(yList)).FOffs = (.KeyOffsets(X) - &HFF000)
                                            ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                        End If
                                        Exit For
                                    End If
                                Next Y
                            End If
                        Next X
                    End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
        Case 3 '8th
            Do
                Get #ff, Coffs, IO156
                Coffs = Coffs + 156
                With IO156
                    If (Abs(.LongNameOffset - .ShortNameOffset) Mod 8) <> 0 Then
                        Exit Do
                    End If
                    If .LongNameOffset <> 0 Then
                        Csong = ApplyFilter(LoadTitle((.LongNameOffset - &H100000) + &H80))
                    End If
                    If .Difficulty(0) > 0 Then
                        For X = 0 To 15
                            If .SetNumbers(X) >= 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    'Debug.Print Len(IO156)
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                        Debug.Print .SetNumbers(X), Csong
                                    Else
                                        X = X
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        For X = 0 To 7
                            If .KeyOffsets(X) <> 0 Then
                                For Y = 0 To UBound(yList)
                                    If Y = UBound(yList) Then
                                        Select Case X
                                            Case 0: yList(UBound(yList)).SongName = Csong + " [7key] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                            Case 1: yList(UBound(yList)).SongName = Csong + " [14key] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                            Case 2: yList(UBound(yList)).SongName = Csong + " [Another7] (10) b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                            Case 3: yList(UBound(yList)).SongName = Csong + " [Another14] (10) b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                            Case 4: yList(UBound(yList)).SongName = Csong + " [Light7] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                            Case 5: yList(UBound(yList)).SongName = Csong + " [Light14] (" + CStr(.Difficulty(3)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                            Case 6: yList(UBound(yList)).SongName = Csong + " [Beginner7] (" + CStr(.Difficulty(4)) + ") b" + CStr(.SetNumbers(14)) + " k" + CStr(.SetNumbers(12))
                                            Case 7: yList(UBound(yList)).SongName = Csong + " [Beginner14] (" + CStr(.Difficulty(5)) + ") b" + CStr(.SetNumbers(15)) + " k" + CStr(.SetNumbers(13))
                                        End Select
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " b-1", "")
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " k-1", "")
                                        Debug.Print Len(IO156)
                                        Debug.Print yList(UBound(yList)).SongName
                                        If (.KeyOffsets(X) And &HFFFF&) < UBound(xList) Then
                                            yList(UBound(yList)).FOffs = xList((.KeyOffsets(X) And &HFFFF&)).FOffs '.KeyOffsets(X) - &HFF000
                                            ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                        End If
                                        Exit For
                                    End If
                                Next Y
                            End If
                        Next X
                    End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
        Case 4 'IIDX 9-10
            Do
                Get #ff, Coffs, IN364
                Coffs = Coffs + 364
                With IN364
                    Csong = ApplyFilter(IN364.SongName)
                    If InStr(Csong, Chr(0)) > 0 Then
                        Csong = Left(Csong, InStr(Csong, Chr(0)) - 1)
                    End If
                    If Len(Csong) = 0 Then
                        Exit Do
                    End If
                    If .Difficulty(0) > 0 Then
                        For X = 0 To 15
                            If .SetNumbers(X) >= 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                        Debug.Print .SetNumbers(X), Csong
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        For X = 0 To 7
                            If .KeyOffsets(X) <> 0 Then
                                For Y = 0 To UBound(yList)
                                    If Y = UBound(yList) Then
                                        Select Case X
                                            Case 0: yList(UBound(yList)).SongName = Csong + " [7key] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                            Case 1: yList(UBound(yList)).SongName = Csong + " [14key] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                            Case 2: yList(UBound(yList)).SongName = Csong + " [Another7] (10) b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                            Case 3: yList(UBound(yList)).SongName = Csong + " [Another14] (10) b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                            Case 4: yList(UBound(yList)).SongName = Csong + " [Light7] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                            Case 5: yList(UBound(yList)).SongName = Csong + " [Light14] (" + CStr(.Difficulty(3)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                            Case 6: yList(UBound(yList)).SongName = Csong + " [Beginner7] (" + CStr(.Difficulty(5)) + ") b" + CStr(.SetNumbers(14)) + " k" + CStr(.SetNumbers(12))
                                            Case 7: yList(UBound(yList)).SongName = Csong + " [Beginner14] (1) b" + CStr(.SetNumbers(15)) + " k" + CStr(.SetNumbers(13))
                                        End Select
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " b-1", "")
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " k-1", "")
                                        If .KeyOffsets(X) < UBound(xList) Then
                                            yList(UBound(yList)).FOffs = xList(.KeyOffsets(X)).FOffs
                                            yList(UBound(yList)).fLink = .KeyOffsets(X)
                                            ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                        End If
                                        Exit For
                                    End If
                                Next Y
                            End If
                        Next X
                    End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
        Case 5 'BM US
            Do
                Get #ff, Coffs, IN372
                Coffs = Coffs + 372
                With IN372
                    Csong = ApplyFilter(.SongName)
                    If InStr(Csong, Chr(0)) > 0 Then
                        Csong = Left(Csong, InStr(Csong, Chr(0)) - 1)
                    End If
                    If Len(Csong) = 0 Then
                        Exit Do
                    End If
                    If InStr(Csong, "_") > 0 Then
                        Csong = Mid(Csong, InStr(Csong, "_") + 1)
                    End If
                    Csong = Trim(Csong)
                    If .Difficulty(0) > 0 Then
                        For X = 0 To 15
                            If .SetNumbers(X) >= 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                        Debug.Print .SetNumbers(X), Csong
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        For X = 0 To 7
                            If .KeyOffsets(X) > 0 Then
                                For Y = 0 To UBound(yList)
                                    If Y = UBound(yList) Then
                                        Select Case X
                                            Case 0
                                                If .Difficulty(1) = 0 And .Difficulty(2) = 0 And .Difficulty(3) = 0 Then
                                                    yList(UBound(yList)).SongName = Csong + " [5key] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                                Else
                                                    yList(UBound(yList)).SongName = Csong + " [7key] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                                End If
                                            Case 1
                                                If .Difficulty(5) = 0 And .Difficulty(6) = 0 And .Difficulty(3) = 0 Then
                                                    yList(UBound(yList)).SongName = Csong + " [10key] (" + CStr(.Difficulty(4)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                                Else
                                                    yList(UBound(yList)).SongName = Csong + " [14key] (" + CStr(.Difficulty(5)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                                End If
                                            Case 2: yList(UBound(yList)).SongName = Csong + " [Light7] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                            Case 3: yList(UBound(yList)).SongName = Csong + " [Light14] (" + CStr(.Difficulty(4)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                            Case 4: yList(UBound(yList)).SongName = Csong + " [Another7] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                            Case 5: yList(UBound(yList)).SongName = Csong + " [Another14] (" + CStr(.Difficulty(6)) + ") b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                            Case 6: yList(UBound(yList)).SongName = Csong + " [Beginner7] (" + CStr(.Difficulty(3)) + ") b" + CStr(.SetNumbers(14)) + " k" + CStr(.SetNumbers(12))
                                            Case 7: yList(UBound(yList)).SongName = Csong + " [Beginner14] (" + CStr(.Difficulty(7)) + ") b" + CStr(.SetNumbers(15)) + " k" + CStr(.SetNumbers(13))
                                        End Select
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " b-1", "")
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " k-1", "")
                                        If .KeyOffsets(X) < UBound(xList) Then
                                            yList(UBound(yList)).FOffs = xList(.KeyOffsets(X)).FOffs
                                            yList(UBound(yList)).fLink = .KeyOffsets(X)
                                            ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                        End If
                                        Exit For
                                    End If
                                Next Y
                            End If
                        Next X
                    End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
        Case 6 'RED
            Do
                Get #ff, Coffs, IN320
                Coffs = Coffs + 320
                With IN320
                    Csong = ApplyFilter(.SongName)
                    If InStr(Csong, Chr(0)) > 0 Then
                        Csong = Left(Csong, InStr(Csong, Chr(0)) - 1)
                    End If
                    If Len(Csong) = 0 Then
                        Exit Do
                    End If
                    If .Difficulty(0) > 0 Then
                        For X = 0 To 15
                            If .SetNumbers(X) >= 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                        Debug.Print .SetNumbers(X), Csong
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        For X = 0 To 7
                            If .KeyOffsets(X) <> 0 Then
                                For Y = 0 To UBound(yList)
                                    If Y = UBound(yList) Then
                                        Select Case X
                                            Case 0: yList(UBound(yList)).SongName = Csong + " [7key] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                            Case 1: yList(UBound(yList)).SongName = Csong + " [14key] (" + CStr(.Difficulty(3)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                            Case 2: yList(UBound(yList)).SongName = Csong + " [Another7] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                            Case 3: yList(UBound(yList)).SongName = Csong + " [Another14] (" + CStr(.Difficulty(5)) + ") b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                            Case 4: yList(UBound(yList)).SongName = Csong + " [Light7] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                            Case 5: yList(UBound(yList)).SongName = Csong + " [Light14] (" + CStr(.Difficulty(4)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                            Case 6: yList(UBound(yList)).SongName = Csong + " [Beginner7] (" + CStr(.Difficulty(7)) + ") b" + CStr(.SetNumbers(14)) + " k" + CStr(.SetNumbers(12))
                                            Case 7: yList(UBound(yList)).SongName = Csong + " [Beginner14] (1) b" + CStr(.SetNumbers(15)) + " k" + CStr(.SetNumbers(13))
                                        End Select
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " b-1", "")
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " k-1", "")
                                        If .KeyOffsets(X) < UBound(xList) Then
                                            yList(UBound(yList)).FOffs = xList(.KeyOffsets(X)).FOffs
                                            yList(UBound(yList)).fLink = .KeyOffsets(X)
                                            ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                        End If
                                        Exit For
                                    End If
                                Next Y
                            End If
                        Next X
                    End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
        Case 9 'HAPPYSKY
            Do
                Get #ff, Coffs, IN324
                Coffs = Coffs + 324
                With IN324
                    Csong = ApplyFilter(.SongName)
                    If InStr(Csong, Chr(0)) > 0 Then
                        Csong = Left(Csong, InStr(Csong, Chr(0)) - 1)
                    End If
                    If Len(Csong) = 0 Then
                        Exit Do
                    End If
                    If .Difficulty(0) > 0 Then
                        For X = 0 To 15
                            .SetNumbers(X) = .SetNumbers(X) - NumberAdjust
                            If .SetNumbers(X) < 0 Then
                                .SetNumbers(X) = -1
                            End If
                            If .SetNumbers(X) >= 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                        Debug.Print .SetNumbers(X), Csong
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        For X = 0 To 7
                            .KeyOffsets(X) = .KeyOffsets(X) - NumberAdjust
                            If .KeyOffsets(X) >= 0 Then
                                For Y = 0 To UBound(yList)
                                    If Y = UBound(yList) Then
                                        Select Case X
                                            Case 0: yList(UBound(yList)).SongName = Csong + " [7key] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                            Case 1: yList(UBound(yList)).SongName = Csong + " [14key] (" + CStr(.Difficulty(3)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                            Case 2: yList(UBound(yList)).SongName = Csong + " [Another7] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                            Case 3: yList(UBound(yList)).SongName = Csong + " [Another14] (" + CStr(.Difficulty(5)) + ") b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                            Case 4: yList(UBound(yList)).SongName = Csong + " [Light7] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                            Case 5: yList(UBound(yList)).SongName = Csong + " [Light14] (" + CStr(.Difficulty(4)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                            Case 6: yList(UBound(yList)).SongName = Csong + " [Beginner7] (" + CStr(.Difficulty(7)) + ") b" + CStr(.SetNumbers(14)) + " k" + CStr(.SetNumbers(12))
                                            Case 7: yList(UBound(yList)).SongName = Csong + " [Beginner14] (1) b" + CStr(.SetNumbers(15)) + " k" + CStr(.SetNumbers(13))
                                        End Select
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " b-1", "")
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " k-1", "")
                                        
                                        If .KeyOffsets(X) < UBound(xList) Then
                                            yList(UBound(yList)).FOffs = xList(.KeyOffsets(X)).FOffs
                                            yList(UBound(yList)).fLink = .KeyOffsets(X)
                                            ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                        End If
                                        Exit For
                                    End If
                                Next Y
                            Else
                                X = X
                            End If
                        Next X
                    End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
        Case 7 '3rd style
            ExtractFileObj.OpenFile xListFile
            Do
                Get #ff, Coffs, IO124
                Coffs = Coffs + 124
                With IO124
                    If (Abs(.LongNameOffset - .ShortNameOffset) Mod 8) <> 0 Then
                        Exit Do
                    End If
                    If .LongNameOffset <> 0 Then
                        Csong = ApplyFilter(LoadTitle(.LongNameOffset - &HFF000))
                    End If
                    If .Difficulty(1) > 0 Then
                        For X = 0 To 7
                            If .SetNumbers(X) > 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        DSong = Csong
                        
                        For X = 0 To 5
                            If .KeyOffsets(X) > 0 Then
                                Select Case X
                                    Case 0: DSong = Csong + " [7key] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                    Case 1: DSong = Csong + " [14key] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                    Case 2: DSong = Csong + " [Another7] (10) b" + CStr(.SetNumbers(5)) + " k" + CStr(.SetNumbers(4))
                                    Case 3: DSong = Csong + " [Another14] (10) b" + CStr(.SetNumbers(5)) + " k" + CStr(.SetNumbers(4))
                                    Case 4: DSong = Csong + " [Light7] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                    Case 5: DSong = Csong + " [Light14] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                End Select
                                DSong = Replace(DSong, " b-1", "")
                                DSong = Replace(DSong, " k-1", "")
                                Debug.Print DSong
                                If bAutoRipCharts Then
                                    ChartExtract.DecodeBemani1 ExtractFileObj, ExtractBuffer(), .KeyOffsets(X) - &HFF000
                                    Open OutFolder + DSong + ".cs2" For Binary As #95
                                    Put #95, 1, ExtractBuffer
                                    Close #95
                                End If
                            End If
                        Next X

                    End If
                    
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
            ExtractFileObj.CloseFile
        Case 10 '4th style
            ExtractFileObj.OpenFile xListFile
            Do
                Get #ff, Coffs, IO144
                Coffs = Coffs + 144
                With IO144
                    If (Abs(.LongNameOffset - .ShortNameOffset) Mod 8) <> 0 Then
                        Exit Do
                    End If
                    If .LongNameOffset <> 0 Then
                        Csong = ApplyFilter(LoadTitle(.LongNameOffset - &HFF000))
                    End If
                    If .Difficulty(1) > 0 Then
                        For X = 0 To 15
                            If .SetNumbers(X) > 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        DSong = Csong
                        
                        For X = 0 To 5
                            If .KeyOffsets(X) > 0 Then
                                Select Case X
                                    Case 0: DSong = Csong + " [7key] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                    Case 1: DSong = Csong + " [14key] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                    Case 2: DSong = Csong + " [Another7] (10) b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                    Case 3: DSong = Csong + " [Another14] (10) b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                    Case 4: DSong = Csong + " [Light7] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                    Case 5: DSong = Csong + " [Light14] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                End Select
                                DSong = Replace(DSong, " b-1", "")
                                DSong = Replace(DSong, " k-1", "")
                                If bAutoRipCharts Then
                                    ChartExtract.DecodeBemani1 ExtractFileObj, ExtractBuffer(), (.KeyOffsets(X) - &HFF000)
                                    Open OutFolder + DSong + ".cs2" For Binary As #95
                                    Put #95, 1, ExtractBuffer
                                    Close #95
                                End If
                            End If
                        Next X
                        Debug.Print Coffs - 144, Csong

                    End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
            ExtractFileObj.CloseFile
        Case 11 'DistorteD
            Do
                Get #ff, Coffs, IN280
                'Debug.Print Len(IN280)
                Coffs = Coffs + 280
                With IN280
                    Csong = ApplyFilter(.SongName)
                    If InStr(Csong, Chr(0)) > 0 Then
                        Csong = Left(Csong, InStr(Csong, Chr(0)) - 1)
                    End If
                    If Len(Csong) = 0 Then
                        Exit Do
                    End If
                    'If .Difficulty(0) > 0 Then
                        For X = 0 To 15
                            .SetNumbers(X) = .SetNumbers(X) - NumberAdjust
                            If .SetNumbers(X) < 0 Then
                                .SetNumbers(X) = -1
                            End If
                            If .SetNumbers(X) >= 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                        Debug.Print .SetNumbers(X), Csong
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        For X = 0 To 7
                            .KeyOffsets(X) = .KeyOffsets(X) - NumberAdjust
                            If .KeyOffsets(X) >= 0 Then
                                For Y = 0 To UBound(yList)
                                    If Y = UBound(yList) Then
                                        Select Case X
                                            Case 0: yList(UBound(yList)).SongName = Csong + " [7key] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                            Case 1: yList(UBound(yList)).SongName = Csong + " [14key] (" + CStr(.Difficulty(3)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                            Case 2: yList(UBound(yList)).SongName = Csong + " [Another7] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                            Case 3: yList(UBound(yList)).SongName = Csong + " [Another14] (" + CStr(.Difficulty(5)) + ") b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                            Case 4: yList(UBound(yList)).SongName = Csong + " [Light7] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                            Case 5: yList(UBound(yList)).SongName = Csong + " [Light14] (" + CStr(.Difficulty(4)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                            Case 6: yList(UBound(yList)).SongName = Csong + " [Beginner7] (1) b" + CStr(.SetNumbers(14)) + " k" + CStr(.SetNumbers(12))
                                            Case 7: yList(UBound(yList)).SongName = Csong + " [Beginner14] (1) b" + CStr(.SetNumbers(15)) + " k" + CStr(.SetNumbers(13))
                                        End Select
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " b-1", "")
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " k-1", "")
                                        
                                        If .KeyOffsets(X) <= UBound(xList) Then
                                            yList(UBound(yList)).FOffs = xList(.KeyOffsets(X)).FOffs
                                            yList(UBound(yList)).fLink = .KeyOffsets(X)
                                            ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                        End If
                                        Exit For
                                    End If
                                Next Y
                            End If
                        Next X
                    'End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
        Case 12 'GOLD
            Do
                Get #ff, Coffs, IN284
                'Debug.Print Len(IN280)
                'Coffs = Coffs - 64
                Coffs = Coffs + 284
                With IN284
                    Csong = ApplyFilter(.SongName)
                    If InStr(Csong, Chr(0)) > 0 Then
                        Csong = Left(Csong, InStr(Csong, Chr(0)) - 1)
                    End If
                    If Len(Csong) = 0 Then
                        Exit Do
                    End If
                    'If .Difficulty(0) > 0 Then
                        For X = 0 To 15
                            .SetNumbers(X) = .SetNumbers(X) - NumberAdjust
                            If .SetNumbers(X) < 0 Then
                                .SetNumbers(X) = -1
                            End If
                            If .SetNumbers(X) >= 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                        Debug.Print .SetNumbers(X), Csong
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        For X = 0 To 7
                            .KeyOffsets(X) = .KeyOffsets(X) - NumberAdjust
                            If .KeyOffsets(X) >= 0 Then
                                For Y = 0 To UBound(yList)
                                    If Y = UBound(yList) Then
                                        Select Case X
                                            Case 0: yList(UBound(yList)).SongName = Csong + " [Normal7] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                            Case 1: yList(UBound(yList)).SongName = Csong + " [Normal14] (" + CStr(.Difficulty(3)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                            Case 2: yList(UBound(yList)).SongName = Csong + " [Another7] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                            Case 3: yList(UBound(yList)).SongName = Csong + " [Another14] (" + CStr(.Difficulty(5)) + ") b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                            Case 4: yList(UBound(yList)).SongName = Csong + " [Hyper7] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                            Case 5: yList(UBound(yList)).SongName = Csong + " [Hyper14] (" + CStr(.Difficulty(4)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                            Case 6: yList(UBound(yList)).SongName = Csong + " [Beginner7] (1) b" + CStr(.SetNumbers(14)) + " k" + CStr(.SetNumbers(12))
                                            Case 7: yList(UBound(yList)).SongName = Csong + " [Beginner14] (1) b" + CStr(.SetNumbers(15)) + " k" + CStr(.SetNumbers(13))
                                        End Select
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " b-1", "")
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " k-1", "")
                                        
                                        If .KeyOffsets(X) <= UBound(xList) Then
                                            yList(UBound(yList)).FOffs = xList(.KeyOffsets(X)).FOffs
                                            yList(UBound(yList)).fLink = .KeyOffsets(X)
                                            ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                        End If
                                        Exit For
                                    End If
                                Next Y
                            End If
                        Next X
                    'End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
        Case 13 'DJTROOPERS
            Do
                Get #ff, Coffs, IN308
                'Debug.Print Len(IN308)
                'Coffs = Coffs - 64
                Coffs = Coffs + 308
                With IN308
                    Csong = ApplyFilter(.SongName)
                    If InStr(Csong, Chr(0)) > 0 Then
                        Csong = Left(Csong, InStr(Csong, Chr(0)) - 1)
                    End If
                    If Len(Csong) = 0 Then
                        Exit Do
                    End If
                    'If .Difficulty(0) > 0 Then
                        For X = 0 To 15
                            .SetNumbers(X) = .SetNumbers(X) - NumberAdjust
                            If .SetNumbers(X) < 0 Then
                                .SetNumbers(X) = -1
                            End If
                            If .SetNumbers(X) >= 0 Then
                                If .SetNumbers(X) < UBound(xList) Then
                                    If xList(.SetNumbers(X)).SongName = "" Then
                                        xList(.SetNumbers(X)).SongName = Csong
                                        Debug.Print .SetNumbers(X), Csong
                                    End If
                                End If
                            End If
                        Next X
                        
                        Csong = ApplyFilter(Csong)
                        
                        For X = 0 To 7
                            .KeyOffsets(X) = .KeyOffsets(X) - NumberAdjust
                            If .KeyOffsets(X) >= 0 Then
                                For Y = 0 To UBound(yList)
                                    If Y = UBound(yList) Then
                                        Select Case X
                                            Case 0: yList(UBound(yList)).SongName = Csong + " [Normal7] (" + CStr(.Difficulty(0)) + ") b" + CStr(.SetNumbers(2)) + " k" + CStr(.SetNumbers(0))
                                            Case 1: yList(UBound(yList)).SongName = Csong + " [Normal14] (" + CStr(.Difficulty(3)) + ") b" + CStr(.SetNumbers(3)) + " k" + CStr(.SetNumbers(1))
                                            Case 2: yList(UBound(yList)).SongName = Csong + " [Another7] (" + CStr(.Difficulty(2)) + ") b" + CStr(.SetNumbers(6)) + " k" + CStr(.SetNumbers(4))
                                            Case 3: yList(UBound(yList)).SongName = Csong + " [Another14] (" + CStr(.Difficulty(5)) + ") b" + CStr(.SetNumbers(7)) + " k" + CStr(.SetNumbers(5))
                                            Case 4: yList(UBound(yList)).SongName = Csong + " [Hyper7] (" + CStr(.Difficulty(1)) + ") b" + CStr(.SetNumbers(10)) + " k" + CStr(.SetNumbers(8))
                                            Case 5: yList(UBound(yList)).SongName = Csong + " [Hyper14] (" + CStr(.Difficulty(4)) + ") b" + CStr(.SetNumbers(11)) + " k" + CStr(.SetNumbers(9))
                                            Case 6: yList(UBound(yList)).SongName = Csong + " [Beginner7] (1) b" + CStr(.SetNumbers(14)) + " k" + CStr(.SetNumbers(12))
                                            Case 7: yList(UBound(yList)).SongName = Csong + " [Beginner14] (1) b" + CStr(.SetNumbers(15)) + " k" + CStr(.SetNumbers(13))
                                        End Select
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " b-1", "")
                                        yList(UBound(yList)).SongName = Replace(yList(UBound(yList)).SongName, " k-1", "")
                                        
                                        If .KeyOffsets(X) <= UBound(xList) Then
                                            yList(UBound(yList)).FOffs = xList(.KeyOffsets(X)).FOffs
                                            yList(UBound(yList)).fLink = .KeyOffsets(X)
                                            ReDim Preserve yList(0 To UBound(yList) + 1) As OffsetUniversal
                                        End If
                                        Exit For
                                    End If
                                Next Y
                            End If
                        Next X
                    'End If
                End With
            Loop
            bUseList = True
            For X = 0 To UBound(xList)
                With xList(X)
                    If .SongName = "" Then
                        .SongName = "untitled" + CStr(X + 1000)
                    End If
                End With
            Next X
        Case 100 'DDR SUPERNOVA 2 PS2
            Do
                Get #ff, Coffs, DDRSN168
                Coffs = Coffs + 168
            Loop
        Case 101 'DDR X
            Do
                Get #ff, Coffs, DDRX152
                Coffs = Coffs + 152
            Loop
        Case Else
    End Select
    Set OldExtract = Nothing
    Set ExtractFileObj = Nothing
End Sub

Public Sub SetListFile(fName As String)
    If ff > 0 Then
        Close #ff
    End If
    If Dir(fName) = "" Then
        xListFile = ""
        Exit Sub
    End If
    ff = FreeFile
    Open fName For Binary As #ff
    xListFile = fName
End Sub

Private Function ApplyFilter(InString As String) As String
    Dim FilterString As String
    Dim X As Long
    FilterString = "*?/><|" + Chr(34)
    ApplyFilter = InString
    For X = 1 To Len(FilterString)
        ApplyFilter = Replace(ApplyFilter, Mid(FilterString, X, 1), "_")
    Next X
    ApplyFilter = Replace(ApplyFilter, Chr(0), "")
    ApplyFilter = Trim(ApplyFilter)
End Function

Public Sub CloseListFile()
    Close #ff
    ff = 0
End Sub

Public Property Get CSOffset(index As Long) As Double
    If bUseList Then
        CSOffset = yList(index).FOffs
    End If
End Property

Public Property Get CSLink(index As Long) As Long
    If bUseList Then
        CSLink = yList(index).fLink
    End If
End Property

Public Property Get CSName(index As Long) As String
    If bUseList Then
        CSName = yList(index).SongName
    End If
End Property

Public Property Get CSCount() As Long
    CSCount = UBound(yList)
End Property

Public Property Get ListLength(index As Long) As Double
    If bUseList Then
        ListLength = xList(index).FLength
    End If
End Property

Public Property Get ListOffset(index As Long) As Double
    If bUseList Then
        ListOffset = xList(index).FOffs
    End If
End Property

Public Property Get ListName(index As Long) As String
    If bUseList Then
        ListName = xList(index).SongName
    Else
        ListName = CStr(index + 1000)
    End If
End Property

Public Sub SetFileObject(FStream As clsFileStream)
    FileStream = FStream
End Sub

Private Function LoadTitle(offs As Long) As String
    Dim TitleStack() As Byte
    ReDim TitleStack(0 To 95) As Byte
    If offs > 0 Then
        Get #ff, offs + 1, TitleStack
    End If
    LoadTitle = StringFromNullArray(TitleStack())
End Function

Public Function Ripped(index As Long) As Boolean
    If bUseList Then
        Ripped = xList(index).bRipped
    Else
        Ripped = False
    End If
End Function

Public Sub SetRipped(index As Long)
    If bUseList Then
        xList(index).bRipped = True
    End If
End Sub

Function StringFromNullArray(inArr() As Byte) As String
    Dim X As Long
    Dim s As String
    If inArr(LBound(inArr)) = 0 Then
        Exit Function
    End If
    For X = LBound(inArr) To UBound(inArr)
        If inArr(X) = 0 Then
            ReDim Preserve inArr(LBound(inArr) To X - 1) As Byte
            Exit For
        End If
    Next X
    s = StrConv(inArr(), vbUnicode)
    StringFromNullArray = Trim(s)
End Function

Private Sub Class_Initialize()
    ReDim xList(0) As OffsetUniversal
    ReDim yList(0) As OffsetUniversal
End Sub

Private Sub Class_Terminate()
    Close #ff
End Sub
