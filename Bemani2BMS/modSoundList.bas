Attribute VB_Name = "modSoundList"
' L's sound list format implementation
' -- saxxonpike 2006
'
' +SongDB
' -- saxxonpike 2007

Public Type xSoundList
    SongID As String
    Title As String
    Genre As String
    MainTitle As String
    SubTitle As String
    Artist As String
    BPM As String
    Difficulty(0 To 6) As String
    VideoFile As String
    VideoInfoFS As String
    VideoInfoCol As String
    VideoInfoDly As String
    VideoInfoExtra As String
    SongVersion As String
End Type

Public Type xSongDB
    Title As String
    Artist As String
    Genre As String
End Type

Public SoundList() As xSoundList
Public SongDB() As xSongDB

Public Sub LoadSongDB(fname As String)
    If Dir(fname) = "" Then
        Exit Sub
    End If
    Dim SLcount As Long
    Dim SLbuffer() As Byte
    Dim f As Long
    Dim currentoffset As Long
    Dim currentitem As Long
    Dim cb As Byte
    ReDim SongDB(0 To 1) As xSongDB
    SLcount = 0
    f = FreeFile
    Open fname For Binary As #f
    ReDim SLbuffer(1 To LOF(f)) As Byte
    Get #f, 1, SLbuffer
    Close #f
    If SLbuffer(UBound(SLbuffer)) <> &HA Then
        ReDim Preserve SLbuffer(1 To UBound(SLbuffer) + 2) As Byte
        SLbuffer(UBound(SLbuffer)) = &HA
        SLbuffer(UBound(SLbuffer) - 1) = &HD
    End If
    currentitem = 0
    currentoffset = 1
    SLcount = 1
    Do While currentoffset < UBound(SLbuffer)
        cb = SLbuffer(currentoffset)
        currentoffset = currentoffset + 1
        If cb <> 9 And cb <> 13 Then
            With SongDB(SLcount)
                Select Case currentitem
                    Case 0
                        .Title = .Title + Chr(cb)
                    Case 1
                        .Artist = .Artist + Chr(cb)
                    Case 2
                        .Genre = .Genre + Chr(cb)
                End Select
            End With
        ElseIf cb = 13 Then
            With SongDB(SLcount)
                .Title = Trim$(.Title)
                .Artist = Trim$(.Artist)
                .Genre = Trim$(.Genre)
            End With
            SLcount = SLcount + 1
            ReDim Preserve SongDB(0 To SLcount) As xSongDB
            currentitem = 0
            currentoffset = currentoffset + 1
        Else
            currentitem = currentitem + 1
        End If
    Loop
End Sub

Public Sub LoadSoundList(fname As String)
    Dim x(0 To 1) As Byte
    If Dir(fname) = "" Then
        Exit Sub
    End If
    f = FreeFile
    Open fname For Binary As #f
    Get #f, 1, x
    Close #f
    If x(0) = &H49 And x(1) = &H44 Then
        LoadSoundListNew fname
    Else
        LoadSoundListOld fname
    End If
End Sub

Public Sub LoadSoundListOld(fname As String)
    Dim SLcount As Long
    Dim SLbuffer() As Byte
    Dim f As Long
    Dim currentoffset As Long
    Dim currentitem As Long
    Dim cb As Byte
    ReDim SoundList(0 To 1) As xSoundList
    SLcount = 0
    f = FreeFile
    Open fname For Binary As #f
    ReDim SLbuffer(1 To LOF(f)) As Byte
    Get #f, 1, SLbuffer
    Close #f
    If SLbuffer(UBound(SLbuffer)) <> &HA Then
        ReDim Preserve SLbuffer(1 To UBound(SLbuffer) + 2) As Byte
        SLbuffer(UBound(SLbuffer)) = &HA
        SLbuffer(UBound(SLbuffer) - 1) = &HD
    End If
    currentitem = 0
    currentoffset = 1
    SLcount = 1
    Do While currentoffset < UBound(SLbuffer)
        cb = SLbuffer(currentoffset)
        currentoffset = currentoffset + 1
        If cb <> 9 And cb <> 13 Then
            With SoundList(SLcount)
                Select Case currentitem
                    Case 0
                        .SongID = .SongID + Chr(cb)
                    Case 1
                        .Title = .Title + Chr(cb)
                    Case 2
                        .Genre = .Genre + Chr(cb)
                    Case 3
                        .MainTitle = .MainTitle + Chr(cb)
                    Case 4
                        .SubTitle = .SubTitle + Chr(cb)
                    Case 5
                        .Artist = .Artist + Chr(cb)
                    Case 6, 7, 8, 9, 10, 11, 12
                        .Difficulty(currentitem - 6) = .Difficulty(currentitem - 6) + Chr(cb)
                    Case 13
                        .VideoFile = .VideoFile + Chr(cb)
                    Case 14
                        .VideoInfoFS = .VideoInfoFS + Chr(cb)
                    Case 15
                        .VideoInfoCol = .VideoInfoCol + Chr(cb)
                    Case 16
                        .VideoInfoDly = .VideoInfoDly + Chr(cb)
                    Case 17
                        .VideoInfoExtra = .VideoInfoExtra + Chr(cb)
                    Case 18
                        .SongVersion = .SongVersion + Chr(cb)
                End Select
            End With
        ElseIf cb = 13 Then
            'Debug.Print Trim(SoundList(SLcount).MainTitle + " " + SoundList(SLcount).SubTitle); ",", ;
            'Debug.Print SoundList(SLcount).Artist; ",", ;
            'Debug.Print SoundList(SLcount).Genre
            SLcount = SLcount + 1
            ReDim Preserve SoundList(0 To SLcount) As xSoundList
            currentitem = 0
            currentoffset = currentoffset + 1
        Else 'cb = 9
            If currentitem = 0 And Len(SoundList(SLcount).SongID) = 3 Then
                SoundList(SLcount).SongID = "0" + SoundList(SLcount).SongID
            End If
            currentitem = currentitem + 1
        End If
    Loop
End Sub

Public Sub LoadSoundListNew(fname As String)
    Dim SLcount As Long
    Dim SLbuffer() As Byte
    Dim f As Long
    Dim currentoffset As Long
    Dim currentitem As Long
    Dim cb As Byte
    ReDim SoundList(0 To 1) As xSoundList
    SLcount = 0
    f = FreeFile
    Open fname For Binary As #f
    ReDim SLbuffer(1 To LOF(f)) As Byte
    Get #f, 1, SLbuffer
    Close #f
    If SLbuffer(UBound(SLbuffer)) <> 44 Then
        ReDim Preserve SLbuffer(1 To UBound(SLbuffer) + 2) As Byte
        SLbuffer(UBound(SLbuffer)) = 44
        SLbuffer(UBound(SLbuffer) - 1) = &HD
    End If
    currentitem = 0
    currentoffset = 1
    SLcount = 1
    Do While currentoffset < UBound(SLbuffer)
        cb = SLbuffer(currentoffset)
        currentoffset = currentoffset + 1
        If (cb <> 44 And cb <> 13) And (currentitem < 5 Or cb <> 95) Then
            If cb = Asc("|") Then
                cb = Asc(",")
            End If
            With SoundList(SLcount)
                Select Case currentitem
                    Case 0
                        .SongID = .SongID + Chr(cb)
                    Case 1
                        .Title = .Title + Chr(cb)
                        .MainTitle = .Title
                    Case 2
                        .Artist = .Artist + Chr(cb)
                    Case 3
                        .Genre = .Genre + Chr(cb)
                    Case 4
                        .BPM = .BPM + Chr(cb)
                    Case 5
                        If .BPM = "" Then
                            .Difficulty(6) = .Difficulty(6) + Chr(cb)
                        End If
                    Case 6, 7, 8, 9, 10, 11, 12
                        .Difficulty(currentitem - 6) = .Difficulty(currentitem - 6) + Chr(cb)
                End Select
            End With
        ElseIf cb = 13 Then
            'Debug.Print Trim(SoundList(SLcount).MainTitle + " " + SoundList(SLcount).SubTitle); ",", ;
            'Debug.Print SoundList(SLcount).Artist; ",", ;
            'Debug.Print SoundList(SLcount).Genre
            Debug.Print SoundList(SLcount).SongID
            If Val(SoundList(SLcount).SongID) > 0 Then
                SLcount = SLcount + 1
                ReDim Preserve SoundList(0 To SLcount) As xSoundList
            Else
                SoundList(SLcount) = SoundList(0)
            End If
            currentitem = 0
            currentoffset = currentoffset + 1
        Else 'cb = 44
            If currentitem = 0 And Len(SoundList(SLcount).SongID) = 3 Then
                SoundList(SLcount).SongID = "0" + SoundList(SLcount).SongID
                SoundList(SLcount).VideoFile = SoundList(SLcount).SongID
            End If
            currentitem = currentitem + 1
        End If
    Loop
End Sub

