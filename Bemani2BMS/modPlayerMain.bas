Attribute VB_Name = "modPlayerMain"
Public Type FormatInternalX
    OffsetMS As Double
    OffsetMetric As Double
    value As Double
    OffsetInternal As Double
    Lane As Double
    xType As Double
    OffsetBMSMetric As Double
    temp0 As Long
    temp1 As Integer
    xEnabled As Integer
End Type

Public CS As New clsCSFile
Public NextEvent As FormatInternalX
Public bPlaying As Boolean
Public bRunning As Boolean
Public CurrentEvent As Long
Public CurrentFrame As Long
Public CurrentMS As Currency
Public PlayTimer As New clsHPTimer

Public DefaultNotes(0 To 64) As Double
Public SampleHandles(0 To 1295) As Long

Sub Main()
    Load Form1
    Form1.Show
    bRunning = True
    Do While bRunning
        If bPlaying Then
            CurrentMS = PlayTimer.TimeElapsed - 2000
            CurrentFrame = CurrentMS / CS.FrameTiming
            Do While CurrentFrame >= NextEvent.OffsetInternal 'NextEvent.OffsetMS
                Debug.Print CurrentEvent, NextEvent.OffsetMS, NextEvent.xType, NextEvent.Lane, NextEvent.value
                Select Case NextEvent.xType
                    Case 1 'measure
                    Case 2 'bpm
                    Case 3 'bpm03
                    Case 4 'note
                        If SampleHandles(NextEvent.value) <> 0 Then
                            If NextEvent.Lane < 64 Then
                                If NextEvent.value <> 0 Then
                                    DefaultNotes(NextEvent.Lane) = NextEvent.value
                                End If
                                BASS_ChannelPlay BASS_SampleGetChannel(SampleHandles(DefaultNotes(NextEvent.Lane)), BASSFALSE), BASSTRUE
                                Form1.PlayChan NextEvent.Lane + 0
                            Else 'for bgm
                                BASS_ChannelPlay BASS_SampleGetChannel(SampleHandles(NextEvent.value), BASSFALSE), BASSTRUE
                            End If
                        End If
                    Case 5 'notechange
                        DefaultNotes(NextEvent.Lane) = NextEvent.value
                    Case 6 'endsong
                         bPlaying = False
                         Exit Do
                    Case 7 'metronome
                    Case 255
                End Select
                CurrentEvent = CurrentEvent + 1
                If CS.GetSimFileData(VarPtr(NextEvent), CurrentEvent) = False Then
                    bPlaying = False
                    Exit Do
                End If
            Loop
        End If
        Sleep 0
        Sleep 1
        DoEvents
    Loop
    Unload Form1
    BASS_Free
    End
End Sub
