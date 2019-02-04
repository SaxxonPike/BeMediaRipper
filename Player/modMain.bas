Attribute VB_Name = "modMain"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sound As New clsSound
Private Display As New frmFlex
Private FrameTimer As New clsHPTimer
Private SoundTimer As New clsHPTimer
Private SoundDelay As Currency

Public bRunning As Boolean
Public bQuit As Boolean
Public bPaint As Boolean

Sub Main()
    Dim PaintCount As Long
    
    bRunning = True
    
    
    'sound up
    Sound.Init
    'Sound.BufferSize = 441
    SoundDelay = Sound.BufferDelay
    Sound.Randomize
    'Debug.Print SoundDelay
    
    'video up
    Display.SetTitle "BeMedia Player"
    Display.SetSize 250, 520
    Display.Show
    
    'timers up
    FrameTimer.Init
    SoundTimer.Init
    
    'main
    Do While bRunning
        If bQuit Then
            Exit Do
        End If
        If bPaint Then
            'draw form
            bPaint = False
        End If
        If SoundTimer.TimeElapsed > SoundDelay Then
            SoundTimer.Tick 'Inc SoundDelay
            Sound.Randomize
            Sound.Play
        End If
        DoEvents
        Sleep 0
    Loop
    
    'close
    End
    
End Sub
