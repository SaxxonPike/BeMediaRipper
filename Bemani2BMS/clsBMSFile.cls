VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBMSFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================
' clsBMSFile
' SaxxonPike 2oo7
'
' for loading, processing and converting BMS/BME simfiles.
'
' last updated January 14, 2009
'================================================================

Private Type InternalNoteBMS
    xOffsetMetric As Double
    xOffsetMS As Double
    xValue As Long              'note value
    xLane As Long               '0-7 keys, 8 scratch, +16 for 2P side
End Type

Public Function LoadBMS(fname As String) As Boolean
    
    '----------------------------------------------------------
    
    Dim f As Long
    
    Dim s As String                 'main string
    Dim v As String                 'data string
    Dim t As String                 'info string
    
    Dim nr As Long                  'value
    Dim nl As Long                  'lane
    Dim nm As Long                  'measure
    Dim nd As String
    Dim nf As Long
    Dim nt As Long
    
    Dim keysounds(1 To 1295) As String
    Dim measuresize(1 To 999) As Double
    Dim maxmeasure As Long
    
    '----------------------------------------------------------
    
    f = FreeFile
    Open fname For Input As #f
    Do While Not EOF(f)
        Line Input #f, s
        If Left$(s, 1) = "#" Then
            If Mid$(s, 7, 1) = ":" Then
                'metric command (notes)
                s = Mid$(s, 2)
                nm = Val(Mid$(s, 1, 3))
                nf = Val(Mid$(s, 4, 2))
                nd = Mid$(s, 7)
                Select Case nf
                    Case 1                      'bgm note
                        nt = 7
                        nf = 8
                    Case 2                      'measure size
                        nt = -1
                        nf = 0
                    Case 3                      'old style bpm
                        nt = 4
                        nf = 100
                    Case 8                      'new style bpm
                        nt = 4
                        nf = 100
                    Case 11, 12, 13, 14, 15     '1p
                        nt = 0
                        nf = nf - 11
                    Case 18, 19                 '1p
                        nt = 0
                        nf = nf - 13
                    Case 16                     '1p
                        nt = 0
                        nf = 7
                    Case 21, 22, 23, 24, 25     '2p
                        nt = 1
                        nf = nf - 11
                    Case 28, 29                 '2p
                        nt = 1
                        nf = nf - 13
                    Case 26                     '2p
                        nt = 1
                        nf = 7
                    Case Else
                        nt = -1
                        nf = 0
                End Select
            Else
                'data command (wav, bpm, etc)
                s = Mid$(s, 2)
                If InStr(s, " ") > 0 Then
                    t = Left$(s, InStr(s, " ") - 1)
                    v = Mid$(s, InStr(s, " ") + 1)
                Else
                    t = s
                    v = ""
                End If
                
                
            End If
        End If
    Loop

    '----------------------------------------------------------

End Function

Public Function RenderAC1(fname As String, destArr() As Byte) As Boolean
    'renders the stored BMS to a .1 file usable by arcade
    Dim fTiming As Double
    fTiming = (1 / 59.94) * 1000
End Function

