VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeVAG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
'
' Playstation VAG audio decompression class
'  Copyright © SaxxonPike 2oo4-2oo8
'
'  Written for use exclusive with DDR Utility (as it's lacking a few
'   other things like stereo support which isn't needed for pBAV anyway)
'
'  Requires BeMedia2.dll
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

Private Declare Sub DecodeVAG Lib "BeMedia2.dll" (in16 As Any, out28 As Any, ByRef p1 As Long, ByRef p2 As Long)

Private Type SoundPackTable1
    xNumber As Long
    xOffset As Long
    xLength As Long
End Type
Private Type SoundPackTable2
    xPanning As Byte
    xVolume As Byte
    xInfo As Byte
    xINumber As Byte
    xSNumber As Byte
    xFlags1 As Byte
    xFlags2 As Byte
    xFlags3 As Byte
    xParam1 As Long
End Type
Private Type SoundPackHead2
    xUnk0 As Long
    xUnk1 As Long
    xUnk2 As Long
End Type

Public Function VAGtoWAV(inFile As String, outFile As String, ByVal lOffset As Long, ByVal lLength As Long, ByVal lFrequency As Long) As Boolean
    Dim p1 As Long
    Dim p2 As Long
    Dim f As Long
    Dim s() As Byte
    Dim sbuf() As Byte
    Dim x As Long
    Dim lOutOffset As Long
    If lLength = 0 Then
        VAGtoWAV = True
        Exit Function
    End If
    lOffset = lOffset + 1
    f = FreeFile
    Open inFile For Binary As #f
    'use -1 to auto-set the length of the data
    If lLength < 0 Then
        lLength = LOF(f) - (lOffset - 1)
    End If
    ReDim sbuf(0 To lLength - 1) As Byte
    ReDim s(0 To ((lLength / 16) * 56) - 1) As Byte
    Get #f, lOffset, sbuf
    Close #f
    lOffset = 0
    lOutOffset = 0
    p1 = 0
    p2 = 0
    For x = 0 To lLength - 1 Step 16
        If (sbuf(lOffset + 1) And 4) Then
            'filter reset bit
            p1 = 0
            p2 = 0
        End If
        If (sbuf(lOffset + 1) And 1) = 0 Then
            'only process audio if the end stream bit is not set
            DecodeVAG sbuf(lOffset), s(lOutOffset), p1, p2
        End If
        lOffset = lOffset + 16
        lOutOffset = lOutOffset + 56
    Next x
    f = FreeFile
    'write header and save
    Open outFile For Binary As #f
    lOutOffset = WriteWaveHeader(f, UBound(s) + 1, 16, lFrequency)
    Put #f, lOutOffset + 1, s
    Close #f
End Function

Private Function WriteWaveHeader(fNumber As Long, ByVal lDataSize As Long, ByVal lBits As Long, ByVal lFrequency As Long) As Long
    Dim s As String * 4
    Dim l As Long
    Dim Ident As Long
    s = "RIFF": Put #fNumber, 1, s              'strings
    s = "WAVE": Put #fNumber, 9, s
    s = "fmt ": Put #fNumber, 13, s
    s = "data": Put #fNumber, 37, s
    Put #fNumber, 41, lDataSize                 'total data bytes
    lDataSize = lDataSize + 32
    Put #fNumber, 5, lDataSize                  'total data bytes + header
    l = 16
    Put #fNumber, 17, l                         'info header size
    Ident = &H10001: Put #fNumber, 21, Ident    'format identifier
    Ident = &H2 Or (&H10000 * lBits)
    Put #fNumber, 33, Ident                     'playback information
    Put #fNumber, 25, lFrequency                'playback frequency
    lFrequency = lFrequency * (lBits \ 8)
    Put #fNumber, 29, lFrequency                'playback byterate/sec
    WriteWaveHeader = 44
End Function

Public Function ExtractSoundPack(inFile As String, outFile As String, ByVal lOffset As Long, ByVal lLength As Long) As Boolean
    Dim SPT1() As SoundPackTable1
    Dim SPT2() As SoundPackTable2
    'Dim SPH1 As SoundPackHead1
    Dim SPH2 As SoundPackHead2
    Dim x As Long
    Dim c As Long
    Dim fr As Long
    If lLength = 0 Then
        ExtractSoundPack = True
        Exit Function
    End If
    lOffset = lOffset + 1
    f = FreeFile
    Open inFile For Binary As #f
    'use -1 to auto-set the length of the data
    If lLength < 0 Then
        lLength = LOF(f) - (lOffset - 1)
    End If
    Get #f, lOffset, c
    If c <= 0 Then
        Close #f
        ExtractSoundPack = True
        Exit Function
    End If
    ReDim SPT1(1 To c) As SoundPackTable1
    ReDim SPT2(1 To c) As SoundPackTable2
    Get #f, lOffset + 4, SPT1
    lOffset = lOffset + (c * 12) + 4
    Get #f, lOffset, SPT2
    lOffset = lOffset + (c * 12) + 16
    For x = 1 To c
        Select Case SPT2(x).xFlags1
            Case 57
                fr = 16000
            Case Else
                fr = 22050
        End Select
        If VAGtoWAV(inFile, outFile + "." + CStr(x + 100) + ".wav", (lOffset - 1) + SPT1(x).xOffset, SPT1(x).xLength, fr) Then
            ExtractSoundPack = True
        End If
    Next x
    Close #f
End Function
