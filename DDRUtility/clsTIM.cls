VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
'
' Playstation TIM/TIM2 image conversion class
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

Private Type BMPheader
    xSig As Integer
    xFileSize As Long
    xRes0 As Integer
    xRes1 As Integer
    xImgStart As Long
    xBMIHeaderSize As Long 'must be 40
    xWidth As Long
    xHeight As Long
    xPlanes As Integer 'must be 1
    xBits As Integer
    xCompression As Long
    xImgSize As Long
    xImgPPMH As Long
    xImgPPMV As Long
    xImgColors As Long
    xImgImpCol As Long
End Type

Private Type CLUTheader
    xPoX As Integer
    xPoY As Integer
    xColorsPer As Integer
    xCount As Integer
End Type

Private Type TIMheader
    xImageBytes As Long
    xImageOrX As Integer
    xImageOrY As Integer
    xImageWidth As Integer
    xImageHeight As Integer
End Type

Public Function TIMtoBMP(inFile As String, outFile As String) As Boolean
    'temporary solution for now, since I haven't yet written a PNG conversion subroutine.
    Dim f As Long
    Dim f2 As Long
    'Dim ThisColor(0 To 2) As Byte
    Dim TimBPP As Long
    Dim TimBits As Long
    Dim TimIdent As Long
    Dim BMP As BMPheader
    Dim CLUT As CLUTheader
    Dim TIM As TIMheader
    Dim MaxPixel As Long
    Dim ThisPixel As Long
    Dim ThisColor As Long
    Dim ThisBMPPixel As Long
    Dim PixelData() As Byte
    Dim Pal() As Long
    Dim o As Long
    Dim i As Integer
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Boolean
    Dim b1 As Byte
    On Local Error Resume Next
    f = FreeFile
    o = 1
    e = False
    d = 0
    Open inFile For Binary As #f
    f2 = FreeFile
    Do While e = False
        With BMP
            Get #f, o, TimIdent
            If TimIdent <> &H10 Then
                Debug.Print "clsTIM: invalid TIM image"
                TIMtoBMP = True
                Close #f
                Exit Function
            End If
            o = o + 4
            Get #f, o, TimBPP
            o = o + 4
            Select Case TimBPP
                Case 0, 1
                    Debug.Print "clsTIM: can't convert this image because it has no color table."
                    TIMtoBMP = True
                    Close #f
                    Exit Function
                Case 8
                    TimBits = 4
                    Get #f, o, x
                    Get #f, o + 4, CLUT
                    If CLUT.xColorsPer = 0 Then
                        CLUT.xColorsPer = 16
                    End If
                    ReDim Pal(0 To CLUT.xColorsPer - 1) As Long
                    For y = 0 To CLUT.xColorsPer - 1
                        Get #f, , i
                        Pal(y) = Convert16BitColorRGB(i)
                    Next y
                    o = o + x
                Case 9
                    TimBits = 8
                    Get #f, o, x
                    Get #f, o + 4, CLUT
                    If CLUT.xColorsPer = 0 Then
                        CLUT.xColorsPer = 256
                    End If
                    ReDim Pal(0 To CLUT.xColorsPer - 1) As Long
                    For y = 0 To CLUT.xColorsPer - 1
                        Get #f, , i
                        Pal(y) = Convert16BitColorRGB(i)
                    Next y
                    o = o + x
                Case 2 '16 bit
                    TimBits = 16
                    'no color table, pixel color data isn't indexed
                Case 3 '24 bit
                    TimBits = 24
                    'no need for any color conversion
                Case Else
                    Debug.Print "clsTIM: invalid TIM image type:"; TimBPP
                    TIMtoBMP = True
                    Close #f
                    Exit Function
            End Select
            Get #f, o, TIM
            a = o + TIM.xImageBytes '((CLng(TIM.xImageWidth) * 2) * CLng(TIM.xImageHeight)) + 12
            o = o + Len(TIM)
            BMP.xWidth = TIM.xImageWidth / (TimBits / 16)
            BMP.xHeight = TIM.xImageHeight
            MaxPixel = (BMP.xWidth * BMP.xHeight) - 1
            If TIM.xImageWidth <= 0 Or TIM.xImageHeight <= 0 Then
                Debug.Print "clsTIM: bad image dimensions"
                Close #f
                TIMtoBMP = True
                Exit Function
            End If
            ReDim PixelData(0 To ((BMP.xWidth * BMP.xHeight) * 3) - 1) As Byte
            ThisPixel = 0
            ThisBMPPixel = UBound(PixelData) + 1 + (BMP.xWidth * 3)
            Do While ThisPixel <= MaxPixel
                If (ThisPixel Mod BMP.xWidth) = 0 Then
                    ThisBMPPixel = ThisBMPPixel - ((BMP.xWidth * 3) * 2)
                End If
                Select Case TimBPP
                    Case 8
                        Get #f, o, b1
                        ThisColor = Pal(b1 And 15)
                        PixelData(ThisBMPPixel) = (ThisColor And &HFF0000) \ &H10000
                        PixelData(ThisBMPPixel + 1) = (ThisColor And &HFF00&) \ &H100&
                        PixelData(ThisBMPPixel + 2) = (ThisColor And &HFF&)
                        ThisColor = Pal((b1 And 240) \ 16)
                        PixelData(ThisBMPPixel + 3) = (ThisColor And &HFF0000) \ &H10000
                        PixelData(ThisBMPPixel + 4) = (ThisColor And &HFF00&) \ &H100&
                        PixelData(ThisBMPPixel + 5) = (ThisColor And &HFF&)
                        ThisBMPPixel = ThisBMPPixel + 6
                        ThisPixel = ThisPixel + 2
                        o = o + 1
                    Case 9
                        Get #f, o, b1
                        ThisColor = Pal(b1)
                        PixelData(ThisBMPPixel) = (ThisColor And &HFF0000) \ &H10000
                        PixelData(ThisBMPPixel + 1) = (ThisColor And &HFF00&) \ &H100&
                        PixelData(ThisBMPPixel + 2) = (ThisColor And &HFF&)
                        ThisBMPPixel = ThisBMPPixel + 3
                        ThisPixel = ThisPixel + 1
                        o = o + 1
                    Case 2
                        Get #f, o, i
                        ThisColor = Convert16BitColorRGB(i)
                        PixelData(ThisBMPPixel) = (ThisColor And &HFF0000) \ &H10000
                        PixelData(ThisBMPPixel + 1) = (ThisColor And &HFF00&) \ &H100&
                        PixelData(ThisBMPPixel + 2) = (ThisColor And &HFF&)
                        ThisBMPPixel = ThisBMPPixel + 3
                        ThisPixel = ThisPixel + 1
                        o = o + 2
                    Case 3
                        Get #f, o, PixelData(ThisBMPPixel)
                        Get #f, o + 1, PixelData(ThisBMPPixel + 1)
                        Get #f, o + 2, PixelData(ThisBMPPixel + 2)
                        ThisBMPPixel = ThisBMPPixel + 3
                        ThisPixel = ThisPixel + 1
                        o = o + 3
                End Select
            Loop
            .xBits = 24
            .xBMIHeaderSize = 40
            .xCompression = 0
            .xFileSize = 54 + UBound(PixelData) + 1
            .xImgColors = 0
            .xImgImpCol = 0
            .xImgPPMH = 0
            .xImgPPMV = 0
            .xImgSize = UBound(PixelData) + 1
            .xImgStart = 54
            .xPlanes = 1
            .xRes0 = 0
            .xRes1 = 0
            .xSig = &H4D42
            o = a
            If e = False Then
                If o > LOF(f) Then
                    e = True
                Else
                    Get #f, o, c
                    If c <> 16 Then
                        e = True
                    Else
                        x = x
                    End If
                End If
            End If
            If (d > 0) Or (e = False) Then
                If InStr(outFile, Chr$(1)) Then
                    Open Replace$(outFile, Chr$(1), CStr(d)) For Binary As #f2
                    d = d + 1
                Else
                    Open outFile For Binary As #f2
                    e = True
                End If
            Else
                If InStr(outFile, Chr$(1)) Then
                    Open Replace$(outFile, Chr$(1), "") For Binary As #f2
                    d = d + 1
                Else
                    Open outFile For Binary As #f2
                End If
                e = True
            End If
            Put #f2, 1, BMP
            Put #f2, 55, PixelData
            Close #f2
        End With
    Loop
    Close #f
End Function

Public Function Convert16BitColorRGB(i2 As Integer) As Long
    Dim i As Long
    i = (CLng(i2) + 65536) Mod 65536
    If (i And &H7FFF&) <> 0 Then
        Convert16BitColorRGB = RGB(((i And 31) * 8) + 7, (((i \ 32) And 31) * 8) + 7, (((i \ 1024) And 31) * 8) + 7)
    Else
        Convert16BitColorRGB = 0
    End If
End Function

Public Function TIMtoPNG() As Boolean
    'todo
End Function

Public Function TIM2toPNG() As Boolean
    'todo
End Function

Public Function TGCDtoPNGs() As Boolean
    'todo
End Function


