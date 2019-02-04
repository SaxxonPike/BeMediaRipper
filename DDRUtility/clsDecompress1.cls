VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDecompress1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
'
' Bemani LZ decompression class
'  Copyright © SaxxonPike 2oo5-2oo8
'
' requires clsFileStream
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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Function DecompressBemani1(ByRef inFS As clsFileStream, ByRef OutData() As Byte, ByVal Offs As Double, ByRef retBytesUsed As Long) As Long
    
    
    On Error GoTo 616
    
    'internal
    Dim OldOffset As Double
    Dim BytesDecompressd As Long
    'Dim Decompressd() As Byte
    Dim EndFlag As Boolean
    Dim OldAdvanceSetting As Boolean
    Dim Decompressd() As Byte
    Dim DecompressOffs As Double
    'ReDim Decompressd(0) As Byte
    
    'Decompressr
    Dim i As Long
    Dim flags As Long
    Dim j As Long
    Dim c As Byte
    Dim t As Byte
    Dim src As Long
    Dim dst As Long
    Dim xloop As Boolean
    
    'code
    OldAdvanceSetting = inFS.AdvanceOffset
    OldOffset = inFS.Offset
    DecompressOffs = OldOffset
    inFS.AdvanceOffset = False
    If Offs <> -1 Then
        inFS.Offset = Offs
        DecompressOffs = Offs
    End If
    
    Do
        Do
            flags = ShiftRight(flags, 1)
            If (flags And &H100&) = 0 Then
                flags = inFS.ReadFileByteFast(DecompressOffs) Or &HFF00&
                DecompressOffs = DecompressOffs + 1
            End If
            c = inFS.ReadFileByteFast(DecompressOffs)
            If (flags And 1&) = 0 Then
                AddByte Decompressd(), c
                DecompressOffs = DecompressOffs + 1
                Exit Do
            End If
            If (c And &H80&) = 0 Then
                i = inFS.ReadFileByteFast(DecompressOffs + 1)
                DecompressOffs = DecompressOffs + 2
                i = i Or ShiftLeft(c And 3, 8)
                j = ShiftRight(c + 0, 2) + 2
                xloop = True
            End If
            If Not xloop Then
                DecompressOffs = DecompressOffs + 1
                If (c And &H40&) = 0 Then
                    i = (c And 15) + 1
                    j = (ShiftRight(c + 0, 4)) + 1 - 8
                    xloop = True
                End If
            End If
            If xloop Then
                xloop = False
                Do
                    If (UBound(Decompressd) + 1) - i > 0 And (UBound(Decompressd) + 1) - i < UBound(Decompressd) Then
                        AddByte Decompressd(), Decompressd((UBound(Decompressd) + 1) - i)
                    ElseIf (UBound(Decompressd) + 1) - i = UBound(Decompressd) Then
                        AddByte Decompressd(), Decompressd(UBound(Decompressd))
                    Else
                        AddByte Decompressd(), 0
                    End If
                    j = j - 1
                Loop While j >= 0
                Exit Do
            End If
            If c = 255 Then
                EndFlag = True
                Exit Do
            End If
            j = c - &HC0& + 7
            Do
                AddByte Decompressd(), inFS.ReadFileByteFast(DecompressOffs)
                DecompressOffs = DecompressOffs + 1
                j = j - 1
            Loop While j >= 0
        Loop
    Loop While EndFlag = False
    BytesDecompressd = UBound(Decompressd) + 1
    
    ReDim OutData(LBound(Decompressd) To UBound(Decompressd))
    CopyMemory OutData(LBound(Decompressd)), Decompressd(LBound(Decompressd)), BytesDecompressd
    
617 inFS.Offset = OldOffset
    inFS.AdvanceOffset = OldAdvanceSetting
    DecompressBemani1 = BytesDecompressd
    retBytesUsed = (DecompressOffs - OldOffset)
    Exit Function
    
616 Debug.Print "clsPSXDecompress.DecompressBemani1: ERR DecOffs:" + CStr(DecompressOffs)
    retBytesUsed = (DecompressOffs - OldOffset)
    BytesDecompressd = 0
    inFS.Offset = OldOffset
    inFS.AdvanceOffset = OldAdvanceSetting
    On Error GoTo 0
End Function

Private Function ShiftRight(ByVal value As Long, ByVal ShiftCount As Long) As Long
' from VBSPEED
' by Jost Schwider, jost@schwider.de, 20011001
  If value And &H80000000 Then
    Select Case ShiftCount
    Case 0&:  ShiftRight = value
    Case 1&:  ShiftRight = &H40000000 Or (value And &H7FFFFFFF) \ &H2&
    Case 2&:  ShiftRight = &H20000000 Or (value And &H7FFFFFFF) \ &H4&
    Case 3&:  ShiftRight = &H10000000 Or (value And &H7FFFFFFF) \ &H8&
    Case 4&:  ShiftRight = &H8000000 Or (value And &H7FFFFFFF) \ &H10&
    Case 5&:  ShiftRight = &H4000000 Or (value And &H7FFFFFFF) \ &H20&
    Case 6&:  ShiftRight = &H2000000 Or (value And &H7FFFFFFF) \ &H40&
    Case 7&:  ShiftRight = &H1000000 Or (value And &H7FFFFFFF) \ &H80&
    Case 8&:  ShiftRight = &H800000 Or (value And &H7FFFFFFF) \ &H100&
    Case 9&:  ShiftRight = &H400000 Or (value And &H7FFFFFFF) \ &H200&
    Case 10&: ShiftRight = &H200000 Or (value And &H7FFFFFFF) \ &H400&
    Case 11&: ShiftRight = &H100000 Or (value And &H7FFFFFFF) \ &H800&
    Case 12&: ShiftRight = &H80000 Or (value And &H7FFFFFFF) \ &H1000&
    Case 13&: ShiftRight = &H40000 Or (value And &H7FFFFFFF) \ &H2000&
    Case 14&: ShiftRight = &H20000 Or (value And &H7FFFFFFF) \ &H4000&
    Case 15&: ShiftRight = &H10000 Or (value And &H7FFFFFFF) \ &H8000&
    Case 16&: ShiftRight = &H8000& Or (value And &H7FFFFFFF) \ &H10000
    Case 17&: ShiftRight = &H4000& Or (value And &H7FFFFFFF) \ &H20000
    Case 18&: ShiftRight = &H2000& Or (value And &H7FFFFFFF) \ &H40000
    Case 19&: ShiftRight = &H1000& Or (value And &H7FFFFFFF) \ &H80000
    Case 20&: ShiftRight = &H800& Or (value And &H7FFFFFFF) \ &H100000
    Case 21&: ShiftRight = &H400& Or (value And &H7FFFFFFF) \ &H200000
    Case 22&: ShiftRight = &H200& Or (value And &H7FFFFFFF) \ &H400000
    Case 23&: ShiftRight = &H100& Or (value And &H7FFFFFFF) \ &H800000
    Case 24&: ShiftRight = &H80& Or (value And &H7FFFFFFF) \ &H1000000
    Case 25&: ShiftRight = &H40& Or (value And &H7FFFFFFF) \ &H2000000
    Case 26&: ShiftRight = &H20& Or (value And &H7FFFFFFF) \ &H4000000
    Case 27&: ShiftRight = &H10& Or (value And &H7FFFFFFF) \ &H8000000
    Case 28&: ShiftRight = &H8& Or (value And &H7FFFFFFF) \ &H10000000
    Case 29&: ShiftRight = &H4& Or (value And &H7FFFFFFF) \ &H20000000
    Case 30&: ShiftRight = &H2& Or (value And &H7FFFFFFF) \ &H40000000
    Case 31&: ShiftRight = &H1&
    End Select
  Else
    Select Case ShiftCount
    Case 0&:  ShiftRight = value
    Case 1&:  ShiftRight = value \ &H2&
    Case 2&:  ShiftRight = value \ &H4&
    Case 3&:  ShiftRight = value \ &H8&
    Case 4&:  ShiftRight = value \ &H10&
    Case 5&:  ShiftRight = value \ &H20&
    Case 6&:  ShiftRight = value \ &H40&
    Case 7&:  ShiftRight = value \ &H80&
    Case 8&:  ShiftRight = value \ &H100&
    Case 9&:  ShiftRight = value \ &H200&
    Case 10&: ShiftRight = value \ &H400&
    Case 11&: ShiftRight = value \ &H800&
    Case 12&: ShiftRight = value \ &H1000&
    Case 13&: ShiftRight = value \ &H2000&
    Case 14&: ShiftRight = value \ &H4000&
    Case 15&: ShiftRight = value \ &H8000&
    Case 16&: ShiftRight = value \ &H10000
    Case 17&: ShiftRight = value \ &H20000
    Case 18&: ShiftRight = value \ &H40000
    Case 19&: ShiftRight = value \ &H80000
    Case 20&: ShiftRight = value \ &H100000
    Case 21&: ShiftRight = value \ &H200000
    Case 22&: ShiftRight = value \ &H400000
    Case 23&: ShiftRight = value \ &H800000
    Case 24&: ShiftRight = value \ &H1000000
    Case 25&: ShiftRight = value \ &H2000000
    Case 26&: ShiftRight = value \ &H4000000
    Case 27&: ShiftRight = value \ &H8000000
    Case 28&: ShiftRight = value \ &H10000000
    Case 29&: ShiftRight = value \ &H20000000
    Case 30&: ShiftRight = value \ &H40000000
    Case 31&: ShiftRight = &H0&
    End Select
  End If
End Function

Private Function ShiftLeft(ByVal value As Long, ByVal ShiftCount As Long) As Long
' from VBSPEED
' by Jost Schwider, jost@schwider.de, 20011001
  Select Case ShiftCount
  Case 0&
    ShiftLeft = value
  Case 1&
    If value And &H40000000 Then
      ShiftLeft = (value And &H3FFFFFFF) * &H2& Or &H80000000
    Else
      ShiftLeft = (value And &H3FFFFFFF) * &H2&
    End If
  Case 2&
    If value And &H20000000 Then
      ShiftLeft = (value And &H1FFFFFFF) * &H4& Or &H80000000
    Else
      ShiftLeft = (value And &H1FFFFFFF) * &H4&
    End If
  Case 3&
    If value And &H10000000 Then
      ShiftLeft = (value And &HFFFFFFF) * &H8& Or &H80000000
    Else
      ShiftLeft = (value And &HFFFFFFF) * &H8&
    End If
  Case 4&
    If value And &H8000000 Then
      ShiftLeft = (value And &H7FFFFFF) * &H10& Or &H80000000
    Else
      ShiftLeft = (value And &H7FFFFFF) * &H10&
    End If
  Case 5&
    If value And &H4000000 Then
      ShiftLeft = (value And &H3FFFFFF) * &H20& Or &H80000000
    Else
      ShiftLeft = (value And &H3FFFFFF) * &H20&
    End If
  Case 6&
    If value And &H2000000 Then
      ShiftLeft = (value And &H1FFFFFF) * &H40& Or &H80000000
    Else
      ShiftLeft = (value And &H1FFFFFF) * &H40&
    End If
  Case 7&
    If value And &H1000000 Then
      ShiftLeft = (value And &HFFFFFF) * &H80& Or &H80000000
    Else
      ShiftLeft = (value And &HFFFFFF) * &H80&
    End If
  Case 8&
    If value And &H800000 Then
      ShiftLeft = (value And &H7FFFFF) * &H100& Or &H80000000
    Else
      ShiftLeft = (value And &H7FFFFF) * &H100&
    End If
  Case 9&
    If value And &H400000 Then
      ShiftLeft = (value And &H3FFFFF) * &H200& Or &H80000000
    Else
      ShiftLeft = (value And &H3FFFFF) * &H200&
    End If
  Case 10&
    If value And &H200000 Then
      ShiftLeft = (value And &H1FFFFF) * &H400& Or &H80000000
    Else
      ShiftLeft = (value And &H1FFFFF) * &H400&
    End If
  Case 11&
    If value And &H100000 Then
      ShiftLeft = (value And &HFFFFF) * &H800& Or &H80000000
    Else
      ShiftLeft = (value And &HFFFFF) * &H800&
    End If
  Case 12&
    If value And &H80000 Then
      ShiftLeft = (value And &H7FFFF) * &H1000& Or &H80000000
    Else
      ShiftLeft = (value And &H7FFFF) * &H1000&
    End If
  Case 13&
    If value And &H40000 Then
      ShiftLeft = (value And &H3FFFF) * &H2000& Or &H80000000
    Else
      ShiftLeft = (value And &H3FFFF) * &H2000&
    End If
  Case 14&
    If value And &H20000 Then
      ShiftLeft = (value And &H1FFFF) * &H4000& Or &H80000000
    Else
      ShiftLeft = (value And &H1FFFF) * &H4000&
    End If
  Case 15&
    If value And &H10000 Then
      ShiftLeft = (value And &HFFFF&) * &H8000& Or &H80000000
    Else
      ShiftLeft = (value And &HFFFF&) * &H8000&
    End If
  Case 16&
    If value And &H8000& Then
      ShiftLeft = (value And &H7FFF&) * &H10000 Or &H80000000
    Else
      ShiftLeft = (value And &H7FFF&) * &H10000
    End If
  Case 17&
    If value And &H4000& Then
      ShiftLeft = (value And &H3FFF&) * &H20000 Or &H80000000
    Else
      ShiftLeft = (value And &H3FFF&) * &H20000
    End If
  Case 18&
    If value And &H2000& Then
      ShiftLeft = (value And &H1FFF&) * &H40000 Or &H80000000
    Else
      ShiftLeft = (value And &H1FFF&) * &H40000
    End If
  Case 19&
    If value And &H1000& Then
      ShiftLeft = (value And &HFFF&) * &H80000 Or &H80000000
    Else
      ShiftLeft = (value And &HFFF&) * &H80000
    End If
  Case 20&
    If value And &H800& Then
      ShiftLeft = (value And &H7FF&) * &H100000 Or &H80000000
    Else
      ShiftLeft = (value And &H7FF&) * &H100000
    End If
  Case 21&
    If value And &H400& Then
      ShiftLeft = (value And &H3FF&) * &H200000 Or &H80000000
    Else
      ShiftLeft = (value And &H3FF&) * &H200000
    End If
  Case 22&
    If value And &H200& Then
      ShiftLeft = (value And &H1FF&) * &H400000 Or &H80000000
    Else
      ShiftLeft = (value And &H1FF&) * &H400000
    End If
  Case 23&
    If value And &H100& Then
      ShiftLeft = (value And &HFF&) * &H800000 Or &H80000000
    Else
      ShiftLeft = (value And &HFF&) * &H800000
    End If
  Case 24&
    If value And &H80& Then
      ShiftLeft = (value And &H7F&) * &H1000000 Or &H80000000
    Else
      ShiftLeft = (value And &H7F&) * &H1000000
    End If
  Case 25&
    If value And &H40& Then
      ShiftLeft = (value And &H3F&) * &H2000000 Or &H80000000
    Else
      ShiftLeft = (value And &H3F&) * &H2000000
    End If
  Case 26&
    If value And &H20& Then
      ShiftLeft = (value And &H1F&) * &H4000000 Or &H80000000
    Else
      ShiftLeft = (value And &H1F&) * &H4000000
    End If
  Case 27&
    If value And &H10& Then
      ShiftLeft = (value And &HF&) * &H8000000 Or &H80000000
    Else
      ShiftLeft = (value And &HF&) * &H8000000
    End If
  Case 28&
    If value And &H8& Then
      ShiftLeft = (value And &H7&) * &H10000000 Or &H80000000
    Else
      ShiftLeft = (value And &H7&) * &H10000000
    End If
  Case 29&
    If value And &H4& Then
      ShiftLeft = (value And &H3&) * &H20000000 Or &H80000000
    Else
      ShiftLeft = (value And &H3&) * &H20000000
    End If
  Case 30&
    If value And &H2& Then
      ShiftLeft = (value And &H1&) * &H40000000 Or &H80000000
    Else
      ShiftLeft = (value And &H1&) * &H40000000
    End If
  Case 31&
    If value And &H1& Then
      ShiftLeft = &H80000000
    Else
      ShiftLeft = &H0&
    End If
  End Select
End Function


Private Sub AddByte(ByRef inArr() As Byte, ByVal va As Byte)
    On Local Error GoTo 1
    inArr(0) = inArr(0)
    On Local Error GoTo 0
    ReDim Preserve inArr(LBound(inArr) To UBound(inArr) + 1) As Byte
    inArr(UBound(inArr)) = va
    Exit Sub
1   ReDim inArr(0) As Byte
    inArr(0) = va
    On Local Error GoTo 0
End Sub


