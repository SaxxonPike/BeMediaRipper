VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------
'  SaxxonPike's File I/O Stream API Wrapper Class for VB6
'   2006 full revision
'
'  Permits file access to VB up to 4 GB.
'
'  CountBase: when converting from VB's built-in file
'             access, you probably won't want to go through
'             with having to change every offset that you
'             coded into the program. Change CountBase
'             to 1 to make the first byte of the file
'             interpreted as 1.
'
'  RetryCount: number of times a disk access operation is
'              attempted before failing out.
'
'  AdvanceOffset: when accessing data on disk, setting this
'                 to FALSE will prevent the file pointer
'                 from being moved. When set to TRUE, the
'                 pointer is moved according to how many
'                 bytes were read/written.
'
'  UseReadBuffer: with this set to TRUE, an internal buffer will be
'                 used. You can set the buffer size to what you wish.
'                 If any requested bytes fall within the bounds of the
'                 buffer, they are read from memory instead of from
'                 the stream. If not, the buffer is refreshed. Note that
'                 because Windows already has a built-in stream buffering
'                 system, you will only want to use this feature when
'                 you plan on using offsets that jump back often.
'                 That's because the Windows buffer only works when
'                 scanning forward.
'
' -------------------------------------------------------------------------

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const MAX_PATH = 255
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        Offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private iRetryCount As Long
Private iOverlapped As OVERLAPPED
Private iHandle As Long
Private iOffset As Double
Private iBase As Double
Private iFind As WIN32_FIND_DATA
Private iFileInfo As BY_HANDLE_FILE_INFORMATION
Private bAdvanceOffset As Boolean
Private HeldOffset As Double
Private bBufferedReads As Boolean
Private bBR() As Byte
Private xBufferSize As Long

Private xBufferOffset As Double

Public Property Get BufferSize() As Long
    BufferSize = xBufferSize
End Property
Public Property Let BufferSize(newsize As Long)
    ReDim bBR(0 To newsize - 1) As Byte
    xBufferSize = newsize
End Property
Public Property Get UseReadBuffer() As Boolean
    UseReadBuffer = bBufferedReads
End Property
Public Property Let UseReadBuffer(i As Boolean)
    bBufferedReads = i
    xBufferOffset = -1
End Property

Public Sub CloseFile()
    lClose iHandle
    iHandle = 0
End Sub

Public Property Get FileHandle() As Long
    FileHandle = iHandle
End Property

Public Property Get FileSize() As Double
    Dim tempfs As Double
    Dim x As Long
    tempfs = iFileInfo.nFileSizeLow
    If tempfs < 0 Then
        tempfs = tempfs + 2147483648#
        tempfs = tempfs + 2147483648#
        'tempfs = 2147483648# - tempfs
    End If
    For x = 1 To iFileInfo.nFileSizeHigh
        tempfs = tempfs + 2147483648#
        tempfs = tempfs + 2147483648#
    Next x
    FileSize = tempfs
End Property

Public Property Get BytesLeft() As Double
    BytesLeft = FileSize - (iOffset - iBase)
End Property

Public Property Get AdvanceOffset() As Boolean
    AdvanceOffset = bAdvanceOffset
End Property

Public Property Let AdvanceOffset(aoflag As Boolean)
    bAdvanceOffset = aoflag
End Property

Public Property Get CountBase() As Double
    CountBase = iBase
End Property

Public Property Let CountBase(ByVal Base As Double)
    Dim tempbc As Long
    tempbc = Base - iBase
    If Base < 0 Then
        Base = 0
    End If
    iOffset = iOffset + tempbc
    iBase = Base
End Property

Public Property Get Offset() As Double
    Offset = iOffset
End Property

Public Property Let Offset(Offs As Double)
    iOffset = Offs
    SeekFile iOffset
End Property

Public Property Get RetryCount() As Long
    RetryCount = iRetryCount
End Property

Public Property Let RetryCount(c As Long)
    iRetryCount = RetryCount
End Property

Private Sub SeekFile(ByVal Offs As Double)
    Dim templong As Long
    Offs = Offs - iBase
    If Offs < 0 Then Offs = 0
    iOffset = Offs
    If Offs >= 2147483648# Then
        Offs = Offs - 2147483648#
        If Offs > 2147483648# Then
            Offs = 0
        Else
            Offs = -2147483648# + Offs
        End If
    End If
    iOverlapped.Offset = Offs
    iOverlapped.OffsetHigh = 0
End Sub

Public Sub SeekFileRel(ByVal Offs As Double)
    SeekFile iOffset + Offs + iBase
End Sub

Public Function ReadFileByteFast(Optional ByVal Offs As Double = -1) As Byte
    ReadFileObject VarPtr(ReadFileByteFast), 1, Offs
End Function

Public Function ReadFileBinary(ByRef TargetArray() As Byte, Optional ByVal Bytes As Long = -1, Optional ByVal Offs As Double = -1, Optional ByVal AutoDefine As Boolean = False) As Boolean
    Dim ret As Long
    Dim erc As Long
    If iHandle = 0 Then
        ReadFileBinary = True
        Exit Function
    End If
    If Bytes = -1 And AutoDefine = False Then
        Bytes = (UBound(TargetArray) - LBound(TargetArray)) + 1
    ElseIf Bytes <> -1 And AutoDefine = True Then
        If Bytes > 0 Then
        
        ReDim TargetArray(0 To Bytes - 1) As Byte
        End If
    Else
        Debug.Print "ReadFileBinary: Can't autodefine a read with unknown byte count"
        Exit Function
    End If
    If Offs - iBase > -1 Then
        SeekFile Offs
    End If
    Do While ReadBytes(VarPtr(TargetArray(LBound(TargetArray))), Bytes, Offs, ret) = 0
        Sleep 1
        If erc = iRetryCount Then
            ret = -1
            Exit Do
        End If
        erc = erc + 1
        Debug.Print "Failed BINARY Read:"; Bytes; "bytes at "; iOverlapped.Offset
    Loop
    If ret <> Bytes Then
        ReadFileBinary = True
    Else
        If bAdvanceOffset Then
            SeekFileRel Bytes + 0
        End If
    End If
End Function

Public Function ReadFileObject(ByVal ObjPointer As Long, ByVal Bytes As Long, Optional ByVal Offs As Double = -1) As Boolean
    Dim ret As Long
    Dim erc As Long
    If iHandle = 0 Then
        ReadFileObject = True
        Exit Function
    End If
    If Offs > -1 Then
        SeekFile Offs
    End If
    Do While ReadBytes(ObjPointer, Bytes, Offs, ret) = 0
        Sleep 1
        If erc = iRetryCount Then
            ret = -1
            Exit Do
        End If
        erc = erc + 1
        Debug.Print "Failed OBJECT Read:"; Bytes; "bytes at "; iOverlapped.Offset
    Loop
    If ret <> Bytes Then
        ReadFileObject = True
    Else
        If bAdvanceOffset Then
            SeekFileRel Bytes + 0
        End If
    End If
End Function

Public Function OpenFile(ByVal sFileName As String, Optional ByVal bReadAccess As Boolean = True, Optional ByVal bWriteAccess As Boolean = False) As Boolean
    Dim tempac As Long
    If bReadAccess Then
        tempac = tempac Or GENERIC_READ
    End If
    If bWriteAccess Then
        tempac = tempac Or GENERIC_WRITE
    End If
    If Dir(sFileName) = "" Then
        OpenFile = True
        Exit Function
    End If
    If iHandle <> 0 Then
        CloseFile
    End If
    iOverlapped.Offset = 0
    iHandle = CreateFile(sFileName, tempac, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, 0)
    GetFileInformationByHandle iHandle, iFileInfo
    Debug.Print "OpenFile: handle"; iHandle; " / Filesize:"; iFileInfo.nFileSizeLow
End Function

Private Sub Class_Initialize()
    iRetryCount = 3
    bAdvanceOffset = True
    ReDim bBR(0 To 511) As Byte
    xBufferSize = 512
    xBufferOffset = -1
End Sub

Private Sub Class_Terminate()
    If iHandle <> 0 Then
        CloseFile
    End If
End Sub

Public Sub PushOffset(Optional inOffs As Double = -1)
    If inOffs <> -1 Then
        HeldOffset = inOffs
    Else
        HeldOffset = Offset
    End If
End Sub

Public Sub PopOffset()
    Offset = HeldOffset
    HeldOffset = 0
End Sub

Public Property Get OffsetHeld() As Double
    OffsetHeld = HeldOffset
End Property

Private Function ReadBytes(vptr As Long, bytecount As Long, ByVal Offs As Double, ByRef ret As Long)
    Dim erc As Long
    'If (bBufferedReads = True) And (bytecount <= xBufferSize) Then
    '    If Offs = -1 Then
    '        Offs = Offset
    '    End If
    '    If (Offs < xBufferOffset) Or (Offs >= xBufferOffset + xBufferSize) Or (xBufferOffset = -1) Then
    '        xBufferOffset = Offs
    '        Do While ReadFile(iHandle, bBR(0), xBufferSize, ret, ByVal VarPtr(iOverlapped)) = 0
    '            Sleep 1
    '            If erc = iRetryCount Then
    '                ret = -1
    '                Exit Do
    '            End If
    '            erc = erc + 1
    '            Debug.Print "Failed BUFFER Read:"; Bytes; "bytes at "; iOverlapped.Offset
    '        Loop
    '        'Debug.Print "clsFileStream: Buffer read ", CStr(Offs)
    '    End If
    '    CopyMemory ByVal vptr, bBR((Offs - xBufferOffset)), bytecount
    '    ReadBytes = 1
    '    ret = bytecount
    'Else
        ReadBytes = ReadFile(iHandle, ByVal vptr, bytecount, ret, ByVal VarPtr(iOverlapped))
        'Debug.Print "clsFileStream: Main read ", CStr(Offs)
    'End If
End Function

Public Sub QuickExtract(ByVal FName As String, ByVal xSize As Double, Optional ByVal Offs As Double = -1)
    Dim bBQ() As Byte
    Dim OldOffset As Double
    OldOffset = Offset
    If Offs < 0 Then
        Offs = OldOffset
    End If
    If xSize <= 0 Then Exit Sub
    ReDim bBQ(0 To xSize - 1) As Byte
    Do While ReadFile(iHandle, bBQ(0), xSize, ret, ByVal VarPtr(iOverlapped)) = 0
        Sleep 1
        If erc = iRetryCount Then
            ret = -1
            Exit Do
        End If
        erc = erc + 1
        Debug.Print "Failed QUICKEXTRACT Read:"; Bytes; "bytes at "; iOverlapped.Offset
    Loop
    Dim f As Long
    f = FreeFile
    Open FName For Binary As #f
    Put #f, 1, bBQ
    Close #f
    Offset = OldOffset
End Sub
