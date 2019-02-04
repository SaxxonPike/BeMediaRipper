VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Konami Decompressor v1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Target"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4455
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Text            =   "out.dat"
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Source"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4455
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Offset"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format Select"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmDecoderMain.frx":0000
         Left            =   120
         List            =   "frmDecoderMain.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PSXDecode As New clsPSXDecode

Private Sub Command1_Click()
    Dim Decoded() As Byte
    Dim InFile As New clsFileStream
    InFile.OpenFile Text1.Text, True, False
    Select Case Combo1.ListIndex
        Case 0 'tool.c: decode()
            PSXDecode.DecodeBemani1 InFile, Decoded(), CDbl(Val(Text2.Text))
            Open Text3.Text For Output As #1
            Close #1
            Open Text3.Text For Binary As #1
            Put #1, 1, Decoded
            Close #1
    End Select
    InFile.CloseFile
    Set InFile = Nothing
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Text = Data.Files(1)
End Sub

Private Sub Text3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text3.Text = Data.Files(1)
End Sub

