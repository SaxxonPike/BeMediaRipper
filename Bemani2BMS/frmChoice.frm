VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Multi-chart naming"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Mode Here"
      Height          =   2655
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   2295
      Begin VB.ListBox lstModes 
         Height          =   2205
         ItemData        =   "frmChoice.frx":0000
         Left            =   120
         List            =   "frmChoice.frx":0025
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chart Info"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3255
      Begin VB.Label Label4 
         Caption         =   "Song Length (units):"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "BPM:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Note Count:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label Label1 
      Caption         =   "If auto-naming for multi-charts is disabled, the chart naming must be done by the user for each mode found."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ChooseMode(modenum As Long, notecount As Long, bpm As Long, songlengthunits As Long)
    Dim x As Long
    Frame1.Caption = Frame1.Caption + "[" + CStr(modenum) + "]"
    Label2 = Label2 + " " + CStr(notecount)
    Label3 = Label3 + " " + CStr(bpm)
    Label4 = Label4 + " " + CStr(songlengthunits)
    For x = 0 To lstModes.ListCount - 1
        lstModes.List(x) = frmMain.lstModes.List(x)
    Next x
    Me.Show
End Sub

Private Sub Command1_Click()
    If lstModes.ListIndex >= 0 Then
        frmMain.ChoiceCallback lstModes.ListIndex
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    frmMain.ChoiceCallback -3
    Unload Me
End Sub

Private Sub lstModes_DblClick()
    Command1_Click
End Sub
