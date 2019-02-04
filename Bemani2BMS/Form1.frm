VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bemani2BMS v3"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Sound List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   52
      Top             =   6720
      Width           =   2295
      Begin VB.TextBox txtSoundList 
         Height          =   285
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      Height          =   375
      Left            =   5640
      TabIndex        =   37
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Frame Frame6 
      Caption         =   "Type && Alignment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5640
      TabIndex        =   47
      Top             =   5280
      Width           =   2295
      Begin VB.ComboBox cmbACTiming 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form1.frx":038A
         Left            =   1080
         List            =   "Form1.frx":039A
         TabIndex        =   58
         Text            =   "0"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtQuant 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "192"
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "Form1.frx":03B8
         Left            =   120
         List            =   "Form1.frx":03DD
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkACTiming 
         Caption         =   "Set FPS"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Quant"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   630
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Keysound Naming"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5640
      TabIndex        =   45
      Top             =   3480
      Width           =   2295
      Begin VB.CheckBox chkLineRedux 
         Caption         =   "BMS Line Reduction"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkTTLFN 
         Caption         =   "#playlevel from filename"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkKeepFile 
         Caption         =   "Keep Filenames"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkNumberFiles 
         Caption         =   "Add #s to filenames"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox cmbKeyNaming 
         Height          =   315
         ItemData        =   "Form1.frx":0476
         Left            =   120
         List            =   "Form1.frx":0486
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Modes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5640
      TabIndex        =   44
      Top             =   120
      Width           =   2295
      Begin VB.ComboBox cmbModeNames 
         Height          =   315
         ItemData        =   "Form1.frx":04C6
         Left            =   960
         List            =   "Form1.frx":04D6
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkAutoMulti 
         Caption         =   "Autoname Multi-charts"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.ListBox lstModes 
         Enabled         =   0   'False
         Height          =   2310
         ItemData        =   "Form1.frx":04FB
         Left            =   120
         List            =   "Form1.frx":051D
         Style           =   1  'Checkbox
         TabIndex        =   32
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox chkModeNames 
         Caption         =   "Modes:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   525
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Additional Tags"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   43
      Top             =   4800
      Width           =   5415
      Begin VB.TextBox txtSync 
         Height          =   285
         Left            =   4800
         TabIndex        =   61
         Text            =   "8"
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkSync 
         Caption         =   "Sync"
         Height          =   255
         Left            =   3960
         TabIndex        =   60
         Top             =   270
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkAddInfo 
         Caption         =   "Add info header"
         Height          =   255
         Left            =   2400
         TabIndex        =   49
         Top             =   270
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkAddHeader 
         Caption         =   "Add Bemani2BMS header"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   270
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.TextBox txtAddTags 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   600
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   39
      Top             =   2520
      Width           =   5415
      Begin VB.TextBox txtTitle 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   27
         ToolTipText     =   "Extra2"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   26
         ToolTipText     =   "Extra1"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   25
         ToolTipText     =   "Another14"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   24
         ToolTipText     =   "14key"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   23
         ToolTipText     =   "Light14"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   22
         ToolTipText     =   "Beginner14"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   21
         ToolTipText     =   "Another7"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   20
         ToolTipText     =   "7key"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   19
         ToolTipText     =   "Light7"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   600
         MaxLength       =   2
         TabIndex        =   18
         ToolTipText     =   "Beginner7"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   17
         ToolTipText     =   "Extra2"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   16
         ToolTipText     =   "Extra1"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3960
         MaxLength       =   2
         TabIndex        =   15
         ToolTipText     =   "Another14"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   14
         ToolTipText     =   "14key"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   13
         ToolTipText     =   "Light14"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   12
         ToolTipText     =   "Beginner14"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   11
         ToolTipText     =   "Another7"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   10
         ToolTipText     =   "7key"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   9
         ToolTipText     =   "Light7"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtPlayLevel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   600
         MaxLength       =   2
         TabIndex        =   8
         ToolTipText     =   "Beginner7"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtGenre 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox txtArtist 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "prefix"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1710
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "level"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1350
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "#GENRE"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "#ARTIST"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "#TITLE"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input File(s)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmdFill 
         Caption         =   "** Fill **"
         Height          =   375
         Left            =   4440
         TabIndex        =   54
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "Info"
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdClearFiles 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdRemoveFile 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdAddFile 
         Caption         =   "Add..."
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   975
      End
      Begin VB.ListBox lstFiles 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         ItemData        =   "Form1.frx":057B
         Left            =   120
         List            =   "Form1.frx":057D
         OLEDropMode     =   1  'Manual
         TabIndex        =   0
         Top             =   240
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Dim x As Long
Dim choice As Long
Dim CSFile As New clsCSFile
Dim bConverting As Boolean
Dim bSilent As Boolean
Dim DefaultTitle As String

Private Sub chkACTiming_Click()
    cmbACTiming.Enabled = chkACTiming.value
    If chkACTiming.value = 0 Then
        cmbACTiming = "0"
    End If
End Sub

Private Sub chkModeNames_Click()
    lstModes.Enabled = (chkModeNames.value = 1)
End Sub

Private Sub cmbType_Click()
    If cmbType.ListIndex = 1 Or cmbType.ListIndex = 3 Then
        chkModeNames.value = 0
        chkModeNames.Enabled = False
    Else
        chkModeNames.Enabled = True
    End If
    If cmbType.ListIndex = 0 Then
        Frame1.Enabled = False
        Frame2.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        Frame7.Enabled = False
    Else
        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
        Frame4.Enabled = True
        Frame7.Enabled = True
    End If
    
    If cmbType.ListIndex = 1 Then
        chkKeepFile.value = 0
    End If
    
    chkTTLFN.Enabled = (cmbType.ListIndex = 2)
    If chkTTLFN.Enabled = False Then
        chkTTLFN.value = 0
    End If
End Sub

Private Sub cmdAddFile_Click()
    Dim x As String
    Dim y As String
    Dim z As String
    Dim s() As String
    Dim c As Long
    If CD_ShowOpen_Save(Me.hWnd, OFN_EXPLORER + OFN_ALLOWMULTISELECT, x, , , "Open Files", , z, True) Then
        If InStr(Left$(x, Len(x) - 2), Chr$(0)) > 0 Then
            'multiple files
            ReDim s(1 To 1) As String
            c = 0
            y = Left$(x, InStr(x, Chr$(0)) - 1) + "\"
            x = Mid$(x, InStr(x, Chr$(0)) + 1)
            Do While Len(x) > 1
                c = c + 1
                z = y + Left$(x, InStr(x, Chr$(0)) - 1)
                x = Mid$(x, InStr(x, Chr$(0)) + 1)
                ReDim Preserve s(1 To c) As String
                s(c) = z
            Loop
        Else
            'single file
            ReDim s(1 To 1) As String
            s(1) = Left$(x, Len(x) - 2)
        End If
        AddFile s()
    End If
End Sub

Private Sub cmdBegin_Click()
    Dim ttitle As String
    Dim tartist As String
    Dim tgenre As String
    Dim tdifficulty As String
    Dim tmode As Long
    Dim smode As Long
    Dim amodes As Long
    Dim mnumber As String
    Dim outfile As String
    Dim inFile As String
    Dim fstring() As Byte
    Dim tprefix As String
    Dim fname As String
    Dim fname2 As String
    Dim fname3 As String
    
    If bConverting Then
        If MsgBox("There is a conversion in process. Do you wish to end it prematurely?", vbYesNo) = vbYes Then
            bConverting = False
            Me.Caption = DefaultTitle
        End If
        Exit Sub
    End If
    
    tmode = 0
    amodes = 0
    
    If txtQuant = "" Then
        txtQuant = 192
    End If
    
    'need to have a type
    If cmbType.ListIndex <= 0 Then
        If Not bSilent Then MsgBox "No file type selected.", , "Error"
        Exit Sub
    End If
    
    'can't do much if there's no files
    If lstFiles.ListCount = 0 Then
        If Not bSilent Then MsgBox "No files to convert.", , "Error"
        Exit Sub
    End If
    
    'Pop'n can't use mode names... yet
    If cmbType.ListIndex = 5 And chkModeNames.value = 1 Then
        If Not bSilent Then MsgBox "Can't use POP'N ripping mode with Mode Names. Disable it to continue."
        Exit Sub
    End If
    
    'check to see if we have enough modes in the box
    If cmbType.ListIndex <> 1 And cmbType.ListIndex <> 3 And cmbType.ListIndex <> 5 Then
        If chkModeNames.value = 1 And lstModes.SelCount <> lstFiles.ListCount Then
            If Not bSilent Then MsgBox "File count and mode count don't match.", , "Error"
            Exit Sub
        End If
    Else
        If chkModeNames.value = 1 And cmbType.ListIndex = 1 Then
            Open Replace(lstFiles.List(0), "/", "\") For Binary As #1
            Get #1, 1, tmode
            If tmode <> &H60 Then
                If Not bSilent Then MsgBox "Incorrect ID: this may not be a .1 file.", , "Error"
                Close #1
                Exit Sub
            End If
            For x = 0 To 11
                Get #1, 1 + (x * 8), tmode
                If tmode <> 0 Then
                    Get #1, , tmode
                    If tmode <> 0 Then
                        amodes = amodes + 1
                    End If
                End If
            Next x
            If amodes <> lstModes.SelCount Then
                If Not bSilent Then MsgBox "File count and mode count don't match. (This file has " + CStr(amodes) + " modes.)" + vbCrLf + "You might want to turn off the MODES checkbox and retry.", , "Error"
                Close #1
                Exit Sub
            End If
        End If
    End If
    
    bConverting = True
    'conversion
    
    
    tmode = 0
    amodes = 1
    If cmbType.ListIndex = 1 Then
        amodes = 12
    End If
    If cmbType.ListIndex = 3 Then
        amodes = 6
    End If
    If cmbType.ListIndex = 5 Then
        amodes = 5
    End If
    If cmbType.ListIndex = 10 Then
        amodes = 7
    End If
    For x = 0 To lstFiles.ListCount - 1
        
        lstFiles.ListIndex = x
        
        If cmdFill.Caption = "(Auto-Fill)" Then
            fname3 = lstFiles.List(x)
            fname3 = Mid$(fname3, InStrRev(fname3, "\") + 1)
            If InStr(fname3, " [") > 0 Then
                fname3 = Left$(fname3, InStr(fname3, " [") - 1)
            End If
            If cmbType.ListIndex <> 1 Then
                txtTitle.Text = fname3
            End If
            Call cmdFill_MouseDown(1, 0, 0, 0)
            DoEvents
        End If
        
        If amodes > 1 Then
            tmode = 0
        End If
        
        For smode = 0 To amodes - 1
        
            If (Not bConverting) Then
                Exit For
            End If
            
            '-------------------------------------------------------------
            
            If cmbType.ListIndex <> 1 And cmbType.ListIndex <> 3 Then
                If chkModeNames.value = 1 Then
                    Do Until lstModes.Selected(tmode) = True
                        tmode = tmode + 1
                        
                    Loop
                    ttitle = txtTitle.Text + " [" + lstModes.List(tmode) + "]"
                Else
                    ttitle = txtTitle.Text
                End If
            Else
                ttitle = txtTitle.Text
            End If
            
            '-------------------------------------------------------------
            
            inFile = Replace(lstFiles.List(x), "/", "\")
            mnumber = CStr(tmode)
            If Len(mnumber) = 1 Then
                mnumber = "0" + mnumber
            End If
            
            '-------------------------------------------------------------
            
            If InStr(inFile, "\") > 0 Then
                outfile = Left$(inFile, InStrRev(inFile, "\"))
            Else
                outfile = ""
            End If
            
            If chkKeepFile.value = 0 Then
                If chkNumberFiles.value = 1 Or (chkModeNames.value = 0 And cmbType.ListIndex <> 1 And cmbType.ListIndex <> 3) Then  'Or cmbType.ListIndex = 1 Or cmbType.ListIndex = 3 Then
                    outfile = outfile + mnumber + "_"
                End If
            
                If txtTitle.Text = "" Then
                    fname2 = Mid$(inFile, InStrRev(inFile, "\") + 1)
                    If InStr(fname2, ".") > 0 Then
                        fname2 = Left$(fname2, InStr(fname2, ".") - 1)
                    End If
                    outfile = outfile + fname2
                Else
                    outfile = outfile + txtTitle.Text
                End If
                If chkModeNames.value = 1 Then
                    outfile = outfile + " [" + lstModes.List(tmode) + "]"
                End If
            Else
                If InStr(lstFiles.List(x), ".") > 0 Then
                    outfile = Left$(lstFiles.List(x), InStrRev(lstFiles.List(x), ".") - 1)
                Else
                    outfile = lstFiles.List(x)
                End If
                If InStr(Mid$(outfile, InStrRev(outfile, "\")), "]") > 0 Then
                    outfile = Left$(outfile, InStrRev(outfile, "]"))
                End If
            End If
            
            
            '-------------------------------------------------------------
            
            tartist = txtArtist.Text
            tgenre = txtGenre.Text
            If x < 10 Then
                tdifficulty = txtPlayLevel(x)
                tprefix = txtPlayLevel(x + 10)
            Else
                tdifficulty = ""
                tprefix = ""
            End If
            Me.Caption = "Loading: " + inFile
            DoEvents
            CSFile.LoadFile inFile, cmbType.ListIndex, smode, CDbl(cmbACTiming.Text), IIf((chkSync.value = 1), CLng(txtSync.Text), -1)
            If CSFile.IsLoaded Then
                If cmbType.ListIndex = 1 Or cmbType.ListIndex = 3 Then
                    tdifficulty = txtPlayLevel(tmode)
                    If chkAutoMulti.value = 0 Then
                        choice = -2
                        DoChoice smode, CSFile.TotalNotes, CSFile.MainBPM, CSFile.TotalSongLength
                        Do: DoEvents: Sleep 1: Loop While choice = -2
                        If choice = -3 Then
                            bConverting = False
                            Exit Sub
                        End If
                    Else
                        Select Case smode
                            Case 0
                                choice = 2
                            Case 1
                                choice = 1
                            Case 2
                                choice = 3
                            Case 3
                                choice = 0
                            Case 4, 5
                                choice = 8
                            Case 6
                                choice = 6
                            Case 7
                                choice = 5
                            Case 8
                                choice = 7
                            Case 9
                                choice = 4
                            Case 10, 11
                                choice = 9
                        End Select
                    End If
                    tmode = tmode + 1
                    outfile = outfile + " [" + lstModes.List(choice) + "]"
                    tprefix = txtPlayLevel(choice + 10)
                    ttitle = ttitle + " [" + lstModes.List(choice) + "]"
                ElseIf cmbType.ListIndex = 5 Then
                    tmode = tmode + 1
                ElseIf cmbType.ListIndex = 10 Then
                    If chkAutoMulti.value = 0 Then
                        choice = -2
                        DoChoice smode, CSFile.TotalNotes, CSFile.MainBPM, CSFile.TotalSongLength
                        Do: DoEvents: Sleep 1: Loop While choice = -2
                        If choice = -3 Then
                            bConverting = False
                            Exit Sub
                        End If
                    Else
                        If CSFile.CSSize = 393216 Then 'GOLD CS
                            Select Case smode
                                Case 0
                                    choice = 1
                                Case 1
                                    choice = 2
                                Case 2
                                    choice = 3
                                Case 3
                                    choice = 5
                                Case 4
                                    choice = 6
                                Case 5
                                    choice = 7
                            End Select
                        ElseIf CSFile.CSSize = 458752 Then 'TROOPERS CS
                            Select Case smode
                                Case 0
                                    choice = 1
                                Case 1
                                    choice = 2
                                Case 2
                                    choice = 3
                                Case 3
                                    choice = -3
                                Case 4
                                    choice = 5
                                Case 5
                                    choice = 6
                                Case 6
                                    choice = 7
                            End Select
                        End If
                    End If
                    tmode = tmode + 1
                    If choice > 0 Then
                        tdifficulty = txtPlayLevel(choice)
                        outfile = outfile + " [" + lstModes.List(choice) + "]"
                        tprefix = txtPlayLevel(choice + 10)
                        ttitle = ttitle + " [" + lstModes.List(choice) + "]"
                    Else
                        outfile = ""
                    End If
                End If
                If outfile <> "" Then
                    outfile = outfile + ".bme"
                    outfile = Replace(outfile, """", "''")
                    Me.Caption = "Saving: " + outfile
                    DoEvents
                    CSFile.WriteBMS outfile, (chkAddInfo.value = 1), (chkAddHeader.value = 1), ttitle, tartist, tgenre, tdifficulty, txtAddTags.Text, tprefix, cmbKeyNaming.ListIndex, CStr(txtQuant.Text), (chkTTLFN.value = 1), (chkLineRedux.value = 1)
                End If
            End If
            
            
            '-------------------------------------------------------------
        Next smode
        tmode = tmode + 1
        If (Not bConverting) Then
            Exit For
        End If
    Next x
    
    If Not bSilent Then MsgBox "Finished."
    bConverting = False
    Me.Caption = DefaultTitle
    
    If cmdClearFiles.Caption <> "Clear" Then
        cmdClearFiles_Click
    End If
    
End Sub

Private Sub cmdClearFiles_Click()
    lstFiles.Clear
End Sub

Private Sub cmdClearFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If cmdClearFiles.Caption = "Clear" Then
            cmdClearFiles.Caption = "(Auto-clear)"
        Else
            cmdClearFiles.Caption = "Clear"
        End If
    End If
End Sub

Private Sub cmdFill_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim smode As Long
    Dim fname As String
    
    If Button = 2 Then
        If cmdFill.Caption = "** Fill **" Then
            cmdFill.Caption = "(Auto-Fill)"
        Else
            cmdFill.Caption = "** Fill **"
        End If
        Exit Sub
    End If
    
    If cmbType.ListIndex = 1 And txtSoundList.Text <> "" Then
    
        If lstFiles.ListCount = 0 Then
            Exit Sub
        End If
        
        LoadSoundList App.Path + "\" + txtSoundList.Text
        If lstFiles.ListIndex = -1 Then
            lstFiles.ListIndex = 0
        End If
        fname = lstFiles.List(lstFiles.ListIndex)
        If InStr(fname, "\") Then
            fname = Mid$(fname, InStrRev(fname, "\") + 1)
        End If

    
        For smode = 0 To UBound(SoundList) 'refers to soundlist in modSoundList
            If (SoundList(smode).SongID + ".1") = fname Then
                With SoundList(smode)
                    txtTitle = Trim(.MainTitle + " " + .SubTitle)
                    txtArtist = Trim(.Artist)
                    txtGenre = Trim(.Genre)
                    txtPlayLevel(0) = Trim(.Difficulty(6))
                    txtPlayLevel(1) = Trim(.Difficulty(0))
                    txtPlayLevel(2) = Trim(.Difficulty(1))
                    txtPlayLevel(3) = Trim(.Difficulty(2))
                    txtPlayLevel(4) = ""
                    txtPlayLevel(5) = Trim(.Difficulty(3))
                    txtPlayLevel(6) = Trim(.Difficulty(4))
                    txtPlayLevel(7) = Trim(.Difficulty(5))
                    txtPlayLevel(8) = ""
                    txtPlayLevel(9) = ""
                    txtAddTags = ""
                    If .VideoFile <> "" Then
                        txtAddTags = "#VIDEOFILE " + .VideoFile + ".4" + vbCrLf + _
                            .VideoInfoCol + vbCrLf + _
                            .VideoInfoDly + vbCrLf + _
                            .VideoInfoFS
                    End If
                End With
                Exit For
            End If
        Next smode
    Else
        
        If Button = 10 Then
            If lstFiles.ListCount > 0 Then
                fname = lstFiles.List(0)
                fname = Mid$(fname, InStrRev(fname, "\") + 1)
                If InStr(fname, " [") > 0 Then
                    fname = Left$(fname, InStr(fname, " [") - 1)
                End If
                txtTitle.Text = fname
            End If
        End If
        
        'use songDB
        txtArtist.Text = ""
        txtGenre.Text = ""
        For smode = LBound(SongDB) To UBound(SongDB)
            fname = UCase$(SongDB(smode).Title)
            fname = Replace$(fname, "*", "_")
            fname = Replace$(fname, "/", "_")
            fname = Replace$(fname, ":", "_")
            fname = Replace$(fname, "?", "_")
            fname = Replace$(fname, "\", "_")
            fname = Replace$(fname, "<", "_")
            fname = Replace$(fname, ">", "_")
            If UCase$(Trim$(txtTitle.Text)) = fname Then
                txtArtist.Text = SongDB(smode).Artist
                txtGenre.Text = SongDB(smode).Genre
                txtTitle.Text = SongDB(smode).Title
                Exit For
            End If
        Next smode
    
    End If
End Sub

Private Sub cmdInfo_Click()
    Dim temphead As Long
    Dim modecount As Long
    Dim x As Long
    If lstFiles.ListCount = 0 Then
        Exit Sub
    End If
    If lstFiles.ListIndex = -1 Then
        lstFiles.ListIndex = 0
    End If
    Dim f As Long
    f = FreeFile
    Open lstFiles.List(lstFiles.ListIndex) For Binary As #f
    Get #f, 1, temphead
    If temphead = 8 Then
        MsgBox "Type: .CS (2nd version)" + vbCrLf + "Modes: 1", , "Info"
    ElseIf temphead = &H60 Then
        For x = 0 To 11
            Get #f, 1 + (x * 8), temphead
            If temphead <> 0 Then
                Get #f, , temphead
                If temphead <> 0 Then
                    modecount = modecount + 1
                End If
            End If
        Next x
        MsgBox "Type: .1 (arcade)" + vbCrLf + "Modes: " + CStr(modecount)
    Else
        MsgBox "Type: Unknown / No header", , "Info"
    End If
    Close #f
End Sub

Private Sub cmdRemoveFile_Click()
    If lstFiles.ListIndex > -1 Then
        lstFiles.RemoveItem lstFiles.ListIndex
    End If
End Sub

Private Sub cmbModeNames_Click()
    Dim x As Long
    Select Case cmbModeNames.ListIndex
        Case 0
            lstModes.List(0) = "Beginner7"
            lstModes.List(1) = "Light7"
            lstModes.List(2) = "7key"
            lstModes.List(3) = "Another7"
            lstModes.List(4) = "Beginner14"
            lstModes.List(5) = "Light14"
            lstModes.List(6) = "14key"
            lstModes.List(7) = "Another14"
        Case 1
            lstModes.List(0) = "Beginner7"
            lstModes.List(1) = "Normal7"
            lstModes.List(2) = "Hyper7"
            lstModes.List(3) = "Another7"
            lstModes.List(4) = "Beginner14"
            lstModes.List(5) = "Normal14"
            lstModes.List(6) = "Hyper14"
            lstModes.List(7) = "Another14"
        Case 2
            lstModes.List(0) = ""
            lstModes.List(1) = ""
            lstModes.List(2) = ""
            lstModes.List(3) = ""
            lstModes.List(4) = ""
            lstModes.List(5) = ""
            lstModes.List(6) = ""
            lstModes.List(7) = ""
        Case 3
            lstModes.List(0) = ""
            lstModes.List(1) = ""
            lstModes.List(2) = ""
            lstModes.List(3) = ""
            lstModes.List(4) = ""
            lstModes.List(5) = ""
            lstModes.List(6) = ""
            lstModes.List(7) = ""
    End Select
    For x = 0 To 9
        If lstModes.List(x) <> "" Then
            txtPlayLevel(x).ToolTipText = lstModes.List(x)
            txtPlayLevel(x).Visible = True
            txtPlayLevel(x + 10).Visible = True
        Else
            txtPlayLevel(x).Visible = False
            txtPlayLevel(x + 10).Visible = flase
        End If
    Next x
End Sub

Private Sub Form_Load()
    ReDim SongDB(0) As xSongDB
    LoadSongDB App.Path + "\songDB.txt"
    Dim x As Integer
    Dim y As String
    Dim c As String
    c = Trim$(Command$)
    Do While InStr(c, Chr$(34)) > 0
        c = Mid$(c, InStr(c, Chr$(34)) + 1)
        y = Left$(c, InStr(c, Chr$(34)) - 1)
        lstFiles.AddItem y
        c = Mid$(c, InStr(c, Chr$(34)) + 1)
    Loop
    On Error Resume Next
    If Dir(App.Path + "\bemedia3.cfg") = "" Then
        Open App.Path + "\bemedia3.cfg" For Output As #1
        Print #1, 0
        Print #1, 0
        Print #1, 1
        Print #1, "output.txt"
        Print #1, 1
        Print #1, 1
        Close #1
    End If
    Open App.Path + "\bemedia3.cfg" For Input As #1
    Input #1, x: cmbKeyNaming.ListIndex = x
    Input #1, x: cmbType.ListIndex = x
    Input #1, x: cmbModeNames.ListIndex = x
    Input #1, y: txtSoundList.Text = y
    Input #1, x: If x = 1 Then cmdClearFiles_MouseDown 2, 0, 0, 0
    Input #1, x: If x = 1 Then cmdFill_MouseDown 2, 0, 0, 0
    Close #1
    On Error GoTo 0
    Me.Caption = Me.Caption + "." + CStr(App.Minor) + " (Release " + CStr(App.Revision) + ")"
    DoEvents
    DefaultTitle = Me.Caption
    c = Trim$(Command$)
    If c <> "" Then
        y = Left$(c, InStr(c, Chr$(34)) - 1)
        If Left$(y, 2) = "-D" Then
            bSilent = True
            cmdBegin_Click
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Open App.Path + "\bemedia3.cfg" For Output As #1
    Print #1, CStr(cmbKeyNaming.ListIndex)
    Print #1, CStr(cmbType.ListIndex)
    Print #1, CStr(cmbModeNames.ListIndex)
    Print #1, txtSoundList.Text
    If cmdClearFiles.Caption = "Clear" Then
        Print #1, "0"
    Else
        Print #1, "1"
    End If
    Close #1
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim a As Long
    Dim z() As String
    If Data.files.Count > 0 Then
        ReDim z(1 To Data.files.Count)
        For a = 1 To UBound(z)
            z(a) = Data.files(a)
        Next a
        AddFile z()
    End If
End Sub

Private Sub AddFile(files() As String)
    Dim s As String
    Dim f As String
    Dim n As Long
    If cmdClearFiles.Caption <> "Clear" Then
        lstFiles.Clear
    End If
    If UCase$(Right$(files(1), 4)) = ".TXT" Then
        lstFiles.Clear
        n = FreeFile
        Open files(1) For Input As #n
        Do While Not EOF(n)
            Line Input #n, s
            lstFiles.AddItem s
        Loop
        Close #n
        Exit Sub
    End If
    For x = LBound(files) To UBound(files)
        f = files(x)
        s = Dir(f)
        If s = "" Then
            f = f + "\*.*"
            s = Dir(f)
            Do While s <> ""
                
                lstFiles.AddItem files(x) + "\" + s 'Data.Files(x)
                s = Dir
            Loop
        Else
            lstFiles.AddItem f
        End If
    Next x
    If cmdFill.Caption = "(Auto-Fill)" Then
        Call cmdFill_MouseDown(10, 0, 0, 0)
    End If
End Sub

Private Sub lstModes_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button And 2) = 2 Then
        For x = 0 To lstModes.ListCount - 1
            lstModes.Selected(x) = False
        Next x
    End If
End Sub

Private Sub txtAddTags_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = 2 Then
        txtAddTags.Text = ""
    End If
End Sub

Private Sub txtArtist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = 2 Then
        txtArtist.Text = ""
    End If
End Sub

Private Sub txtGenre_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = 2 Then
        txtGenre.Text = ""
    End If
End Sub

Private Sub txtPlayLevel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = 2 Then
        For x = 0 To txtPlayLevel.Count - 1
            txtPlayLevel(x).Text = ""
        Next x
    End If
    lstModes.ListIndex = Index Mod 10
End Sub

Private Sub txtQuant_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = 2 Then
        txtQuant = 192
    End If
End Sub

Private Sub txtTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shift = 2 Then
        txtTitle.Text = ""
    End If
End Sub

Public Sub ChoiceCallback(ch As Long)
    choice = ch
    Me.Enabled = True
End Sub

Private Sub DoChoice(modenumber As Long, notecount As Long, BPM As Long, songlen As Long)
    Me.Enabled = False
    Form1.ChooseMode modenumber, notecount, BPM, songlen
End Sub
