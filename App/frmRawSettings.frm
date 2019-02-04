VERSION 5.00
Begin VB.Form frmRawSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Raw VAG settings"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3690
   ControlBox      =   0   'False
   Icon            =   "frmRawSettings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Remove Leading Silence"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   7
      Text            =   "44100"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Text            =   "1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Text            =   "0"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Frequency"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Block Size"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Channel Count"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Offset"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmRawSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
