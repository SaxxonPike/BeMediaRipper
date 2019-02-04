VERSION 5.00
Begin VB.Form frmRipList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rip List"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unselect All"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select All"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   4320
      IntegralHeight  =   0   'False
      ItemData        =   "frmRipList.frx":0000
      Left            =   120
      List            =   "frmRipList.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmRipList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim X As Long
    For X = List1.ListCount - 1 To 0 Step -1
        List1.Selected(X) = True
    Next X
End Sub

Private Sub Command2_Click()
    Dim X As Long
    For X = List1.ListCount - 1 To 0 Step -1
        List1.Selected(X) = False
    Next X
End Sub

Private Sub Command3_Click()
    Me.Visible = False
End Sub

Public Function IsSelected(iIndex As Long) As Boolean
    If (iIndex - 1000) >= List1.ListCount Then
        IsSelected = True
        Exit Function
    End If
    If (iIndex - 1000) > 0 Then
        IsSelected = (List1.Selected(iIndex - 1000) = True)
    End If
End Function

