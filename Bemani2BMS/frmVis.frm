VERSION 5.00
Begin VB.Form frmVis 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2475
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1
   ScaleMode       =   0  'User
   ScaleWidth      =   8
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrVis 
      Interval        =   1
      Left            =   120
      Top             =   0
   End
End
Attribute VB_Name = "frmVis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
