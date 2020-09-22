VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private qBrush As Long

Private Sub Form_Load()
qBrush = GetGridBrush(19, Me.hdc)

Me.ScaleMode = vbPixels

SelectObject Me.hdc, qBrush

End Sub

Private Sub Form_Paint()
Dim r As RECT

SetRect r, 0, 0, Me.ScaleWidth, Me.ScaleHeight
FillRect Me.hdc, r, qBrush

End Sub
