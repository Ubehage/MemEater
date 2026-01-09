VERSION 5.00
Begin VB.UserControl SeperatorLine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "SeperatorLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Refresh()
  UserControl.Cls
  DrawLine
End Sub

Private Sub DrawLine()
  Dim LineRect As RECT
  With LineRect
    .Left = 0
    .Top = 0
    .Right = (UserControl.ScaleWidth - Screen.TwipsPerPixelX)
    UserControl.Line (.Left, .Top)-(.Right, .Top), COLOR_OUTLINE_LIGHT
    .Top = (.Top + Screen.TwipsPerPixelY)
    UserControl.Line (.Left, .Top)-(.Right, .Top), COLOR_OUTLINE
  End With
End Sub

Private Sub UserControl_Resize()
  Dim h As Long
  h = ((UserControl.Height - UserControl.ScaleHeight) + (Screen.TwipsPerPixelY * 2))
  If UserControl.Height <> h Then UserControl.Height = h: Exit Sub
  Refresh
End Sub
