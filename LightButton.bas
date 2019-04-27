Attribute VB_Name = "Module1"
Option Explicit

Type POINTAPI
        x As Long
        y As Long
End Type

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

