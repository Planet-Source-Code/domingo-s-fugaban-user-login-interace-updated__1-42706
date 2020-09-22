Attribute VB_Name = "modMouseOver"
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function WindowFromPoint Lib "user32" ( _
    ByVal xPoint As Long, _
    ByVal yPoint As Long) As Long
    
Public Declare Function GetCursorPos Lib "user32" ( _
    lpPoint As POINTAPI) As Long
