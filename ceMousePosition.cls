VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ceMousePosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Response As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Property Get Top() As Long
    Top = MousePosition.Y * 0.75
End Property

Public Property Get Left() As Long
    Left = MousePosition.X * 0.75
End Property

Private Property Get MousePosition() As POINTAPI
    Response = GetCursorPos(MousePosition)
End Property

