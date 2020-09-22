Attribute VB_Name = "Module1"
Option Explicit

Public Const HTCAPTION As Integer = 2
Public Const WM_NCLBUTTONDOWN As Integer = &HA1

Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const LWA_ALPHA As Long = &H2&
Public Const LWA_COLORKEY As Integer = &H1

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Incr As Integer
Public Frame As Integer

Public Sub Rotate()

    Frame = Frame + Incr

    If Frame = 137 Then
        Frame = 101
    End If

    If Frame = 100 Then
        Frame = 136
    End If
    
    Form1.Picture = LoadResPicture(Frame, 0)

End Sub

