VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3285
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2025
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Pt As POINTAPI

Private Capture As Boolean
Private xx As Long
Private yy As Long
Private rx As Long
Private ry As Long

Private Sub Combo1_Click()

    If Combo1.Text = 0 Then
        Exit Sub
    End If
    
    Incr = Combo1.Text
    Rotate
    Combo1.ListIndex = 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
      Case 37, 38
        Incr = 1
        Rotate
      Case 39, 40
        Incr = -1
        Rotate
      Case 27
        Unload Me
        End
    End Select

End Sub

Private Sub Form_Load()

  Dim Ret As Long

    Combo1.AddItem 1
    Combo1.AddItem 0
    Combo1.AddItem -1
    Combo1.Left = -1000

    Frame = 111
    
    Rotate

    Move Screen.Width / 2 - Width / 2, Screen.Height - Height - 300

    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_COLORKEY Or LWA_ALPHA

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    xx = X
    yy = Y
    Capture = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Capture Then
        GetCursorPos Pt
        Move Pt.X * Screen.TwipsPerPixelX - xx, Top
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Capture = False

End Sub

Private Sub Form_Resize()

    Form1.CurrentX = 0
    Form1.CurrentY = 0
    Form1.Print "Use Mouse Wheel" & vbCrLf; "  or Cursor KEYS" & vbCrLf & "     to ROTATE"
    ForeColor = &H80FFFF
    Form1.Print "DRAG TO MOVE"
    ForeColor = &HFFFFFF
    Form1.Print "   ESC TO EXIT"
    Combo1.SetFocus

End Sub

