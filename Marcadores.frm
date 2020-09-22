VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   6375
      Index           =   0
      Left            =   420
      Picture         =   "Marcadores.frx":0000
      Top             =   440
      Width           =   6390
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MK_LBUTTON As Long = &H1
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
    
Private Const MK_RBUTTON As Long = &H2
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205

Private Const WM_LBUTTONDBLCLK As Long = &H203
    
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Sub Form_Load()
    Dim NormalWindowStyle As Long
    Dim col As Long
    Dim Ret As Long
    
    NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, 50, LWA_ALPHA

    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    col = RGB(0, 0, 0)
    SetLayeredWindowAttributes Me.hwnd, col, 50, LWA_COLORKEY

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nMousePosition As Long
    
    nMousePosition = MakeDWord(X, Y)
    Select Case Button
      Case 1
        PostMessage Form1.hwnd, WM_LBUTTONDOWN, Button, nMousePosition
      Case 2
        PostMessage Form1.hwnd, WM_RBUTTONDOWN, Button, nMousePosition
    End Select
    
End Sub

Public Function MakeDWord(LoWord As Single, HiWord As Single) As Long

    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)

End Function


