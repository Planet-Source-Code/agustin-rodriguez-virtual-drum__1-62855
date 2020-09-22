VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu nome 
         Caption         =   "Agustin Rodriguez"
         Index           =   0
         Begin VB.Menu opt 
            Caption         =   "E-Mail"
            Index           =   1
         End
         Begin VB.Menu opt 
            Caption         =   "Home PAge"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Msg As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1

Private Sub opt_Click(Index As Integer)

On Error Resume Next
Select Case Index
Case 1
    ShellExecute hWnd, "open", "mailto:virtual_guitar_1@hotmail.com", vbNullString, vbNullString, conSwNormal
Case 2
    ShellExecute hWnd, "open", "http://geocities.com/virtual_quality/", vbNullString, vbNullString, conSwNormal
Case 3
    Frame1.Visible = False
End Select

End Sub
