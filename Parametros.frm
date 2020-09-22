VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4860
   ClientLeft      =   -45
   ClientTop       =   -435
   ClientWidth     =   5145
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Parametros.frx":0000
   ScaleHeight     =   4860
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   2880
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   840
      MouseIcon       =   "Parametros.frx":50232
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Image7 
      Height          =   970
      Left            =   4560
      MouseIcon       =   "Parametros.frx":5053C
      MousePointer    =   99  'Custom
      Picture         =   "Parametros.frx":50846
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   195
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   4560
      MouseIcon       =   "Parametros.frx":51698
      MousePointer    =   99  'Custom
      Picture         =   "Parametros.frx":519A2
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   195
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   4560
      MouseIcon       =   "Parametros.frx":527F4
      MousePointer    =   99  'Custom
      Picture         =   "Parametros.frx":52AFE
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   1350
      Index           =   5
      Left            =   7920
      Picture         =   "Parametros.frx":53950
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   1350
      Index           =   4
      Left            =   7560
      Picture         =   "Parametros.frx":547A2
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   1320
      Index           =   3
      Left            =   7200
      Picture         =   "Parametros.frx":555F4
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   1320
      Index           =   2
      Left            =   6840
      Picture         =   "Parametros.frx":563F6
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   1350
      Index           =   1
      Left            =   6480
      Picture         =   "Parametros.frx":571F8
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   1350
      Index           =   0
      Left            =   6120
      Picture         =   "Parametros.frx":5804A
      Top             =   3120
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   840
      Index           =   4
      Left            =   6840
      Picture         =   "Parametros.frx":58E9C
      Top             =   1560
      Width           =   225
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RESOLUTION"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   120
      Index           =   2
      Left            =   3975
      TabIndex        =   24
      Top             =   1755
      Width           =   660
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   4000
      TabIndex        =   23
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   840
      Index           =   3
      Left            =   6480
      Picture         =   "Parametros.frx":58FAB
      Top             =   1680
      Width           =   225
   End
   Begin VB.Image Image3 
      Height          =   735
      Index           =   2
      Left            =   6120
      Picture         =   "Parametros.frx":590B5
      Top             =   1680
      Width           =   225
   End
   Begin VB.Image Image3 
      Height          =   735
      Index           =   1
      Left            =   5880
      Picture         =   "Parametros.frx":591B3
      Top             =   1680
      Width           =   225
   End
   Begin VB.Image Image3 
      Height          =   720
      Index           =   0
      Left            =   5520
      Picture         =   "Parametros.frx":592A9
      Top             =   1800
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   4200
      Picture         =   "Parametros.frx":5938D
      Top             =   2000
      Width           =   225
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   1
      Left            =   480
      MouseIcon       =   "Parametros.frx":5948B
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Load"
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   0
      Left            =   480
      MouseIcon       =   "Parametros.frx":59795
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "120"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4000
      TabIndex        =   20
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BPM"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4120
      TabIndex        =   19
      Top             =   3000
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   7
      Left            =   1320
      Top             =   4080
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   6
      Left            =   1320
      Top             =   3600
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   5
      Left            =   1320
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   4
      Left            =   1320
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   3
      Left            =   1320
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   2
      Left            =   1320
      Top             =   1680
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   1
      Left            =   1320
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   0
      Left            =   1320
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000001&
      Height          =   345
      Index           =   0
      Left            =   1800
      TabIndex        =   18
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "36"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000001&
      Height          =   345
      Index           =   1
      Left            =   1800
      TabIndex        =   17
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "37"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000001&
      Height          =   345
      Index           =   2
      Left            =   1800
      TabIndex        =   16
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "38"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000001&
      Height          =   345
      Index           =   3
      Left            =   1800
      TabIndex        =   15
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "39"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000001&
      Height          =   345
      Index           =   4
      Left            =   1800
      TabIndex        =   14
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000001&
      Height          =   345
      Index           =   5
      Left            =   1800
      TabIndex        =   13
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "41"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000001&
      Height          =   345
      Index           =   6
      Left            =   1800
      TabIndex        =   12
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "42"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000001&
      Height          =   345
      Index           =   7
      Left            =   1800
      TabIndex        =   11
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000001&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000001&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000001&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000001&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000001&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000001&
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000001&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000001&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4200
      MouseIcon       =   "Parametros.frx":59A9F
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Inst_name 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   9
      Top             =   765
      Width           =   1455
   End
   Begin VB.Label Inst_name 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Label Inst_name 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   7
      Top             =   1740
      Width           =   1455
   End
   Begin VB.Label Inst_name 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   6
      Top             =   2205
      Width           =   1455
   End
   Begin VB.Label Inst_name 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   5
      Top             =   2685
      Width           =   1455
   End
   Begin VB.Label Inst_name 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   4
      Top             =   3165
      Width           =   1455
   End
   Begin VB.Label Inst_name 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   3
      Top             =   3645
      Width           =   1455
   End
   Begin VB.Label Inst_name 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   2
      Top             =   4125
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   8
      Left            =   4000
      TabIndex        =   1
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BANK"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      Top             =   3840
      Width           =   435
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Mouse_sobre As Integer

Private Sub Form_Load()

Dim NormalWindowStyle As Long
Dim ret As Long
Dim col As Long

NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
SetLayeredWindowAttributes Me.hwnd, 0, 50, LWA_ALPHA

ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
ret = ret Or WS_EX_LAYERED
SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
col = RGB(0, 0, 0)
SetLayeredWindowAttributes Me.hwnd, col, 50, LWA_COLORKEY

Form1.Show , Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MidiClose
End Sub

Private Sub Form_Resize()
Move Form1.left + Form1.Width - 1300
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Cor_atual = Shape1(Index).FillColor
Debug.Print Cor_atual
End Sub



Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3_MouseMove 8, Button, Shift, X, Y
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6_MouseMove Button, Shift, X, Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static direção As Integer




If Button Then

    xx = (Label1 + Sgn(direção - Y))
    If xx < 40 Then xx = 40
    If xx > 240 Then xx = 240
Image5.Picture = Image4(xx Mod 4).Picture
    Label1 = xx
    Form1.Timer1.Interval = 60000 / ((2 ^ Divisão) * Label1)
End If
direção = Y
End Sub

Private Sub Label2_Click()
Unload Form4
Unload Form3
Unload Form4
Unload Me
End Sub


Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Static direção As Integer, delay As Integer, roda As Integer

If Button Then
    If delay <> 10 Then
        delay = delay + 1
        Exit Sub
        Else
        delay = 0
    End If
    
    
roda = roda + Sgn(direção - Y)
If roda = 7 Then roda = 0
If roda < 0 Then roda = 7

Image6.Picture = Image4(roda Mod 4).Picture

volte:
    xx = (Label3(Index) + Sgn(direção - Y)) And 127
    Label3(Index).Caption = xx

If Bank_util(xx) = False And Index = 8 Then
GoTo volte
End If


    direção = Y
    
    If Index < 8 Then
            Inst_name(Index) = Drum_Name(Label3(Index), Label3(8))
    Else
            For i = 0 To 7
                Inst_name(i) = Drum_Name(Label3(i), Label3(8))
             Next
             ShortMessage &HC9, Label3(8), 0
    End If
End If





End Sub

Private Sub Label5_Click(Index As Integer)
Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String
Dim sSave As SelectedFile
    
    On Error GoTo e_trap
 
 If Index = 0 Then
    FileDialog.sFilter = "Drum Psttern(*.Drm)" & Chr$(0) & "*.drm" & Chr$(0)
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
    FileDialog.sDlgTitle = "Open File"
    FileDialog.sInitDir = App.Path & "\Patterns"
    sOpen = ShowOpen(Me.hwnd, True)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        FileList = sOpen.sLastDirectory
        For Count = 1 To sOpen.nFilesSelected
            FileList = FileList & sOpen.sFiles(Count)
        Next Count
        arquivo = FileList
        If arquivo <> "" Then Open_arquivo (arquivo)
    End If

Else
    
    FileDialog.sFilter = "Drum Pattern(*.Drm)" & Chr$(0) & "*.Drm" & Chr$(0)
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.sTemplateName = "Sem Titulo"
    FileDialog.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
    FileDialog.sDlgTitle = "Save Map"
    FileDialog.sInitDir = App.Path & "\Sounds Map"
    FileDialog.sDefFileExt = ".Map"
    sSave = ShowSave(Me.hwnd, True)
    If Err.Number <> 32755 And sSave.bCanceled = False Then
        FileList = sSave.sLastDirectory
        For Count = 1 To sSave.nFilesSelected
            FileList = FileList & sSave.sFiles(Count)
        Next Count
        arquivo = FileList
        If arquivo <> "" Then Save_arquivo (arquivo)
    End If
    
End If
sair:
    Exit Sub
e_trap:
    
    Resume sair
End Sub

Private Sub Label5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Label5(Index).ForeColor = &HC0FFFF
Label5(Index Xor 1).ForeColor = 1
Mouse_sobre = Index
Timer1.Enabled = True

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static direção As Integer, delay As Integer, xx As Integer

If Button Then
    If delay <> 10 Then
        delay = delay + 1
        Exit Sub
        Else
        delay = 0
    End If
    xx = (Divisão + Sgn(direção - Y))
    If xx > 4 Then xx = 4
    If xx < 0 Then xx = 0
    Divisão = xx
    Image7.Picture = Image4(xx Mod 4).Picture
    Image1.Picture = Image3(Divisão)
    Form1.Timer1.Interval = 60000 / ((2 ^ Divisão) * Label1)
End If
direção = Y
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

PopupMenu Form4.Menu

End Sub

Private Sub Timer1_Timer()
Label5(Mouse_sobre).ForeColor = 1
End Sub

Private Sub Open_arquivo(arquivo)
Dim X As Long, Free As Integer
Free = FreeFile

Open arquivo For Binary As Free
    Get #1, 1, Dados 'Dados em cores
    
    Get #1, , X
    If X Then
        Form1.Label2_DblClick (0)
    Else
        Form1.Label2_DblClick (1)
    End If
    
    For i = 0 To 8
        Get #1, , X
        Label3(i) = X 'Instrumentos e Bank
    Next
    Get #1, , X
    
    Label1 = X 'BMP
    
    For i = 0 To 7
        Inst_name(i) = Drum_Name(Label3(i), Label3(8))
    Next
    ShortMessage &HC9, Label3(8), 0
    
    Get #1, , Divisão
    Image7.Picture = Image4(xx Mod 4).Picture
    Image1.Picture = Image3(Divisão)
    
    Form1.Timer1.Interval = 60000 / ((2 ^ Divisão) * Label1)
        
Close Free
If Form1.Shape1(0).FillColor = &HFF Then
Form1.Pintar44
Else
Form1.Pintar34
End If
End Sub

Private Sub Save_arquivo(arquivo As String)
Dim X As Long, Free As Integer
Free = FreeFile
If Form1.Shape1(0).FillColor = &HFF Then
    Form1.capture44
Else
    Form1.capture34
End If

Open arquivo For Binary As Free
    Put #1, 1, Dados 'Dados em cores
    X = Form1.Shape1(0).FillColor = &HFF ' 'se 3/4 ou 4/4
    Put #1, , X
    For i = 0 To 8
        X = Label3(i) 'Instrumentos e Bank
        Put #1, , X
    Next
    X = Label1 'BMP
    Put #1, , X
    Put #1, , Divisão
Close Free

End Sub

