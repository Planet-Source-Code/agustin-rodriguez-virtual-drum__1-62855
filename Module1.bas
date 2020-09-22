Attribute VB_Name = "Module1"
Option Explicit
Public Drum_Name(127, 127) As String
Public Bank_util(0 To 127) As Integer
Public Dados(256) As Long
Public Divisão As Integer
Public ag As Double
Public xt As Double
Public yt As Double
Public CentroX  As Double
Public CentroY As Double
Public Raio As Double
    
Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Public Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Public Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Public Declare Function MIDIOutOpen Lib "winmm.dll" Alias "midiOutOpen" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

Public ang As Single

Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_SHOWWINDOW As Long = &H40

Public Declare Function apiSetWindowPos Lib "User32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST As Integer = -1
Public Const HWND_NOTOPMOST As Integer = -2
Public Const SWP_NOMOVE As Integer = &H2
Public Const SWP_NOSIZE As Integer = &H1

Public Const MAXPNAMELEN As Integer = 32

Public Type MIDIOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    wTechnology As Integer
    wVoices As Integer
    wNotes As Integer
    wChannelMask As Integer
    dwSupport As Long
End Type

Public MidiCaps As MIDIOUTCAPS
Public hMidi As Long
Public Dev_OUT As Long
Public rc As Long
Public midimsg As Long
Public IsOpen As Integer

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal i As Long, ByVal i As Long, ByVal w As Long, ByVal i As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const LWA_ALPHA As Long = &H2&
Public Const LWA_COLORKEY As Integer = &H1

Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Radiano As Single
Public Type Circulo
    tamanho As Integer
    cor As Long
    Raio As Integer
End Type

Public Cor_atual As Long

Public Obj(0 To 8) As Circulo
Public Const FLOODFILLSURFACE As Long = 1

Public Sub ShortMessage(status As Integer, dado1 As Integer, dado2 As Integer)

  Dim X As String, Valor As Long

    X = "&h" + Right$("00" & Hex$(dado2), 2) + Right$("00" & Hex$(dado1), 2) + Right$("00" & Hex$(status), 2)
    Valor = Val(X)
    midiOutShortMsg hMidi, Valor

End Sub

Public Sub MidiOpen()

    rc = MIDIOutOpen(hMidi, Dev_OUT, 0, 0, 0)
    If rc <> 0 Then
        MsgBox "Open MIDI Out failed"
    End If
    If rc = 0 Then
        IsOpen = True
    End If

End Sub

Public Sub MidiClose()

    If IsOpen = False Then
        Exit Sub
    End If
        
    rc = midiOutClose(hMidi)
    If rc <> 0 Then
        MsgBox "Close MIDI Out failed"
    End If
        
    If rc = 0 Then
        IsOpen = False
    End If

End Sub

Public Sub SetOnTop(Frm As Form, OnTop As Long)

    If OnTop = -1 Then
     
        apiSetWindowPos Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
      Else
        apiSetWindowPos Frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    End If

End Sub


