Attribute VB_Name = "TransparenteMod"
Option Explicit

#If VBA7 Then
'// 64 Bits
'// Declara��es DLL para alterar ou apar�ncia do UserForm
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

#Else
'// 32 Bits
'// Declara��es DLL para alterar ou apar�ncia do UserForm
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

#End If

'// Constantes windows para barra de t�tulo
Private Const GWL_STYLE As Long = (-16)           '//The offset of a window's style
Private Const GWL_EXSTYLE As Long = (-20)         '//The offset of a window's extended style
Private Const WS_CAPTION As Long = &HC00000       '//Style to add a titlebar
Private Const WS_EX_DLGMODALFRAME As Long = &H1   '//Controls if the window has an icon
 
'// Constantes windows para transpar�ncia
Private Const WS_EX_LAYERED = &H80000             '//cor
Private Const LWA_COLORKEY = &H1                  '//Chroma key for fading a certain color on your Form
Private Const LWA_ALPHA = &H2                     '//Only needed if you want to fade the entire userform


Function HideTitleBarAndBordar(frm As Object)

'// Ocultar barra de t�tulo e borda em torno do formul�rio
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindow(vbNullString, frm.Caption)
'// Build window and set window until you remove the caption, title bar and frame around the window
'// Cria a janela e define a janela at� remover a legenda, a barra de t�tulo e o quadro ao redor da janela
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl

End Function

Function MakeUserformTransparent(frm As Object, Optional Color As Variant)

'//set transparencies on userform***********************************
Dim formhandle As Long
Dim bytOpacity As Byte

formhandle = FindWindow(vbNullString, frm.Caption)
If IsMissing(Color) Then Color = RGB(107, 23, 18)      '//rgbWhite
bytOpacity = 0

SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED

frm.BackColor = Color
SetLayeredWindowAttributes formhandle, Color, bytOpacity, LWA_COLORKEY

End Function
