Attribute VB_Name = "TextLimit"


Option Explicit
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Const EM_LIMITTEXT = &HC5
Dim Imwindow
Dim RTB1
Dim RTB2
Dim Xhwnd

Public Sub Change_limit()
Imwindow = FindWindow("IMWindowClass", vbNullString)
RTB1 = FindWindowEx(Imwindow, 0, "RichEdit20A", vbNullString)
If RTB1 = 0 Then RTB1 = FindWindowEx(Imwindow, 0, "RichEdit20W", vbNullString)
RTB2 = FindWindowEx(Imwindow, RTB1, "RichEdit20A", vbNullString)
If RTB2 = 0 Then RTB2 = FindWindowEx(Imwindow, RTB1, "RichEdit20W", vbNullString)
Xhwnd = RTB2
SendMessage Xhwnd, EM_LIMITTEXT, 1200, 0
End Sub


