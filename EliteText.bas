Attribute VB_Name = "EliteText"
'Api's
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)



Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETTEXT = &HC
Public Const WM_CHAR = &H102
Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2

Dim RTBox As Long
Dim RTBox2 As Long
Dim OldText As String


Function MSN_Text()
Dim ParentHwnd As Long
Dim TempText As String

ParentHwnd = FindWindow("IMWindowClass", vbNullString)

If ParentHwnd = 0 Then Exit Function

RTBox = FindWindowEx(ParentHwnd, 0&, "RichEdit20A", vbNullString)
RTBox2 = FindWindowEx(ParentHwnd, RTBox, "RichEdit20A", vbNullString)


If OldText = GetText(RTBox2) Then Exit Function

TempText = ConvertText(GetText(RTBox2)) '

Call SendMessageByString(RTBox2, WM_SETTEXT, 0&, "")
SendText TempText, RTBox2

For X = 1 To Len(TempText)
SendKeys "{Right}"
Next
End Function

Function ConvertText(Text As String) As String
Dim Lett$
Dim str As String

For a = 1 To Len(Text)
  Lett = LCase(Mid(Text, a, 1))
  Select Case Lett
  Case "a"
    X = Int(Rnd * 12) + 1
    Select Case X
    Case 1
      Lett = "@"
    Case 2
      Lett = "ª"
    Case 3
      Lett = "å"
    Case 4
      Lett = "Å"
    Case 5
      Lett = "ã"
    Case 6
      Lett = "Ã"
    Case 7
      Lett = "â"
    Case 8
      Lett = "Â"
    Case 9
      Lett = "á"
    Case 10
      Lett = "Á"
    Case 11
      Lett = "à"
    Case 12
      Lett = "À"
    End Select
  Case "b"
    Lett = "ß"
  Case "c"
    X = Int(Rnd * 4) + 1
    Select Case X
    Case 1
      Lett = "Ç"
    Case 2
      Lett = "ç"
    Case 3
      Lett = "¢"
    Case 4
      Lett = "©"
    End Select
  Case "d"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "Ð"
    Case 2
      Lett = "ð"
    End Select
  Case "e"
    X = Int(Rnd * 9) + 1
    Select Case X
    Case 1
      Lett = "£"
    Case 2
      Lett = "Ë"
    Case 3
      Lett = "Ê"
    Case 4
      Lett = "É"
    Case 5
      Lett = "È"
    Case 6
      Lett = "è"
    Case 7
      Lett = "é"
    Case 8
      Lett = "ê"
    Case 9
      Lett = "ë"
    End Select
  Case "f"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "F"
    Case 2
      Lett = "f"
    End Select
  Case "g"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "G"
    Case 2
      Lett = "g"
    End Select
  Case "h"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "H"
    Case 2
      Lett = "h"
    End Select
  Case "i"
    X = Int(Rnd * 10) + 1
    Select Case X
    Case 1
      Lett = "|"
    Case 2
      Lett = "¦"
    Case 3
      Lett = "Ï"
    Case 4
      Lett = "Î"
    Case 5
      Lett = "Í"
    Case 6
      Lett = "Ì"
    Case 7
      Lett = "ì"
    Case 8
      Lett = "í"
    Case 9
      Lett = "î"
    Case 10
      Lett = "ï"
    End Select
  Case "j"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "J"
    Case 2
      Lett = "j"
    End Select
  Case "k"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "K"
    Case 2
      Lett = "k"
    End Select
  Case "l"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "L"
    Case 2
      Lett = "l"
    End Select
  Case "m"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "M"
    Case 2
      Lett = "m"
    End Select
  Case "n"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "Ñ"
    Case 2
      Lett = "ñ"
    End Select
  Case "o"
    X = Int(Rnd * 13) + 1
    Select Case X
    Case 1
      Lett = "0"
    Case 2
      Lett = "ø"
    Case 3
      Lett = "Ø"
    Case 4
      Lett = "º"
    Case 5
      Lett = "°"
    Case 6
      Lett = "õ"
    Case 7
      Lett = "Õ"
    Case 8
      Lett = "Ô"
    Case 9
      Lett = "ô"
    Case 10
      Lett = "ó"
    Case 11
      Lett = "Ó"
    Case 12
      Lett = "Ò"
    Case 13
      Lett = "ò"
    End Select
  Case "p"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "þ"
    Case 2
      Lett = "Þ"
    End Select
  Case "q"
    Lett = "¶"
  Case "r"
    Lett = "®"
  Case "s"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "$"
    Case 2
      Lett = "§"
    End Select
  Case "t"
    Lett = "+"
  Case "u"
    X = Int(Rnd * 7) + 1
    Select Case X
    Case 1
      Lett = "µ"
    Case 2
      Lett = "Ù"
    Case 3
      Lett = "ù"
    Case 4
      Lett = "ú"
    Case 5
      Lett = "Ú"
    Case 6
      Lett = "Û"
    Case 7
      Lett = "û"
    End Select
  Case "v"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "V"
    Case 2
      Lett = "v"
    End Select
  Case "w"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "W"
    Case 2
      Lett = "w"
    End Select
  Case "x"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "X"
    Case 2
      Lett = "x"
    End Select
  Case "y"
    X = Int(Rnd * 4) + 1
    Select Case X
    Case 1
      Lett = "Ý"
    Case 2
      Lett = "ÿ"
    Case 3
      Lett = "ý"
    Case 4
      Lett = "¥"
    End Select
  Case "z"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "Z"
    Case 2
      Lett = "z"
    End Select
  Case "1"
    Lett = "¹"
  Case "2"
    Lett = "²"
  Case "3"
    Lett = "³"
  Case "!"
    Lett = "¡"
  Case "?"
    Lett = "¿"
  Case "-"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "¬"
    Case 2
      Lett = "·"
    End Select
  Case "("
    X = Int(Rnd * 3) + 1
    Select Case X
    Case 1
      Lett = "{"
    Case 2
      Lett = "["
    Case 3
      Lett = "<"
    End Select
  Case ")"
    X = Int(Rnd * 3) + 1
    Select Case X
    Case 1
      Lett = "]"
    Case 2
      Lett = "}"
    Case 3
      Lett = ">"
    End Select
  Case "<"
    Lett = "«"
  Case ">"
    Lett = "»"
  Case ","
    Lett = "¸"
  Case "0"
    X = Int(Rnd * 2) + 1
    Select Case X
    Case 1
      Lett = "O"
    Case 2
      Lett = "o"
    End Select
  Case Else
    Lett = Lett
End Select
str = str & Lett
Next a
ConvertText = str
End Function

Public Function GetText(ByRef WindowHwnd As Long) As String
Dim Buffer As String, TextLength As Long

TextLength = SendMessage(WindowHwnd, WM_GETTEXTLENGTH, 0&, 0&)


Buffer = String(TextLength, 0&)


Call SendMessageByString(WindowHwnd, WM_GETTEXT, TextLength + 1, Buffer)


GetText = Buffer
End Function


Public Function SendText(Text As String, hwnd As Long)

Call SendMessageByString(hwnd, WM_SETTEXT, 0&, Text$)

Call SendMessageByNum(hwnd, WM_CHAR, 13, 0&)
End Function



