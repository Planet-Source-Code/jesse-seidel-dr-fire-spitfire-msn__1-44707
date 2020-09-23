Attribute VB_Name = "SurpressWindow"
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Function SupressWindow()
Dim hiddenwindowclass As Long
Dim msblpopupmsgwclass As Long
hiddenwindowclass& = FindWindow("hiddenwindowclass", vbNullString)
msblpopupmsgwclass& = FindWindow("msblpopupmsgwclass", vbNullString)
Call ShowWindow(msblpopupmsgwclass&, SW_HIDE)
End Function



