VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2865
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   600
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   1080
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(Test Message)"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   180
         Left            =   1680
         MouseIcon       =   "frmAlert.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmAlert.frx":030A
         Top             =   120
         Width           =   195
      End
      Begin VB.Label lblAlert 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Alert Message"
         Height          =   735
         Left            =   90
         TabIndex        =   1
         Top             =   600
         Width           =   1815
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer tmrAlert 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   120
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Const SM_CXFULLSCREEN = 16
Const SM_CYFULLSCREEN = 17
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10


Private ClsGradient As New CGradient
Private fX As Long
Private fY As Long
Private lngScaleX As Long
Private lngScaleY As Long
Private AlertIndex As Long


Private Sub Image1_Click()
Me.Hide
End Sub


Private Sub lblAlert_Click()
    MsgBox "This is just a demo"
End Sub

Private Sub lblAlert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblAlert.FontUnderline = False Then
        lblAlert.FontUnderline = True
        lblAlert.ForeColor = RGB(0, 0, 255)
    End If
    lblAlert.BorderStyle = 1
End Sub

Private Sub picBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lblAlert.FontUnderline = True Then
        lblAlert.FontUnderline = False
        lblAlert.ForeColor = &H0
    End If
    lblAlert.BorderStyle = 0
End Sub

Private Sub tmrAlert_Timer()
    tmrAlert.Enabled = False
    tmrClose.Enabled = True
End Sub

Private Sub tmrClose_Timer()
    Dim curHeight As Long
    curHeight = Me.Height
    If curHeight > 120 Then
        Me.Height = curHeight - 30
        Me.Top = Me.Top + 30
    Else
        If AlertCount = AlertIndex Then AlertCount = 0
        Unload Me
    End If
End Sub

Private Sub tmrOpen_Timer()
    Dim curHeight As Long
    Dim newHeight As Long
    curHeight = Me.Height
    If curHeight < picBackground.Height + lngScaleY Then
        newHeight = curHeight + 30
        If newHeight > picBackground.Height + lngScaleY Then newHeight = picBackground.Height + lngScaleY
        Me.Height = Me.Height + (newHeight - curHeight)
        Me.Top = Me.Top - (newHeight - curHeight)
    Else
        tmrOpen.Enabled = False
        tmrAlert.Enabled = True
    End If
End Sub

Public Sub DisplayAlert(MessageText As String, Duration As Long)

    Dim wFlags As Long, X As Long

    AlertCount = AlertCount + 1
    AlertIndex = AlertCount

    lblAlert.Caption = MessageText

    tmrAlert.Interval = Duration

    fX = GetSystemMetrics(SM_CXFULLSCREEN)
    fY = GetSystemMetrics(SM_CYFULLSCREEN)
    lngScaleX = Me.Width - Me.ScaleWidth
    lngScaleY = Me.Height - Me.ScaleHeight
    
    Me.Height = 90
    Me.Width = picBackground.Width + lngScaleX
    Me.Left = fX * Screen.TwipsPerPixelX - Me.Width
    Me.Top = (fY * Screen.TwipsPerPixelY) - ((picBackground.Height + lngScaleY) * (AlertCount - 1)) + 300
    Me.Show
    
    wFlags = SND_ASYNC Or SND_NODEFAULT

    With ClsGradient
        .Angle = -100
        .Color1 = RGB(61, 149, 255)
        .Color2 = RGB(255, 255, 255)
        .Draw picBackground
    End With
    picBackground.Refresh

    tmrOpen.Enabled = True

End Sub
