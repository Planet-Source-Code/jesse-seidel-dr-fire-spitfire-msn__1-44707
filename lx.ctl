VERSION 5.00
Begin VB.UserControl lx 
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   ScaleHeight     =   2580
   ScaleWidth      =   3030
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Image xpush 
      Height          =   255
      Left            =   1800
      Picture         =   "lx.ctx":0000
      Top             =   1560
      Width           =   495
   End
   Begin VB.Image xgray 
      Height          =   255
      Left            =   1080
      Picture         =   "lx.ctx":06E6
      Top             =   1560
      Width           =   495
   End
   Begin VB.Image minpush 
      Height          =   255
      Left            =   240
      Picture         =   "lx.ctx":0DCC
      Top             =   1560
      Width           =   495
   End
   Begin VB.Image mingray 
      Height          =   255
      Left            =   840
      Picture         =   "lx.ctx":14B2
      Top             =   960
      Width           =   495
   End
   Begin VB.Image controls1 
      Height          =   255
      Left            =   120
      Picture         =   "lx.ctx":1B98
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "lx.ctx":227E
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "lx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = minpush.Picture
frmMin.Show
frmMin.Timer1.Enabled = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = mingray.Picture
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = controls1.Picture
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = xgray.Picture
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = controls1.Picture
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
End
End Sub

Private Sub UserControl_Initialize()
UserControl.Height = "255"
UserControl.Width = "495"
End Sub

Private Sub UserControl_Resize()
UserControl.Height = "255"
UserControl.Width = "495"
End Sub
