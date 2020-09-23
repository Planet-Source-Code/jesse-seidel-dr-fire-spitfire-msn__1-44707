VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3135
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1200
      Top             =   2040
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Made By: Jesse Seidel"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   4560
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading... Please wait"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Line Line3 
      X1              =   4560
      X2              =   4560
      Y1              =   3120
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4560
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3120
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   840
      Picture         =   "frmSplash.frx":15F942
      Top             =   120
      Width           =   2955
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
frmContacts.Visible = True
End Sub
