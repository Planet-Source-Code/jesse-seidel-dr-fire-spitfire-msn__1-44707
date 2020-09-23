VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmContacts 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "SpitFire MSN"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   ForeColor       =   &H8000000E&
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   840
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   2040
      Top             =   960
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1080
      Top             =   1080
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   -2147483634
      TabCaption(0)   =   "Auto-Messages"
      TabPicture(0)   =   "main.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(1)=   "Line5"
      Tab(0).Control(2)=   "Line6"
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(4)=   "Check1"
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(6)=   "Check2"
      Tab(0).Control(7)=   "Text2"
      Tab(0).Control(8)=   "List1"
      Tab(0).Control(9)=   "Text5"
      Tab(0).Control(10)=   "Command7"
      Tab(0).Control(11)=   "Command8"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Popup Options"
      TabPicture(1)   =   "main.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(1)=   "Text3"
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(3)=   "Command3"
      Tab(1).Control(4)=   "Check3"
      Tab(1).Control(5)=   "Timer1"
      Tab(1).Control(6)=   "Text6"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Text Options"
      TabPicture(2)   =   "main.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Timer2"
      Tab(2).Control(1)=   "Check4"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "User Options"
      TabPicture(3)   =   "main.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(1)=   "Line4"
      Tab(3).Control(2)=   "Line7"
      Tab(3).Control(3)=   "Line8"
      Tab(3).Control(4)=   "Line9"
      Tab(3).Control(5)=   "Command1"
      Tab(3).Control(6)=   "Text4"
      Tab(3).Control(7)=   "Command6"
      Tab(3).Control(8)=   "Command9"
      Tab(3).Control(9)=   "Timer4"
      Tab(3).Control(10)=   "Check5"
      Tab(3).Control(11)=   "Timer6"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "Commands"
      TabPicture(4)   =   "main.frx":093A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label9"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label10"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label11"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label12"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label13"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Text7"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Text8"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Timer7"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "commandtxt"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Combo1"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "msg1"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).ControlCount=   11
      Begin VB.TextBox msg1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         TabIndex        =   41
         Text            =   "Message text goes here"
         Top             =   2640
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "main.frx":0956
         Left            =   960
         List            =   "main.frx":0963
         TabIndex        =   40
         Text            =   "--------------------------------------------------------"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox commandtxt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   38
         Text            =   "/(COMMAND)"
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Timer Timer7 
         Interval        =   1000
         Left            =   2160
         Top             =   1320
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Text            =   "Saturday, April 12-------------- I just added a whole new feature to my program, its where you can use commands!"
         Top             =   720
         Width           =   6495
      End
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -73800
         Top             =   2760
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Change max text limit to 1200 instead of 400"
         Height          =   255
         Left            =   -74760
         TabIndex        =   31
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -73920
         TabIndex        =   30
         Text            =   "Text6"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Timer Timer4 
         Interval        =   1
         Left            =   -69000
         Top             =   720
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Current Name"
         Height          =   255
         Left            =   -72480
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Change Name"
         Height          =   255
         Left            =   -71280
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72480
         TabIndex        =   27
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set name as now"
         Height          =   255
         Left            =   -74640
         TabIndex        =   25
         Top             =   840
         Width           =   1935
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -73680
         Top             =   2040
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Save &As"
         Height          =   255
         Left            =   -69240
         TabIndex        =   24
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Clear"
         Height          =   255
         Left            =   -71160
         TabIndex        =   23
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   2175
         Left            =   -71160
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   22
         Top             =   720
         Width           =   2775
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   -74880
         TabIndex        =   20
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74880
         TabIndex        =   18
         Text            =   "Please dont close my window"
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Dont allow people to close your windows"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74880
         TabIndex        =   16
         Text            =   "Im Currently Away From The Computer"
         Top             =   720
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Enable Away Message"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H8000000A&
         Caption         =   "Convert all the text I type into an Instant Message to Elite Text"
         Height          =   255
         Left            =   -74040
         TabIndex        =   14
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -69240
         Top             =   1200
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H8000000B&
         Caption         =   "Disable Popups"
         Height          =   255
         Left            =   -70920
         TabIndex        =   13
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Send &Real Popup"
         Height          =   255
         Left            =   -72720
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Show &Demo"
         Height          =   255
         Left            =   -74160
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74160
         TabIndex        =   9
         Text            =   "This is a test popup"
         Top             =   1560
         Width           =   5535
      End
      Begin VB.Label Label13 
         Caption         =   "What it does:"
         Height          =   255
         Left            =   960
         TabIndex        =   39
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Custom Command:"
         Height          =   255
         Left            =   720
         TabIndex        =   37
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "More commands coming soon..."
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
         Left            =   1200
         TabIndex        =   36
         Top             =   3240
         Width           =   4935
      End
      Begin VB.Label Label10 
         Caption         =   "/ip:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "/get-news:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   240
         Top             =   -2520
         Width           =   255
      End
      Begin VB.Line Line9 
         X1              =   -69960
         X2              =   -69960
         Y1              =   1200
         Y2              =   360
      End
      Begin VB.Line Line8 
         X1              =   -72600
         X2              =   -69960
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line7 
         X1              =   -72600
         X2              =   -72600
         Y1              =   1200
         Y2              =   360
      End
      Begin VB.Line Line4 
         X1              =   -74880
         X2              =   -72600
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Away Messages"
         Height          =   255
         Left            =   -71160
         TabIndex        =   21
         Top             =   480
         Width           =   2775
      End
      Begin VB.Line Line6 
         X1              =   -71280
         X2              =   -71280
         Y1              =   1080
         Y2              =   2880
      End
      Begin VB.Line Line5 
         X1              =   -74880
         X2              =   -71280
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "People who closed your window:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Message:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   1560
         Width           =   855
      End
   End
   Begin Project1.lx lx1 
      Height          =   255
      Left            =   6450
      TabIndex        =   7
      Top             =   10
      Width           =   495
      _extentx        =   873
      _extenty        =   450
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Display Name"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&E-mail"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.ListBox LstBuddylist 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6615
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   120
      Top             =   3960
      Width           =   975
   End
   Begin VB.Image statoffline 
      Height          =   240
      Left            =   3240
      Picture         =   "main.frx":09B5
      Top             =   360
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image statbusy 
      Height          =   240
      Left            =   2760
      Picture         =   "main.frx":0B39
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image statbrb 
      Height          =   240
      Left            =   2280
      Picture         =   "main.frx":0D93
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image statonline 
      Height          =   240
      Left            =   1800
      Picture         =   "main.frx":0FEF
      Top             =   360
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image statmain 
      Height          =   255
      Left            =   70
      Top             =   380
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SpitFire MSN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   4200
      Picture         =   "main.frx":13DA
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image Image17 
      Height          =   300
      Left            =   2400
      Picture         =   "main.frx":1439
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line3 
      X1              =   6960
      X2              =   6960
      Y1              =   6600
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6960
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   6600
   End
   Begin VB.Image Image16 
      Height          =   300
      Left            =   0
      Picture         =   "main.frx":1498
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msnobj As MsgrObject
Dim number As Integer
Dim Polo As Integer
Dim InitHeight As String
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim MsnApp As IMessengerApp
Dim WithEvents respond As MsgrObject
Attribute respond.VB_VarHelpID = -1
 Dim s As Integer
    Dim dta As String
Dim a As Integer
Dim msn As IMessenger
Dim X As String

Private Sub Check3_Click()
If Check3.Value = 1 Then Timer1.Enabled = True Else Timer1.Enabled = False
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then Timer2.Enabled = True Else Timer1.Enabled = False
End Sub

Private Sub Combo1_Change()
If Combo1.Text = "Send text" Then msg1.Enabled = True
If Combo1.Text = "Send special message from Jesse" Then msg1.Visible = False
End Sub

Private Sub Command1_Click()
msnobj.Services.PrimaryService.FriendlyName = msnobj.LocalFriendlyName & " - " & Now
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then Timer6.Enabled = True Else Timer6.Enabled = False
End Sub

Private Sub Command10_Click()
If List2.Text = "Online" Then
msnobj.LocalState = MSTATE_ONLINE
End If
End Sub

Private Sub Command2_Click()
Alert Text3.Text
End Sub

Private Sub Command4_Click()
Dim user As IMsgrUser
LstBuddylist.Clear
For Each user In msnobj.List(MLIST_CONTACT)
LstBuddylist.AddItem (user.EmailAddress)
Next
Label7.Caption = "User Count: " & LstBuddylist.ListCount
End Sub

Private Sub Command5_Click()
Dim user As IMsgrUser
LstBuddylist.Clear
For Each user In msnobj.List(MLIST_CONTACT)
LstBuddylist.AddItem (user.FriendlyName & " - " & user.EmailAddress)
Next
Label7.Caption = "User Count: " & LstBuddylist.ListCount
End Sub

Private Sub Command6_Click()
msnobj.Services.PrimaryService.FriendlyName = Text4.Text
End Sub

Private Sub Command7_Click()
Text5.Text = ""
End Sub

Private Sub Command9_Click()
Text4.Text = msnobj.LocalFriendlyName
End Sub

Private Sub Form_Load()
On Error GoTo err:
Set msnobj = New MsgrObject
Dim user As IMsgrUser
Set respond = New MsgrObject
 Set MsnApp = CreateObject("messenger.messengerapp")
Polo = True
InitHeight = Me.Height
Label2.Caption = " " & msnobj.LocalFriendlyName
For Each user In msnobj.List(MLIST_CONTACT)
LstBuddylist.AddItem (user.EmailAddress)
Next
Label7.Caption = "User Count: " & LstBuddylist.ListCount
Text4.Text = msnobj.LocalFriendlyName
Exit Sub
err:
MsgBox ("You are not signed-in to MSN Messenger." & vbCrLf & "Please sign-in and try again.")
End
End Sub

Private Sub Label1_DblClick()
If Polo = True Then
    Me.Height = 300
    Polo = False
Else
    Me.Height = InitHeight
    Polo = True
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF012&, 0
End Sub

Private Sub Label9_Click()

End Sub

Private Sub respond_OnTextReceived(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal pSourceUser As Messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)
If Check1.Value = 1 Then
pIMSession.SendText "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: EF=; CO=0000FF; CS=0; PF=12" & vbCrLf & vbCrLf, "" & Text1.Text, MMSGTYPE_NORESULT
Text5.Text = Text5.Text + pSourceUser.FriendlyName + " sayed: " + bstrMsgText
End If
If bstrMsgText = "/ip" Then
pIMSession.SendText "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: EF=; CO=0000FF; CS=0; PF=12" & vbCrLf & vbCrLf, "" & msnobj.LocalFriendlyName & " IP Address: " & Winsock1.LocalIP, MMSGTYPE_NORESULT
End If
If bstrMsgText = "/get-news" Then
pIMSession.SendText "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: EF=; CO=0000FF; CS=0; PF=12" & vbCrLf & vbCrLf, "" & "News: " & Text7.Text, MMSGTYPE_NORESULT
End If
If bstrMsgText = commandtxt.Text Then
If Combo1.Text = "Send text" Then pIMSession.SendText "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: EF=; CO=0000FF; CS=0; PF=12" & vbCrLf & vbCrLf, "" & msg1.Text, MMSGTYPE_NORESULT
If Combo1.Text = "Send special message from Jesse" Then pIMSession.SendText "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: EF=; CO=0000FF; CS=0; PF=12" & vbCrLf & vbCrLf, "" & msnobj.LocalFriendlyName & " is using SpitFire MSN, by Jesse Seidel. Ask him/her for the file and you can use it too!", MMSGTYPE_NORESULT
End If
End Sub

Private Sub respond_OnUserLeave(ByVal pIMsgrUser As Messenger.IMsgrUser, ByVal pIMSession As Messenger.IMsgrIMSession)
If Check2.Value = 1 Then
pIMsgrUser.SendText "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: EF=; CO=0000FF; CS=0; PF=12" & vbCrLf & vbCrLf, "" & Text2.Text, MMSGTYPE_NORESULT
List1.AddItem (pIMsgrUser.EmailAddress)
End If
End Sub

Private Sub Alert(Text As String)
    Dim AlertBox As frmAlert
    Set AlertBox = New frmAlert
    AlertBox.DisplayAlert Text, 3000
    Me.SetFocus
End Sub

Private Sub Timer1_Timer()
SupressWindow
End Sub

Private Sub Timer2_Timer()
If Check4.Value = 0 Then Exit Sub
MSN_Text
End Sub

Private Sub Timer3_Timer()
Set msnobj = New MsgrObject
Label2.Caption = " " & msnobj.LocalFriendlyName
    dta = Label2.Caption & Space$(41)
    s = s + 1
    Label2.Caption = Mid(dta, 1, s)
    If Len(Label2.Caption) >= 42 Then Label2.Caption = Right(Label2.Caption, 41)


    If s = Len(dta) Then
        Label2.Caption = ""
        s = 0
    End If
End Sub

Private Sub Timer5_Timer()
If msnobj.LocalState = MSTATE_ONLINE Then
    statmain.Picture = statonline.Picture
End If
If msnobj.LocalState = MSTATE_AWAY Then
    statmain.Picture = statbrb.Picture
End If
If msnobj.LocalState = MSTATE_BE_RIGHT_BACK Then
    statmain.Picture = statbrb.Picture
End If
If msnobj.LocalState = MSTATE_BUSY Then
    statmain.Picture = statbusy.Picture
End If
If msnobj.LocalState = MSTATE_IDLE Then
    statmain.Picture = statonline.Picture
End If
If msnobj.LocalState = MSTATE_INVISIBLE Then
    statmain.Picture = statoffline.Picture
End If
If msnobj.LocalState = MSTATE_OFFLINE Then
    statmain.Picture = statoffline.Picture
End If
If msnobj.LocalState = MSTATE_ON_THE_PHONE Then
    statmain.Picture = statbrb.Picture
End If
If msnobj.LocalState = MSTATE_OUT_TO_LUNCH Then
    statmain.Picture = statbrb.Picture
End If
End Sub

Private Sub Timer4_Timer()
Label4.Caption = Now
End Sub

Private Sub Timer6_Timer()
If Check5.Value = 0 Then Exit Sub
Change_limit
End Sub

Private Sub Timer7_Timer()
Text8.Text = Winsock1.LocalIP
End Sub
