VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTTTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe - Server"
   ClientHeight    =   3030
   ClientLeft      =   3090
   ClientTop       =   2025
   ClientWidth     =   2415
   Icon            =   "frmTTTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   2415
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2175
      Begin VB.PictureBox Box1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Box2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Box3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Box4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Box5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Box6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Box7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   5
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox Box8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   4
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox Box9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   3
         Top             =   1680
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1440
         X2              =   1440
         Y1              =   240
         Y2              =   2160
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   720
         X2              =   720
         Y1              =   240
         Y2              =   2160
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   0
         X2              =   2160
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   0
         X2              =   2160
         Y1              =   840
         Y2              =   840
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2775
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Wating..."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "You are O's"
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
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmTTTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Box1_Click()
frmServer.server.SendData "TTTBOX1"
box1s = frmServer.txtUsername.Text
Box1.Picture = LoadPicture(App.Path & "\O.jpg")
Box1.Enabled = False
Frame1.Enabled = False
StatusBar1.SimpleText = "Waiting for " & usernames & "."
allboxs = allboxs + 1
calcwin
End Sub

Private Sub Box2_Click()
frmServer.server.SendData "TTTBOX2"
box2s = frmServer.txtUsername.Text
Box2.Picture = LoadPicture(App.Path & "\O.jpg")
Box2.Enabled = False
Frame1.Enabled = False
StatusBar1.SimpleText = "Waiting for " & usernames & "."
allboxs = allboxs + 1
calcwin
End Sub

Private Sub Box3_Click()
frmServer.server.SendData "TTTBOX3"
box3s = frmServer.txtUsername.Text
Box3.Picture = LoadPicture(App.Path & "\O.jpg")
Box3.Enabled = False
Frame1.Enabled = False
StatusBar1.SimpleText = "Waiting for " & usernames & "."
allboxs = allboxs + 1
calcwin
End Sub

Private Sub Box4_Click()
frmServer.server.SendData "TTTBOX4"
box4s = frmServer.txtUsername.Text
Box4.Picture = LoadPicture(App.Path & "\O.jpg")
Box4.Enabled = False
Frame1.Enabled = False
StatusBar1.SimpleText = "Waiting for " & usernames & "."
allboxs = allboxs + 1
calcwin
End Sub

Private Sub Box5_Click()
frmServer.server.SendData "TTTBOX5"
box5s = frmServer.txtUsername.Text
Box5.Picture = LoadPicture(App.Path & "\O.jpg")
Box5.Enabled = False
Frame1.Enabled = False
StatusBar1.SimpleText = "Waiting for " & usernames & "."
allboxs = allboxs + 1
calcwin
End Sub

Private Sub Box6_Click()
frmServer.server.SendData "TTTBOX6"
box6s = frmServer.txtUsername.Text
Box6.Picture = LoadPicture(App.Path & "\O.jpg")
Box6.Enabled = False
Frame1.Enabled = False
StatusBar1.SimpleText = "Waiting for " & usernames & "."
allboxs = allboxs + 1
calcwin
End Sub

Private Sub Box7_Click()
frmServer.server.SendData "TTTBOX7"
box7s = frmServer.txtUsername.Text
Box7.Picture = LoadPicture(App.Path & "\O.jpg")
Box7.Enabled = False
Frame1.Enabled = False
StatusBar1.SimpleText = "Waiting for " & usernames & "."
allboxs = allboxs + 1
calcwin
End Sub

Private Sub Box8_Click()
frmServer.server.SendData "TTTBOX8"
box8s = frmServer.txtUsername.Text
Box8.Picture = LoadPicture(App.Path & "\O.jpg")
Box8.Enabled = False
Frame1.Enabled = False
StatusBar1.SimpleText = "Waiting for " & usernames & "."
allboxs = allboxs + 1
calcwin
End Sub

Private Sub Box9_Click()
frmServer.server.SendData "TTTBOX9"
box9s = frmServer.txtUsername.Text
Box9.Picture = LoadPicture(App.Path & "\O.jpg")
Box9.Enabled = False
Frame1.Enabled = False
StatusBar1.SimpleText = "Waiting for " & usernames & "."
allboxs = allboxs + 1
calcwin
End Sub

Public Sub calcwin()
If box1s = frmServer.txtUsername.Text And box2s = frmServer.txtUsername.Text And box3s = frmServer.txtUsername.Text Then
    StatusBar1.SimpleText = "You win!"
    Frame1.Enabled = False
ElseIf box4s = frmServer.txtUsername.Text And box5s = frmServer.txtUsername.Text And box6s = frmServer.txtUsername.Text Then
    StatusBar1.SimpleText = "You win!"
    Frame1.Enabled = False
ElseIf box7s = frmServer.txtUsername.Text And box8s = frmServer.txtUsername.Text And box9s = frmServer.txtUsername.Text Then
    StatusBar1.SimpleText = "You win!"
    Frame1.Enabled = False
ElseIf box1s = frmServer.txtUsername.Text And box4s = frmServer.txtUsername.Text And box7s = frmServer.txtUsername.Text Then
    StatusBar1.SimpleText = "You win!"
    Frame1.Enabled = False
ElseIf box2s = frmServer.txtUsername.Text And box5s = frmServer.txtUsername.Text And box8s = frmServer.txtUsername.Text Then
    StatusBar1.SimpleText = "You win!"
    Frame1.Enabled = False
ElseIf box3s = frmServer.txtUsername.Text And box6s = frmServer.txtUsername.Text And box9s = frmServer.txtUsername.Text Then
    StatusBar1.SimpleText = "You win!"
    Frame1.Enabled = False
ElseIf box1s = frmServer.txtUsername.Text And box5s = frmServer.txtUsername.Text And box9s = frmServer.txtUsername.Text Then
    StatusBar1.SimpleText = "You win!"
    Frame1.Enabled = False
ElseIf box3s = frmServer.txtUsername.Text And box5s = frmServer.txtUsername.Text And box7s = frmServer.txtUsername.Text Then
    StatusBar1.SimpleText = "You win!"
    Frame1.Enabled = False
ElseIf box1s = usernames And box2s = usernames And box3s = usernames Then
    StatusBar1.SimpleText = "You lose!"
    Frame1.Enabled = False
ElseIf box4s = usernames And box5s = usernames And box6s = usernames Then
    StatusBar1.SimpleText = "You lose!"
    Frame1.Enabled = False
ElseIf box7s = usernames And box8s = usernames And box9s = usernames Then
    StatusBar1.SimpleText = "You lose!"
    Frame1.Enabled = False
ElseIf box1s = usernames And box4s = usernames And box7s = usernames Then
    StatusBar1.SimpleText = "You lose!"
    Frame1.Enabled = False
ElseIf box2s = usernames And box5s = usernames And box8s = usernames Then
    StatusBar1.SimpleText = "You lose!"
    Frame1.Enabled = False
ElseIf box3s = usernames And box6s = usernames And box9s = usernames Then
    StatusBar1.SimpleText = "You lose!"
    Frame1.Enabled = False
ElseIf box1s = usernames And box5s = usernames And box9s = usernames Then
    StatusBar1.SimpleText = "You lose!"
    Frame1.Enabled = False
ElseIf box3s = usernames And box5s = usernames And box7s = usernames Then
    StatusBar1.SimpleText = "You lose!"
    Frame1.Enabled = False
ElseIf allboxs = 9 Then
    StatusBar1.SimpleText = "It's a tie!"
    Frame1.Enabled = False
End If
End Sub

Private Sub Form_Load()
Box1.Enabled = True
Box2.Enabled = True
Box3.Enabled = True
Box4.Enabled = True
Box5.Enabled = True
Box6.Enabled = True
Box7.Enabled = True
Box8.Enabled = True
Box9.Enabled = True
box1s = ""
box2s = ""
box3s = ""
box4s = ""
box5s = ""
box6s = ""
box7s = ""
box8s = ""
box9s = ""
allboxs = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmServer.server.SendData "TTTCLOSE"
End Sub
