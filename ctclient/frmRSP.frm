VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRSP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rock Scissors Paper - Client"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   Icon            =   "frmRSP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Shoot!!"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3015
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1140
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Make a Selection"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton Option3 
         Caption         =   "Paper"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Scissors"
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rock"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then rsp = "rock"
If Option2.Value = True Then rsp = "scissors"
If Option3.Value = True Then rsp = "paper"

frmClient.client.SendData "101RSP" + rsp
Command1.Enabled = False
Frame1.Enabled = False
StatusBar1.SimpleText = "Waiting for " & usernamec & "."

If Len(rsp2) > 0 Then Call checkwin

End Sub

Private Sub Form_Load()
rsp = ""
rsp2 = ""
StatusBar1.SimpleText = "Waiting for " & usernamec & "."
End Sub

Public Sub checkwin()
If rsp = "rock" And rps2 = "rock" Then
    StatusBar1.SimpleText = "Rock vs Rock - No Go!"
ElseIf rsp = "rock" And rsp2 = "scissors" Then
    StatusBar1.SimpleText = "Rock vs Scissors - You Win!"
ElseIf rsp = "rock" And rsp2 = "paper" Then
    StatusBar1.SimpleText = "Rock vs Paper - You Lose!"
ElseIf rsp = "scissors" And rsp2 = "rock" Then
    StatusBar1.SimpleText = "Scissors vs Rock - You Lose!"
ElseIf rsp = "scissors" And rsp2 = "scissors" Then
    StatusBar1.SimpleText = "Scissors vs Scissors - No Go!"
ElseIf rsp = "scissors" And rsp2 = "paper" Then
    StatusBar1.SimpleText = "Scissors vs Paper - You Win!"
ElseIf rsp = "paper" And rsp2 = "rock" Then
    StatusBar1.SimpleText = "Paper vs Rock - You Win!"
ElseIf rsp = "paper" And rsp2 = "scissors" Then
    StatusBar1.SimpleText = "Paper vs Scissors - You Lose!"
ElseIf rsp = "paper" And rsp2 = "paper" Then
    StatusBar1.SimpleText = "Paper vs Paper - No Go!"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClient.client.SendData "101RSPCLOSE"
End Sub
