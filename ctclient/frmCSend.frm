VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ChatThing - Client - Send File"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   3000
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3795
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   441
            MinWidth        =   441
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   441
            MinWidth        =   441
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "Enter IP"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4440
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "File Size -"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "% Complete"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "frmCSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
Command2.Enabled = True
Command1.Enabled = False
Command3.Enabled = True
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = 1002
Winsock1.Bind 1003
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
If File1.FileName = "" Then
    MsgBox "Select a file dummy!", vbOKOnly, "Duh!"
Else
    Winsock1.SendData "Filename" + File1.FileName & "; " & Label3.Caption
    Command2.Enabled = False
    File1.Enabled = False
End If

End Sub

Private Sub Command3_Click()
Command3.Enabled = False
Command1.Enabled = True
Winsock1.Close
StatusBar1.Panels(2).Text = ""
StatusBar1.Panels(3).Text = ""
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
On Error GoTo errorhandler
If Len(Dir1.Path) <= 3 Then Open Dir1.Path + File1.FileName For Binary Access Read As #1
If Len(Dir1.Path) > 3 Then Open Dir1.Path + "\" + File1.FileName For Binary Access Read As #1
Label3 = LOF(1) & " bytes"
errorhandler:
Close #1
End Sub

Private Sub Form_Load()
Text1.Text = frmClient.txtRemotehost.Text
File1.Path = Dir1.Path
frmClient.client.SendData "101SEND" + Winsock1.LocalIP
End Sub

Private Sub Timer1_Timer()
Winsock1.SendData "101CONNECT"
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim receivedata2 As String

Winsock1.GetData receivedata2

If receivedata2 = "DECLINE" Then
    MsgBox "File was declined", vbOKOnly, "No Go"
    Command2.Enabled = True
ElseIf receivedata2 = "101CONNECT" Then
    StatusBar1.Panels(3).Text = "X"
    Winsock1.SendData "101GOTIT"
ElseIf receivedata2 = "101GOTIT" Then
    Timer1.Enabled = False
    StatusBar1.Panels(2).Text = "X"
ElseIf receivedata2 = "ACCEPT" Then
    
'Error Trapping
If File1.FileName = "" Then msg = MsgBox("No Files Selected", 16, ""): Exit Sub

'If Dir1.path less than 3
If Len(Dir1.Path) <= 3 Then
Open Dir1.Path + File1.FileName For Binary Access Read As #1
Winsock1.SendData "OpenFile" + File1.FileName
End If

'if Dir1.path more than 3
If Len(Dir1.Path) > 3 Then
Open Dir1.Path + "\" + File1.FileName For Binary Access Read As #1
Winsock1.SendData "OpenFile" + File1.FileName
End If

'If File Is Smaller Than 2044
If LOF(1) <= 2044 Then
buffer = Space(LOF(1))
Get #1, , buffer
Winsock1.SendData "°°°°" + buffer
Winsock1.SendData "Close"
GoTo Closer
End If

'if File Is Greater Than 2044
If LOF(1) > 2044 Then
'Send In Chunks Of 2044
For SendChunk = 1 To LOF(1) / 2044
percent = SendChunk / (LOF(1) / 2044) * 100
Winsock1.SendData "PercentComplete" + Str(percent)
Label1.Caption = percent & "% Complete"
ProgressBar1.Value = percent
buffer = Space(2044)
Get #1, , buffer
Winsock1.SendData "°°°°" + buffer
Sleep 200
Next

'If There Is Any Left Over
If LOF(1) Mod 2044 <> 0 Then
buffer = Space(LOF(1) Mod 2044)
Get #1, , buffer
Winsock1.SendData "°°°°" + buffer
End If
Winsock1.SendData "Close"
GoTo Closer
End If

Exit Sub
Closer:
Close #1
Exit Sub
End If
End Sub

