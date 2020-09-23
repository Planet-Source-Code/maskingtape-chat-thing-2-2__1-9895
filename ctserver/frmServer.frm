VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Thing - Server"
   ClientHeight    =   4470
   ClientLeft      =   5115
   ClientTop       =   3330
   ClientWidth     =   4710
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4710
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   4215
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3176
            MinWidth        =   3176
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Object.ToolTipText     =   "Current Version"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "6:09 PM"
            Object.ToolTipText     =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "7/19/00"
            Object.ToolTipText     =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   3240
      MaxLength       =   15
      TabIndex        =   8
      Text            =   "Username"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show IP"
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock server 
      Left            =   0
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Settings"
      TabPicture(0)   =   "frmServer.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CommonDialog1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtLocalport"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Check1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Timer1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Timer3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Chat!"
      TabPicture(1)   =   "frmServer.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ProgressBar1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtChatwindow"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Timer2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtChat"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame2 
         Caption         =   "Other Options"
         Height          =   2175
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1575
         Begin VB.CheckBox Check2 
            Caption         =   "Scroll Title Bar"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Value           =   1  'Checked
            Width           =   1335
         End
      End
      Begin VB.Timer Timer3 
         Interval        =   500
         Left            =   2760
         Top             =   1920
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3240
         Width           =   975
      End
      Begin RichTextLib.RichTextBox txtChat 
         Height          =   300
         Left            =   -74880
         TabIndex        =   17
         Top             =   3000
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   529
         _Version        =   393217
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmServer.frx":047A
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   -74880
         Top             =   480
      End
      Begin RichTextLib.RichTextBox txtChatwindow 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4260
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmServer.frx":0528
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   3360
         Top             =   1920
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sound Enabled"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sound File"
         Height          =   615
         Left            =   1200
         TabIndex        =   11
         Top             =   2880
         Width           =   2895
         Begin VB.TextBox txtSoundfile 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Change"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Clear"
         Height          =   255
         Left            =   -72000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Clear The Chat Window"
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Stop Server"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Start Server"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Set Port"
         Height          =   255
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtLocalport 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Text            =   "1001"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         Height          =   255
         Left            =   -71880
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3360
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2280
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Wave Files | *.wav"
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   -74880
         TabIndex        =   15
         Top             =   3360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Max             =   5
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Your buddy is typing..."
         Height          =   255
         Left            =   -73800
         TabIndex        =   18
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu loadbrowser 
         Caption         =   "Load Browser"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu timestamp 
         Caption         =   "&Time Stamp"
      End
      Begin VB.Menu clear 
         Caption         =   "Clear"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu smallw 
         Caption         =   "Small Window"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu play 
      Caption         =   "Games"
      Enabled         =   0   'False
      Begin VB.Menu tttload 
         Caption         =   "Tic Tac Toe"
      End
      Begin VB.Menu rspload 
         Caption         =   "Rock Scissors Paper"
      End
   End
   Begin VB.Menu color 
      Caption         =   "&Color"
      Begin VB.Menu formback 
         Caption         =   "Form Background"
      End
      Begin VB.Menu tabback 
         Caption         =   "Tab Text"
      End
      Begin VB.Menu buttons 
         Caption         =   "Buttons"
         Begin VB.Menu allbuttons 
            Caption         =   "All Buttons"
         End
         Begin VB.Menu line2 
            Caption         =   "-"
         End
         Begin VB.Menu setportb 
            Caption         =   "'Set Port' "
         End
         Begin VB.Menu startserverb 
            Caption         =   "'Start Server' "
         End
         Begin VB.Menu stopserverb 
            Caption         =   "'Stop Server'"
         End
         Begin VB.Menu changesoundb 
            Caption         =   "'Change Sound'"
         End
         Begin VB.Menu clearb 
            Caption         =   "'Clear'"
         End
         Begin VB.Menu sendb 
            Caption         =   "'Send'"
         End
         Begin VB.Menu showipb 
            Caption         =   "'Show IP'"
         End
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu savecolor 
         Caption         =   "Save"
      End
      Begin VB.Menu loadcolor 
         Caption         =   "Load"
      End
   End
   Begin VB.Menu rightclick1 
      Caption         =   "rightclick "
      Visible         =   0   'False
      Begin VB.Menu smallw2 
         Caption         =   "Small Window  F2 "
      End
   End
   Begin VB.Menu rightclickchat 
      Caption         =   "rightclickchat"
      Visible         =   0   'False
      Begin VB.Menu copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu insert 
         Caption         =   "Insert"
         Begin VB.Menu insertip 
            Caption         =   "IP Address"
         End
         Begin VB.Menu htmllink 
            Caption         =   "HTML Link"
         End
         Begin VB.Menu lastthing 
            Caption         =   "Last thing you said"
         End
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub allbuttons_Click()
CommonDialog1.CancelError = True
 On Error GoTo errHandler
 CommonDialog1.Flags = &H1
 CommonDialog1.ShowColor
 Command1.BackColor = CommonDialog1.color
 Command2.BackColor = CommonDialog1.color
 Command3.BackColor = CommonDialog1.color
 Command4.BackColor = CommonDialog1.color
 Command5.BackColor = CommonDialog1.color
 Command6.BackColor = CommonDialog1.color
 Command7.BackColor = CommonDialog1.color
errHandler:

End Sub

Private Sub changesoundb_Click()
Command7.BackColor = thecolor()
End Sub

Private Sub Check2_Click()
If Check2.Value = Unchecked Then
    Timer3.Enabled = False
    Me.Caption = oldcaption
Else: Check2.Value = Checked
    Timer3.Enabled = True
End If
End Sub

Private Sub clear_Click()
txtChatwindow.Text = ""
End Sub

Private Sub clearb_Click()
Command6.BackColor = thecolor()
End Sub

Private Sub Command1_Click()
On Error GoTo errorhandler

If txtChat.Text = "" Then
Else
    lastsaid = txtChat.Text
    If nochats = 0 Then
    Else
    nochats = nochats - 1
    ProgressBar1.Value = nochats
    Timer2.Enabled = True
    txtChatwindow.SelStart = Len(txtChatwindow)
    txtChatwindow.SelColor = vbBlue
    txtChatwindow.SelBold = True
    server.SendData txtChat.Text
    txtChatwindow.SelText = txtChatwindow.SelText & vbCrLf & txtUsername & " - "
    txtChatwindow.SelBold = False
    txtChatwindow.SelColor = vbBlack
    txtChatwindow.SelText = txtChatwindow.SelText & txtChat.Text
    txtChat.Text = ""
    txtChatwindow.SelStart = (Len(txtChatwindow))
    istyping = False
    End If
End If

errorhandler:
    Select Case Err
    Case Is = 40006
     MsgBox "You are not connected to anyone!"
     server.Close
    End Select
End Sub

Private Sub Command2_Click()
Let txtIP.Text = server.LocalIP
End Sub

Private Sub Command3_Click()
Let server.LocalPort = txtLocalport.Text
End Sub

Private Sub Command4_Click()

If txtUsername.Text = "Username" Then
    MsgBox "Please enter a username", vbOKOnly, "New Username"
Else
    server.Listen
    Command5.Enabled = True
    Command4.Enabled = False
End If
End Sub

Private Sub Command5_Click()
server.Close
txtUsername.Enabled = True
StatusBar1.Panels(1).Text = "Not Connected..."
Command5.Enabled = False
Command4.Enabled = True
End Sub

Private Sub Command6_Click()
txtChatwindow.Text = ""
End Sub

Private Sub Command7_Click()
CommonDialog1.ShowOpen
Let txtSoundfile.Text = CommonDialog1.FileName
End Sub

Private Sub Command8_Click()
txtSoundfile.Text = ""
End Sub

Private Sub copy_Click()
Clipboard.SetText txtChat.SelText
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
Me.Height = 5130
server.Close
StatusBar1.Panels(2).Text = "Version " & App.Major & "." & App.Minor
nochats = 5
ProgressBar1.Value = nochats
server.LocalPort = 1001
oldcaption = Me.Caption

On Error GoTo errorhandler
Open App.Path + "\ctserver.cfg" For Input As #2
Input #2, var1, var2, var3, var4, var5, var6
Close #2

Let txtUsername.Text = var1
Let txtSoundfile.Text = var2
Let Check1.Value = var3
Let Me.Top = var4
Let Me.Left = var5
Let Check2.Value = var6

errorhandler:
Close #2
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu rightclick1
End Sub

Private Sub Form_Resize()
Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open App.Path + "\ctserver.cfg" For Output As #2
Write #2, txtUsername.Text, txtSoundfile, Check1.Value, Me.Top, Me.Left, Check2.Value
Close #2

  Dim counter As Integer
  Dim i As Integer
  counter = Me.Height
  Do: DoEvents
    counter = counter - 50
    Me.Height = counter
    Me.Top = (Screen.Height - Me.Height) / 2
  Loop Until counter <= 50
End

End Sub

Private Sub formback_Click()
CommonDialog1.CancelError = True
 On Error GoTo errHandler
 CommonDialog1.Flags = &H1
 CommonDialog1.ShowColor
 frmServer.BackColor = CommonDialog1.color
 SSTab1.BackColor = CommonDialog1.color
errHandler:
End Sub

Private Sub htmllink_Click()
Dim rc As String

rc = InputBox("Enter the URL", "HTML Link")

txtChat.Text = "link:" & rc
End Sub

Private Sub insertip_Click()
Let txtChat.Text = txtChat.Text & server.LocalIP
End Sub

Private Sub lastthing_Click()
If lastsaid = "" Then Else txtChat.Text = lastsaid
End Sub

Private Sub loadbrowser_Click()
BrowserMDIS.Show
BrowserMDIS.WindowState = 2
End Sub

Private Sub loadcolor_Click()

Dim tmp1 As String
Dim tmp2 As String
Dim tmp3 As String
Dim tmp4 As String
Dim tmp5 As String
Dim tmp6 As String
Dim tmp7 As String
Dim tmp8 As String
Dim tmp9 As String
Dim tmp10 As String


Open App.Path + "/colors.dat" For Input As #1
Input #1, tmp1, tmp2, tmp3, tmp4, tmp5, tmp6, tmp7, tmp8, tmp9, tmp10
Close #1

frmServer.BackColor = tmp1
SSTab1.BackColor = tmp2
SSTab1.ForeColor = tmp3
Command1.BackColor = tmp4
Command2.BackColor = tmp5
Command3.BackColor = tmp6
Command4.BackColor = tmp7
Command5.BackColor = tmp8
Command6.BackColor = tmp9
Command7.BackColor = tmp10

End Sub

Private Sub paste_Click()
txtChat.Text = Clipboard.GetText
End Sub

Private Sub rspload_Click()
server.SendData "101RSPLOAD"
frmRSP.Show
End Sub

Private Sub savecolor_Click()
Open App.Path + "/colors.dat" For Output As #1
Write #1, frmServer.BackColor, SSTab1.BackColor, SSTab1.ForeColor, Command1.BackColor, Command2.BackColor, Command3.BackColor, Command4.BackColor, Command5.BackColor, Command6.BackColor, Command7.BackColor
Close #1
End Sub

Private Sub sendb_Click()
Command1.BackColor = thecolor()
End Sub

Private Sub server_Close()
server.Close
Command4.Enabled = True
Command5.Enabled = False
StatusBar1.Panels(1).Text = "Not Connected..."
txtUsername.Enabled = True
Me.Caption = "Chat Thing - Server"
oldcaption = Me.Caption
txtChatwindow.SelBold = True
txtChatwindow.SelColor = vbRed
txtChatwindow.SelText = txtChatwindow.SelText & vbCrLf & usernames & " left the chat!"
txtChatwindow.SelBold = False
txtChatwindow.SelColor = vbBlack
usernames = ""
End Sub

Private Sub server_ConnectionRequest(ByVal requestID As Long)
If server.State <> sckClosed Then server.Close
server.Accept requestID
StatusBar1.Panels(1).Text = "Connected..."
SSTab1.Tab = 1
txtUsername.Enabled = False
server.SendData "1001NAME" + txtUsername.Text
play.Enabled = True
End Sub

Private Sub server_DataArrival(ByVal bytesTotal As Long)
Dim sound As String
Dim sound2 As Long
Dim Chat2 As String
server.GetData Chat2


If InStr(1, Chat2, "1001NAME") <> 0 Then
    usernames = Right$(Chat2, Len(Chat2) - 8)
    Me.Caption = "Chat Thing - Server" & " (" & usernames & ")"
    oldcaption = Me.Caption
ElseIf InStr(1, Chat2, "link:") <> 0 Then
    Dim rc As String
    rc = MsgBox(usernames & " wants you to visit " & Right$(Chat2, Len(Chat2) - 5) & ". Do you want to go?", vbYesNo, "Web Link")
    If rc = vbYes Then
        BrowserMDIS.Show
        BrowserMDIS.WindowState = 2
        frmBrowserS.cboAddress1.Text = Right$(Chat2, Len(Chat2) - 5)
        frmBrowserS.cboAddress1_Click
    End If
    Label1.Visible = False
ElseIf Chat2 = "101RSPLOAD" Then
    frmRSP.Show
ElseIf Chat2 = "101RSPCLOSE" Then
    Unload frmRSP
ElseIf Chat2 = "ISTYPING" Then
    Label1.Visible = True
    If smallw.Checked = True Then StatusBar1.Panels(1).Text = "Buddy is Typing.."
ElseIf InStr(1, Chat2, "101RSP") <> 0 Then
    rsp2 = Right$(Chat2, Len(Chat2) - 6)
    frmRSP.StatusBar1.SimpleText = "Waiting for you!"
    If Len(rsp) <> 0 Then Call frmRSP.checkwin
ElseIf Chat2 = "101MIN" Then
    Me.WindowState = 1
ElseIf Chat2 = "101CLEAR" Then
    txtChatwindow.Text = ""
ElseIf Chat2 = "TTTLOAD" Then
    frmTTTS.Show
ElseIf Chat2 = "TTTCLOSE" Then
    Unload frmTTTS
ElseIf Chat2 = "TTTBOX1" Then
    frmTTTS.Box1.Picture = LoadPicture(App.Path & "\X.jpg")
    frmTTTS.Box1.Enabled = False
    allboxs = allboxs + 1
    box1s = usernames
    frmTTTS.Frame1.Enabled = True
    frmTTTS.StatusBar1.SimpleText = "Choose your square."
    frmTTTS.calcwin
ElseIf Chat2 = "TTTBOX2" Then
    frmTTTS.Box2.Picture = LoadPicture(App.Path & "\X.jpg")
    frmTTTS.Box2.Enabled = False
    allboxs = allboxs + 1
    box2s = usernames
    frmTTTS.Frame1.Enabled = True
    frmTTTS.StatusBar1.SimpleText = "Choose your square."
    frmTTTS.calcwin
ElseIf Chat2 = "TTTBOX3" Then
    frmTTTS.Box3.Picture = LoadPicture(App.Path & "\X.jpg")
    frmTTTS.Box3.Enabled = False
    allboxs = allboxs + 1
    box3s = usernames
    frmTTTS.Frame1.Enabled = True
    frmTTTS.StatusBar1.SimpleText = "Choose your square."
    frmTTTS.calcwin
ElseIf Chat2 = "TTTBOX4" Then
    frmTTTS.Box4.Picture = LoadPicture(App.Path & "\X.jpg")
    frmTTTS.Box4.Enabled = False
    allboxs = allboxs + 1
    box4s = usernames
    frmTTTS.Frame1.Enabled = True
    frmTTTS.StatusBar1.SimpleText = "Choose your square."
    frmTTTS.calcwin
ElseIf Chat2 = "TTTBOX5" Then
    frmTTTS.Box5.Picture = LoadPicture(App.Path & "\X.jpg")
    frmTTTS.Box5.Enabled = False
    allboxs = allboxs + 1
    box5s = usernames
    frmTTTS.Frame1.Enabled = True
    frmTTTS.StatusBar1.SimpleText = "Choose your square."
    frmTTTS.calcwin
ElseIf Chat2 = "TTTBOX6" Then
    frmTTTS.Box6.Picture = LoadPicture(App.Path & "\X.jpg")
    frmTTTS.Box6.Enabled = False
    allboxs = allboxs + 1
    box6s = usernames
    frmTTTS.Frame1.Enabled = True
    frmTTTS.StatusBar1.SimpleText = "Choose your square."
    frmTTTS.calcwin
ElseIf Chat2 = "TTTBOX7" Then
    frmTTTS.Box7.Picture = LoadPicture(App.Path & "\X.jpg")
    frmTTTS.Box7.Enabled = False
    allboxs = allboxs + 1
    box7s = usernames
    frmTTTS.Frame1.Enabled = True
    frmTTTS.StatusBar1.SimpleText = "Choose your square."
    frmTTTS.calcwin
ElseIf Chat2 = "TTTBOX8" Then
    frmTTTS.Box8.Picture = LoadPicture(App.Path & "\X.jpg")
    frmTTTS.Box8.Enabled = False
    allboxs = allboxs + 1
    box8s = usernames
    frmTTTS.Frame1.Enabled = True
    frmTTTS.StatusBar1.SimpleText = "Choose your square."
    frmTTTS.calcwin
ElseIf Chat2 = "TTTBOX9" Then
    frmTTTS.Box9.Picture = LoadPicture(App.Path & "\X.jpg")
    frmTTTS.Box9.Enabled = False
    allboxs = allboxs + 1
    box9s = usernames
    frmTTTS.Frame1.Enabled = True
    frmTTTS.StatusBar1.SimpleText = "Choose your square."
    frmTTTS.calcwin
ElseIf timestamp.Checked = True Then
    txtChatwindow.SelColor = vbRed
    txtChatwindow.SelBold = True
    txtChatwindow.SelText = txtChatwindow.SelText & vbCrLf & "(" & Time & ") " & usernames
    txtChatwindow.SelBold = False
    txtChatwindow.SelColor = vbBlack
    txtChatwindow.SelText = txtChatwindow.SelText & " - " & Chat2
    txtChatwindow.SelStart = Len(txtChatwindow)
    
    If Check1.Value = Checked Then
        Let sound = txtSoundfile.Text
        sound2 = sndPlaySound(sound, 1)
    End If
    
    Label1.Visible = False
    
    If Me.WindowState = 1 Then Timer1.Enabled = True
    
    If smallw.Checked = True Then
        Let smallmsgs = smallmsgs + 1
        StatusBar1.Panels(1).Text = "You got " & smallmsgs & " msgs."
    End If
Else
    txtChatwindow.SelStart = Len(txtChatwindow)
    txtChatwindow.SelColor = vbRed
    txtChatwindow.SelBold = True
    txtChatwindow.SelText = txtChatwindow.SelText & vbCrLf & usernames
    txtChatwindow.SelBold = False
    txtChatwindow.SelColor = vbBlack
    txtChatwindow.SelText = txtChatwindow.SelText & " - " & Chat2
    txtChatwindow.SelStart = (Len(txtChatwindow))
    
    If Check1.Value = Checked Then
        Let sound = txtSoundfile.Text
        sound2 = sndPlaySound(sound, 1)
    End If

    Label1.Visible = False
    
    If Me.WindowState = 1 Then Timer1.Enabled = True
    
    If smallw.Checked = True Then
        Let smallmsgs = smallmsgs + 1
        StatusBar1.Panels(1).Text = "You got " & smallmsgs & " msgs."
    End If
End If

End Sub

Private Sub setportb_Click()
Command3.BackColor = thecolor()
End Sub

Private Sub showipb_Click()
Command2.BackColor = thecolor()
End Sub

Private Sub smallw_Click()
smallmsgs = 0

If smallw.Checked = True Then
Me.Height = 5130
smallw.Checked = False
    If usernames = "" Then
        StatusBar1.Panels(1).Text = "Not Connected..."
    ElseIf Len(usernames) > 0 Then
        StatusBar1.Panels(1).Text = "Connected..."
    End If
Else
Me.Height = 915
smallw.Checked = True
StatusBar1.Panels(1).Text = "You got " & smallmsgs & " msgs."
End If
End Sub

Private Sub smallw2_Click()
Call smallw_Click
End Sub

Private Sub startserverb_Click()
Command4.BackColor = thecolor()
End Sub

Private Sub stopserverb_Click()
Command5.BackColor = thecolor()
End Sub

Private Sub tabback_Click()
SSTab1.ForeColor = thecolor()
End Sub

Private Sub Timer1_Timer()
FlashWindow hwnd, 1
End Sub

Private Sub Timer2_Timer()
If nochats = 5 Then
Timer2.Enabled = False
Else
nochats = nochats + 1
ProgressBar1.Value = nochats
End If

End Sub

Private Sub Timer3_Timer()
Dim caption1 As String

If Me.Caption = "" Then
    Me.Caption = oldcaption
Else
    caption1 = Me.Caption
    Me.Caption = Right$(caption1, Len(caption1) - 1)
End If

End Sub

Private Sub timestamp_Click()
If timestamp.Checked = True Then
    timestamp.Checked = False
Else
    timestamp.Checked = True
End If
End Sub

Private Sub tttload_Click()
frmTTTS.Show
server.SendData "TTTLOAD"
frmTTTS.Frame1.Enabled = True
frmTTTS.StatusBar1.SimpleText = "Choose your square."
End Sub

Private Sub txtChat_Change()
Command1.Default = True

If Len(txtChat.Text) > 0 And Len(usernames) > 0 And istyping = False Then server.SendData "ISTYPING": istyping = True

End Sub

Private Function thecolor() As Long

CommonDialog1.CancelError = True
 On Error GoTo errHandler
 CommonDialog1.Flags = &H1
 CommonDialog1.ShowColor
 thecolor = CommonDialog1.color

errHandler:

End Function

Private Sub txtChat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    If lastsaid = "" Then Else txtChat.Text = lastsaid
End If
End Sub

Private Sub txtChat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If txtChat.SelText = "" Then copy.Enabled = False Else copy.Enabled = True
    If Clipboard.GetText = "" Then paste.Enabled = False Else paste.Enabled = True
    PopupMenu rightclickchat
End If
End Sub
