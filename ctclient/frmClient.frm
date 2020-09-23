VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Thing - Client"
   ClientHeight    =   4470
   ClientLeft      =   5445
   ClientTop       =   3330
   ClientWidth     =   4710
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4710
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4200
      Top             =   120
   End
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
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show IP"
      Height          =   255
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock client 
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
      TabPicture(0)   =   "frmClient.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtRemoteport"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtRemotehost"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Timer1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CommonDialog1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Timer2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command8"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Timer4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Chat!"
      TabPicture(1)   =   "frmClient.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "ProgressBar1"
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(4)=   "txtChatwindow"
      Tab(1).Control(5)=   "txtChat"
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame2 
         Caption         =   "Other Options"
         Height          =   2175
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   1575
         Begin VB.CheckBox Check2 
            Caption         =   "Scroll Title Bar"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Value           =   1  'Checked
            Width           =   1335
         End
      End
      Begin VB.Timer Timer4 
         Interval        =   500
         Left            =   1920
         Top             =   2280
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3240
         Width           =   975
      End
      Begin RichTextLib.RichTextBox txtChat 
         Height          =   295
         Left            =   -74880
         TabIndex        =   18
         Top             =   3000
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   529
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"frmClient.frx":047A
      End
      Begin RichTextLib.RichTextBox txtChatwindow 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4260
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmClient.frx":0528
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   1920
         Top             =   1320
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1800
         Width           =   1695
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
      Begin VB.CommandButton Command6 
         Caption         =   "Change"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3000
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2880
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select Sound File"
         Filter          =   "Wave Files | *.wav"
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear"
         Height          =   255
         Left            =   -72000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Clear The Chat Window"
         Top             =   0
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1920
         Top             =   1800
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Connect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtRemotehost 
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Text            =   "IP Address"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Set Port"
         Height          =   255
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtRemoteport 
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Text            =   "1001"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Send"
         Height          =   255
         Left            =   -71880
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3360
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   -74880
         TabIndex        =   16
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
         TabIndex        =   19
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu loadbrowser 
         Caption         =   "Load &Browser"
      End
      Begin VB.Menu line6 
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
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu smallw 
         Caption         =   "Small Window"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu play 
      Caption         =   "&Games"
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
         Begin VB.Menu buttonback 
            Caption         =   "All Buttons"
         End
         Begin VB.Menu line2 
            Caption         =   "-"
         End
         Begin VB.Menu setportb 
            Caption         =   "'Set Port' Button"
         End
         Begin VB.Menu connectb 
            Caption         =   "'Connect' Button"
         End
         Begin VB.Menu disconnectb 
            Caption         =   "'Disconnect' Button"
         End
         Begin VB.Menu changesb 
            Caption         =   "'Change Sound' Button"
         End
         Begin VB.Menu clearb 
            Caption         =   "'Clear' Button"
         End
         Begin VB.Menu sendb 
            Caption         =   "'Send' Button"
         End
         Begin VB.Menu showipb 
            Caption         =   "'Show IP' Button"
         End
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu colorsave 
         Caption         =   "Save"
      End
      Begin VB.Menu colorload 
         Caption         =   "Load"
      End
   End
   Begin VB.Menu rightclick1 
      Caption         =   "rightclick"
      Visible         =   0   'False
      Begin VB.Menu smallw2 
         Caption         =   "Small Window  F2"
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
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub buttonback_Click()

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

Private Sub changesb_Click()
Command6.BackColor = thecolor()
End Sub

Private Sub Check2_Click()
If Check2.Value = Unchecked Then
    Timer4.Enabled = False
    Me.Caption = oldcaption
Else: Check2.Value = Checked
    Timer4.Enabled = True
End If
End Sub

Private Sub clear_Click()
txtChatwindow.Text = ""
End Sub

Private Sub clearb_Click()
Command5.BackColor = thecolor()
End Sub

Private Sub client_Close()
client.Close
txtRemotehost.Enabled = True
StatusBar1.Panels(1).Text = "Not Connected..."
Command7.Enabled = False
Command4.Enabled = True
txtUsername.Enabled = True
Me.Caption = "Chat Thing - Client"
oldcaption = Me.Caption
txtChatwindow.SelBold = True
txtChatwindow.SelColor = vbRed
txtChatwindow.SelText = txtChatwindow.SelText & vbCrLf & usernamec & " left the chat!"
txtChatwindow.SelBold = False
txtChatwindow.SelColor = vbBlack
usernamec = ""
End Sub

Private Sub client_Connect()
txtRemotehost.Enabled = False
StatusBar1.Panels(1).Text = "Connected..."
Timer1.Enabled = False
SSTab1.Tab = 1
txtUsername.Enabled = False
Command7.Enabled = True
client.SendData "1001NAME" + txtUsername.Text
play.Enabled = True
End Sub

Private Sub client_DataArrival(ByVal bytesTotal As Long)
Dim Chat As String
Dim sound As String
Dim sound2 As Long
client.GetData Chat

If Chat = "101MIN" Then
    Me.WindowState = 1
ElseIf Chat = "101CLEAR" Then
    txtChatwindow.Text = ""
ElseIf Chat = "ISTYPING" Then
    Label1.Visible = True
    If smallw.Checked = True Then StatusBar1.Panels(1).Text = "Buddy is Typing.."
ElseIf InStr(1, Chat, "1001NAME") <> 0 Then
    usernamec = Right$(Chat, Len(Chat) - 8)
    Me.Caption = "Chat Thing - Client" & " (" & usernamec & ")"
    oldcaption = Me.Caption
ElseIf InStr(1, Chat, "link:") <> 0 Then
    Dim rc As String
    rc = MsgBox(usernamec & " wants you to visit " & Right$(Chat, Len(Chat) - 5) & ". Do you want to go?", vbYesNo, "Web Link")
    If rc = vbYes Then
        BrowserMDIC.Show
        BrowserMDIC.WindowState = 2
        frmBrowserC.cboAddress.Text = Right$(Chat, Len(Chat) - 5)
        frmBrowserC.cboAddress_Click
    End If
    Label1.Visible = False
ElseIf Chat = "101RSPLOAD" Then
    frmRSP.Show
ElseIf Chat = "101RSPCLOSE" Then
    Unload frmRSP
ElseIf InStr(1, Chat, "101RSP") <> 0 Then
    rsp2 = Right$(Chat, Len(Chat) - 6)
    frmRSP.StatusBar1.SimpleText = "Waiting for you!"
    If Len(rsp) <> 0 Then Call frmRSP.checkwin
ElseIf Chat = "TTTLOAD" Then
    frmTTTC.Show
ElseIf Chat = "TTTCLOSE" Then
    Unload frmTTTC
ElseIf Chat = "TTTBOX1" Then
    frmTTTC.Box1.Picture = LoadPicture(App.Path & "\O.jpg")
    frmTTTC.Box1.Enabled = False
    allboxc = allboxc + 1
    box1c = usernamec
    frmTTTC.Frame1.Enabled = True
    frmTTTC.StatusBar1.SimpleText = "Choose your square."
    Call frmTTTC.calcwin
ElseIf Chat = "TTTBOX2" Then
    frmTTTC.Box2.Picture = LoadPicture(App.Path & "\O.jpg")
    frmTTTC.Box2.Enabled = False
    allboxc = allboxc + 1
    box2c = usernamec
    frmTTTC.Frame1.Enabled = True
    frmTTTC.StatusBar1.SimpleText = "Choose your square."
    Call frmTTTC.calcwin
ElseIf Chat = "TTTBOX3" Then
    frmTTTC.Box3.Picture = LoadPicture(App.Path & "\O.jpg")
    frmTTTC.Box3.Enabled = False
    allboxc = allboxc + 1
    box3c = usernamec
    frmTTTC.Frame1.Enabled = True
    frmTTTC.StatusBar1.SimpleText = "Choose your square."
    Call frmTTTC.calcwin
ElseIf Chat = "TTTBOX4" Then
    frmTTTC.Box4.Picture = LoadPicture(App.Path & "\O.jpg")
    frmTTTC.Box4.Enabled = False
    allboxc = allboxc + 1
    box4c = usernamec
    frmTTTC.Frame1.Enabled = True
    frmTTTC.StatusBar1.SimpleText = "Choose your square."
    Call frmTTTC.calcwin
ElseIf Chat = "TTTBOX5" Then
    frmTTTC.Box5.Picture = LoadPicture(App.Path & "\O.jpg")
    frmTTTC.Box5.Enabled = False
    allboxc = allboxc + 1
    box5c = usernamec
    frmTTTC.Frame1.Enabled = True
    frmTTTC.StatusBar1.SimpleText = "Choose your square."
    Call frmTTTC.calcwin
ElseIf Chat = "TTTBOX6" Then
    frmTTTC.Box6.Picture = LoadPicture(App.Path & "\O.jpg")
    frmTTTC.Box6.Enabled = False
    allboxc = allboxc + 1
    box6c = usernamec
    frmTTTC.Frame1.Enabled = True
    frmTTTC.StatusBar1.SimpleText = "Choose your square."
    Call frmTTTC.calcwin
ElseIf Chat = "TTTBOX7" Then
    frmTTTC.Box7.Picture = LoadPicture(App.Path & "\O.jpg")
    frmTTTC.Box7.Enabled = False
    allboxc = allboxc + 1
    box7c = usernamec
    frmTTTC.Frame1.Enabled = True
    frmTTTC.StatusBar1.SimpleText = "Choose your square."
    Call frmTTTC.calcwin
ElseIf Chat = "TTTBOX8" Then
    frmTTTC.Box8.Picture = LoadPicture(App.Path & "\O.jpg")
    frmTTTC.Box8.Enabled = False
    allboxc = allboxc + 1
    box8c = usernamec
    frmTTTC.Frame1.Enabled = True
    frmTTTC.StatusBar1.SimpleText = "Choose your square."
    Call frmTTTC.calcwin
ElseIf Chat = "TTTBOX9" Then
    frmTTTC.Box9.Picture = LoadPicture(App.Path & "\O.jpg")
    frmTTTC.Box9.Enabled = False
    allboxc = allboxc + 1
    box9c = usernamec
    frmTTTC.Frame1.Enabled = True
    frmTTTC.StatusBar1.SimpleText = "Choose your square."
    Call frmTTTC.calcwin
ElseIf timestamp.Checked = True Then
    txtChatwindow.SelColor = vbRed
    txtChatwindow.SelBold = True
    txtChatwindow.SelText = txtChatwindow.SelText & vbCrLf & "(" & Time & ") " & usernamec
    txtChatwindow.SelBold = False
    txtChatwindow.SelColor = vbBlack
    txtChatwindow.SelText = txtChatwindow.SelText & " - " & Chat
    txtChatwindow.SelStart = (Len(txtChatwindow))
    
    If Check1.Value = Checked Then
        Let sound = txtSoundfile.Text
        sound2 = sndPlaySound(sound, 1)
    End If
    
    Label1.Visible = False
    
    If Me.WindowState = 1 Then Timer2.Enabled = True
    
    If smallw.Checked = True Then
        Let smallmsgs = smallmsgs + 1
        StatusBar1.Panels(1).Text = "You got " & smallmsgs & " msgs."
    End If
Else
    txtChatwindow.SelColor = vbRed
    txtChatwindow.SelBold = True
    txtChatwindow.SelText = txtChatwindow.SelText & vbCrLf & usernamec
    txtChatwindow.SelBold = False
    txtChatwindow.SelColor = vbBlack
    txtChatwindow.SelText = txtChatwindow.SelText & " - " & Chat
    txtChatwindow.SelStart = (Len(txtChatwindow))
    
    If Check1.Value = Checked Then
        Let sound = txtSoundfile.Text
        sound2 = sndPlaySound(sound, 1)
    End If
    
    Label1.Visible = False

    If Me.WindowState = 1 Then Timer2.Enabled = True

    If smallw.Checked = True Then
        Let smallmsgs = smallmsgs + 1
        StatusBar1.Panels(1).Text = "You got " & smallmsgs & " msgs."
    End If
End If

End Sub

Private Sub colorload_Click()

Open App.Path + "/colors.dat" For Input As #1
Input #1, tmp1, tmp2, tmp3, tmp4, tmp5, tmp6, tmp7, tmp8, tmp9, tmp10
Close #1

frmClient.BackColor = tmp1
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

Private Sub colorsave_Click()
Open App.Path + "/colors.dat" For Output As #1
Write #1, frmClient.BackColor, SSTab1.BackColor, SSTab1.ForeColor, Command1.BackColor, Command2.BackColor, Command3.BackColor, Command4.BackColor, Command5.BackColor, Command6.BackColor, Command7.BackColor
Close #1
End Sub

Private Sub Command1_Click()
Let txtIP.Text = client.LocalIP
End Sub

Private Sub Command2_Click()
On Error GoTo errorhandler

If txtChat.Text = "" Then
Else
    If nochatc = 0 Then
    Else
    lastsaid = txtChat.Text
    nochatc = nochatc - 1
    ProgressBar1.Value = nochatc
    Timer3.Enabled = True
    client.SendData txtChat.Text
    txtChatwindow.SelColor = vbBlue
    txtChatwindow.SelBold = True
    txtChatwindow.SelText = txtChatwindow.SelText & vbCrLf & txtUsername
    txtChatwindow.SelBold = False
    txtChatwindow.SelColor = vbBlack
    txtChatwindow.SelText = txtChatwindow.SelText & " - " & txtChat.Text
    txtChat.Text = ""
    txtChatwindow.SelStart = Len(txtChatwindow)
    istyping = False
    End If
End If

errorhandler:
    Select Case Err
    Case Is = 40006
     MsgBox "You are not connected to anyone!"
     client.Close
    End Select
End Sub

Private Sub Command3_Click()
Let client.RemotePort = txtRemoteport.Text
End Sub

Private Sub Command4_Click()
On Error GoTo errorhandler
timercount = 0

If txtUsername.Text = "Username" Then
    MsgBox "Please enter a username", vbOKOnly, "New Username"
Else
Let client.RemoteHost = txtRemotehost.Text
client.Connect
Command4.Enabled = False
Timer1.Enabled = True
End If
errorhandler:
End Sub

Private Sub Command5_Click()
txtChatwindow.Text = ""
End Sub

Private Sub Command6_Click()
CommonDialog1.ShowOpen
Let txtSoundfile.Text = CommonDialog1.FileName

End Sub

Private Sub Command7_Click()
client.Close
txtUsername.Enabled = True
StatusBar1.Panels(1).Text = "Not Connected..."
Command4.Enabled = True
Command7.Enabled = False
End Sub

Private Sub Command8_Click()
txtSoundfile.Text = ""
End Sub

Private Sub connectb_Click()
Command4.BackColor = thecolor()
End Sub

Private Sub copy_Click()
Clipboard.SetText txtChat.SelText
End Sub

Private Sub disconnectb_Click()
Command7.BackColor = thecolor()
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
Me.Height = 5130
client.Close
StatusBar1.Panels(2).Text = "Version " & App.Major & "." & App.Minor

nochatc = 5
ProgressBar1.Value = nochatc
Let client.RemotePort = 1001

oldcaption = Me.Caption

On Error GoTo errorhandler
Open App.Path + "\" + "ctclient.cfg" For Input As #1
Input #1, var1, var2, var3, var4, var5, var6
Close #1

Let txtUsername.Text = var1
Let txtSoundfile.Text = var2
Let Check1.Value = var3
Let Me.Top = var4
Let Me.Left = var5
Let Check2.Value = var6

errorhandler:
Close #1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu rightclick1
End Sub

Private Sub Form_Resize()
Timer2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open App.Path + "\" + "ctclient.cfg" For Output As #1
Write #1, txtUsername.Text, txtSoundfile, Check1.Value, Me.Top, Me.Left, Check2.Value
Close #1

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
 frmClient.BackColor = CommonDialog1.color
 SSTab1.BackColor = CommonDialog1.color
errHandler:
End Sub

Private Sub htmllink_Click()
Dim rc As String

rc = InputBox("Enter the URL", "HTML Link")

txtChat.Text = "link:" & rc
End Sub

Private Sub insertip_Click()
Let txtChat.Text = txtChat.Text & client.LocalIP
End Sub

Private Sub lastthing_Click()
If lastsaid = "" Then Else txtChat.Text = lastsaid
End Sub

Private Sub loadbrowser_Click()
BrowserMDIC.Show
BrowserMDIC.WindowState = 2
End Sub

Private Sub mp3player_Click()
frmMP3.Show
End Sub

Private Sub paste_Click()
txtChat.Text = Clipboard.GetText
End Sub

Private Sub rspload_Click()
client.SendData "101RSPLOAD"
frmRSP.Show
End Sub

Private Sub sendb_Click()
Command2.BackColor = thecolor()
End Sub

Private Sub setportb_Click()
Command3.BackColor = thecolor()
End Sub

Private Sub showipb_Click()
Command1.BackColor = thecolor()
End Sub

Private Sub smallw_Click()
smallmsgs = 0

If smallw.Checked = True Then
Me.Height = 5130
smallw.Checked = False
    If usernamec = "" Then
        StatusBar1.Panels(1).Text = "Not Connected..."
    ElseIf Len(usernamec) > 0 Then
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

Private Sub tabback_Click()
SSTab1.ForeColor = thecolor()
End Sub

Private Sub Timer2_Timer()
FlashWindow hwnd, 1
End Sub

Private Sub Timer3_Timer()
If nochatc = 5 Then
Timer3.Enabled = False
Else
nochatc = nochatc + 1
ProgressBar1.Value = nochatc
End If
End Sub

Private Sub Timer4_Timer()
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
frmTTTC.Show
client.SendData "TTTLOAD"
frmTTTC.Frame1.Enabled = True
frmTTTC.StatusBar1.SimpleText = "Choose your square."
End Sub

Private Sub txtChat_Change()
Command2.Default = True

If Len(txtChat.Text) > 0 And Len(usernamec) > 0 And istyping = False Then client.SendData "ISTYPING": istyping = True

End Sub

Private Sub Timer1_Timer()
Let timercount = timercount + 1
client.Close

If timercount = 10 Then
    StatusBar1.Panels(1).Text = "Connection failed..."
    Timer1.Enabled = False
    Command4.Enabled = True
    client.Close
Else
    client.Connect
End If
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

Private Sub txtRemotehost_Change()
Command4.Enabled = True
Command4.Default = True
End Sub

