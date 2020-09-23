VERSION 5.00
Begin VB.MDIForm BrowserMDIC 
   BackColor       =   &H8000000C&
   Caption         =   "Browser - Client"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   Icon            =   "BrowserMDIC.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "BrowserMDIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
frmBrowserC.Show
    frmBrowserC.WindowState = 2
    frmBrowserC.cboAddress.Text = ""
End Sub
