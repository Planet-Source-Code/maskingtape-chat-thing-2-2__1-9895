VERSION 5.00
Begin VB.MDIForm BrowserMDIS 
   BackColor       =   &H8000000C&
   Caption         =   "Browser - Server"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   Icon            =   "BrowserMDIC.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "BrowserMDIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
frmBrowserS.Show
    frmBrowserS.WindowState = 2
    frmBrowserS.cboAddress1.Text = ""
End Sub
