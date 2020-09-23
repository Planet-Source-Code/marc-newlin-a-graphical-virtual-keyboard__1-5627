VERSION 5.00
Begin VB.Form splash 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   5115
   ClientTop       =   3015
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "splash.frx":0000
   ScaleHeight     =   3405
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2520
      Top             =   1440
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Initialize()
Timer1.Enabled = True
splash.Show
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload splash
vkb.Show
End Sub
