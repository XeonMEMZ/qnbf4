VERSION 5.00
Begin VB.Form Form36 
   BorderStyle     =   0  'None
   Caption         =   "Form36"
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   LinkTopic       =   "Form36"
   Picture         =   "Form36.frx":0000
   ScaleHeight     =   2085
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
cdspeed = 100
maxl = Screen.Width - 8985
maxt = Screen.Height - 2100
x = cdspeed
y = cdspeed
End Sub

Private Sub Timer1_Timer()
Form36.Left = Form36.Left + x
Form36.Top = Form36.Top + y
If Form36.Left < 0 Then
 x = cdspeed
End If
If Form36.Left > maxl Then
 x = cdspeed * -1
End If
If Form36.Top < 0 Then
 y = cdspeed
End If
If Form36.Top > maxt Then
 y = cdspeed * -1
End If
End Sub

