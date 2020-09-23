VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4485
   ClientLeft      =   2070
   ClientTop       =   2610
   ClientWidth     =   7560
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   2160
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Top = 2265
Left = 2010
ExplodeForm Form2, 3500
End Sub

Private Sub Timer1_Timer()
Form1.Show
Unload Me
End Sub
