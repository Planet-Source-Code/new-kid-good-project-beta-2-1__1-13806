VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Edit1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor 2000: No File"
   ClientHeight    =   6090
   ClientLeft      =   2640
   ClientTop       =   1590
   ClientWidth     =   7980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   Icon            =   "Edit1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6090
   ScaleWidth      =   7980
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6360
      Top             =   6345
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6975
      Top             =   6255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5010
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8837
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Edit1.frx":0442
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3660
      TabIndex        =   1
      Top             =   5730
      Width           =   885
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   195
      Picture         =   "Edit1.frx":04FC
      Top             =   5595
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   6165
      Left            =   -135
      Picture         =   "Edit1.frx":0D3E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8250
   End
   Begin VB.Menu All 
      Caption         =   "All"
      Visible         =   0   'False
      Begin VB.Menu open 
         Caption         =   "&Open"
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
      End
      Begin VB.Menu help 
         Caption         =   "&Help"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Edit1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
Form2.Show
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1 = Time
Left = 200
Top = Form1.Height + 200
ExplodeForm Edit1, 3500
End Sub

Private Sub help_Click()
Path = App.Path
retval = Shell("winhelp.exe Edit.HLP", 1)
End Sub


Private Sub Image2_Click()
PopupMenu All
End Sub


Private Sub saveas_Click()

End Sub


Private Sub open_Click()
CommonDialog1.Filter = "Text Files|*.txt|Yube Text File|*.YB|All Files|*.*"
CommonDialog1.ShowOpen
RichTextBox1.FileName = CommonDialog1.FileName
Form1.Caption = "Editor 2000: " + CommonDialog1.FileName
End Sub


Private Sub save_Click()
CommonDialog1.Filter = "Yube Text File|*.YB|Text Files|*.txt|All Files|*.*"
CommonDialog1.ShowSave
'CommonDialog1.filename
Edit1.RichTextBox1.SaveFile CommonDialog1.FileName
End Sub


Private Sub Timer1_Timer()
Label1 = Time
End Sub


