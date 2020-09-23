VERSION 5.00
Object = "{A48A1D86-6463-4F3A-B152-6051F2FA65A1}#1.0#0"; "GURHANCOOLBUTTON.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About Good Project Beta 2"
   ClientHeight    =   6000
   ClientLeft      =   2505
   ClientTop       =   2070
   ClientWidth     =   7560
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   4141.307
   ScaleMode       =   0  'User
   ScaleWidth      =   7099.231
   Begin Gurhan_Cool_Button.GurhanCoolButton GurhanCoolButton2 
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      ToolTipText     =   "Opens the Read Me File"
      Top             =   5490
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   450
      CaptionTitle    =   "Read Me"
      MouseIcon       =   "frmAbout.frx":6F7F2
      ButtonBackColor =   0
      UseBorders      =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonTextColor =   16744576
   End
   Begin Gurhan_Cool_Button.GurhanCoolButton GurhanCoolButton1 
      Height          =   615
      Left            =   6600
      TabIndex        =   2
      ToolTipText     =   "Close the About Window"
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
      CaptionTitle    =   "Close"
      MouseIcon       =   "frmAbout.frx":6FB0C
      Picture         =   "frmAbout.frx":6FE26
      ButtonBackColor =   0
      UseBorders      =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonTextColor =   16777215
   End
   Begin Gurhan_Cool_Button.GurhanCoolButton GurhanCoolButton3 
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      ToolTipText     =   "Click to Visit Yube on the net"
      Top             =   5160
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   450
      CaptionTitle    =   "WWW.GOODPROJECT.BIZLAND.COM"
      MouseIcon       =   "frmAbout.frx":71470
      ButtonBackColor =   0
      UseBorders      =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonTextColor =   16744576
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "*This doesn't include the BAS files. Look in the Read Me for the names of the BAS file makers."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   6255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Email:   Bfuzz@mbox.com.au                  Website:    WWW.YUBE.BIZLAND.COM"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"frmAbout.frx":7178A
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   7575
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   0
      Top             =   4200
      Width           =   135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Form_Load()
ExplodeForm frmAbout, 3500
A$ = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion"
B$ = GetStringValue(A$, "Version")
Label4.Caption = "OS Version: " & B$
End Sub

Private Sub GurhanCoolButton1_CLICKED()
Unload Me
End Sub


Private Sub GurhanCoolButton2_CLICKED()
Edit1.Show
Edit1.RichTextBox1.FileName = App.Path & "\readme.txt"
End Sub


Private Sub GurhanCoolButton3_CLICKED()
FileExecutor Me.hwnd, "WWW.GOODPROJECT.BIZLAND.COM", "Open"
End Sub


