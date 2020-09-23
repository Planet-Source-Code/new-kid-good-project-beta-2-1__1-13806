VERSION 5.00
Object = "{A48A1D86-6463-4F3A-B152-6051F2FA65A1}#1.0#0"; "GURHANCOOLBUTTON.OCX"
Object = "{D2303D7E-7DAC-11D3-8F49-00104B4D60E0}#6.0#0"; "SOSGENERALCONTROLS.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PicView1 
   Caption         =   "Picture Viewer 1.0"
   ClientHeight    =   5805
   ClientLeft      =   2190
   ClientTop       =   1590
   ClientWidth     =   7665
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   Picture         =   "PicView1.frx":0000
   ScaleHeight     =   5805
   ScaleWidth      =   7665
   Begin sosGeneralControls.sosOption sosOption1 
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      Caption         =   "Resize Image On/Off"
      BackColor       =   9453343
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   5325
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Gurhan_Cool_Button.GurhanCoolButton GurhanCoolButton3 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   105
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   450
      CaptionTitle    =   ""
      MouseIcon       =   "PicView1.frx":91242
      Picture         =   "PicView1.frx":9155C
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
      ButtonTextColor =   65280
   End
   Begin Gurhan_Cool_Button.GurhanCoolButton GurhanCoolButton2 
      Height          =   390
      Left            =   3720
      TabIndex        =   1
      Top             =   5235
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   688
      CaptionTitle    =   "Close"
      MouseIcon       =   "PicView1.frx":A9ADE
      Picture         =   "PicView1.frx":A9DF8
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
   End
   Begin Gurhan_Cool_Button.GurhanCoolButton GurhanCoolButton1 
      Height          =   390
      Left            =   2265
      TabIndex        =   0
      Top             =   5235
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   688
      CaptionTitle    =   "Locate File"
      MouseIcon       =   "PicView1.frx":ABE76
      Picture         =   "PicView1.frx":AC190
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
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4710
      Left            =   315
      Stretch         =   -1  'True
      Top             =   405
      Width           =   6870
   End
End
Attribute VB_Name = "PicView1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Stretch = sosOption1.Value = True
Left = 200
Top = Form1.Height + 200
ExplodeForm PicView1, 3500
End Sub

Private Sub GurhanCoolButton1_CLICKED()
CommonDialog1.Filter = "All Images|*.BMP;*.GIF"
CommonDialog1.ShowOpen
GurhanCoolButton3.CaptionTitle = CommonDialog1.FileName
Image1.Picture = LoadPicture(GurhanCoolButton3.CaptionTitle)

End Sub

Private Sub GurhanCoolButton2_CLICKED()
Image1.Picture = LoadPicture("")
Unload Me
End Sub

Private Sub GurhanCoolButton3_CLICKED()
If GurhanCoolButton3.CaptionTitle = "" Then
Else
FileExecutor Me.hWnd, GurhanCoolButton3.CaptionTitle, "Open"
End If
End Sub

Private Sub sosOption1_Click()
Image1.Stretch = sosOption1.Value
Image1.Width = 6870
Image1.Height = 4710
End Sub
