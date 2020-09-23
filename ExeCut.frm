VERSION 5.00
Object = "{A48A1D86-6463-4F3A-B152-6051F2FA65A1}#1.0#0"; "GURHANCOOLBUTTON.OCX"
Begin VB.Form ExeCut 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Execute"
   ClientHeight    =   945
   ClientLeft      =   3840
   ClientTop       =   5790
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4440
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   240
      Left            =   1515
      TabIndex        =   3
      Top             =   1470
      Width           =   900
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "ExeCut.frx":0000
      Left            =   0
      List            =   "ExeCut.frx":0013
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin Gurhan_Cool_Button.GurhanCoolButton GurhanCoolButton2 
      Height          =   360
      Left            =   1815
      TabIndex        =   2
      Top             =   300
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      CaptionTitle    =   "OK"
      Picture         =   "ExeCut.frx":0050
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
   Begin Gurhan_Cool_Button.GurhanCoolButton GurhanCoolButton1 
      Height          =   330
      Left            =   1815
      TabIndex        =   1
      Top             =   615
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      CaptionTitle    =   "Cancel"
      Picture         =   "ExeCut.frx":2682
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
   Begin VB.Image Image2 
      Height          =   690
      Left            =   -960
      Picture         =   "ExeCut.frx":4CB4
      Top             =   285
      Width           =   5475
   End
   Begin VB.Image Image1 
      Height          =   2820
      Left            =   0
      Picture         =   "ExeCut.frx":111E6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "ExeCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Command1_Click()
FileExecutor Me.hwnd, Combo1.Text, "Open"
Unload Me
End Sub

Private Sub Form_Load()
Combo1.Text = Left(App.Path, 2)
Left = 200
Top = Form1.Height + 200
ExplodeForm ExeCut, 3500
End Sub

Private Sub GurhanCoolButton1_CLICKED()
Unload Me
End Sub

Private Sub GurhanCoolButton2_CLICKED()
Select Case Combo1.Text
    Case "drive a"
    Combo1.Text = "A:"
    Case "drive c"
    Combo1.Text = "C:"
' Drives End
    Case "my documents"
    Combo1.Text = "C:\mydocu~1"
    Case "program files"
    Combo1.Text = "C:\progra~1"
    Case "explorer"
    Combo1.Text = "explorer"
    Case Else
    FileExecutor Me.hwnd, Combo1.Text, "Open"
End Select
Unload Me
End Sub


