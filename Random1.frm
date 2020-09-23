VERSION 5.00
Object = "{D2303D7E-7DAC-11D3-8F49-00104B4D60E0}#6.0#0"; "SOSGENERALCONTROLS.OCX"
Begin VB.Form Random1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Match the Image 1.0"
   ClientHeight    =   4500
   ClientLeft      =   4335
   ClientTop       =   2535
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4500
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Default         =   -1  'True
      Height          =   1575
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "How to Win"
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   4935
      Begin VB.Image Image22 
         Height          =   615
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image Image21 
         Height          =   615
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image Image20 
         Height          =   615
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Image19 
         Height          =   615
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Image18 
         Height          =   615
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image Image17 
         Height          =   615
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image Image16 
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Image15 
         Height          =   615
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image Image14 
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image Image13 
         Height          =   615
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image Image12 
         Height          =   615
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image Image11 
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image Image10 
         Height          =   615
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image Image9 
         Height          =   615
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image Image8 
         Height          =   615
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Image7 
         Height          =   615
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   615
      End
      Begin VB.Image Image6 
         Height          =   615
         Left            =   840
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Image5 
         Height          =   615
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "To win get 3 of the same, or one of the following."
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
   End
   Begin sosGeneralControls.sosButton sosButton1 
      Height          =   1575
      Left            =   3600
      TabIndex        =   0
      ToolTipText     =   "Click to Play the Game"
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2778
      Caption         =   "Play"
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   2670
      Index           =   5
      Left            =   11280
      Picture         =   "Random1.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   1335
      Index           =   4
      Left            =   11640
      Picture         =   "Random1.frx":145AA
      Top             =   3600
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   1425
      Index           =   3
      Left            =   10920
      Picture         =   "Random1.frx":19D88
      Top             =   9600
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image Image1 
      Height          =   3375
      Index           =   2
      Left            =   10440
      Picture         =   "Random1.frx":1F0EA
      Top             =   5280
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   1605
      Index           =   1
      Left            =   8400
      Picture         =   "Random1.frx":3D654
      Top             =   480
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   6810
      Index           =   0
      Left            =   5880
      Picture         =   "Random1.frx":45E02
      Top             =   3000
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "Random1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Randomize Timer
A$ = Int((Rnd * 6))
B$ = Int((Rnd * 6))
C$ = Int((Rnd * 6))
Image2.Picture = Image1(C$).Picture
Image3.Picture = Image1(B$).Picture
Image4.Picture = Image1(A$).Picture
End Sub


Private Sub Form_Load()
Left = 200
Top = Form1.Height + 200
ExplodeForm Random1, 3500
Random1.Image5.Picture = Image1(1).Picture
Random1.Image6.Picture = Image1(1).Picture
Random1.Image7.Picture = Image1(2).Picture
Random1.Image8.Picture = Image1(2).Picture
Random1.Image9.Picture = Image1(3).Picture
Random1.Image10.Picture = Image1(3).Picture
Random1.Image11.Picture = Image1(4).Picture
Random1.Image12.Picture = Image1(4).Picture
Random1.Image13.Picture = Image1(5).Picture
Random1.Image14.Picture = Image1(5).Picture
Random1.Image15.Picture = Image1(0).Picture
Random1.Image16.Picture = Image1(0).Picture
Random1.Image17.Picture = Image1(1).Picture
Random1.Image18.Picture = Image1(2).Picture
Random1.Image19.Picture = Image1(3).Picture
Random1.Image10.Picture = Image1(4).Picture
Random1.Image21.Picture = Image1(5).Picture
Random1.Image22.Picture = Image1(0).Picture
Random1.Image20.Picture = Image1(0).Picture

End Sub

Private Sub sosButton1_Click(Button As Integer)
Randomize Timer
A$ = Int((Rnd * 6))
B$ = Int((Rnd * 6))
C$ = Int((Rnd * 6))
Image2.Picture = Image1(C$).Picture
Image3.Picture = Image1(B$).Picture
Image4.Picture = Image1(A$).Picture
End Sub
