VERSION 5.00
Object = "{D2303D7E-7DAC-11D3-8F49-00104B4D60E0}#6.0#0"; "SOSGENERALCONTROLS.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Vote1 
   BorderStyle     =   0  'None
   Caption         =   "Thank you for your Vote"
   ClientHeight    =   3405
   ClientLeft      =   1980
   ClientTop       =   4410
   ClientWidth     =   9990
   LinkTopic       =   "Form3"
   ScaleHeight     =   3405
   ScaleWidth      =   9990
   Begin sosGeneralControls.sosProgress sosProgress2 
      Height          =   255
      Left            =   0
      Top             =   3120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   450
      Caption         =   "Youe Vote will help me help others in the Visual Basic area of programing. Visit my Website @ www.yube.bizland.com"
      Value           =   100
      CaptionStyle    =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   5
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Index           =   2
      Left            =   9240
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin sosGeneralControls.sosProgress sosProgress1 
      Height          =   255
      Left            =   0
      Top             =   2880
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   450
      Caption         =   "Thank you for your help. You must connect to the web to Vote for me."
      Value           =   100
      CaptionStyle    =   1
   End
   Begin sosGeneralControls.sosFrame sosFrame1 
      Height          =   2895
      Left            =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5106
      Caption         =   "Click on the Banner to Vote. This will help me continue my work in Visual Basic"
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   255
         Index           =   0
         Left            =   9240
         TabIndex        =   1
         Top             =   0
         Width           =   735
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2655
         Left            =   0
         TabIndex        =   0
         Top             =   240
         Width           =   9975
         ExtentX         =   17595
         ExtentY         =   4683
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
End
Attribute VB_Name = "Vote1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Top = Form1.Height
Left = 0
ExplodeForm Vote1, 3500
Vote1.WebBrowser1.Navigate "www.goodproject.bizland.com\vote.htm"
End Sub

Private Sub sosFrame1_Click()

End Sub
