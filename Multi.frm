VERSION 5.00
Begin VB.Form Multi1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Enable"
   ClientHeight    =   1785
   ClientLeft      =   2205
   ClientTop       =   2505
   ClientWidth     =   6300
   Icon            =   "Multi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Multi.frx":08CA
   ScaleHeight     =   1785
   ScaleWidth      =   6300
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5880
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   795
      Left            =   5400
      Picture         =   "Multi.frx":25C98
      ScaleHeight     =   735
      ScaleWidth      =   810
      TabIndex        =   3
      Top             =   480
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enable The F8 Button to Activate Hot Key for This Form"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Lock Controls"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Unlock Controls"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Li 
      Caption         =   "1"
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   4800
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Resolution:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Window Version:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "Multi1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvparam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Const WM_SETHOTKEY = &H32
Const VK_PAUSE = &H77
Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
End Function

Sub GetHiLoWord(X As Long, LoWord As Integer, HiWord As Integer)
    LoWord = CInt(X And &HFFFF&)
    HiWord = CInt(X \ &H10000)
End Sub
Sub GetHiLoByte(X As Integer, LoByte As Integer, HiByte As Integer)
    LoByte = X And &HFF&
    HiByte = X \ &H100
End Sub







Private Sub Command1_Click()
    Dim i As Long
    i = SendMessage(Me.hWnd, WM_SETHOTKEY, VK_PAUSE, 0)
    MsgBox "Press The F8 Key to Activate"
End Sub


Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    Left = 200
Top = Form1.Height + 200
ExplodeForm Multi1, 3500
    Dim lpvparam As Boolean
    Dim X As Long

    X = SystemParametersInfo(97, False, lpvparam, 0)
'Computer Name
    Dim PCName As String
    Dim P As Long
    P = NameOfPC(PCName)
    Label2.Caption = PCName
    
'Res
    Dim XTwips As Long
    Dim YTwips As Long
    Dim XPixels As Long
    Dim YPixels As Long

    XTwips = Screen.TwipsPerPixelX
    YTwips = Screen.TwipsPerPixelY

    YPixels = Screen.Height / YTwips
    XPixels = Screen.Width / XTwips

    Label3.Caption = Str$(XPixels) + " x " + Str$(YPixels)

'Win Ver
    Dim WinMajor As Integer
    Dim WinMinor As Integer
    Dim DosMajor As Integer
    Dim DosMinor As Integer
    Dim RetLong As Long
    Dim LoWord As Integer
    Dim HiWord As Integer
    RetLong = GetVersion()
    Call GetHiLoWord(RetLong, LoWord, HiWord)
    Call GetHiLoByte(LoWord, WinMajor, WinMinor)
    Call GetHiLoByte(HiWord, DosMinor, DosMajor)
    Label4.Caption = WinMajor & "." & WinMinor
'    Text2.Text = "DOS version:" & DosMajor & "." & DosMinor


End Sub


Private Sub Option1_Click()
 X = SystemParametersInfo(97, False, lpvparam, 0)
End Sub


Private Sub Option2_Click()
 X = SystemParametersInfo(97, True, lpvparam, 0)
End Sub




Private Sub Timer1_Timer()
Select Case Li.Caption
Case "1"
Shape1.FillColor = RGB(255, 0, 0)
Li.Caption = "2"
Case "2"
Shape1.FillColor = RGB(0, 255, 0)
Li.Caption = "3"
Case "3"
Shape1.FillColor = RGB(0, 0, 255)
Li.Caption = "1"
End Select
End Sub


Private Sub Picture1_DblClick()
MsgBox "WOW"
End Sub


