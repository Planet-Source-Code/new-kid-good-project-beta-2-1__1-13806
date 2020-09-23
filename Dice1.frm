VERSION 5.00
Begin VB.Form Dice1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dice Roller"
   ClientHeight    =   3870
   ClientLeft      =   2190
   ClientTop       =   3600
   ClientWidth     =   5970
   Icon            =   "Dice1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Dice1.frx":030A
   ScaleHeight     =   3870
   ScaleWidth      =   5970
   Begin VB.CommandButton Command2 
      Caption         =   "&Roll"
      Height          =   360
      Left            =   1215
      TabIndex        =   8
      Top             =   3165
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Roll"
      Height          =   1215
      Left            =   15240
      TabIndex        =   4
      Top             =   11040
      Width           =   3975
   End
   Begin VB.TextBox Sides 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Text            =   "6"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Rolls 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Text            =   "1"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Modifier 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Text            =   "0"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Answer 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Text            =   "0"
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   960
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "Dice1.frx":4BD50
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Modifier :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   3
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rolls :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sides :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
      Visible         =   0   'False
      Begin VB.Menu clear 
         Caption         =   "&Clear Answer Box"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Dice1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub about_Click()
Msg = "Version 2.00.0001"
Style = vbOKOnly
Title = "About Dice"
help = "WOW"
Ctxt = 1000
Response = MsgBox(Msg, Style, Title, help, Ctxt)
End Sub

Private Sub clear_Click()
Answer = "0"
End Sub

Private Sub Command1_Click()
Random = 0
Answer = 0
total = 0
Result = 0

start:

Random = Int(Rnd(1) * Sides + 1)
total = total + 1
Result = Random + Result

If total = Rolls Then
GoTo finish
Else
GoTo start
End If
finish:
Answer = Result + Modifier

End Sub


Private Sub Command2_Click()
Randomize Timer
Rol = 0
Answer = 0
total = 0
Result = 0
'Look for No Rolls and Finish
If Rolls < 0 Then
Answer = Modifier
GoTo DONE
End If
'Look for No Sides and Finish
If Sides = 0 Then
Answer = Modifier
GoTo DONE
End If
'Roll The Dice
Roll:
Result = Int((Rnd * Sides) + 1)
If Rol = Rolls Then
GoTo AddMod
Else
Rol = Rol + 1
GoTo Roll
End If
AddMod:
Result = Result + Modifier
GoTo DONE
Error:
Result = "Error, Try Something else."
DONE:
Answer = Result
End Sub

Private Sub exit_Click()
A$ = Form1.Top
B$ = Form1.Left
Open "DIDAT.EYE" For Output As #1
Write #1, A$
Write #1, B$
Close #1
Unload Me
End Sub

Private Sub Form_Load()
Left = 200
Top = Form1.Height + 200
ExplodeForm Dice1, 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
A$ = Form1.Top
B$ = Form1.Left
Open "DIDAT.EYE" For Output As #1
Write #1, A$
Write #1, B$
Close #1
End Sub

Private Sub Image1_Click()
On Error GoTo finish3
Random = 0
Answer = 0
total = 0
Result = 0

start:
If Rolls < 1 Then
    GoTo finish3
End If
If Sides = 0 Then
    Answer = Modifier
    GoTo DONE
End If
    Randomize Timer
    Random = Int(Rnd(1) * Sides + 1)
    total = total + 1
    Result = Random + Result
If total = Rolls Then
    GoTo finish
    Else
    GoTo start
End If
finish:
    Answer = Result + Modifier
    GoTo DONE
finish3:
    Answer = "There was a error"
DONE:
End Sub





Private Sub Image3_Click()
PopupMenu Home
End Sub


Private Sub Label3_Click()
Label2.ForeColor = &HC0&
Random = 0
Answer = 0
total = 0
Result = 0

start:
If Sides = 0 Then
Answer = Modifier
GoTo finish3
End If
On Error GoTo finish2
If Rolls < 1 Then
GoTo finish2
End If
Randomize Timer
Random = Int(Rnd(1) * Sides + 1)
total = total + 1
Result = Random + Result

If total = Rolls Then
GoTo finish
Else
GoTo start
End If
finish:
Answer = Result + Modifier
GoTo finish3

finish2:
Answer = "ERROR"

finish3:

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Label2.ForeColor = &HE0E0E0
Label2.ForeColor = &HC0C0&
End Sub


