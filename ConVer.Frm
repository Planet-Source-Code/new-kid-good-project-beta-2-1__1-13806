VERSION 5.00
Object = "{A48A1D86-6463-4F3A-B152-6051F2FA65A1}#1.0#0"; "GURHANCOOLBUTTON.OCX"
Begin VB.Form ConVer 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conversions"
   ClientHeight    =   3480
   ClientLeft      =   2220
   ClientTop       =   4110
   ClientWidth     =   8160
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "ConVer.frx":0000
   ScaleHeight     =   3480
   ScaleWidth      =   8160
   Begin Gurhan_Cool_Button.GurhanCoolButton GurhanCoolButton2 
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   3110
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      CaptionTitle    =   "Close"
      Picture         =   "ConVer.frx":5C0E2
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
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      CaptionTitle    =   "Convert"
      Picture         =   "ConVer.frx":5F014
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
   Begin VB.OptionButton Option8 
      BackColor       =   &H00000000&
      Caption         =   "Mile to Kilometer"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   1440
      Width           =   2895
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00000000&
      Caption         =   "Yard to Meter"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3720
      TabIndex        =   8
      Top             =   1200
      Width           =   3015
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00000000&
      Caption         =   "Foot to Meter"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Width           =   3015
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00000000&
      Caption         =   "Inch to Millimeter"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Width           =   3015
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00000000&
      Caption         =   "Millimeter to Inch"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   3135
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Kilometer to Mile"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Text            =   "Your Answer"
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "1"
      Top             =   360
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Meters to Yards"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Centimeter to Inch"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3135
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   3000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Type Here:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Read the Output Here:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   3360
      Y1              =   1560
      Y2              =   600
   End
End
Attribute VB_Name = "ConVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload ConVer
End Sub


Private Sub Command2_Click()
If Option5 = True Then
Text3 = Text1 * 25.4
GoTo Next4305
End If
If Option6 = True Then
Text3 = Text1 * 0.3048
GoTo Next4305
End If
If Option7 = True Then
Text3 = Text1 * 0.9144
GoTo Next4305
End If
If Option8 = True Then
Text3 = Text1 * 1.609
GoTo Next4305
End If
If Option1 = True Then
Text3 = Text1 * 0.394
GoTo Next4305
End If
If Option3 = True Then
Text3 = Text1 * 0.6214
GoTo Next4305
End If
If Option4 = True Then
Text3 = Text1 * 0.039
GoTo Next4305
End If
If Option2 = True Then
Text3 = Text1 * 1.094
Else
MsgBox "Nothing to do"
End If
Next4305:
End Sub


Private Sub Image2_Click()
On Error GoTo Next4306
If Option5 = True Then
Text3 = Text1 * 25.4
GoTo Next4305
End If
If Option6 = True Then
Text3 = Text1 * 0.3048
GoTo Next4305
End If
If Option7 = True Then
Text3 = Text1 * 0.9144
GoTo Next4305
End If
If Option8 = True Then
Text3 = Text1 * 1.609
GoTo Next4305
End If
If Option1 = True Then
Text3 = Text1 * 0.394
GoTo Next4305
End If
If Option3 = True Then
Text3 = Text1 * 0.6214
GoTo Next4305
End If
If Option4 = True Then
Text3 = Text1 * 0.039
GoTo Next4305
End If
If Option2 = True Then
Text3 = Text1 * 1.094
Else
MsgBox "Nothing to do"
End If
Next4305:
GoTo Next4307
Next4306:
MsgBox "There was a Error. Restart the Program"
End
Return
Next4307:
End Sub


Private Sub Image3_Click()
Unload ConVer
End Sub


Private Sub Image4_Click()
CommonDialog1.ShowOpen
Open CommonDialog1.FileName For Input As #1
Input #1, A$
Input #1, B$
Input #1, C$
Input #1, d$
Input #1, e$
Close #1
Label3 = A$
Label4 = B$
Label5 = C$
Label6 = d$
Label7 = e$
End Sub


Private Sub Form_Load()
Left = 200
Top = Form1.Height + 200
ExplodeForm ConVer, 3500
End Sub

Private Sub GurhanCoolButton1_CLICKED()
On Error GoTo Next4306
If Option5 = True Then
Text3 = Text1 * 25.4
GoTo Next4305
End If
If Option6 = True Then
Text3 = Text1 * 0.3048
GoTo Next4305
End If
If Option7 = True Then
Text3 = Text1 * 0.9144
GoTo Next4305
End If
If Option8 = True Then
Text3 = Text1 * 1.609
GoTo Next4305
End If
If Option1 = True Then
Text3 = Text1 * 0.394
GoTo Next4305
End If
If Option3 = True Then
Text3 = Text1 * 0.6214
GoTo Next4305
End If
If Option4 = True Then
Text3 = Text1 * 0.039
GoTo Next4305
End If
If Option2 = True Then
Text3 = Text1 * 1.094
Else
MsgBox "Nothing to do"
End If
Next4305:
GoTo Next4307
Next4306:
MsgBox "There was a Error. Restart the Program"
End
Return
Next4307:
End Sub

Private Sub GurhanCoolButton2_CLICKED()
Unload Me
End Sub
