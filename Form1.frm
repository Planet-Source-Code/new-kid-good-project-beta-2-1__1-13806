VERSION 5.00
Object = "{D2303D7E-7DAC-11D3-8F49-00104B4D60E0}#6.0#0"; "SOSGENERALCONTROLS.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Good Project Beta 2"
   ClientHeight    =   840
   ClientLeft      =   1635
   ClientTop       =   3165
   ClientWidth     =   12270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   12270
   Begin sosGeneralControls.sosFrame sosFrame1 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   1720
      Caption         =   "Good Project Beta 2"
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3720
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0D1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1172
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":15C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3702
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3B56
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3FAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":43FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4852
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4CA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":50FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":554E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":59A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5DF6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   0
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "run"
               Object.ToolTipText     =   "Run"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "conver"
               Object.ToolTipText     =   "Converts Imperial to Metric"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "dice"
               Object.ToolTipText     =   "A Dice Roller"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "text"
               Object.ToolTipText     =   "A Notepad Replacement"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "fileexp"
               Object.ToolTipText     =   "File Explorer"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "game1"
               Object.ToolTipText     =   "Match the Image 1.0"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "image"
               Object.ToolTipText     =   "Image Viewer 1.0a"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "multi"
               Object.ToolTipText     =   "Multi On 1.0"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "pic"
               Object.ToolTipText     =   "Project Information Center 1.0"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tip"
               Object.ToolTipText     =   "Tip of the Day"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "about"
               Object.ToolTipText     =   "About Good Project"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "exit"
               Object.ToolTipText     =   "Exit Good Project"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "vote"
               Object.ToolTipText     =   "Click here to Help me continue my work"
               ImageIndex      =   14
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "Form1.frx":624A
         Begin VB.Timer Timer1 
            Interval        =   5000
            Left            =   11640
            Top             =   120
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H00CFA77F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   11840
         TabIndex        =   2
         Top             =   0
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10800
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
On Error GoTo endsube
'Form1.Animation1.open "C:\mp.avi"
Top = 0
Left = 0
Width = 12120
ExplodeForm Form1, 3500
Form1.Label1.Caption = Time
Form1.Label2.Caption = " " & Right(Time, 2)
frmTip.Show
endsube:
End Sub

Private Sub Timer1_Timer()
Form1.Label1.Caption = Time
Form1.Label2.Caption = " " & Right(Time, 2)
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "exit"
    End
    Case "game1"
    Random1.Show
    Case "vote"
    Vote1.Show
    Case "image"
    PicView1.Show
    Case "pic"
    Pic1.Show
    Case "conver"
    ConVer.Show
    Case "tip"
    Hid.Label1.Caption = "1"
    frmTip.Show
    Case "fileexp"
    FileExp.Show
    Case "run"
    ExeCut.Show
    Case "text"
    Edit1.Show
    Case "dice"
    Dice1.Show
    Case "multi"
    Multi1.Show
    Case "about"
    frmAbout.Show
    Case Else
    'Do nothing
End Select
End Sub
