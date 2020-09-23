VERSION 5.00
Object = "{6340742D-986A-11D3-9155-00104B47E7E6}#48.0#0"; "EX.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ExList example"
   ClientHeight    =   8640
   ClientLeft      =   3195
   ClientTop       =   1860
   ClientWidth     =   8280
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   8280
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   255
      Left            =   2280
      TabIndex        =   44
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   6240
      TabIndex        =   34
      Top             =   0
      Width           =   1935
      Begin VB.CheckBox Check2 
         Caption         =   "Hidden Files Ghosted"
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Auto update List"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Show System"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Show Hidden"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Show Archive"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Show ReadOnly"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Folder"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command11 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   33
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "About..."
      Height          =   375
      Left            =   5280
      TabIndex        =   32
      Top             =   6240
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form1.frx":1272
      Left            =   6480
      List            =   "Form1.frx":127C
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   6600
      Width           =   1695
   End
   Begin ex.ExList ExList1 
      Height          =   5535
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9763
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HideSelection   =   0   'False
      MouseIcon       =   "Form1.frx":1291
      View            =   3
      HiddenGhosted   =   -1  'True
      AutoUpdateList  =   -1  'True
      SortHeaderClick =   -1  'True
      Language        =   1
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Open with Associated"
      Height          =   375
      Left            =   1680
      TabIndex        =   25
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Properties"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Sort 'Size'"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6480
      TabIndex        =   21
      Top             =   7725
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Columnheader Text 'Name'"
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   7680
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Columnheader Width 'Name'"
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   7200
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6480
      TabIndex        =   18
      Top             =   7245
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove Subitem ""Größe"""
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete to Recycle Bin"
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Style"
      Height          =   3135
      Left            =   6240
      TabIndex        =   9
      Top             =   2520
      Width           =   1935
      Begin VB.CheckBox Check8 
         Caption         =   "Track Selected"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Flat Columnheader"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Two Click Activate"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check5 
         Caption         =   "One Click Activate"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Fullrowselect"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Gridlines"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "You should do this at runtime."
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2640
         Width           =   1455
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":12AD
      Left            =   6480
      List            =   "Form1.frx":12BD
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   300
      Left            =   7800
      TabIndex        =   3
      ToolTipText     =   "Browse"
      Top             =   8190
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Top             =   8205
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   5760
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":12E9
      Left            =   120
      List            =   "Form1.frx":12F9
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      Height          =   195
      Left            =   4080
      TabIndex        =   43
      Top             =   8250
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Bytes free"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   6240
      Width           =   870
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Language"
      Height          =   195
      Left            =   6480
      TabIndex        =   31
      Top             =   6360
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Associated EXE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   6720
      Width           =   1425
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Columnheader:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   6480
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Filename:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   6000
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "View"
      Height          =   195
      Left            =   6480
      TabIndex        =   8
      Top             =   5760
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fileextenshion"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   7080
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   6
      Top             =   5760
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Files:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------
'- Name: Ascher Stefan
'- Email: s.ascher@tirol.com
'- Web: http://stefan.scriptmania.com
'- Company:
'- Date/Time: 11.11.99 19:09:16
'----------------------------------------
'- Modulname: Form1 (Code)
'----------------------------------------
'- Notes:
'
'----------------------------------------

Option Explicit

Private Sub Check1_Click()
    'Should Folders be displayed in the list
    ExList1.ShowFolder = -Check1.Value
    
End Sub

Private Sub Check10_Click()
    'Show readonly files
    ExList1.ShowReadOnly = -Check10.Value
End Sub

Private Sub Check11_Click()
    'Show files with the attribute "Archive"
    ExList1.ShowArchive = -Check11.Value

End Sub

Private Sub Check12_Click()
    'Show hidden files
    ExList1.ShowHidden = -Check12.Value

End Sub

Private Sub Check13_Click()
    'Show files with the attrubute "system"
    ExList1.ShowSystem = -Check13.Value

End Sub

Private Sub Check2_Click()
    'Should hidden Files be ghosted like since IE4
    ExList1.HiddenFilesGhosted = -Check2.Value
    
End Sub

Private Sub Check3_Click()
    'Should Gridlines be displayed
    ExList1.GridLines = -Check3.Value
End Sub

Private Sub Check4_Click()
    'Fullrowselect
    ExList1.FullRowSelect = -Check4.Value
End Sub

Private Sub Check5_Click()
    'Style
    ExList1.OneClickActivate = -Check5.Value
End Sub

Private Sub Check6_Click()
    'Style
    ExList1.TwoClickActivate = -Check6.Value
End Sub

Private Sub Check7_Click()
    'Style
    ExList1.FlatColumnHeader = -Check7.Value
End Sub

Private Sub Check8_Click()
    'Style
    ExList1.TrackSelected = -Check8.Value
End Sub

Private Sub Check9_Click()
    ExList1.AutoUpdateList = -Check9.Value
End Sub

Private Sub Combo1_Click()
    'View, large-, small icons, list, report
    ExList1.View = Combo1.ListIndex
End Sub

Private Sub Combo2_Change()
    'File extenshion (Pattern)
    ExList1.FileSpez = Combo2.Text
End Sub

Private Sub Combo2_Click()
    'Changing the Filefilter
    ExList1.FileSpez = Combo2.Text
    ExList1.RefreshList
End Sub

Private Sub Combo3_Click()
    'Select the language, either german, or english.
    ExList1.Language = Combo3.ListIndex
End Sub

Private Sub Command1_Click()
    'Browse for Folder Dialog
    Dim Path As String
    Path = ExList1.BrowserForFolder(Me.hWnd, "Browser for folder")
    If Len(Path) > 0 Then
        Text1 = Path
        ExList1.Path = Text1
        ExList1.RefreshList
    End If
End Sub

Private Sub Command10_Click()
    ExList1.ShowAbout
End Sub

Private Sub Command11_Click()
    Unload Me
End Sub

Private Sub Command12_Click()
    If Form1.ExList1.FileName = vbDirectory Then
    MsgBox "dir"
    End If
End Sub

Private Sub Command2_Click()
    'Refresh the List, it could take a little time (it's not C/C++)
    ExList1.Path = Text1.Text
    ExList1.RefreshList

End Sub

Private Sub Command3_Click()
    'You should do this before you fill the list with files,
    ExList1.RemoveSubitem exlSize
    'then you had not to do this.
    ExList1.RefreshList
End Sub

Private Sub Command4_Click()
    'Delete selected Item to rcicle bin.
    ExList1.DelSelFileToRecicleBin
End Sub

Private Sub Command5_Click()
    'Set the width of the Columnheader 'Name'.
    ExList1.ColumnHeaderWidth(exlName) = CLng(Text2.Text)
End Sub

Private Sub Command6_Click()
    'Set the Text of the Columnheader 'Name' (if it should not be in german).
    'OK 'Name' is the same in english, but you can do this with each header.
    ExList1.ColumnHeaderText(exlName) = Text3.Text
End Sub

Private Sub Command7_Click()
    'It is not funny to sort Dates and Times, so it don't work how it
    'should be.
    ExList1.Sort exlSize, True
End Sub

Private Sub Command8_Click()
    'Shows the Explorer Properties Dialog.
    ExList1.ShowSelFileProperties
End Sub

Private Sub Command9_Click()
    'Opens the selected File with the Associated EXE.
    ExList1.OpenSelWithEXE
End Sub

Private Sub ExList1_BeginRefresh(ValidPath As Boolean)
    'Is it a path, that exists?
    Debug.Print "Valid Path: " & ValidPath
    'Start Refreshing, it can take a little time.
    If ValidPath Then Me.MousePointer = 11

End Sub

Private Sub ExList1_ColumnClick(Column As ex.exlHeaderConst)
    'Display the index of the Columnheader, you clicked.
    Label7.Caption = "Columnheader: " & Column

End Sub

Private Sub ExList1_ItemClick(FileName As String)
    'File selected
    Label6.Caption = "Filename: " & FileName
    Label8.Caption = "Associated EXE: " & _
    IIf(Len(ExList1.SelAssosiatedEXE) = 0, "None", ExList1.SelAssosiatedEXE)
End Sub

Private Sub ExList1_ItemDblClick(FileName As String)
    'It occured a doubleclick on an Item, so lets open the file.
    ExList1.OpenSelWithEXE

End Sub

Private Sub ExList1_RefreshReady(Time As Long)
    'Ready with refreshing, and displaying the time it taked
    Label3.Caption = "Time: " & Time & " ms"
    Label4.Caption = "Files: " & ExList1.FileCount & " (" & Format$(ExList1.BytesUsed / 1024, "0.0#") & " KB)"
    Label10.Caption = Format$(ExList1.FreeByteOnDisk("C:\") / 1024, "0.0#") & " KB free on C:\"
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    'Initialize some stuff
    ExList1.FileSpez = "*.*"
    ExList1.Path = App.Path
    Text1 = ExList1.Path
    Combo1.ListIndex = 3
    Combo2.ListIndex = 0
    Combo3.ListIndex = ExList1.Language
    
    Check1.Value = -ExList1.ShowFolder
    Check2.Value = -ExList1.HiddenFilesGhosted
    Check3.Value = -ExList1.GridLines
    Check4.Value = -ExList1.FullRowSelect
    Check5.Value = -ExList1.OneClickActivate
    Check6.Value = -ExList1.TwoClickActivate
    Check7.Value = -ExList1.FlatColumnHeader
    Check8.Value = -ExList1.TrackSelected
    Check9.Value = -ExList1.AutoUpdateList
    Check10.Value = -ExList1.ShowReadOnly
    Check11.Value = -ExList1.ShowArchive
    Check12.Value = -ExList1.ShowHidden
    Check13.Value = -ExList1.ShowSystem
    Text2.Text = ExList1.ColumnHeaderWidth(exlName)
    Text3.Text = ExList1.ColumnHeaderText(exlName)
    ExList1.FileSpez = Combo2.Text
End Sub


