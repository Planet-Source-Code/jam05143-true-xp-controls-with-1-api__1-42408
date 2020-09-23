VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2580
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Width           =   2475
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1635
      Left            =   6600
      TabIndex        =   19
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2884
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   1635
      Left            =   4740
      TabIndex        =   18
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2884
      _Version        =   327682
      PathSeparator   =   ""
      Style           =   7
      Appearance      =   1
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   2820
      TabIndex        =   17
      Top             =   3660
      Width           =   1755
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   240
      TabIndex        =   16
      Top             =   4140
      Width           =   2355
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   3660
      Width           =   2355
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1740
      TabIndex        =   14
      ToolTipText     =   "Other Buttons are white"
      Top             =   300
      Width           =   1215
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   5445
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "XP Demo!"
            TextSave        =   "XP Demo!"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   12224
            Text            =   "This Program Will Show WinXP Controls on a machine running XP with the XP theme enabled     "
            TextSave        =   "This Program Will Show WinXP Controls on a machine running XP with the XP theme enabled     "
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
      EndProperty
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   855
      Left            =   2160
      TabIndex        =   12
      Top             =   1380
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   300
      TabIndex        =   11
      Top             =   1980
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1155
      Left            =   7320
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3060
      Top             =   1320
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   975
      Left            =   5640
      TabIndex        =   8
      Top             =   900
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1720
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   3120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   556
      _Version        =   327682
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   2700
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   1500
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   3600
      TabIndex        =   5
      Top             =   840
      Width           =   1875
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   300
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   1020
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3180
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Default Button Is Blue"
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   435
      Left            =   5820
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "Even The MessageBoxes Are WinXP Style!", vbExclamation Or vbOKOnly, "XP Demo"
End Sub

Private Sub Command2_Click()
    Dim RetStr As String
    RetStr = InputBox("Even The InputBoxes Are XP Style!", "XP Demo")
    
End Sub

Private Sub Form_Initialize()
     'Call the sub in here, befor the form is shown
     InitXP
End Sub

Private Sub Form_Load()
    Me.Show
End Sub

Private Sub Timer1_Timer()
    
    'Demonstration of the progressbar control
    Static i As Integer
    i = i + 1
    If i >= 1500 Then
        Timer1.Enabled = False
        Exit Sub
    End If
    ProgressBar1.Value = i
    
End Sub
