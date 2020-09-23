VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOpts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmopts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin TabDlg.SSTab SSTab1 
      Height          =   2220
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   3916
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Default web"
      TabPicture(0)   =   "frmopts.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FR(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Interface"
      TabPicture(1)   =   "frmopts.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FR(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Documents"
      TabPicture(2)   =   "frmopts.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FR(3)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Miscellaneous"
      TabPicture(3)   =   "frmopts.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FR(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame FR 
         Caption         =   "Preferences"
         Height          =   1725
         Index           =   3
         Left            =   -74910
         TabIndex        =   22
         Top             =   405
         Width           =   4110
         Begin VB.TextBox txUser 
            Height          =   315
            Left            =   90
            TabIndex        =   28
            Top             =   450
            Width           =   3930
         End
         Begin VB.CheckBox chImage 
            Caption         =   "Use Internal &Image viewer for supported formats"
            Height          =   195
            Left            =   90
            TabIndex        =   27
            Top             =   1440
            Value           =   1  'Checked
            Width           =   3840
         End
         Begin VB.OptionButton opVM 
            Caption         =   "No &Wrap"
            Height          =   195
            Index           =   0
            Left            =   1035
            TabIndex        =   26
            Top             =   900
            Width           =   960
         End
         Begin VB.OptionButton opVM 
            Caption         =   "Wo&rd wrap"
            Height          =   195
            Index           =   1
            Left            =   1980
            TabIndex        =   25
            Top             =   900
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton opVM 
            Caption         =   "P&rinter"
            Height          =   195
            Index           =   2
            Left            =   3105
            TabIndex        =   24
            Top             =   900
            Width           =   825
         End
         Begin VB.CheckBox chDoc 
            Caption         =   "Display the D&ocument's HTML tags in the outline"
            Height          =   195
            Left            =   90
            TabIndex        =   23
            Top             =   1215
            Width           =   3750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "A&uthor name to insert in META tags:"
            Height          =   195
            Left            =   90
            TabIndex        =   30
            Top             =   225
            Width           =   2610
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "View &Mode:"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   900
            Width           =   825
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Preferences"
         Height          =   1725
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Top             =   405
         Width           =   4110
         Begin VB.TextBox txDefW 
            Height          =   315
            Left            =   1125
            TabIndex        =   19
            Top             =   540
            Width           =   2850
         End
         Begin VB.Label Label2 
            Caption         =   "Specify a web to be loaded automatically on startup. Leave this box blank to avoid using this feature."
            ForeColor       =   &H8000000C&
            Height          =   420
            Left            =   135
            TabIndex        =   21
            Top             =   900
            Width           =   3840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Def&ault web:"
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   585
            Width           =   930
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Preferences"
         Height          =   1725
         Index           =   1
         Left            =   -74910
         TabIndex        =   11
         Top             =   405
         Width           =   4110
         Begin VB.ComboBox cbSizes 
            Height          =   315
            Left            =   3375
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   360
            Width           =   600
         End
         Begin VB.CheckBox chTB 
            Caption         =   "Show &Toolbar in main MDI window"
            Height          =   195
            Left            =   135
            TabIndex        =   14
            Top             =   900
            Value           =   1  'Checked
            Width           =   2805
         End
         Begin VB.CheckBox chFL 
            Caption         =   "Highlight Text When Found"
            Height          =   195
            Left            =   135
            TabIndex        =   13
            Top             =   1125
            Value           =   1  'Checked
            Width           =   2670
         End
         Begin VB.CheckBox chSB 
            Caption         =   "Show &StatusBar on the bottom edge"
            Height          =   195
            Left            =   135
            TabIndex        =   12
            Top             =   1350
            Value           =   1  'Checked
            Width           =   2940
         End
         Begin MSComctlLib.ImageCombo imFonts 
            Height          =   330
            Left            =   1035
            TabIndex        =   16
            Top             =   360
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "Select Editor Font"
            ImageList       =   "imlFonts"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Editor F&ont:"
            Height          =   195
            Left            =   135
            TabIndex        =   17
            Top             =   435
            Width           =   855
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Preferences"
         Height          =   1725
         Index           =   2
         Left            =   -74910
         TabIndex        =   5
         Top             =   405
         Width           =   4110
         Begin VB.TextBox txBAt 
            Height          =   315
            Left            =   90
            TabIndex        =   8
            Text            =   "alink=""#FF0000"" vlink=""#000080"" link=""#0000FF"""
            Top             =   405
            Width           =   3930
         End
         Begin VB.TextBox txComments 
            Height          =   315
            Left            =   90
            TabIndex        =   7
            Top             =   945
            Width           =   3930
         End
         Begin VB.CheckBox chAC 
            Caption         =   "&Automatically convert plain URLs to hyperlinks"
            Height          =   195
            Left            =   90
            TabIndex        =   6
            Top             =   1395
            Value           =   2  'Grayed
            Width           =   3930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Pre-sp&ecified <BODY> attributes:"
            Height          =   195
            Left            =   90
            TabIndex        =   10
            Top             =   215
            Width           =   2430
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Code comm&ents to insert in header:"
            Height          =   195
            Left            =   90
            TabIndex        =   9
            Top             =   750
            Width           =   2565
         End
      End
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2280
      Width           =   960
   End
   Begin VB.CommandButton cmNo 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2280
      Width           =   960
   End
   Begin VB.ComboBox cbSort 
      Height          =   315
      Left            =   90
      Locked          =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2295
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MSComctlLib.ImageList imlFonts 
      Left            =   5160
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmopts.frx":007C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmopts.frx":0618
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmopts.frx":0BB4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Test 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   5160
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "frmOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'######################################
'Designer 1.2 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
Option Explicit

Private Sub cmNo_Click()

    Unload Me

End Sub

Private Sub cmOK_Click()

    On Error Resume Next
      If imFonts.SelectedItem Is Nothing Then imFonts.SelectedItem = imFonts.ComboItems(1) ':( Expand Structure
      SaveValue "DocumentTree", CBool(chDoc.value)
      SaveValue "FontName", imFonts.SelectedItem.text
      SaveValue "FontSize", cbSizes.text
      SaveValue "SelectFind", CBool(chFL.value)
      SaveValue "ToolBar", chTB.value
      SaveValue "StatusBar", chSB.value
      SaveValue "BodyAttrib", txBAt.text
      SaveValue "Comments", txComments.text
      SaveValue "Author", txUser.text
      SaveValue "ImageViewer", chImage.value
      SaveValue "ViewMode", GetVM()
      SaveValue "AutoComplete", CBool(chAC.value)

      'SaveSetting App.Title, "Options", "DocumentTree", CBool(chDoc.Value)
      'SaveSetting App.Title, "Options", "FontName", imFonts.Index
      'SaveSetting App.Title, "Options", "FontSize", cbSizes.text
      'SaveSetting App.Title, "Options", "SelectFind", CBool(chFL.Value)
      'SaveSetting App.Title, "Options", "StatusBar", chSB.Value
      'SaveSetting App.Title, "Options", "BodyAttrib", txBAt.text
      'SaveSetting App.Title, "Options", "Comments", txComments.text
      'SaveSetting App.Title, "Options", "Author", txUser.text
      'SaveSetting App.Title, "Options", "ImageViewer", chImage.Value
      'SaveSetting App.Title, "Options", "ViewMode", GetVM()
      'SaveSetting App.Title, "Options", "AutoComplete", CBool(chAC.Value)

      Unload Me
      frmMain.SetPrefs
      frmMain.GetPrefs

End Sub ':( On Error Resume still active

Private Sub Form_Load()

    On Error Resume Next
      Call AddFonts
      SetFont Me

      chAC.value = CBinary(ReadValue("AutoComplete", False))
      chDoc.value = CBinary(ReadValue("DocumentTree"))
      chImage.value = ReadValue("ImageViewer", 1)
      cbSizes.text = CInt(ReadValue("FontSize"))
      imFonts.SelectedItem = imFonts.ComboItems(ReadValue("FontName"))
      txDefW.text = ReadValue("WebDefault")
      chFL.value = CBinary(ReadValue("SelectFind"))
      chTB.value = CBinary(ReadValue("ToolBar"))
      chSB.value = CBinary(ReadValue("StatusBar"))
      txBAt.text = ReadValue("BodyAttrib")
      txComments.text = ReadValue("Comments", "<!-- Created with Designer Pro //-->")
      txUser.text = ReadValue("Author")
      opVM(CInt(ReadValue("ViewMode"))).value = True

End Sub ':( On Error Resume still active

Sub AddFonts()

  Dim i As Integer, IC As Integer

    For i = 0 To Screen.FontCount - 1
        cbSort.AddItem Screen.Fonts(i)
    Next i

    For i = 8 To 72 Step 2
        cbSizes.AddItem i
    Next i

    For i = 0 To cbSort.ListCount - 1
        Test.FontName = cbSort.List(i)
        Test.FontSize = 72.75
        If Test.FontSize = 72.75 Then IC = 2 Else IC = 0 ':( Expand Structure
        imFonts.ComboItems.Add i + 1, cbSort.List(i), cbSort.List(i), IC
    Next i

End Sub

Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)

    FR(Node.index - 1).ZOrder vbBringToFront

End Sub

Function GetVM() As Integer

  Dim i As Integer

    For i = 0 To 2
        If opVM(i).value Then GetVM = i: Exit Function ':( Expand Structure
    Next i

End Function

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:54 PM) 8 + 102 = 110 Lines
