VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmChild 
   AutoRedraw      =   -1  'True
   Caption         =   "Untitled"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmchild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   7665
   Tag             =   "This is the Editor."
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6360
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtWebSite 
      Height          =   285
      Left            =   6960
      TabIndex        =   16
      Top             =   2505
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUrl 
      Height          =   285
      Left            =   6960
      TabIndex        =   15
      Top             =   2220
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtBusiness 
      Height          =   285
      Left            =   6960
      TabIndex        =   14
      Top             =   1935
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   6960
      TabIndex        =   13
      Top             =   1650
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCompany 
      Height          =   285
      Left            =   6960
      TabIndex        =   12
      Top             =   1365
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   6960
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6360
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtIndent 
      Height          =   285
      Left            =   6360
      TabIndex        =   9
      Text            =   "&nbsp;&nbsp;&nbsp;&nbsp;"
      Top             =   3495
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6360
      TabIndex        =   8
      Top             =   3210
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6360
      TabIndex        =   7
      Top             =   2925
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   6720
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox p2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   945
      Picture         =   "frmchild.frx":0E42
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2745
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   450
      Picture         =   "frmchild.frx":1C84
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   2745
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer tmTimer 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox pD 
      AutoRedraw      =   -1  'True
      Height          =   3030
      Left            =   1320
      ScaleHeight     =   2970
      ScaleWidth      =   4410
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "This is the margin selection box. It helps to keep things neat."
      Top             =   120
      Width           =   4470
      Begin RichTextLib.RichTextBox RTF1 
         Height          =   2580
         Left            =   405
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4551
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmchild.frx":2AC6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   360
         Y1              =   540
         Y2              =   2250
      End
   End
   Begin Designer.TabCtl TabCtl1 
      Height          =   240
      Left            =   45
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Switch between Source and Preview mode (Ctrl+Space)"
      Top             =   3060
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   423
   End
   Begin SHDocVwCtl.WebBrowser IE1 
      Height          =   3255
      Left            =   840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   5741
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
   Begin VB.Label lbM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   240
      Left            =   4365
      TabIndex        =   6
      Top             =   3195
      Width           =   270
   End
   Begin VB.Menu mnuWhatever 
      Caption         =   "&What"
      Visible         =   0   'False
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Update"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilenew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFilenewWeb 
         Caption         =   "New Web"
      End
      Begin VB.Menu mnuFilenewWebBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "C&lose"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "S&ave as..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Sa&ve all"
      End
      Begin VB.Menu mnuFileSepBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRevert 
         Caption         =   "Re&vert..."
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrinterSetup 
         Caption         =   "Print s&etup..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSepMRU 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFTP 
         Caption         =   "FTP"
         Enabled         =   0   'False
         Begin VB.Menu mnuSetupFTP 
            Caption         =   "Setup FTP Info"
         End
         Begin VB.Menu mnuSendFTP 
            Caption         =   "Send This Document"
         End
      End
      Begin VB.Menu mnuFileDetails 
         Caption         =   "My Details"
      End
      Begin VB.Menu FTPBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "E&dit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "U&ndo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "Cle&ar"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select a&ll"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDefinition 
         Caption         =   "De&finition"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "R&eplace..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditClean 
         Caption         =   "&Mark clean"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuDocument 
      Caption         =   "D&ocument"
      Begin VB.Menu mnuDocStruct 
         Caption         =   "Document Structure"
         Begin VB.Menu mnuDocCleanup 
            Caption         =   "Cleanup Code"
         End
         Begin VB.Menu mnuDocCompact 
            Caption         =   "Co&mpact Code"
         End
         Begin VB.Menu mnuRestructure 
            Caption         =   "Restructure Code"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuConfigFormat 
            Caption         =   "Configure Formatting"
         End
      End
      Begin VB.Menu mnuDocBar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuDocBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Color Scheme"
      End
      Begin VB.Menu mnuDocBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDocumentConvert 
         Caption         =   "&Convert HTML to Text"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuConvertTextToHTML 
         Caption         =   "Convert &Text to HTML"
      End
   End
   Begin VB.Menu mnuEditors 
      Caption         =   "Editors"
      Begin VB.Menu mnuEditorJS 
         Caption         =   "JavaScript"
      End
      Begin VB.Menu mnuEditorCGI 
         Caption         =   "Perl/CGI"
      End
      Begin VB.Menu mnuEditorCSS 
         Caption         =   "Style Sheets"
      End
      Begin VB.Menu mnuEditorXML 
         Caption         =   "XML"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertDateTime 
         Caption         =   "D&ate/Time"
         Begin VB.Menu mnuInsertDateTimeLong 
            Caption         =   "&Long"
         End
         Begin VB.Menu mnuInsertDateTimeShort 
            Caption         =   "&Short"
         End
         Begin VB.Menu mnuBoth 
            Caption         =   "Date and Time"
         End
         Begin VB.Menu mnuInsertDateTimeSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInsertDateTimeWeekday 
            Caption         =   "&Weekday"
         End
      End
      Begin VB.Menu mnuInsertActiveXObject 
         Caption         =   "Object"
         Begin VB.Menu mnuInsertObject 
            Caption         =   "ActiveX Object"
         End
         Begin VB.Menu mnuInsertApplet 
            Caption         =   "Java Applet"
         End
      End
      Begin VB.Menu mnuInsertBar36 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertPic 
         Caption         =   "Picture"
         Begin VB.Menu mnuInsertClipart 
            Caption         =   "Clipart (Need CD)"
         End
         Begin VB.Menu mnuInsertPicFile 
            Caption         =   "From File"
         End
         Begin VB.Menu mnuInsertBar22 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInsertScanner 
            Caption         =   "From Scanner"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuInsertSymbol 
         Caption         =   "&Symbol..."
      End
      Begin VB.Menu mnuInsertStream 
         Caption         =   "S&tream"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInsertbar45 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinsertQuikScripts 
         Caption         =   "&Quick Codes"
         Begin VB.Menu mnuHtml 
            Caption         =   "HTML"
            Begin VB.Menu mnuinsertCopyrightInfo 
               Caption         =   "Copyright Information"
            End
            Begin VB.Menu mnuInsertCustomFooter 
               Caption         =   "Custom Footer..."
            End
            Begin VB.Menu mnuInsertPageUpdate 
               Caption         =   "Page Update Info"
            End
         End
         Begin VB.Menu mnuDhtml 
            Caption         =   "DHTML"
         End
         Begin VB.Menu mnuJavaScript 
            Caption         =   "JavaScript"
            Begin VB.Menu mnuJSBrowser 
               Caption         =   "Browser Detect"
            End
            Begin VB.Menu mnuJSCountdown 
               Caption         =   "Countdown Script"
            End
            Begin VB.Menu mnuJSShowDate 
               Caption         =   "Show Date Script"
            End
            Begin VB.Menu mnuJSBar34 
               Caption         =   "-"
            End
            Begin VB.Menu mnuSlideShow 
               Caption         =   "Create Slide Show"
            End
            Begin VB.Menu mnuJSMenu 
               Caption         =   "Drop and Go Menu"
            End
         End
         Begin VB.Menu mnuQuickBar 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEmail 
            Caption         =   "E-mail Link"
            Begin VB.Menu mnuPersonalMail 
               Caption         =   "Personal E-mail Address"
            End
            Begin VB.Menu mnuBusinessMail 
               Caption         =   "Business E-mail Address"
            End
         End
         Begin VB.Menu mnuUrl 
            Caption         =   "Web URL"
            Begin VB.Menu mnuPersonalUrl 
               Caption         =   "Personal Website"
            End
            Begin VB.Menu mnuBusUrl 
               Caption         =   "Business Website"
            End
         End
      End
   End
   Begin VB.Menu mnuWeb 
      Caption         =   "&Web"
      Begin VB.Menu mnuNewWeb 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuWebOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuWebClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuasdasd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuSepWeb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebDefault 
         Caption         =   "De&fault"
      End
      Begin VB.Menu mnuWebSp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebMRU 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolBar 
         Caption         =   "&Tool Bar"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "&Status Bar"
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFileTree 
         Caption         =   "File &Tree"
      End
      Begin VB.Menu mnuViewDocuments 
         Caption         =   "&Document"
      End
      Begin VB.Menu mnuViewScripts 
         Caption         =   "&ScriptView"
      End
      Begin VB.Menu mnuViewSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&No wrap"
         Index           =   1
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&Word wrap"
         Index           =   2
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&Printer"
         Index           =   3
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWysiwyg 
         Caption         =   "Preview Mode"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuEditor 
      Caption         =   "&Editor"
      Begin VB.Menu mnuEditFontSize 
         Caption         =   "Font &Size"
         Begin VB.Menu mnuFontSizeEight 
            Caption         =   "8"
         End
         Begin VB.Menu mnuFontSizeTen 
            Caption         =   "10"
         End
         Begin VB.Menu mnuFontSizeTwelve 
            Caption         =   "12"
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Wi&ndow"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascadeWin 
         Caption         =   "&Cascade"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuTileHorizontal 
         Caption         =   "&Tile Horizontal"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowSepX 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowAlignLefts 
         Caption         =   "&Align Lefts"
      End
      Begin VB.Menu mnuWindowAlignTops 
         Caption         =   "Align &Tops"
      End
      Begin VB.Menu mnuWindowSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowMinimizeAll 
         Caption         =   "&Minimize all"
      End
      Begin VB.Menu mnuWindowMaximizeAll 
         Caption         =   "Ma&ximize all"
      End
      Begin VB.Menu mnuRestoreAll 
         Caption         =   "&Restore all"
      End
      Begin VB.Menu mnuWindowUnloadAll 
         Caption         =   "&Unload all"
         Shortcut        =   ^{F9}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTodayTip 
         Caption         =   "&Today's tip"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "A&bout..."
      End
   End
   Begin VB.Menu mnuForms 
      Caption         =   "Forms (Test here)"
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bChanged As Boolean
Dim bUpdated As Boolean
Public Undo As cUNDO
Dim StrTag As String
Dim ThisPath As String
Dim CurrentType As String
Dim canExit As Boolean
Dim Num As Long
Dim q$ ':( Type Suffixes are obsolete
Dim lastN As Long
Dim n As Long
Dim W$ ':( Type Suffixes are obsolete
Dim Msg$ ':( Type Suffixes are obsolete
Dim apppath As String
Dim starttime As Date
Dim tmpchr As String * 1
Dim tmpint As Long
Dim varColorText, varColorTag, varColorProp, varColorPropVal, varColorComment As OLE_COLOR

Private Sub Form_GotFocus()

    On Error Resume Next
      RTF1_GotFocus

End Sub ':( On Error Resume still active

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
      If KeyCode = vbKeySpace And Shift = 2 Then TabCtl1.CycleTabs: RTF1.SetFocus ':( Expand Structure

End Sub ':( On Error Resume still active

Private Sub Form_Load()

    On Error Resume Next
      ThisPath = App.Path
      IE1.Navigate "about:blank"
      SetFont Me
      CopyMRUList
      WindowState = ReadValue("ChildState", 0)
      Set Undo = New cUNDO
      bUpdated = False
      mnuViewMode_Click CInt(ReadValue("ViewMode", 0)) + 1
      RTF1.Font.Name = ReadValue("FontName", "Tahoma")
      RTF1.Font.Size = ReadValue("FontSize", 8)
      RTF1.SelIndent = 45 'just a little
      bChanged = False
      SetMenus
      tmTimer_Timer 'load time
      lbM.Font.Name = "Marlett"
      mnuEdit_Click
      GetDetails
varColorText = vbBlack
varColorTag = &HC00000
varColorProp = &HC000C0
varColorPropVal = &HC000&
varColorComment = &H808080

End Sub ':( On Error Resume still active

Private Sub Form_LostFocus()

    On Error Resume Next
      RTF1_LostFocus

End Sub ':( On Error Resume still active

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next
      If bChanged Then
    Dim bMsg As VbMsgBoxResult ':( Move line to top of current Sub
          bMsg = MsgBox("The current document has changed." & vbNewLine & "Do you want to save changes to it?", vbExclamation + vbYesNoCancel, Caption)
          If bMsg = vbNo Then
              Cancel = False
            ElseIf bMsg = vbYes Then
              If Mid$(Caption, 2, 1) <> ":" Then Cancel = True ':( Expand Structure
              mnuFileSave_Click
            Else
              Cancel = True
              RTF1.SetFocus
          End If
      End If

End Sub ':( On Error Resume still active

Private Sub Form_Resize()

    On Error Resume Next
      If WindowState = 0 Then
          pD.BorderStyle = 1
        Else
          pD.BorderStyle = 0
      End If
      pD.Move 0, 0, ScaleWidth, ScaleHeight - TabCtl1.Height
      RTF1.Move 360, 0, pD.ScaleWidth - 360, pD.ScaleHeight
      IE1.Move pD.left, pD.top, pD.Width, pD.Height
      TabCtl1.Move 345, pD.Height
      Line1.Refresh
      lbM.Move ScaleWidth - lbM.Width, ScaleHeight - lbM.Height
      Line1.Y1 = 0
      Line1.Y2 = RTF1.Height
      Line1.X1 = 345
      Line1.X2 = 345

End Sub ':( On Error Resume still active

Private Sub Form_Terminate()

    On Error Resume Next
      If FormsLeft = 0 Then
          frmMain.SB.Panels(2).text = "Loc: 0, Line: 0 "
          frmMain.SB.Panels(3).text = "Size: 0 KB, Lines: 0 "
      End If
      If Not frmMain.ActiveForm Is Nothing Then
          Outline frmMain.ActiveForm, frmMain.tvD, frmMain.ActiveForm.RTF1.text, Not frmMain.SSTab1.TabVisible(1)
          AddScripts frmMain.ActiveForm.RTF1.text, frmMain.tvS, Not frmMain.SSTab1.TabVisible(2)
        Else
          frmMain.tvD.Nodes.Clear
          frmMain.tvS.Nodes.Remove frmMain.tvS.Nodes("Document").index
      End If

End Sub ':( On Error Resume still active

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
      If Caption = "Untitled" Then GoTo n ':( Expand Structure
    Dim i As Integer ':( Move line to top of current Sub
      For i = 1 To frmMain.tvW.Nodes.Count
          frmMain.tvW.Nodes(i).Bold = False
      Next i
n:       'next
      Kill FullPath(ThisPath, "temp.html")
      frmMain.ActiveForm.SetFocus
      frmMain.ActiveForm.RTF1.SetFocus
      SaveValue "ChildState", WindowState

End Sub ':( On Error Resume still active

Private Sub IE1_DownloadComplete()

    On Error Resume Next
      RTF1.SetFocus
      SendMessage IE1.hWnd, SPI_SETBORDER, ByVal 0, 0&

End Sub ':( On Error Resume still active

Private Sub IE1_StatusTextChange(ByVal text As String)

    frmMain.SB.Panels(1).text = text

End Sub

Private Sub IE1_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)

    MsgBox "Script trying to close window: Access DENIED.", vbExclamation
    Cancel = True

End Sub

Private Sub mnuArrangeIcons_Click()

    frmMain.Arrange vbArrangeIcons

End Sub

Private Sub mnuBoth_Click()

    RTF1.SelText = Now

End Sub

Private Sub mnuBusinessMail_Click()

    RTF1.SelText = "<A HREF=" & Chr$(34) & "mailto:" & txtBusiness.text & Chr$(34) & ">Contact us</A> for more information."

End Sub

Private Sub mnuBusUrl_Click()

    RTF1.SelText = "<A HREF=" & Chr$(34) & txtWebSite.text & Chr$(34) & ">Visit</A> our company's website"

End Sub

Private Sub mnuCascadeWin_Click()

    frmMain.Arrange vbCascade

End Sub

Private Sub mnuColors_Click()

    Load colors
    colors.Show

End Sub

Private Sub mnuConfigFormat_Click()
Load frmConfigFormat
frmConfigFormat.Show
End Sub

Private Sub mnuConvertTextToHTML_Click()

    cdl.ShowOpen
    Text2.text = cdl.fileName
    Text3.text = left$(cdl.fileName, Len(cdl.fileName) - 4) & ".html"

    ConvertTextToHTML

End Sub

Private Sub mnuDelete_Click()

    RTF1.SelText = ""

End Sub

Private Sub mnuDocCleanup_Click()
    RTF1.text = FormatCode(RTF1.text)

End Sub

Private Sub mnuDocCompact_Click()

    RTF1.SelStart = 0
    RTF1.SelLength = Len(RTF1.text)
    RTF1.SelText = CompactFormat(RTF1.SelText)
mnuDocCleanup.Enabled = False
End Sub

Private Sub mnuDocument_Click()

    frmMain.SB.Panels(1).text = "Contains commands for inserting symbols, converting documents, etc."

End Sub

Sub mnuDocumentConvert_Click()
    Dim lpD As New frmChild

    On Error Resume Next
      frmMain.SB.Panels(1).text = "Please wait, converting to ASCII..."
      frmMain.MousePointer = 11
      Load lpD
      Me.SetFocus
      lpD.Icon = lpD.p2.Picture
      lpD.tag = lpD.WindowState
      lpD.WindowState = vbMinimized
      lpD.RTF1.text = ""
      lpD.RTF1.text = RTF1.text
      lpD.Caption = "Converting..."
      CleanUp lpD.RTF1
      ConvEntities lpD.RTF1
      ReplaceStuff lpD.RTF1
      lpD.RTF1.SelStart = 0
      lpD.WindowState = lpD.tag
      lpD.RTF1.SetFocus
      lpD.Caption = "Untitled"
      frmMain.MousePointer = 0
      frmMain.SB.Panels(1).text = "Press F6 for Quick help or F1 for contents."

End Sub ':( On Error Resume still active

Private Sub mnuEdit_Click()

    frmMain.SB.Panels(1).text = "Contains commands for editing documents."
    mnuEditUndo.Enabled = Undo.UndoAvailable
    mnuEditRedo.Enabled = Undo.RedoAvailable
    'mnuEditCut.Enabled = frmMain.TB.Buttons("cut").Enabled
    'mnuEditCopy.Enabled = frmMain.TB.Buttons("copy").Enabled
    'mnuEditPaste.Enabled = frmMain.TB.Buttons("paste").Enabled
    mnuEditClean.Enabled = bChanged
    mnuSelectAll.Enabled = (RTF1.SelLength <> Len(RTF1.text))
    mnuEditClear.Enabled = (RTF1.text <> "")

End Sub

Private Sub mnuEditClean_Click()

    On Error Resume Next
      If MsgBox("Mark this document unchanged?" & vbNewLine & "The whole buffer will be discarded.", vbExclamation + vbYesNo, GetFile(Caption)) = vbYes Then
          bChanged = False
          Undo.ResetAll
          'frmMain.TB.Buttons("undo").Enabled = False
          'frmMain.TB.Buttons("redo").Enabled = False
      End If
      RTF1.SetFocus

End Sub ':( On Error Resume still active

Sub mnuEditClear_Click()

    RTF1.text = ""

End Sub

Sub mnuEditCopy_Click()

    Clipboard.Clear
    Clipboard.SetText RTF1.SelText

End Sub

Sub mnuEditCut_Click()

    Clipboard.Clear
    Clipboard.SetText RTF1.SelText
    RTF1.SelText = ""

End Sub

Private Sub mnuEditDefinition_Click()

    On Error Resume Next
      If CurrentType = "" Then Exit Sub ':( Expand Structure
    Dim lPos As Long, ModeLen As Long ':( Move line to top of current Sub
      lPos = InStr(1, RTF1.text, "function " & CurrentType): ModeLen = 9
      If lPos = 0 Then lPos = InStr(1, RTF1.text, "var " & CurrentType): ModeLen = 4 ':( Expand Structure
      If lPos > 0 Then
          RTF1.SelStart = lPos - 1
          If ReadValue("SelectFind", True) = False Then RTF1.SetFocus Else RTF1.SelLength = Len(CurrentType) + ModeLen ':( Expand Structure
        Else
          MsgBox "This document contains no valid script defining " & "'" & CurrentType & "'.", vbExclamation, "Definition"
          RTF1.SetFocus
      End If

End Sub ':( On Error Resume still active

Sub mnuEditFind_Click()

    frmFind.Show vbModal

End Sub

Private Sub mnuEditorJS_Click()
Load frmEditJS
frmEditJS.Show
End Sub

Sub mnuEditPaste_Click()

    On Error Resume Next
    Dim strD As String ':( Move line to top of current Sub
      strD = Clipboard.GetText(vbCFText)
      RTF1.SelText = strD

End Sub ':( On Error Resume still active

Sub mnuEditRedo_Click()

    Undo.RedoChange RTF1

End Sub

Private Sub mnuEditReplace_Click()

    frmFind.Show vbModal

End Sub

Sub mnuEditUndo_Click()

    On Error GoTo hell
    Undo.UndoChange RTF1
hell:

End Sub

Private Sub mnuFile_Click()

    frmMain.SB.Panels(1).text = "Contains commands for operating this program."

End Sub

Private Sub mnuFileClose_Click()

    Unload Me

End Sub

Private Sub mnuFileDetails_Click()

    Load frmDetails
    frmDetails.Show

End Sub

Private Sub mnuFileExit_Click()

    Unload frmMain

End Sub

Private Sub mnuFileMRU_Click(index As Integer)

    On Error Resume Next
    Dim lpF As New frmChild ':( Move line to top of current Sub
      Load lpF
      lpF.LoadHTMLFile mnuFileMRU(index).tag

End Sub ':( On Error Resume still active

Private Sub mnuFileNew_Click()

    frmMain.NewDocument

End Sub

Sub LoadHTMLFile(lpFileName As String)

    On Error Resume Next
      Select Case LCase$(right$(lpFileName, 3))
        Case "gif", "bmp", "jpg", "ico"
          Unload Me
          LoadImage lpFileName
          Exit Sub '>---> Bottom
        Case "htm", "tml", "asp", "xml"
          Icon = p1.Picture
        Case "css", "ini", "bat", "cfg", "inf", "txt", ".js", "vbs"
          Icon = p2.Picture
        Case Else
          Unload Me
          If ShellExecute(0, "open", lpFileName, "", Up1Level(lpFileName), 10) < 32 Then MsgBox "Could not execute " & GetFile(lpFileName), vbExclamation ':( Expand Structure
          Exit Sub '>---> Bottom
      End Select
      RTF1.LoadFile lpFileName, rtfText
      Caption = lpFileName
      RTF1.SetFocus
      bChanged = False
      AddFileMRU lpFileName
      'Undo.Remove Undo.Count
      'when text was loaded it was recorded as an undo action.
      'trying to undo will end up clearing the RTF.
      'Undo.Remove Undo.Count 'we need to do this twice
      'Undo.AddAction RTF1.text, RTF1.SelStart
      'this is the first action, add it to the buffer

End Sub ':( On Error Resume still active

Private Sub mnuFilenewWeb_Click()

  'frmMDI.mnuNewWeb_Click


End Sub

Private Sub mnuFileOpen_Click()

  Dim lpF As New frmChild

    On Error GoTo hell
    With frmMain.CD
        .ShowOpen
        Load lpF
        lpF.LoadHTMLFile .fileName
    End With 'FRMMAIN.CD
hell:

End Sub

Private Sub mnuFileRevert_Click()

    If left$(Caption, 8) = "Untitled" Or Mid$(Caption, 2, 1) <> ":" Then Exit Sub ':( Expand Structure
    If MsgBox("Do you want to revert to the saved version?", vbExclamation + vbYesNo) = vbYes Then
        LoadHTMLFile Caption
    End If

End Sub

Sub mnuFileSave_Click()

    If left$(Caption, 8) = "Untitled" Or Mid$(Caption, 2, 1) <> ":" Then mnuFileSaveAs_Click: Exit Sub ':( Expand Structure
    SaveHTMLFile Caption
    mnuFileSave.Enabled = False
    'frmMain.TB.Buttons(3).Enabled = False

End Sub

Sub mnuFileSaveAll_Click()

  Dim lpF As Form

    For Each lpF In Forms
        If lpF.BackColor = &H8000000F Then
            lpF.SetFocus
            lpF.mnuFileSave_Click
        End If
    Next lpF

End Sub

Private Sub mnuFileSaveAs_Click()

    On Error GoTo hell
    If left$(RTF1.text, 1) <> "<" Then frmMain.CD.FilterIndex = 3 Else frmMain.CD.FilterIndex = 1 ':( Expand Structure
    frmMain.CD.ShowSave
    SaveHTMLFile frmMain.CD.fileName
hell:
    RTF1.SetFocus

End Sub

Private Sub mnuFontSizeEight_Click()

    RTF1.Font.Size = 8

End Sub

Private Sub mnuFontSizeTen_Click()

    RTF1.Font.Size = 10

End Sub

Private Sub mnuFontSizeTwelve_Click()

    RTF1.Font.Size = 12

End Sub

Private Sub mnuForms_Click()

    Load frmForms
    frmForms.Show

End Sub

Private Sub mnuHelp_Click()

    frmMain.SB.Panels(1).text = "Contains help commands."

End Sub

Private Sub mnuHelpAbout_Click()

    frmAbout.Show vbModal

End Sub

Private Sub mnuHelpContents_Click()

    ShellExecute 0, "open", App.Path & "\help\index.html", "", "", 10

End Sub

Private Sub mnuInsertApplet_Click()

    Load ActiveXobj
    ActiveXobj.Caption = "Designer - Load Java Applet"
    ActiveXobj.Frame5.Visible = True
    ActiveXobj.Frame4.Visible = False
    ActiveXobj.Show

End Sub

Private Sub mnuinsertCopyrightInfo_Click()

  Dim NameResponse As String

    NameResponse = InputBox("Please give your name or company name.", "Author Name", txtName.text)
    RTF1.Find ("</body>")
    RTF1.SelRTF = "<hr><small><center><b>Copyright &copy;2001 by " & NameResponse & ". All Rights Reserved</b></center></small>" & vbCrLf & "</body>"
    SaveSetting App.Title, "Scripts", "Name", NameResponse

End Sub

Private Sub mnuInsertCustomFooter_Click()

    Load frmCustomFooter
    frmCustomFooter.Show

End Sub

Private Sub mnuInsertDateTimeLong_Click()

    RTF1.SelText = mnuInsertDateTimeLong.Caption

End Sub

Private Sub mnuInsertDateTimeShort_Click()

    RTF1.SelText = mnuInsertDateTimeShort.Caption

End Sub

Private Sub mnuInsertDateTimeWeekday_Click()

    RTF1.SelText = mnuInsertDateTimeWeekday.Caption

End Sub

Private Sub mnuInsertObject_Click()

    Load ActiveXobj
    ActiveXobj.Show

End Sub

Private Sub mnuInsertPageUpdate_Click()

  Dim DateResponse As String

    DateResponse = InputBox("Please enter today's date.", "Page Update Info", Now)
    RTF1.Find ("</body>")
    RTF1.SelRTF = "<center><i><small>Page updated: " & DateResponse & "</small></i></center>" & vbCrLf & "</body>"

End Sub

Private Sub mnuInsertStream_Click()

  'MousePointer = 11
  'Dim i As Integer
  'For i = 32 To 255
  'RTF1.SelText = "&#" & i & ";"
  'Next i
  'MousePointer = 0


End Sub

Private Sub mnuInsertSymbol_Click()

    On Error Resume Next
      frmCharacterMap.Show vbModal
      RTF1.SetFocus

End Sub ':( On Error Resume still active

Private Sub mnuJSBrowser_Click()

  Dim msURL, nnURL As String ':( As Variant ?

    msURL = InputBox("If the web surfer is using Internet Explorer," & vbCrLf & "which page should they go to?", "Internet Explorer URL", "http://www.IE_Web_Page.com")
    nnURL = InputBox("If the web surfer is using Netscape Navigator," & vbCrLf & "which page should they go to?", "Netscape Navigator URL", "http://www.Netscape_Web_Page.com")
RTF1.Find ("</head>")
    RTF1.SelText = "<script language = " & Chr$(34) & "JavaScript" & Chr$(34) & ">" & vbCrLf & _
                   "var name = navigator.appName;" & vbCrLf & _
                   "var vers = navigator.appVersion;" & vbCrLf & _
                   "vers = vers.substring(0,1);" & vbCrLf & _
                   "// or 0,4  could return 4.5 instead of just 4" & vbCrLf & _
                   "if (name == " & Chr$(34) & "Microsoft Internet Explorer" & Chr$(34) & ")" & vbCrLf & _
                   "{" & vbCrLf & _
                   "// You can edit this message." & vbCrLf & _
                   "  document.write(" & Chr$(34) & "You are using Internet Explorer." & Chr$(34) & ");" & vbCrLf & _
                   "// If you want to redirect your visitors to a" & vbCrLf & _
                   "// Internet Explorer-friendly version of your" & vbCrLf & _
                   "// page, use this code:" & vbCrLf & _
                   "// window.location=" & msURL & ";" & vbCrLf & _
                   "}" & vbCrLf & _
                   "Else" & vbCrLf & _
                   "{" & vbCrLf & _
                   "// You can edit this message." & vbCrLf & _
                   "  document.write(" & Chr$(34) & "You are using Netscape." & Chr$(34) & ");" & vbCrLf & _
                   "// If you want to redirect your visitors to a" & vbCrLf & _
                   "// Netscape Navigator-friendly version of your" & vbCrLf & _
                   "// page, use this code:" & vbCrLf & _
                   "// window.location=" & nnURL & ";" & vbCrLf & _
                   "}" & vbCrLf & _
                   "</script>" & vbCrLf & "</head>"

End Sub

Private Sub mnuJSCountdown_Click()

  Dim inDate, inMessage As String

    inDate = InputBox("What date am I counting down to?", "Countdown to...", "January 1, 2002")
    inMessage = InputBox("What is the occasion?", "Countdown to...", "until I finish this web page")
    RTF1.SelText = "<!-- Start count down script -->" & vbCrLf & vbCrLf & _
                   "<script language = " & Chr$(34) & "JavaScript" & Chr$(34) & ">" & vbCrLf & _
                   "var now = new Date();" & vbCrLf & _
                   "// set this value to the countdown date." & vbCrLf & _
                   "var then = new Date(" & Chr$(34) & inDate & Chr$(34) & ");" & vbCrLf & _
                   "var gap = then.getTime() - now.getTime();" & vbCrLf & _
                   "gap = Math.floor(gap / (1000 * 60 * 60 * 24));" & vbCrLf & _
                   "document.write(gap);" & vbCrLf & _
                   "</script>" & vbCrLf & _
                   "<!--your message here-->days " & inMessage & "!" & vbCrLf & _
                   "<!-- All done :-) -->"

End Sub

Private Sub mnuJSShowDate_Click()

    RTF1.SelText = "<script language = " & Chr$(34) & "JavaScript" & Chr$(34) & ">" & vbCrLf & _
                   "<!-- " & vbCrLf & _
                   "// Array of day names" & vbCrLf & _
                   "var dayNames = new Array(" & Chr$(34) & "Sunday" & Chr$(34) & "," & Chr$(34) & "Monday" & Chr$(34) & "," & Chr$(34) & "Tuesday" & Chr$(34) & "," & Chr$(34) & "Wednesday" & Chr$(34) & "," & Chr$(34) & "Thursday" & Chr$(34) & "," & Chr$(34) & "Friday" & Chr$(34) & "," & Chr$(34) & "Saturday" & Chr$(34) & ");" & vbCrLf & _
                   "var monthNames = new Array(" & Chr$(34) & "January" & Chr$(34) & "," & Chr$(34) & "February" & Chr$(34) & "," & Chr$(34) & "March" & Chr$(34) & "," & Chr$(34) & "April" & Chr$(34) & "," & Chr$(34) & "May" & Chr$(34) & "," & Chr$(34) & "June" & Chr$(34) & "," & Chr$(34) & "July" & Chr$(34) & "," & Chr$(34) & "August" & Chr$(34) & "," & Chr$(34) & "September" & Chr$(34) & "," & Chr$(34) & "October" & Chr$(34) & "," & Chr$(34) & "November" & Chr$(34) & "," & Chr$(34) & "December" & Chr$(34) & ");" & vbCrLf & _
                   "var dt = new Date();" & vbCrLf & _
                   "var y  = dt.getYear();" & vbCrLf & _
                   "// This is for Y2K compliancy" & vbCrLf & _
                   "if (y < 1000) y +=1900;" & vbCrLf & _
                   "document.write(dayNames[dt.getDay()] + " & Chr$(34) & ", " & Chr$(34) & "+ monthNames[dt.getMonth()] + " & Chr$(34) & " " & Chr$(34) & " + dt.getDate() + " & Chr$(34) & ", " & Chr$(34) & " + y);" & vbCrLf & _
                   "// -->" & vbCrLf & _
                   "</script>" & vbCrLf

End Sub

Private Sub mnuNewWeb_Click()

    On Error GoTo hell
  Dim lpLoc As String ':( Move line to top of current Sub
    lpLoc = SelectDir()
    If lpLoc = "" Then Exit Sub ':( Expand Structure
    If frmMain.IsWebOpen Then frmMain.CloseWeb ':( Expand Structure
    MkDir lpLoc
    frmMain.LoadWeb lpLoc

Exit Sub

hell:
    MsgBox Error, vbExclamation

End Sub

Private Sub mnuPersonalMail_Click()

    RTF1.SelText = "<A HREF=" & Chr$(34) & "mailto:" & txtEmail.text & Chr$(34) & ">Send me an e-mail</A> :-)"

End Sub

Private Sub mnuPersonalUrl_Click()

    RTF1.SelText = "<A HREF=" & Chr$(34) & txtUrl.text & Chr$(34) & ">Check out</A> my website!"

End Sub

Sub mnuPreview_Click()

    mnuFileSave_Click 'save it first
    If left$(Caption, 8) = "Untitled" Or Mid$(Caption, 2, 1) <> ":" Then
        frmMain.SB.Panels(1).text = "The file must be saved before you can preview it."
        Exit Sub '>---> Bottom
    End If
    ShellExecute Me.hWnd, "open", "explorer", Caption, "", 10

End Sub

Private Sub mnuRestoreAll_Click()

  Dim lpF As Form

    For Each lpF In Forms
        If left$(lpF.Caption, 32) <> "Designer 2001 Personal Edition" Then
            lpF.WindowState = vbNormal
        End If
    Next lpF

End Sub

Private Sub mnuRestructure_Click()

    'RTF1.SelStart = 0
    'RTF1.SelLength = Len(RTF1.text)
    'RTF1.SelText = SimpleFormat(RTF1.SelText)

End Sub

Sub mnuSelectAll_Click()

    RTF1.SelStart = 0
    RTF1.SelLength = Len(RTF1.text)

End Sub

Private Sub mnuSendFTP_Click()

    Open "C:\temp.htm" For Output As #1
    Print #1, Text1.text
    Close #1

    Inet1.URL = frmMain.Text1.text
    Inet1.UserName = frmMain.Text2.text
    Inet1.Password = frmMain.Text3.text

    Inet1.Execute , ("PUT C:\temp.htm default.htm")

End Sub

Private Sub mnuSetupFTP_Click()

    Load StartOptions
    StartOptions.Show

End Sub

Private Sub mnuSlideShow_Click()
Load frmJavaScript
frmJavaScript.Show
End Sub

Private Sub mnuTileHorizontal_Click()

    frmMain.Arrange vbHorizontal

End Sub

Private Sub mnuTileVertical_Click()

    frmMain.Arrange vbVertical

End Sub

Private Sub mnuTodayTip_Click()

    On Error Resume Next
      frmTip.Show vbModal

End Sub ':( On Error Resume still active

Private Sub mnuUpdate_Click()

    On Error Resume Next
      If Caption = "Untitled" Or Mid$(Caption, 2, 1) <> ":" Then GoTo y ':( Expand Structure
      If LCase$(right$(Caption, 3)) <> "tml" And LCase$(right$(Caption, 3)) <> "htm" And LCase$(right$(Caption, 3)) <> "asp" And LCase$(right$(Caption, 3)) <> ".js" Then GoTo n ':( Expand Structure
y:       'Outline
      AddScripts RTF1.text, frmMain.tvS, Not frmMain.SSTab1.TabVisible(2)
      frmMain.tvS.Nodes("Document").Expanded = True
      Outline Me, frmMain.tvD, RTF1.text, Not frmMain.SSTab1.TabVisible(1)
      GoTo skipClear
n:       'no outline
      frmMain.tvD.Nodes.Clear
skipClear:
      frmMain.tvD.Nodes.Item("Main").Expanded = True
      frmMain.tvD.Nodes.Item("document").Expanded = True
      frmMain.tvD.Nodes.Item("declare").Expanded = True
      frmMain.tvD.Nodes.Item("title").Expanded = True
      frmMain.tvD.SelectedItem = frmMain.tvD.Nodes.Item("Main")

End Sub ':( On Error Resume still active

Private Sub mnuView_Click()

    mnuViewScripts.Checked = frmMain.SSTab1.TabVisible(2)
    frmMain.SB.Panels(1).text = "Contains commands for manipulating the view."
    mnuViewDocuments.Checked = frmMain.SSTab1.TabVisible(1)
    mnuViewToolBar.Checked = frmMain.CoolBar1.Visible
    mnuViewStatusBar.Checked = frmMain.SB.Visible
    mnuViewFileTree.Checked = frmMain.pLeft.Visible

End Sub

Private Sub mnuViewDocuments_Click()

    mnuViewDocuments.Checked = Not mnuViewDocuments.Checked
    frmMain.SSTab1.TabVisible(1) = mnuViewDocuments.Checked
    SaveValue "DocumentTree", mnuViewDocuments.Checked

End Sub

Private Sub mnuViewFileTree_Click()

    frmMain.mnuViewFileTree_Click

End Sub

Private Sub mnuViewMode_Click(index As Integer)

    MousePointer = 11
  Dim i As Integer ':( Move line to top of current Sub
    For i = 1 To 3
        mnuViewMode(i).Checked = False
    Next i
    mnuViewMode(index).Checked = True
    SetViewMode index - 1, RTF1
    SaveValue "ViewMode", index - 1
    MousePointer = 0

End Sub

Private Sub mnuViewOptions_Click()

    frmOpts.Show vbModal

End Sub

Private Sub mnuViewStatusBar_Click()

    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    frmMain.SB.Visible = mnuViewStatusBar.Checked
    SaveValue "Statusbar", frmMain.SB.Visible

End Sub

Private Sub mnuViewToolBar_Click()

    mnuViewToolBar.Checked = Not mnuViewToolBar.Checked
    frmMain.CoolBar1.Visible = mnuViewToolBar.Checked
    SaveValue "Toolbar", frmMain.CoolBar1.Visible

End Sub

Private Sub mnuWeb_Click()

    On Error GoTo hell
    frmMain.SB.Panels(1).text = "Contains commands for managing webs."
    mnuWebDefault.Enabled = (frmMain.tvW.Nodes.Count > 0)
    mnuWebRefresh.Enabled = mnuWebDefault.Enabled
    mnuWebClose.Enabled = (frmMain.tvW.Nodes.Count > 0)
    mnuWebDefault.Checked = (ReadValue("WebDefault") = frmMain.tvW.Nodes(1).key)

Exit Sub

hell:
    mnuWebDefault.Enabled = False
    Resume Next

End Sub

Private Sub mnuWebClose_Click()

    frmMain.CloseWeb

End Sub

Private Sub mnuWebDefault_Click()

    mnuWebDefault.Checked = Not mnuWebDefault.Checked
    If mnuWebDefault.Checked Then
        SaveValue "WebDefault", frmMain.tvW.Nodes(1).key 'root
      Else
        SaveValue "WebDefault", "" 'nothing
    End If

End Sub

Private Sub mnuWebRefresh_Click()

    frmMain.LoadWeb frmMain.Fil.Path

End Sub

Private Sub mnuWindow_Click()

    frmMain.SB.Panels(1).text = "Contains commands for navigating and arranging windows."

End Sub

Private Sub mnuWindowAlignLefts_Click()

    On Error Resume Next
    Dim lpF As Form ':( Move line to top of current Sub
      For Each lpF In Forms
          If lpF.BackColor = &H8000000F Then
              lpF.left = frmMain.ActiveForm.left
          End If
      Next lpF

End Sub ':( On Error Resume still active

Private Sub mnuWindowAlignTops_Click()

    On Error Resume Next
    Dim lpF As Form ':( Move line to top of current Sub
      For Each lpF In Forms
          If lpF.BackColor = &H8000000F Then
              lpF.top = frmMain.ActiveForm.top
          End If
      Next lpF

End Sub ':( On Error Resume still active

Private Sub mnuWindowMaximizeAll_Click()

    WindowState = vbMaximized 'that's all

End Sub

Private Sub mnuWindowMinimizeAll_Click()

    On Error Resume Next
    Dim lpC As Form ':( Move line to top of current Sub
      For Each lpC In Forms
          If lpC.BackColor = &H8000000F Then 'is not MDI
              lpC.WindowState = vbMinimized
          End If
      Next lpC

End Sub ':( On Error Resume still active

Private Sub mnuWindowUnloadAll_Click()

    On Error Resume Next
    Dim lpC As Form ':( Move line to top of current Sub
      For Each lpC In Forms
          If lpC.BackColor = &H8000000F Then 'is not MDI
              Unload lpC
          End If
      Next lpC

End Sub ':( On Error Resume still active

Private Sub mnuWysiwyg_Click()

    Load frmWYSIWYG
    frmWYSIWYG.rtf.text = RTF1.text
    frmWYSIWYG.Show

End Sub

Private Sub RTF1_Change()

    bChanged = True
    'frmMain.TB.Buttons(3).Enabled = True
    mnuFileSave.Enabled = True
    'Undo.AddAction RTF1.text, RTF1.SelStart

End Sub

Private Sub RTF1_DragDrop(source As Control, x As Single, y As Single)

    frmMain.SB.Panels(1).text = source.Name

End Sub

Private Sub RTF1_DragOver(source As Control, x As Single, y As Single, State As Integer)

    frmMain.SB.Panels(1).text = source.Name & vbTab & State

End Sub

Sub RTF1_GotFocus()

    On Error Resume Next
    Dim i As Integer ':( Move line to top of current Sub
      If Mid$(Caption, 2, 1) <> ":" Then GoTo y ':( Expand Structure
      Select Case right$(LCase$(Caption), 3) 'extension
        Case "htm", "tml", "asp", "led" 'untitled also counts
          'don't do anything
          GoTo y
        Case ".js"
          GoTo scriptsOnly
        Case Else
          'skip the outlining
          frmMain.tvD.Nodes.Clear
          GoTo noOutline
      End Select
y:
      If bUpdated = False Then
          Outline Me, frmMain.tvD, RTF1.text, Not frmMain.SSTab1.TabVisible(1)
scriptsOnly:
          If bUpdated = True Then GoTo noOutline ':( Expand Structure
          AddScripts RTF1.text, frmMain.tvS, Not frmMain.SSTab1.TabVisible(2)
          bUpdated = True
      End If
noOutline:
      'frmMain.TB.Buttons(3).Enabled = bChanged
      mnuFileSave.Enabled = bChanged
      RTF1_SelChange
      If frmMain.tvW.Nodes(Caption).Bold = True Then Exit Sub ':( Expand Structure
      For i = 1 To frmMain.tvW.Nodes.Count
          frmMain.tvW.Nodes(i).Bold = False
          If frmMain.tvW.Nodes(i).key = Caption Then frmMain.tvW.Nodes(i).Bold = True: frmMain.tvW.Nodes(i).EnsureVisible ':( Expand Structure
      Next i

End Sub ':( On Error Resume still active

Private Sub RTF1_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
      Form_KeyDown KeyCode, Shift
      If Shift = 2 And KeyCode = 32 Then KeyCode = 0 ':( Expand Structure
      If KeyCode = vbKeyTab Then
          If RTF1.SelLength > 0 Then
              KeyCode = 0
    Dim StrP() As String, i As Integer ':( Move line to top of current Sub
              StrP = Split(RTF1.SelText, vbNewLine)
              For i = 0 To UBound(StrP) - 1
                  RTF1.SelText = vbTab & StrP(i) & vbNewLine
              Next i
              RTF1.SelText = vbTab & StrP(UBound(StrP))
          End If
      End If
      RTF1.SetFocus

End Sub ':( On Error Resume still active

Private Sub RTF1_KeyPress(KeyAscii As Integer)
    Dim dblOldPosition As Long, MoreInf As String ':( Move line to top of current Sub

    On Error Resume Next
      If ReadValue("AutoComplete", True) = False Then Exit Sub ':( Expand Structure

      dblOldPosition = RTF1.SelStart

      MoreInf = GetMoreInfo(CurrentType)

      If KeyAscii = 32 Or KeyAscii = 13 Then
          If IsLink(CurrentType) Then
              RTF1.SelStart = dblOldPosition - Len(CurrentType)
              RTF1.SelText = "<A href=" & Chr$(34) & MoreInf & CurrentType & Chr$(34) & ">"
              RTF1.SelStart = RTF1.SelStart + Len(CurrentType)
              RTF1.SelText = "</A>"
          End If
          CurrentType = ""
        ElseIf KeyAscii = 8 Then
          If Len(CurrentType) > 1 Then CurrentType = left$(CurrentType, Len(CurrentType) - 1) ':( Expand Structure
        Else
          CurrentType = CurrentType & Chr$(KeyAscii)
      End If

End Sub ':( On Error Resume still active

Private Sub RTF1_LostFocus()

    On Error Resume Next
      If frmMain.ActiveForm.hWnd <> Me.hWnd Then bUpdated = False Else bUpdated = True ':( Expand Structure

End Sub ':( On Error Resume still active

Private Sub RTF1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
      mnuEdit_Click
      If Button = 2 Then PopupMenu mnuEdit, vbPopupMenuCenterAlign + vbPopupMenuLeftAlign, x, y - 2000 ':( Expand Structure

End Sub ':( On Error Resume still active

Sub SaveHTMLFile(lpFileName As String)

    On Error Resume Next
      Open lpFileName For Output As #1
      Print #1, RTF1.text
      Close #1
      Caption = lpFileName
      If frmMain.IsWebOpen Then frmMain.tvW.Nodes.Add Up1Level(lpFileName), tvwChild, lpFileName, GetFile(lpFileName), FileIcon(lpFileName) ':( Expand Structure
      AddFileMRU lpFileName
      frmMain.tvS.Nodes("Document").text = GetFile(Caption)
      mnuUpdate_Click
      bChanged = False

End Sub ':( On Error Resume still active

Private Sub RTF1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    CurrentType = RichWordOver(RTF1, x, y)

End Sub

Private Sub RTF1_SelChange()

    On Error Resume Next

    Dim PC As Integer, lP As POINTAPI ':( Move line to top of current Sub
      frmMain.SB.Panels(2).text = " Loc: " & RTF1.SelStart & ", Line: " & GetCurrentLine(RTF1) & " "
      frmMain.SB.Panels(3).text = " Size: " & Round(Len(RTF1.text) / 1000, 1) & " KB, Lines: " & GetTotalLines(RTF1) & " "

      If left$(frmMain.SB.Panels(1).text, 32) = "Please wait, converting to ASCII" Then
          PC = Round((RTF1.SelStart * 100) / Len(RTF1.text), 0)
          frmMain.SB.Panels(1).text = "Please wait, converting to ASCII (" & PC & "%)"
      End If

      'frmMain.TB.Buttons("cut").Enabled = (RTF1.SelLength > 0)
      'frmMain.TB.Buttons("copy").Enabled = frmMain.TB.Buttons("cut").Enabled
      'frmMain.TB.Buttons("paste").Enabled = Clipboard.GetFormat(vbCFText)
      'frmMain.TB.Buttons("undo").Enabled = Undo.UndoAvailable
      'frmMain.TB.Buttons("redo").Enabled = Undo.RedoAvailable

      GetCaretPos lP
      CurrentType = RichWordOver(RTF1, lP.x * Screen.TwipsPerPixelX + RTF1.left, lP.y * Screen.TwipsPerPixelY + RTF1.top)

End Sub ':( On Error Resume still active

Private Sub TabCtl1_TabChanged(ByVal NewTabIndex As Long)

  Dim i As Integer

    On Error Resume Next
      Select Case NewTabIndex
        Case 1
          pD.ZOrder vbBringToFront
    
          'With frmMain.TB2.Buttons
          'For i = 1 To .Count
          '.Item(i).Enabled = True
          'Next i
          'End With
    
          IE1.Navigate "about:blank"
          frmMain.SB.Panels(1).text = ""
    
          'EnableBar
    
        Case 2
          If Caption = "Untitled" Then ThisPath = App.Path Else ThisPath = Up1Level(Caption) ':( Expand Structure
            
          Open FullPath(ThisPath, "temp.html") For Output As #1
          Print #1, RTF1.text
          Close #1
          'I could have used /temp.html as app.path does not put in /
          'unless it is a root drive, but if i dont, it will be put in the parent
          'folder, but who cares as long as we can access it
          IE1.Navigate FullPath(ThisPath, "temp.html")

          IE1.ZOrder vbBringToFront
    
          'DisableBar

      End Select

End Sub ':( On Error Resume still active

Private Sub tmTimer_Timer()

    On Error Resume Next
    Dim lpMinute As String, AM_PM As String ':( Move line to top of current Sub
      lpMinute = Minute(Time)
      AM_PM = right$(TimeSerial(Hour(Time), Minute(Time), 0), 2)
      If Len(lpMinute) = 1 Then lpMinute = "0" & lpMinute ':( Expand Structure
      mnuInsertDateTimeLong.Caption = Hour(Time) & ":" & lpMinute & " " & AM_PM
      mnuInsertDateTimeShort.Caption = Date
      mnuInsertDateTimeWeekday.Caption = GetDay(Date)

End Sub ':( On Error Resume still active

Function GetSelStart() As Long

  Dim lpS As Long

    lpS = RTF1.Find("<BODY", , , rtfNoHighlight)
    lpS = RTF1.Find(">", lpS + 1, , rtfNoHighlight)
    GetSelStart = lpS + 3

End Function

Sub SetMenus()

    On Error Resume Next
      'Menus if given mnemonics don't work
      mnuEditCut.Caption = mnuEditCut.Caption & vbTab & "Ctrl+X"
      mnuEditCopy.Caption = mnuEditCopy.Caption & vbTab & "Ctrl+C"
      mnuEditPaste.Caption = mnuEditPaste.Caption & vbTab & "Ctrl+V"

End Sub ':( On Error Resume still active

Sub CopyMRUList()

    On Error Resume Next
    Dim i As Integer ':( Move line to top of current Sub

      For i = 1 To 6

          mnuWebMRU(i).Caption = frmMain.mnuWebMRU(i).Caption
          mnuWebMRU(i).tag = frmMain.mnuWebMRU(i).tag
          mnuWebMRU(i).Visible = (Len(mnuWebMRU(i).Caption) > 4)

          mnuFileMRU(i).Caption = frmMain.mnuFileMRU(i).Caption
          mnuFileMRU(i).tag = frmMain.mnuFileMRU(i).tag
          mnuFileMRU(i).Visible = (Len(mnuFileMRU(i).Caption) > 4)

      Next i

End Sub ':( On Error Resume still active

Private Sub mnuWebMRU_Click(index As Integer)

    On Error Resume Next
      frmMain.LoadWeb mnuWebMRU(index).tag

End Sub ':( On Error Resume still active

Private Sub mnuViewScripts_Click()

    mnuViewScripts.Checked = Not mnuViewScripts.Checked
    frmMain.SSTab1.TabVisible(2) = mnuViewScripts.Checked
    SaveValue "ScriptView", mnuViewScripts.Checked

End Sub

Function IsEmptyElement(ByVal Element As String) As Boolean

    If left$(Element, 1) = "<" Then Element = right$(Element, Len(Element) - 1) ':( Expand Structure
    If left$(Element, 1) = "/" Then Element = right$(Element, Len(Element) - 1) ':( Expand Structure
    If right$(Element, 1) = ">" Then Element = left$(Element, Len(Element) - 1) ':( Expand Structure
    If left$(Element, 2) = "!-" Then IsEmptyElement = True: Exit Function ':( Expand Structure

  Dim iPos As Long ':( Move line to top of current Function
    iPos = InStr(1, Element, " ")
    If iPos > 0 Then Element = left$(Element, iPos): Element = Trim$(Element) ':( Expand Structure
    frmMain.SB.Panels(1).text = Element

    Select Case LCase$(Element)
      Case "hr", "img", "br", "input", "button", "bgsound", "base", "meta", "!doctype", "!--", "isindex"
        IsEmptyElement = True
      Case Else
        IsEmptyElement = False
    End Select

End Function

Function GetTag(ByVal tag As String) As String

  Dim iPos As Long

    iPos = InStr(1, tag, " ")
    If iPos > 0 Then
        GetTag = Trim$(left$(tag, iPos))
      Else
        GetTag = tag
    End If

End Function

Sub AddComment(ByVal lpStart As Long, lpEndTag As String)

    RTF1.SelText = ">"
    RTF1.SelStart = lpStart + 1
    RTF1.SelText = vbNewLine & "<!--" & vbNewLine & vbNewLine & "// -->" & vbNewLine & lpEndTag
    RTF1.SelStart = RTF1.SelStart - Len(lpEndTag) - 10

End Sub

Function IsLink(text As String) As Boolean

  Dim i As Integer, Words() As String

    Words = Split(Domains, " ")
    For i = 0 To UBound(Words())
        If InStr(1, text, Words(i)) > 1 Then IsLink = True: Exit Function ':( Expand Structure
    Next i
    IsLink = False 'not found anywhere

End Function

Function GetMoreInfo(text As String) As String

    If InStr(1, text, "@") Then
        GetMoreInfo = "mailto:"
      ElseIf InStr(1, text, "http://") Then
        GetMoreInfo = ""
      ElseIf InStr(1, text, "://") = 0 Then
        GetMoreInfo = "http://"
      Else
        GetMoreInfo = ""
    End If

End Function

Private Sub ConvertTextToHTML()

  Dim q, W, i, Msg, result ':( As Variant ?':( Duplicated Name
  Dim InputTitle As String

    q = vbNullString
    W = vbNullString

    'If ActiveForm Is Nothing Then LoadNewDoc
    On Error GoTo ex
    InputTitle = InputBox("What is the title for this page?", "Page Title", cdl.FileTitle) '"Untitled")
    'off_off
    Open Text2.text For Input As #1
    q = Input$(LOF(1), 1)
    Close #1
    ''''''''''''''''''''''''''
    If fileExist(Text3.text) = True Then ':( Remove Pleonasm
        Msg = "File " & Text3.text & " already exists!" _
              & vbCrLf & "Do you want to overwrite it?"
        result = MsgBox(Msg, 4 + 48, "Warning!")
        If result = vbYes Then
            GoTo continue
          Else
            GoTo stp
        End If
    End If
    ''''''''''''''''''''''''''
continue:
    canExit = False
    'Label3.Caption = "Wait..."
    '*******************
    For i = 1 To Len(q)
        If Mid$(q, i, 2) = vbCrLf Then Num = Num + 1 ':( Expand Structure
    Next ':( Repeat For-Variable: I
    '*******************
    'Command3.Enabled = True

    lastN = 1

    W = "<html>" & vbCrLf & "<head>" & vbCrLf _
        & vbCrLf & "<title>" & InputTitle & "</title>" _
        & vbCrLf & "</head>" & vbCrLf & "<body>" & vbCrLf _
        & vbCrLf & "<br><font color=""navy""><H2>" _
        & InputTitle & "</H2></font>" _
        & vbCrLf & "<p align=""justify"">" & txtIndent.text

    'W = "<html>" & vbCrLf & "<head>" & vbCrLf _
        '& "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">" _
        '& vbCrLf & "<title>" & InputTitle & "</title>" _
        '& vbCrLf & "</head>" & vbCrLf & "<body leftmargin=64 " _
        '& "style=""border-left: medium solid rgb(111,0,221); " _
        '& "border-top: medium solid rgb(84,0,168)"">" _
        '& vbCrLf & "<br><font color=""navy""><H2>" _
        '& InputTitle & "</H2></font>" _
        '& vbCrLf & "<p align=""justify"">" & txtIndent.Text
    Do
        If canExit Then GoTo stp ':( Expand Structure
        DoEvents
        n = InStr(lastN, q, vbCrLf)
        If n Then
            If n <> lastN Then W = W & Mid$(q, lastN, n - lastN) & vbCrLf & "<p>" & Text1.text ':( Expand Structure
            lastN = n + 2
          Else
            W = W & right$(q, Len(q) - lastN + 1) & "</p>" _
                & vbCrLf & "<hr>" & vbCrLf _
                & "<p><p>" _
                & vbCrLf & "</body>" & vbCrLf & "</html>"
        End If
        '''''''''''''
        Num = Num - 1
        Caption = "0 - " & Num
        '''''''''''''
    Loop While n <> 0

    'Text5.Text = W
    RTF1.text = W
    'Open Text3.Text For Output As #1
    '  Print #1, W
    'Close #1

    canExit = True
stp:
    'Label3.Caption = "Done..."
    Caption = "Designer"
    Num = 1
    'on_on
    Beep

Exit Sub

ex:
    If Err.Number = 53 Then
        MsgBox "File not found!" & vbCrLf & Text2.text, vbExclamation, "Warning!"
      Else
        MsgBox Error$ & " Error " & Err, vbExclamation, "Warning!"
    End If
    Close #1
    'on_on

End Sub

Private Sub GetDetails()

    txtName.text = GetSetting(App.Title, "Details", "PersonalName", "Your Name")
    txtCompany.text = GetSetting(App.Title, "Details", "BusinessName", "Company Name")
    txtEmail.text = GetSetting(App.Title, "Details", "PersonalEmail", "Your E-mail")
    txtBusiness.text = GetSetting(App.Title, "Details", "BusinessEmail", "Company E-mail")
    txtUrl.text = GetSetting(App.Title, "Details", "PersonalURL", "Your Website")
    txtWebSite.text = GetSetting(App.Title, "Details", "BusinessURL", "Company Website")

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:19 PM) 14 + 1411 = 1425 Lines
