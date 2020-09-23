VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmForms 
   Caption         =   "Create A Form"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   8160
      TabIndex        =   3
      Top             =   480
      Width           =   3615
      Begin VB.Frame fraButtons 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton Command6 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1920
            TabIndex        =   61
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtCaption 
            Height          =   285
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   2775
         End
         Begin VB.OptionButton optReset 
            Caption         =   "Reset Button"
            Height          =   195
            Left            =   1800
            TabIndex        =   58
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optSubmit 
            Caption         =   "Submit Button"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "What should the caption say?:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Width           =   2160
         End
      End
      Begin VB.Frame fraFile 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton Command5 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1920
            TabIndex        =   55
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtFilesSize 
            Height          =   285
            Left            =   120
            TabIndex        =   54
            Text            =   "45"
            Top             =   915
            Width           =   615
         End
         Begin VB.TextBox txtFilesName 
            Height          =   285
            Left            =   120
            TabIndex        =   52
            Top             =   315
            Width           =   3135
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Size / Max length"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   120
            Width           =   465
         End
      End
      Begin VB.Frame fraMenu 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   3375
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   840
            TabIndex        =   49
            Text            =   "Select Me!"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1920
            TabIndex        =   47
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtList 
            Height          =   495
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   1800
            Width           =   3135
         End
         Begin VB.TextBox txtOption 
            Height          =   285
            Left            =   840
            TabIndex        =   45
            Text            =   "Option1"
            Top             =   915
            Width           =   2415
         End
         Begin VB.TextBox txtSize 
            Height          =   285
            Left            =   120
            TabIndex        =   43
            Text            =   "4"
            Top             =   915
            Width           =   615
         End
         Begin VB.TextBox txtMenuname 
            Height          =   285
            Left            =   120
            TabIndex        =   41
            Text            =   "Menu1"
            Top             =   315
            Width           =   3135
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Text:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   1365
            Width           =   360
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Option:"
            Height          =   195
            Left            =   840
            TabIndex        =   44
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Size:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   345
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Name of Menu:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame fraRadio 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   3375
         Begin VB.TextBox txtRadioName 
            Height          =   285
            Left            =   120
            TabIndex        =   35
            Text            =   "Radio"
            Top             =   360
            Width           =   3135
         End
         Begin VB.CheckBox chkRadio 
            Caption         =   "Default Choice"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtRadioValue 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Text            =   "Choice1"
            Top             =   915
            Width           =   3135
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1920
            TabIndex        =   32
            Top             =   1950
            Width           =   1335
         End
         Begin VB.TextBox txtRadioText 
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Text            =   "Radio Button 1"
            Top             =   1515
            Width           =   3135
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Name of Radio Button:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   1620
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Value:"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   450
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Radio Button Text:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   1335
         End
      End
      Begin VB.Frame fraCheck 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   3375
         Begin VB.TextBox txtDefault 
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Text            =   "Check1"
            Top             =   1515
            Width           =   3135
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1920
            TabIndex        =   27
            Top             =   1950
            Width           =   1335
         End
         Begin VB.TextBox txtChkValue 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Text            =   "Multiple Choice"
            Top             =   915
            Width           =   3135
         End
         Begin VB.CheckBox chkDefault 
            Caption         =   "Default Choice"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtChkName 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Text            =   "Check1"
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Checkbox Text:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Width           =   1125
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Value:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   450
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Name of Checkbox:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   1410
         End
      End
      Begin VB.Frame fraTextArea 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton Command1 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1920
            TabIndex        =   20
            Top             =   1470
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Word Wrap"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtCols 
            Height          =   285
            Left            =   1800
            TabIndex        =   18
            Text            =   "60"
            Top             =   915
            Width           =   615
         End
         Begin VB.TextBox txtRows 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Text            =   "5"
            Top             =   915
            Width           =   615
         End
         Begin VB.TextBox txtTextArea 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Text            =   "TextArea1"
            Top             =   315
            Width           =   3135
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Columns:"
            Height          =   195
            Left            =   1800
            TabIndex        =   16
            Top             =   720
            Width           =   645
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Rows:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Name Of Textarea:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   1350
         End
      End
      Begin VB.Frame fraText 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3375
         Begin VB.CommandButton cmdText 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1920
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtTextLength 
            Height          =   285
            Left            =   1800
            TabIndex        =   10
            Text            =   "20"
            Top             =   915
            Width           =   615
         End
         Begin VB.TextBox txtTextSize 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Text            =   "25"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   315
            Width           =   3135
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Maximum Length:"
            Height          =   195
            Left            =   1800
            TabIndex        =   9
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Size of Textbox:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name of the Textbox:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   1530
         End
      End
   End
   Begin MSComDlg.CommonDialog cdlForm 
      Left            =   2760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Form Templates (*.qcf)|*.qcf|Web Pages (*.htm, *.html)|*.htm;*html"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":049C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":0934
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":0DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":1264
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":16FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":1B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":202C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":24C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":2624
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":2784
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":28E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":2A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":2BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":2D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":2E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":2FC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":3124
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":3284
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":33E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":3544
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForms.frx":36A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   767
      ButtonWidth     =   741
      ButtonHeight    =   609
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New Form"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save As Template"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "textbox"
            Object.ToolTipText     =   "Add A TextBox"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "password"
            Object.ToolTipText     =   "Add A Password Box"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "textarea"
            Object.ToolTipText     =   "Add A Text Area Box"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "checkbox"
            Object.ToolTipText     =   "Add A Checkbox"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "option"
            Object.ToolTipText     =   "Add A Radio Button"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "menu"
            Object.ToolTipText     =   "Add a Drop-Down Menu"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sendFiles"
            Object.ToolTipText     =   "Insert Browse Button"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "submit"
            Object.ToolTipText     =   "Add A Submit or Reset Button"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "send"
            Object.ToolTipText     =   "Send To Editor"
            ImageIndex      =   22
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   8070
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
      Location        =   ""
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmForms.frx":39F8
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Send"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu mnuFormatting 
      Caption         =   "Formatting"
      Begin VB.Menu mnuFormatBold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuFormatItalics 
         Caption         =   "Italics"
      End
      Begin VB.Menu mnuFormatUnderline 
         Caption         =   "Underline"
      End
      Begin VB.Menu Formatbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatLeft 
         Caption         =   "Align Left"
      End
      Begin VB.Menu mnuFormatCenter 
         Caption         =   "Center"
      End
      Begin VB.Menu mnuAlignRight 
         Caption         =   "Align Right"
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Form"
      Begin VB.Menu mnuFormatTextbox 
         Caption         =   "Textbox"
      End
      Begin VB.Menu mnuFormatPassword 
         Caption         =   "Password Box"
      End
      Begin VB.Menu mnuFormatTextArea 
         Caption         =   "Text Area"
      End
      Begin VB.Menu mnuFormatBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatCheckbox 
         Caption         =   "Checkbox"
      End
      Begin VB.Menu mnuRadioButton 
         Caption         =   "Radio Button"
      End
      Begin VB.Menu mnuFormatMenu 
         Caption         =   "Select Menu"
      End
      Begin VB.Menu mnuFormatbar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatBrowse 
         Caption         =   "Browse For Files"
      End
      Begin VB.Menu mnuFormatBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubmitButon 
         Caption         =   "Submit Button"
      End
      Begin VB.Menu mnuFormatReset 
         Caption         =   "Reset Button"
      End
   End
End
Attribute VB_Name = "frmForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdText_Click()

    rtf.SelText = "<input type=" & Chr$(34) & "text" & Chr$(34) & " name=" & Chr$(34) & txtName.text & Chr$(34) & " size=" & Chr$(34) & txtTextSize.text & Chr$(34) & "maxlength=" & Chr$(34) & txtTextLength.text & Chr$(34) & ">" & "<br>" & vbCrLf

End Sub

Private Sub Command1_Click()

    Select Case Check1.value
      Case vbChecked
        rtf.SelText = "<textarea name=" & Chr$(34) & txtTextArea.text & Chr$(34) & " rows=" & Chr$(34) & txtRows.text & Chr$(34) & " cols=" & Chr$(34) & txtCols.text & Chr$(34) & " wrap>Give us your comments here...</textarea></br>" & vbCrLf
      Case vbUnchecked
        rtf.SelText = "<textarea name=" & Chr$(34) & txtTextArea.text & Chr$(34) & " rows=" & Chr$(34) & txtRows.text & Chr$(34) & " cols=" & Chr$(34) & txtCols.text & Chr$(34) & ">Give us your comments here...</textarea></br>" & vbCrLf
    End Select

End Sub

Private Sub Command2_Click()

    Select Case chkDefault.value
      Case vbUnchecked
        rtf.SelText = "<input type=" & Chr$(34) & "checkbox" & Chr$(34) & " name=" & Chr$(34) & txtChkName.text & Chr$(34) & " value=" & Chr$(34) & txtChkValue.text & Chr$(34) & ">" & txtDefault.text & "<br>" & vbCrLf
      Case vbChecked
        rtf.SelText = "<input type=" & Chr$(34) & "checkbox" & Chr$(34) & " name=" & Chr$(34) & txtChkName.text & Chr$(34) & " value=" & Chr$(34) & txtChkValue.text & Chr$(34) & " checked>" & txtDefault.text & "<br>" & vbCrLf
    End Select

End Sub

Private Sub Command3_Click()

    Select Case chkRadio.value
      Case vbUnchecked
        rtf.SelText = "<input type=" & Chr$(34) & "radio" & Chr$(34) & " name=" & Chr$(34) & txtRadioName.text & Chr$(34) & " value=" & Chr$(34) & txtRadioValue.text & Chr$(34) & ">" & txtRadioText.text & "<br>" & vbCrLf
      Case vbChecked
        rtf.SelText = "<input type=" & Chr$(34) & "radio" & Chr$(34) & " name=" & Chr$(34) & txtRadioName.text & Chr$(34) & " value=" & Chr$(34) & txtRadioValue.text & Chr$(34) & " checked>" & txtRadioText.text & "<br>" & vbCrLf
    End Select

End Sub

Private Sub Command4_Click()

    rtf.SelText = "<select name=" & Chr$(34) & txtMenuname.text & Chr$(34) & " size=" & Chr$(34) & txtSize.text & Chr$(34) & ">" & vbCrLf & txtList.text & vbCrLf & "</select>" & vbCrLf

End Sub

Private Sub Command5_Click()

    rtf.SelText = "<input type=" & Chr$(34) & "file" & Chr$(34) & " enctype=" & Chr$(34) & "multipart/form-data" & Chr$(34) & " name=" & Chr$(34) & txtFilesName.text & Chr$(34) & " size=" & Chr$(34) & txtFilesSize.text & Chr$(34) & "><br>" & vbCrLf

End Sub

Private Sub Command6_Click()

    If optSubmit.value = True Then ':( Remove Pleonasm
        rtf.SelText = "<input type=" & Chr$(34) & "submit" & Chr$(34) & " value=" & Chr$(34) & txtCaption.text & Chr$(34) & ">" & vbCrLf
      Else
        rtf.SelText = "<input type=" & Chr$(34) & "reset" & Chr$(34) & " value=" & Chr$(34) & txtCaption.text & Chr$(34) & ">" & vbCrLf
    End If

End Sub

Private Sub Form_Load()

    fraText.Visible = True
    fraButtons.Visible = False
    fraTextArea.Visible = False
    fraFile.Visible = False
    fraMenu.Visible = False
    fraRadio.Visible = False
    fraCheck.Visible = False
    wb.Navigate ("about:blank")

End Sub

Private Sub Form_Resize()

  'rtf.Width = Me.ScaleWidth

    wb.Width = Me.ScaleWidth

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub optReset_Click()

    txtCaption.text = "Clear"

End Sub

Private Sub optSubmit_Click()

    txtCaption.text = "Send"

End Sub

Private Sub rtf_Change()

    Render

End Sub

Private Sub Render()

    wb.Document.Script.Document.Clear
    wb.Document.Script.Document.Write rtf.text
    wb.Document.Script.Document.Close

    Exit Sub

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtList.SelText = Text1.text & vbCrLf
        txtOption.text = ""
        Text1.text = ""
        txtOption.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
      Case "new"
        rtf.text = ""
        Render
      Case "open"
        cdlForm.ShowOpen
        rtf.LoadFile cdlForm.fileName
      Case "save"

      Case "delete"
        rtf.SelText = ""
      Case "cut"
        Clipboard.SetText rtf.SelText
        rtf.SelText = ""
      Case "copy"
        Clipboard.SetText rtf.SelText
      Case "paste"
        rtf.SelText = Clipboard.GetText
      Case "bold"
        rtf.SelText = "<b>" & rtf.SelText & "</b>"
      Case "italic"
        rtf.SelText = "<i>" & rtf.SelText & "</i>"
      Case "underline"
        rtf.SelText = "<u>" & rtf.SelText & "</u>"
      Case "left"
        rtf.SelText = "<div align=" & Chr$(34) & "left" & Chr$(34) & ">" & rtf.SelText & "</div>"
      Case "center"
        rtf.SelText = "<center>" & rtf.SelText & "</center>"
      Case "right"
        rtf.SelText = "<div align=" & Chr$(34) & "right" & Chr$(34) & ">" & rtf.SelText & "</div>"
      Case "textbox"
        fraText.Visible = True
        fraButtons.Visible = False
        fraTextArea.Visible = False
        fraFile.Visible = False
        fraMenu.Visible = False
        fraRadio.Visible = False
        fraCheck.Visible = False
      Case "password"
        fraText.Visible = False
        fraButtons.Visible = False
        fraTextArea.Visible = False
        fraFile.Visible = False
        fraMenu.Visible = False
        fraRadio.Visible = False
        fraCheck.Visible = False
      Case "textarea"
        fraText.Visible = False
        fraButtons.Visible = False
        fraTextArea.Visible = True
        fraFile.Visible = False
        fraMenu.Visible = False
        fraRadio.Visible = False
        fraCheck.Visible = False
      Case "checkbox"
        fraText.Visible = False
        fraButtons.Visible = False
        fraTextArea.Visible = False
        fraFile.Visible = False
        fraMenu.Visible = False
        fraRadio.Visible = False
        fraCheck.Visible = True
      Case "option"
        fraText.Visible = False
        fraButtons.Visible = False
        fraTextArea.Visible = False
        fraFile.Visible = False
        fraMenu.Visible = False
        fraRadio.Visible = True
        fraCheck.Visible = False
      Case "menu"
        fraText.Visible = False
        fraButtons.Visible = False
        fraTextArea.Visible = False
        fraFile.Visible = False
        fraMenu.Visible = True
        fraRadio.Visible = False
        fraCheck.Visible = False
      Case "sendFiles"
        fraText.Visible = False
        fraButtons.Visible = False
        fraTextArea.Visible = False
        fraFile.Visible = True
        fraMenu.Visible = False
        fraRadio.Visible = False
        fraCheck.Visible = False
      Case "submit"
        fraText.Visible = False
        fraButtons.Visible = True
        fraTextArea.Visible = False
        fraFile.Visible = False
        fraMenu.Visible = False
        fraRadio.Visible = False
        fraCheck.Visible = False
      Case "send"
        frmMain.ActiveForm.RTF1.SelText = "<form method=post action=" & Chr$(34) & "Path to CGI or Perl script here." & Chr$(34) & ">" & vbCrLf & _
                                          rtf.text & vbCrLf & "</form>"
    End Select

End Sub

Private Sub txtOption_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtList.SelText = "<option value=" & Chr$(34) & txtOption.text & Chr$(34) & ">" & vbCrLf
        Text1.SetFocus
        KeyAscii = 0
    End If

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:06 PM) 1 + 243 = 244 Lines
