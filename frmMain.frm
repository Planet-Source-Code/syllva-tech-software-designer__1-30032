VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Queen City Software - Designer"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10035
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "frmMain.frx":0E42
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":354C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":429C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":494C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3690
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":50FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5214
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":532C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5444
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":555C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5674
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":578C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F34
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":604C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6164
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1323
      _CBWidth        =   10035
      _CBHeight       =   750
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinHeight1      =   330
      Width1          =   2715
      NewRow1         =   0   'False
      Child2          =   "Combo1"
      MinHeight2      =   315
      Width2          =   1530
      NewRow2         =   -1  'True
      Child3          =   "Toolbar2"
      MinHeight3      =   330
      Width3          =   1275
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   1725
         TabIndex        =   15
         Top             =   390
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "fonts"
               Object.ToolTipText     =   "Insert Font Information"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "image"
               Object.ToolTipText     =   "Insert an Image"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "lrgfont"
               Object.ToolTipText     =   "Enlarge Font"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "smfont"
               Object.ToolTipText     =   "Make Font Smaller"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bullet"
               Object.ToolTipText     =   "Create Bulleted List"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "number"
               Object.ToolTipText     =   "Create numbered List"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "links"
               Object.ToolTipText     =   "Create a Link"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "cgi"
               Object.ToolTipText     =   "Create CGI Script"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "css"
               Object.ToolTipText     =   "Create Style Sheet"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "js"
               Object.ToolTipText     =   "Create A JavaScript"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "vbs"
               Object.ToolTipText     =   "Create a Visual Basic Script"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMain.frx":6A24
         Left            =   165
         List            =   "frmMain.frx":6A31
         TabIndex        =   14
         Text            =   "View"
         Top             =   390
         Width           =   1335
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   13
         Top             =   30
         Width           =   9780
         _ExtentX        =   17251
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   28
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "new"
               Object.ToolTipText     =   "New Document"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "newweb"
               Object.ToolTipText     =   "Create New Web Site"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "open"
               Object.ToolTipText     =   "Open"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "save"
               Object.ToolTipText     =   "Save Document"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "print"
               Object.ToolTipText     =   "Print This Document"
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "printCode"
                     Text            =   "Print Source Code"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "printWebPage"
                     Text            =   "Print Web Page"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "undo"
               Object.ToolTipText     =   "Undo Last Action"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "redo"
               Object.ToolTipText     =   "Redo Last Action"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "delete"
               Object.ToolTipText     =   "Delete"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cut"
               Object.ToolTipText     =   "Cut"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "copy"
               Object.ToolTipText     =   "Copy"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "paste"
               Object.ToolTipText     =   "Paste"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bold"
               Object.ToolTipText     =   "Make Text Bold"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "italic"
               Object.ToolTipText     =   "make Text Italic"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "underline"
               Object.ToolTipText     =   "Underline Text"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "strike"
               Object.ToolTipText     =   "Strike Out Text"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "left"
               Object.ToolTipText     =   "Align Text Left"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "center"
               Object.ToolTipText     =   "Center"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "right"
               Object.ToolTipText     =   "Align Right"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "find"
               Object.ToolTipText     =   "Find (In Top Window)"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "convert"
               Object.ToolTipText     =   "Convert from HTML to Text"
               ImageIndex      =   20
            EndProperty
            BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "CleanCode"
               Object.ToolTipText     =   "Cleanup Code"
               ImageIndex      =   21
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4680
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "html"
      DialogTitle     =   "WonderHTML"
      Filter          =   "HTML Documents (*.htm, *.html)|*.html;*.htm|GIF, JPEG and Bitmap Images (*.gif,*.jpg,*.bmp)|*.gif;*.jpg;*.bmp|All files (*.*)|*.*"
   End
   Begin MSComctlLib.ImageList imlTB2 
      Left            =   4950
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A54
            Key             =   "strike"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6BB0
            Key             =   "center"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D0C
            Key             =   "big"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E68
            Key             =   "small"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6FC4
            Key             =   "left"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7120
            Key             =   "bullets"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":727C
            Key             =   "numbers"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73D8
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7534
            Key             =   "time"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7690
            Key             =   "font"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C2C
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D88
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7EE4
            Key             =   "applet"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D38
            Key             =   "right"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E94
            Key             =   "link"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8FF0
            Key             =   "image"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":914C
            Key             =   "script"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7035
      Left            =   0
      ScaleHeight     =   7035
      ScaleWidth      =   2955
      TabIndex        =   1
      Tag             =   "The FileTree lists all the files and folders in a specified path, for quick access."
      Top             =   750
      Width           =   2955
      Begin VB.TextBox txtCD 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   6360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   5970
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   5685
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   5400
         Visible         =   0   'False
         Width           =   2655
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5280
         Left            =   0
         TabIndex        =   2
         Tag             =   "The Tab lets you switch between web and document view. Double-click it for more information."
         Top             =   0
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   9313
         _Version        =   393216
         Style           =   1
         TabHeight       =   556
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Web Manager"
         TabPicture(0)   =   "frmMain.frx":96E8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "tvW"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Fldr"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Fil"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Document"
         TabPicture(1)   =   "frmMain.frx":9704
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "tvD"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Scripts"
         TabPicture(2)   =   "frmMain.frx":9720
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fSC"
         Tab(2).Control(1)=   "tvS"
         Tab(2).ControlCount=   2
         Begin VB.FileListBox fSC 
            Height          =   285
            Left            =   -72750
            TabIndex        =   9
            Top             =   4905
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.FileListBox Fil 
            Height          =   1260
            Left            =   435
            TabIndex        =   6
            Top             =   3600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.DirListBox Fldr 
            Height          =   990
            Left            =   615
            TabIndex        =   5
            Top             =   2430
            Visible         =   0   'False
            Width           =   1590
         End
         Begin MSComctlLib.TreeView tvW 
            Height          =   4560
            Left            =   90
            TabIndex        =   4
            Tag             =   $"frmMain.frx":973C
            Top             =   405
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   8043
            _Version        =   393217
            Indentation     =   335
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            ImageList       =   "imlTV"
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView tvD 
            Height          =   4635
            Left            =   -74910
            TabIndex        =   3
            Tag             =   $"frmMain.frx":97C7
            Top             =   405
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   8176
            _Version        =   393217
            Indentation     =   335
            LabelEdit       =   1
            Sorted          =   -1  'True
            Style           =   7
            ImageList       =   "imlTVD"
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView tvS 
            Height          =   4635
            Left            =   -74910
            TabIndex        =   8
            Tag             =   $"frmMain.frx":9877
            Top             =   405
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   8176
            _Version        =   393217
            Indentation     =   335
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imlJS"
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox pS 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3030
         Left            =   2340
         MouseIcon       =   "frmMain.frx":9927
         MousePointer    =   99  'Custom
         ScaleHeight     =   3030
         ScaleWidth      =   45
         TabIndex        =   7
         Top             =   720
         Width           =   45
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   10035
      TabIndex        =   10
      Top             =   7785
      Visible         =   0   'False
      Width           =   10035
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   $"frmMain.frx":9A79
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   45
         Width           =   9960
      End
   End
   Begin MSComctlLib.ImageList imlJS 
      Left            =   5520
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9B08
            Key             =   "fp"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B4D4
            Key             =   "file"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C328
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C8C4
            Key             =   "o"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CE60
            Key             =   "image"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D3FC
            Key             =   "!file"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E250
            Key             =   "function"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E7EC
            Key             =   "variable"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E948
            Key             =   "fp2"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tC 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4200
      Top             =   5040
   End
   Begin MSComctlLib.ImageList imlTVD 
      Left            =   3840
      Top             =   2040
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
            Picture         =   "frmMain.frx":10314
            Key             =   "main"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11CE0
            Key             =   "element"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":136AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTV 
      Left            =   3120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15078
            Key             =   "file"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15ECC
            Key             =   "!file"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16D20
            Key             =   "!open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":186EC
            Key             =   "image"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18C88
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19224
            Key             =   "open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTB 
      Left            =   4320
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   55
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":197C0
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A614
            Key             =   "arrh"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B468
            Key             =   "arrv"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C2BC
            Key             =   "ascii"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D110
            Key             =   "casc"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DF64
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EDB8
            Key             =   "find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FC0C
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20A60
            Key             =   "print"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":218B4
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22708
            Key             =   "revert"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2355C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":243B0
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2494C
            Key             =   "new"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":257A0
            Key             =   "opend"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":265F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26750
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26CF4
            Key             =   "web"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":286C0
            Key             =   "html"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A08C
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AEE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B040
            Key             =   "open"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BE94
            Key             =   "sd"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C430
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D284
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D820
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DDBC
            Key             =   "close"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Tag             =   "The status bar displays current status and informs you of errors."
      Top             =   8025
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12065
            Picture         =   "frmMain.frx":2DF18
            Text            =   "Press F6 for quick help or F1 for contents."
            TextSave        =   "Press F6 for quick help or F1 for contents."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileNewWeb 
         Caption         =   "New &Web"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOpenURL 
         Caption         =   "Open From &URL"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "C&lose"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "S&ave as..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Sa&ve all"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSepBar 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileRevert 
         Caption         =   "Re&vert..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrinterSetup 
         Caption         =   "Print s&etup..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Enabled         =   0   'False
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu sepbarMRU 
         Caption         =   "-"
         Visible         =   0   'False
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
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "E&dit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditUndo 
         Caption         =   "U&ndo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "Cle&ar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select a&ll"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDefine 
         Caption         =   "De&finition"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "R&eplace..."
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditClean 
         Caption         =   "&Mark clean"
         Enabled         =   0   'False
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuDocument 
      Caption         =   "D&ocument"
      Visible         =   0   'False
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview..."
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuDocumentConvert 
         Caption         =   "&Convert..."
         Enabled         =   0   'False
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert..."
      Visible         =   0   'False
      Begin VB.Menu mnuInsertObject 
         Caption         =   "Object"
      End
      Begin VB.Menu mnuInsertSymbol 
         Caption         =   "&Symbol..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInsertDateTime 
         Caption         =   "D&ate/Time"
         Begin VB.Menu mnuInsertDateTimeLong 
            Caption         =   "[Time]"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuInsertDateTimeShort 
            Caption         =   "[Date]"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuInsertDateTimeSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInsertDateTimeWeekday 
            Caption         =   "[Day]"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuInsertStream 
         Caption         =   "S&tream..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuWeb 
      Caption         =   "&Web"
      Begin VB.Menu mnuNewWeb 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuOpenWeb 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuWebClose 
         Caption         =   "&Close"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebSepZ 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebRefresh 
         Caption         =   "&Refresh"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebDefault 
         Caption         =   "D&efault"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebSep2 
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
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Wi&ndow"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascadeWin 
         Caption         =   "&Cascade"
         Enabled         =   0   'False
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuTileHorizontal 
         Caption         =   "&Tile Horizontal"
         Enabled         =   0   'False
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuTileVertical 
         Caption         =   "Tile &Vertical"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "&Arrange Icons"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWinSepX 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowAlignLefts 
         Caption         =   "&Align Lefts"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWindowAlignTops 
         Caption         =   "Align &Tops"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWindowSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowMinimizeAll 
         Caption         =   "&Minimize all"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWindowMaximizeAll 
         Caption         =   "Ma&ximize all"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRestoreall 
         Caption         =   "&Restore all"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWindowUnloadAll 
         Caption         =   "&Unload all"
         Enabled         =   0   'False
         Shortcut        =   ^{F9}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpHomepage 
         Caption         =   "&Homepage"
      End
      Begin VB.Menu mnuHelpQuickInfo 
         Caption         =   "QuickInfo"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuHelpLegend 
         Caption         =   "&Legend..."
      End
      Begin VB.Menu mnuHSep 
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
   Begin VB.Menu mnuToolbarPop 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu mnuToolText 
         Caption         =   "Text labels"
      End
   End
   Begin VB.Menu mnuTree 
      Caption         =   "&Tree"
      Visible         =   0   'False
      Begin VB.Menu mnuTreeNewWeb 
         Caption         =   "New Web"
      End
      Begin VB.Menu mnuTreeBar55 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeOpen 
         Caption         =   "&Open file"
      End
      Begin VB.Menu mnuTreeLinkFile 
         Caption         =   "&Link file..."
      End
      Begin VB.Menu mnuTreeSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeCollapse 
         Caption         =   "&Collapse All"
      End
      Begin VB.Menu mnuTreeExpand 
         Caption         =   "&Expand All"
      End
      Begin VB.Menu mnuTreeSep567 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeDelete 
         Caption         =   "&Delete..."
      End
      Begin VB.Menu mnuTreeCopy 
         Caption         =   "&Copy to..."
      End
      Begin VB.Menu mnuTreeMove 
         Caption         =   "&Move to..."
      End
      Begin VB.Menu mnuTreeRename 
         Caption         =   "&Rename..."
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "P&roperties"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public WebMRU As Collection
Public FileMRU As Collection

Dim PreviousWeb As String
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const DRIVE_REMOVABLE = 2 ':( As Integer ?
Private Const DRIVE_FIXED = 3 ':( As Integer ?
Private Const DRIVE_REMOTE = 4 ':( As Integer ?
Private Const DRIVE_CDROM = 5 ':( As Integer ?
Private Const DRIVE_RAMDISK = 6 ':( As Integer ?
Const m_Syntax = True ':( As Integer ?

Private Sub Fldr_Change()

    On Error Resume Next
      Fil.Path = Fldr.Path

End Sub ':( On Error Resume still active

Private Sub MDIForm_Load()

    On Error Resume Next
      GetPrefs
      SetFont Me
      GetFileMRU
      NewTreeView
      GetWebMRU
      AddFlags
      AddScriptFiles
      ChDir App.Path
      LoadCMDLine
      ResizeBar
      If ReadValue("WebDefault") <> "" Then LoadWeb ReadValue("WebDefault") ':( Expand Structure
      Text1.text = GetSetting(App.Title, "FTP", "URL", "Your Name")
      Text2.text = GetSetting(App.Title, "FTP", "Username", "Your Name")
      Text3.text = GetSetting(App.Title, "FTP", "Password", "Your Name")
      FindCD

End Sub ':( On Error Resume still active

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
      mnuView_Click
      mnuFile_Click
      If Button = 2 And Shift = 0 And ActiveForm Is Nothing Then PopupMenu mnuView ':( Expand Structure
      If Button = 2 And Shift = 0 And Not ActiveForm Is Nothing Then PopupMenu ActiveForm.mnuView ':( Expand Structure
      If Button = 2 And Shift = 1 Then PopupMenu mnuFile ':( Expand Structure
      If Button = 2 And Shift = 2 Then PopupMenu mnuWeb ':( Expand Structure

End Sub ':( On Error Resume still active

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If SSTab1.Height <> pLeft.ScaleHeight - 15 Then MDIForm_Resize ':( Expand Structure
    SB.Panels(1).text = "Press F6 for quick help or F1 for contents."

End Sub

Sub MDIForm_Resize()

    On Error Resume Next
      SSTab1.Height = pLeft.ScaleHeight - 15
      tvD.Height = SSTab1.Height - 525
      tvW.Height = ReadValue("TreeHeight", tvD.Height)
      tvS.Height = tvD.Height
      ResizeBar

End Sub ':( On Error Resume still active

Private Sub MDIForm_Unload(Cancel As Integer)

    SetPrefs
    End

End Sub

Private Sub mnuCascadeWin_Click()

    Arrange vbCascade

End Sub

Private Sub mnuDocument_Click()

    SB.Panels(1).text = "Contains commands for inserting text, converting documents, etc."

End Sub

Private Sub mnuEdit_Click()

    SB.Panels(1).text = "Contains commands for editing documents."

End Sub

Private Sub mnuFile_Click()

    SB.Panels(1).text = "Contains commands for operating the program."

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub mnuFileMRU_Click(index As Integer)

    On Error Resume Next
    Dim lpF As New frmChild ':( Move line to top of current Sub
      Load lpF
      lpF.LoadHTMLFile mnuFileMRU(index).tag

End Sub ':( On Error Resume still active

Private Sub mnuFileNew_Click()

    NewDocument

End Sub

Private Sub mnuFilenewWeb_Click()

    mnuNewWeb_Click

End Sub

Private Sub mnuFileOpen_Click()

    OpenDocument

End Sub

Private Sub mnuFileProperties_Click()

    FileInfo FileLen(tvW.SelectedItem.key), tvW.SelectedItem.key

End Sub

Private Sub mnuFileStartOptions_Click()

    Load StartOptions
    StartOptions.Show

End Sub

Private Sub mnuHelp_Click()

    SB.Panels(1).text = "Contains help commands."

End Sub

Private Sub mnuHelpAbout_Click()

    frmAbout.Show vbModal

End Sub

Private Sub mnuHelpContents_Click()

    ShellExecute 0, "open", App.Path & "\help\index.html", "", "", 10

End Sub

Private Sub mnuHelpHomepage_Click()

  'ShellExecute 0, "open", "http://sushantshome.tripod.com/vb/wonder.html", "", "", 10


End Sub

Private Sub mnuHelpQuickInfo_Click()

    On Error Resume Next
      MsgBox ActiveControl.tag, vbInformation

End Sub ':( On Error Resume still active

Private Sub mnuNewWeb_Click()

    On Error GoTo hell
  Dim lpLoc As String ':( Move line to top of current Sub
    lpLoc = SelectDir()
    If lpLoc = "" Then Exit Sub ':( Expand Structure
    If IsWebOpen Then CloseWeb ':( Expand Structure
    MkDir lpLoc
    LoadWeb lpLoc

Exit Sub

hell:
    MsgBox Error, vbExclamation

End Sub

Private Sub mnuOpenURL_Click()

    Load frmBrowse
    frmBrowse.Show

End Sub

Private Sub mnuOpenWeb_Click()

    On Error GoTo hell
  Dim lpLoc As String ':( Move line to top of current Sub
    lpLoc = SelectDir(True)
    If lpLoc = "" Then Exit Sub ':( Expand Structure
    LoadWeb lpLoc

Exit Sub

hell:
    MsgBox Error, vbExclamation

End Sub

Private Sub mnuRestoreAll_Click()

  Dim lpF As Form

    For Each lpF In Forms
        If lpF.Caption <> "Designer 2001 Personal Edition" Then
            lpF.WindowState = vbNormal
        End If
    Next lpF

End Sub

Private Sub mnuTileHorizontal_Click()

    Arrange vbHorizontal

End Sub

Private Sub mnuTileVertical_Click()

    Arrange vbTileVertical

End Sub

Private Sub mnuTodayTip_Click()

    On Error Resume Next
      frmTip.Show vbModal

End Sub ':( On Error Resume still active

'Private Sub mnuToolbarPop_Click()
'mnuToolText.Checked = (TB.Buttons(1).Caption <> "")
'End Sub

Private Sub mnuToolText_Click()

    mnuToolText.Checked = Not mnuToolText.Checked
    'ClearToolBar mnuToolText.Checked
    SaveValue "ToolText", mnuToolText.Checked
    MDIForm_Resize

End Sub

Private Sub mnuTreeCollapse_Click()

  Dim i As Integer

    For i = 1 To tvW.Nodes.Count
        tvW.Nodes(i).Expanded = False
    Next i
    tvW.SelectedItem = tvW.Nodes(1)

End Sub

Private Sub mnuTreeCopy_Click()

    On Error Resume Next
    Dim loc As String ':( Move line to top of current Sub

      loc = SelectDir(True)
      If loc = "" Then Exit Sub ':( Expand Structure

      If right$(loc, 1) <> "\" Then loc = loc & "\" ':( Expand Structure

      MousePointer = 11
      DoEvents
      If CopyFile(tvW.SelectedItem.key, loc & tvW.SelectedItem.text) Then
          tvW.Nodes.Add left$(loc, Len(loc) - 1), tvwChild, loc & tvW.SelectedItem.text, tvW.SelectedItem.text, FileIcon(tvW.SelectedItem.text)
      End If
      DoEvents
      MousePointer = 0

End Sub ':( On Error Resume still active

Private Sub mnuTreeDelete_Click()

    On Error Resume Next
      MousePointer = 11
      DoEvents
      If DeleteFile(tvW.SelectedItem.key) Then tvW.Nodes.Remove tvW.SelectedItem.index ':( Expand Structure
      DoEvents
      MousePointer = 0

End Sub ':( On Error Resume still active

Private Sub mnuTreeExpand_Click()

  Dim i As Integer

    For i = 1 To tvW.Nodes.Count
        tvW.Nodes(i).Expanded = True
    Next i
    tvW.SelectedItem = tvW.Nodes(1)

End Sub

Private Sub mnuTreeLinkFile_Click()

    On Error Resume Next
    Dim Path As String ':( Move line to top of current Sub
      If ActiveForm Is Nothing Then Exit Sub ':( Expand Structure
      If ActiveForm.Caption = "Untitled" Then Path = tvW.SelectedItem.key: GoTo n ':( Expand Structure
      Path = Replace$(tvW.SelectedItem.key, Up1Level(ActiveForm.Caption), "")
      Path = Replace$(Path, "\", "/")
      If left$(Path, 1) = "/" Then Path = right$(Path, Len(Path) - 1) ':( Expand Structure
n:       'next
      ActiveForm.RTF1.SelText = "<A href=" & Chr$(34) & Path & Chr$(34) & ">" & Path & "</A>"
      ActiveForm.RTF1.SelStart = ActiveForm.RTF1.SelStart - Len(Path) - 4
      ActiveForm.RTF1.SetFocus

End Sub ':( On Error Resume still active

Private Sub mnuTreeMove_Click()

    On Error Resume Next
    Dim loc As String ':( Move line to top of current Sub

      loc = SelectDir(True)
      If loc = "" Then Exit Sub ':( Expand Structure

      If right$(loc, 1) <> "\" Then loc = loc & "\" ':( Expand Structure

      MousePointer = 11
      DoEvents
      If MoveFile(tvW.SelectedItem.key, loc & tvW.SelectedItem.text) Then
          tvW.Nodes.Add left$(loc, Len(loc) - 1), tvwChild, loc & tvW.SelectedItem.text, tvW.SelectedItem.text, FileIcon(tvW.SelectedItem.text)
          tvW.Nodes.Remove tvW.SelectedItem.index
      End If
      DoEvents
      MousePointer = 0

End Sub ':( On Error Resume still active

Private Sub mnuTreeNewWeb_Click()

    mnuNewWeb_Click

End Sub

Private Sub mnuTreeOpen_Click()

    OnNodeClick tvW.SelectedItem

End Sub

Private Sub mnuTreeRename_Click()

    tvW.StartLabelEdit

End Sub

Private Sub mnuViewDocuments_Click()

    mnuViewDocuments.Checked = Not mnuViewDocuments.Checked
    SSTab1.TabVisible(1) = mnuViewDocuments.Checked
    SaveValue "DocumentTree", mnuViewDocuments.Checked

End Sub

Private Sub mnuViewOptions_Click()

    frmOpts.Show vbModal

End Sub

Private Sub mnuViewScripts_Click()

    mnuViewScripts.Checked = Not mnuViewScripts.Checked
    SSTab1.TabVisible(2) = mnuViewScripts.Checked
    SaveValue "ScriptView", mnuViewScripts.Checked

End Sub

Private Sub mnuWeb_Click()

    On Error GoTo hell
    SB.Panels(1).text = "Contains commands for manipulating webs."
    mnuWebDefault.Enabled = (tvW.Nodes.Count > 0)
    mnuWebRefresh.Enabled = mnuWebDefault.Enabled
    mnuWebClose.Enabled = (tvW.Nodes.Count > 0)
    mnuWebDefault.Checked = (ReadValue("WebDefault") = frmMain.tvW.Nodes(1).key)

Exit Sub

hell:
    mnuWebDefault.Enabled = False
    Resume Next

End Sub

Private Sub mnuWebClose_Click()

    CloseWeb

End Sub

Private Sub mnuWebDefault_Click()

    mnuWebDefault.Checked = Not mnuWebDefault.Checked
    If mnuWebDefault.Checked Then
        SaveValue "WebDefault", tvW.Nodes(1).key 'root
      Else
        SaveValue "WebDefault", "" 'nothing
    End If

End Sub

Private Sub mnuWebMRU_Click(index As Integer)

    LoadWeb mnuWebMRU(index).tag

End Sub

Private Sub mnuWebRefresh_Click()

    LoadWeb PreviousWeb

End Sub

Private Sub mnuWindow_Click()

    SB.Panels(1).text = "Contains commands for arranging and navigating windows."

End Sub

Private Sub pLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    SSTab1_MouseMove 0, 0, 0, 0 'dummies

End Sub

Private Sub pS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    tC.Enabled = True

End Sub

Private Sub pS_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    tC.Enabled = False
    ResizeBar
    SaveValue "TreeWidth", pLeft.Width

End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If SSTab1.Height <> pLeft.ScaleHeight - 15 Then MDIForm_Resize ':( Expand Structure

End Sub

Private Sub SB_PanelClick(ByVal Panel As MSComctlLib.Panel)

  'KEEP THIS NEXT LINE OF CODE AND SEE HOW TO WORK IT OUT
  'If Panel.Index = 4 Then frmInfo.Show vbModal


End Sub

'Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim lWidth As Long
'On Error GoTo h
'Select Case Button.Index
'Case 5 'save all
'ActiveForm.mnuFileSaveAll_Click
'Case 9 'close
'Unload ActiveForm
'Case 20 'default insert
'lWidth = (TB.ButtonWidth * 11) + (6 * TB.Buttons(4).Width) '7 seps, 12 buttons before 20
'TB.Buttons(20).Value = tbrPressed
'SB.Panels(1).text = "Contains commands for inserting symbols, converting documents, etc."
'PopupMenu ActiveForm.mnuDocument, , lWidth, TB.ButtonHeight + 45
'TB.Buttons(20).Value = tbrUnpressed
'Case 24 'test
'ActiveForm.mnuPreview_Click
'Case 25 'convert
'ActiveForm.mnuDocumentConvert_Click
'End Select
'h:
'If Err.Number = 91 Then SB.Panels(1).text = "No documents are currently open."
'End Sub

'Private Sub TB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 2 Then PopupMenu mnuToolbarPop
'End Sub

Sub NewDocument()

    On Error Resume Next
    Dim lpF As New frmChild ':( Move line to top of current Sub
      Load lpF
      lpF.Show
      lpF.RTF1.text = HTML 'the constant
      lpF.RTF1.SelStart = lpF.GetSelStart  'approximately
      'lpF.Undo.Remove lpF.Undo.Count
      'lpF.Undo.Remove lpF.Undo.Count
      'lpF.Undo.AddAction lpF.RTF1.text, lpF.RTF1.SelStart
      lpF.bChanged = False 'document not changed
      lpF.SetFocus: lpF.RTF1.SetFocus

End Sub ':( On Error Resume still active

Sub OpenDocument()

  Dim lpF As New frmChild

    On Error GoTo hell
    CD.ShowOpen
n:
    Load lpF
    lpF.LoadHTMLFile CD.fileName
hell:

End Sub

Private Sub mnuView_Click()

    SB.Panels(1).text = "Contains commands for manipulating the view."
    mnuViewToolBar.Checked = (frmMain.Toolbar1.Visible)
    mnuViewStatusBar.Checked = frmMain.SB.Visible
    mnuViewFileTree.Checked = frmMain.pLeft.Visible
    mnuViewDocuments.Checked = frmMain.SSTab1.TabVisible(1)
    mnuViewScripts.Checked = frmMain.SSTab1.TabVisible(2)

End Sub

Sub mnuViewFileTree_Click()

    mnuViewFileTree.Checked = Not mnuViewFileTree.Checked
    pLeft.Visible = mnuViewFileTree.Checked
    SaveValue "FileTree", pLeft.Visible

End Sub

Private Sub mnuViewStatusBar_Click()

    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    frmMain.SB.Visible = mnuViewStatusBar.Checked
    SaveValue "Statusbar", SB.Visible

End Sub

Private Sub mnuViewToolBar_Click()

    mnuViewToolBar.Checked = Not mnuViewToolBar.Checked
    frmMain.CoolBar1.Visible = mnuViewToolBar.Checked
    'frmMain.TB2.Visible = mnuViewToolBar.Checked
    SaveValue "Toolbar", frmMain.CoolBar1.Visible

End Sub

Sub GetPrefs()

    On Error Resume Next
      WindowState = ReadValue("WindowState")
      If WindowState = 0 Then
          Width = ReadValue("Width")
          Height = ReadValue("Height")
          top = ReadValue("Top")
          left = ReadValue("Left")
      End If
      pLeft.Width = ReadValue("TreeWidth")
      pLeft.Visible = ReadValue("FileTree", True)
      'TB.Style = ReadValue("FlatBar")
      'TB2.Style = ReadValue("FlatBar")
      SSTab1.TabVisible(1) = ReadValue("DocumentTree", True)
      SSTab1.TabVisible(2) = ReadValue("ScriptView", True)
      frmMain.CoolBar1.Visible = ReadValue("Toolbar", True)
      SB.Visible = ReadValue("Statusbar", True)
      'If ReadValue("ToolText", False) Then ClearToolBar ReadValue("ToolText", False)
      ResizeBar

End Sub ':( On Error Resume still active

Sub SetPrefs()

    On Error Resume Next
      SaveValue "WindowState", WindowState
      If WindowState = 0 Then
          SaveValue "Width", Width
          SaveValue "Height", Height
          SaveValue "Top", top
          SaveValue "Left", left
      End If

End Sub ':( On Error Resume still active

Sub LoadWeb(lpFilePath As String)

    If Dir$(lpFilePath, vbDirectory) = "" Then Exit Sub ':( Expand Structure
    On Error Resume Next

      ChDir lpFilePath
      PreviousWeb = mnuWebMRU(1).tag
      If right$(lpFilePath, 1) = "\" Then lpFilePath = left$(lpFilePath, Len(lpFilePath) - 1) ':( Expand Structure

      SB.Panels(1).text = "Loading " & lpFilePath & " web..."
      MousePointer = 11

    Dim i As Long, ii As Long, iii As Long ':( Move line to top of current Sub

      Fldr.Path = lpFilePath
      Fldr.Refresh: Fil.Refresh

      tvW.Nodes.Clear
      tvW.Nodes.Add , , Fldr.Path, GetFile(Fldr.Path), "!open"

      For i = 0 To Fldr.ListCount - 1
          tvW.Nodes.Add Fldr.Path, tvwChild, Fldr.List(i), GetFile(Fldr.List(i)), "closed"
          tvW.Nodes.Item(Fldr.List(i)).ExpandedImage = "open"
          tvW.Nodes.Add Fldr.List(i), tvwChild, "", ""
      Next i

      Fil.Path = lpFilePath

      For ii = 0 To Fil.ListCount - 1
          tvW.Nodes.Add lpFilePath, tvwChild, FullPath(Fil.Path, Fil.List(ii)), Fil.List(ii), FileIcon(Fil.List(ii))
      Next ii

      SB.Panels(1).text = ""

      AddWebMRU lpFilePath

      tvW.Nodes(1).Expanded = True

      MousePointer = 0

End Sub ':( On Error Resume still active

Private Sub OnNodeClick(ByVal Node As MSComctlLib.Node)

    On Error Resume Next
    Dim lpF As New frmChild, lpD As Form ':( Move line to top of current Sub

      MousePointer = 11

      SB.Panels(1).text = Node.key

      Select Case Node.Image

        Case 5, "closed"
          If Node.Children > 1 And Node.Child.text <> "" Then MousePointer = 0: Exit Sub ':( Expand Structure
          If Node.Child.text = "" Then tvW.Nodes.Remove Node.Child.index ':( Expand Structure
    Dim i As Long ':( Move line to top of current Sub
          Fldr.Path = Node.key
          If Fldr.ListCount = 0 Then GoTo nfil ':( Expand Structure
nfld:
          For i = 0 To Fldr.ListCount - 1
              tvW.Nodes.Add Node.key, tvwChild, Fldr.List(i), GetFile(Fldr.List(i)), "closed"
              tvW.Nodes.Item(Fldr.List(i)).ExpandedImage = "open"
              tvW.Nodes.Add Fldr.List(i), tvwChild, "", ""
          Next i
nfil:
          For i = 0 To Fil.ListCount - 1
              tvW.Nodes.Add Node.key, tvwChild, FullPath(Fil.Path, Fil.List(i)), Fil.List(i), FileIcon(Fil.List(i))
          Next i

          MousePointer = 0
          Exit Sub '>---> Bottom

        Case 4, "!open"
          MousePointer = 0
          Exit Sub '>---> Bottom
    
        Case Else 'files and images

          'loop and find if it's already open if yes then set focus to it
          For Each lpD In Forms
              If lpD.Caption = Node.key Then
                  lpD.SetFocus: lpD.RTF1.SetFocus: lpD.pB.SetFocus: MousePointer = 0: Exit Sub
              End If
          Next lpD
        
          Select Case LCase$(right$(Node.key, 3))
        
              'open based on the extension
            Case "tml", "css", "txt", "asp", "htm", ".js", "vbs", "xml"
        
              Load lpF
              lpF.LoadHTMLFile Node.key
        
            Case "jpg", "gif", "bmp", "ico"
        
              'below line confirms if user wants the internal image viewer
              If ReadValue("ImageViewer", 1) = 0 Then GoTo nope ':( Expand Structure
              LoadImage Node.key
        
            Case Else
nope:
        
              'try to execute otherwise notify that file couldn't be opened
              If ShellExecute(Me.hWnd, "open", Node.key, "", Up1Level(Node.key), 10) < 32 Then MsgBox "Failed to execute " & GetFile(Node.key), vbExclamation ':( Expand Structure
        
          End Select

      End Select

      For i = 1 To tvW.Nodes.Count
          tvW.Nodes(i).Bold = False
      Next i
      Node.Bold = True

      MousePointer = 0

End Sub ':( On Error Resume still active

Sub AddFlags()

  'add flags to common dialog

    CD.Flags = cdlOFNCreatePrompt + cdlOFNFileMustExist + cdlOFNOverwritePrompt + cdlOFNPathMustExist

End Sub

Function IsWebOpen() As Boolean

  'is any web open in the manager

    IsWebOpen = (tvW.Nodes.Count > 0)

End Function

Function CloseWeb()

  'does just what it says

    mnuWebDefault.Checked = False
    tvW.Nodes.Clear

End Function

Private Sub FindNodeText()

  'this is called when user clicks on document outline tree

    On Error Resume Next
    Dim lF As Long ':( Move line to top of current Sub
      lF = InStr(1, ActiveForm.RTF1.text, SB.Panels(1).text)
      If lF = 0 Then Exit Sub ':( Expand Structure
      ActiveForm.RTF1.SelStart = lF - 1
      If ReadValue("SelectFind") = True Then ActiveForm.RTF1.SelLength = Len(SB.Panels(1).text) Else ActiveForm.RTF1.SetFocus ':( Expand Structure

End Sub ':( On Error Resume still active

Private Sub tC_Timer()

  'resize left bar
  
  Dim lpP As POINTAPI

    GetCursorPos lpP
    pLeft.Width = lpP.x * Screen.TwipsPerPixelX

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim lWidth As Long

    On Error GoTo h
    Select Case Button.key
      Case "new" 'new
        NewDocument
      Case "open" 'open
        OpenDocument
      Case "save" 'save
        ActiveForm.mnuFileSave_Click
      Case 5 'save all
        ActiveForm.mnuFileSaveAll_Click
      Case "print" 'print

      Case 9 'close
        Unload ActiveForm
      Case "newweb" 'web
        lWidth = (Toolbar1.ButtonWidth * 2) '+ (3 * Toolbar1.Buttons(4).Width) '4 seps, 6 buttons before 11
        Toolbar1.Buttons(2).value = tbrPressed
        mnuWeb_Click 'enable/disable
        PopupMenu mnuWeb, , lWidth, Toolbar1.ButtonHeight + 15
        'PopupMenu mnuWeb, , Toolbar1.ButtonHeight + 45

        Toolbar1.Buttons(2).value = tbrUnpressed
      Case "undo" 'undo
        ActiveForm.mnuEditUndo_Click
        'TB.Buttons("undo").Enabled = ActiveForm.Undo.UndoAvailable
        'TB.Buttons("redo").Enabled = ActiveForm.Undo.RedoAvailable
      Case "redo" 'redo
        ActiveForm.mnuEditRedo_Click
        'TB.Buttons("redo").Enabled = ActiveForm.Undo.RedoAvailable
        'TB.Buttons("undo").Enabled = ActiveForm.Undo.UndoAvailable
      Case "delete"
        ActiveForm.RTF1.SelText = " "
      Case "cut"
        ActiveForm.mnuEditCut_Click
      Case "bold"
        ActiveForm.RTF1.SelText = "<B>" & ActiveForm.RTF1.SelText & "</B>"
      Case "italic"
        ActiveForm.RTF1.SelText = "<I>" & ActiveForm.RTF1.SelText & "</I>"

      Case "underline"
        ActiveForm.RTF1.SelText = "<U>" & ActiveForm.RTF1.SelText & "</U>"

      Case "strike"
        ActiveForm.RTF1.SelText = "<S>" & ActiveForm.RTF1.SelText & "</S>"

      Case "copy" 'copy
        ActiveForm.mnuEditCopy_Click
      Case "paste" 'paste
        ActiveForm.mnuEditPaste_Click
      Case 20 'default insert
        'lWidth = (TB.ButtonWidth * 11) + (6 * TB.Buttons(4).Width) '7 seps, 12 buttons before 20
        'TB.Buttons(20).Value = tbrPressed
        'SB.Panels(1).text = "Contains commands for inserting symbols, converting documents, etc."
        'PopupMenu ActiveForm.mnuDocument, , lWidth, TB.ButtonHeight + 45
        'TB.Buttons(20).Value = tbrUnpressed
      Case "find" 'find
        ActiveForm.mnuEditFind_Click
      Case 24 'test
        ActiveForm.mnuPreview_Click
      Case "convert" 'convert
        ActiveForm.mnuDocumentConvert_Click
      Case "left"
        ActiveForm.RTF1.SelText = "<DIV align=" & Chr$(34) & "left" & Chr$(34) & ">" & ActiveForm.RTF1.SelText & "</DIV>"

      Case "center"
        ActiveForm.RTF1.SelText = "<DIV align=" & Chr$(34) & "center" & Chr$(34) & ">" & ActiveForm.RTF1.SelText & "</DIV>"

      Case "right"
        ActiveForm.RTF1.SelText = "<DIV align=" & Chr$(34) & "right" & Chr$(34) & ">" & ActiveForm.RTF1.SelText & "</DIV>"
        Case "CleanCode"
        ActiveForm.RTF1.text = FormatCode(ActiveForm.RTF1.text)
    End Select
h:
    If Err.Number = 91 Then SB.Panels(1).text = "No documents are currently open." ':( Expand Structure

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
      Case "fonts"
        Load frmFont
        frmFont.Text1.text = ActiveForm.RTF1.SelText
        frmFont.Text2.SelText = "<FONT SIZE=" & Chr$(34) & frmFont.Combo1.text & Chr$(34) & ">" & frmFont.Text1.text & "</FONT>"
        frmFont.Show

      Case "image"
        Load frmImageTwo
        If ActiveForm Is Nothing Then
            frmImageTwo.Command1.Enabled = False
          Else
            frmImageTwo.Command1.Enabled = True
        End If
        frmImageTwo.Show

      Case "lrgfont"
        ActiveForm.RTF1.SelText = "<BIG>" & ActiveForm.RTF1.SelText & "</BIG>"

      Case "smfont"
        ActiveForm.RTF1.SelText = "<SMALL>" & ActiveForm.RTF1.SelText & "</SMALL>"

      Case "bullet"
        Load frmLists
        frmLists.Show

      Case "number"
        Load frmLists
        frmLists.Show

      Case "links"
        'Load frmLinks
        'frmLinks.Show

      Case "cgi"
        'Load frmCGI
        'frmCGI.Show

      Case "css"
        'Load frmCSS
        'frmCSS.Show

      Case "js"
        Load frmJavaScript
        frmJavaScript.Show

      Case "vbs"
        'Load frmVBScript
        'frmVBScript.Show

    End Select

End Sub

Private Sub tvD_Collapse(ByVal Node As MSComctlLib.Node)

    tvD.SelectedItem = Node

End Sub

Private Sub tvD_DblClick()

    FindNodeText 'find the text

End Sub

Private Sub tvD_Expand(ByVal Node As MSComctlLib.Node)

    tvD.SelectedItem = Node

End Sub

Private Sub tvD_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
      If ActiveForm Is Nothing Then Exit Sub ':( Expand Structure
      If Button = 2 Then PopupMenu ActiveForm.mnuWhatever ':( Expand Structure

End Sub ':( On Error Resume still active

Private Sub tvD_NodeClick(ByVal Node As MSComctlLib.Node)

  Dim iTag As String

    'based on image, if it is a file or folder, accordingly show tag
    If Node.Image > 2 Then iTag = Node.Children & " sub-items under '" & Node.text & "'." ':( Expand Structure
    If Node.Image < 2 Then iTag = Node.Children & " sub-items under '" & Node.text & "'." ':( Expand Structure
    If Node.Image = 2 Then iTag = "<" & Node.tag & ">" ':( Expand Structure
    If iTag <> "<>" Then SB.Panels(1).text = iTag Else SB.Panels(1).text = "" ':( Expand Structure

End Sub

Private Sub tvS_Collapse(ByVal Node As MSComctlLib.Node)

    tvS.SelectedItem = Node

End Sub

Private Sub tvS_DblClick()

    On Error Resume Next
      If tvS.SelectedItem.Image = "fp" Then Exit Sub 'it's root':( Expand Structure
      If tvS.SelectedItem.Image = "fp2" Then Exit Sub 'it's root':( Expand Structure
      If tvS.SelectedItem.Image = "c" Then Exit Sub 'its a folder':( Expand Structure
      If tvS.SelectedItem.Image = "function" Or tvS.SelectedItem.Image = "variable" Then
          ActiveForm.RTF1.SelStart = InStr(1, ActiveForm.RTF1.text, tvS.SelectedItem.key) - 1
          ActiveForm.SetFocus: ActiveForm.RTF1.SetFocus
          Exit Sub '>---> Bottom
      End If
    Dim lpF As New frmChild ':( Move line to top of current Sub
      Load lpF
      lpF.LoadHTMLFile tvS.SelectedItem.key

End Sub ':( On Error Resume still active

Private Sub tvS_Expand(ByVal Node As MSComctlLib.Node)

    tvS.SelectedItem = Node

End Sub

Private Sub tvS_NodeClick(ByVal Node As MSComctlLib.Node)

    SB.Panels(1).text = Node.key

End Sub

Private Sub tvW_AfterLabelEdit(Cancel As Integer, NewString As String)

  'If tvW.SelectedItem.Image = "closed" Or tvW.SelectedItem.Image = "!open" Then NewString = tvW.SelectedItem.Text: Exit Sub

    If Cancel > 0 Then Exit Sub ':( Expand Structure
    If NewString = tvW.SelectedItem.text Then Exit Sub ':( Expand Structure

  Dim StrPS As String ':( Move line to top of current Sub

    If right$(tvW.SelectedItem.key, 1) = "\" Then StrPS = "" Else StrPS = "\" ':( Expand Structure

    MousePointer = 11
    DoEvents
    If Not RenameFile(tvW.SelectedItem.key, tvW.SelectedItem.Parent.key & StrPS & NewString) Then NewString = tvW.SelectedItem.text ':( Expand Structure
    tvW.SelectedItem.key = tvW.SelectedItem.Parent.key & StrPS & NewString
    DoEvents
    MousePointer = 0

End Sub

Private Sub tvW_Collapse(ByVal Node As MSComctlLib.Node)

    tvW.SelectedItem = Node

End Sub

Private Sub tvW_DblClick()

    On Error Resume Next
      OnNodeClick tvW.SelectedItem

End Sub ':( On Error Resume still active

Private Sub tvW_Expand(ByVal Node As MSComctlLib.Node)

    tvW.SelectedItem = Node
    OnNodeClick Node

End Sub

Private Sub tvW_KeyPress(KeyAscii As Integer)

    If KeyAscii = 32 Then OnNodeClick tvW.SelectedItem ':( Expand Structure

End Sub

Private Sub tvW_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 And tvW.Nodes.Count > 0 Then
        PopupMenu mnuTree
    End If

End Sub

Private Sub tvW_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If SSTab1.Height <> pLeft.ScaleHeight - 15 Then MDIForm_Resize ':( Expand Structure

End Sub

'Sub ClearToolBar(lpExpression As Boolean)
'remove or add captions
'Dim i As Integer
'If lpExpression Then
'For i = 1 To TB.Buttons.Count
'    TB.Buttons(i).Caption = TB.Buttons(i).tag
'Next i
'Else
'For i = 1 To TB.Buttons.Count
'    TB.Buttons(i).Caption = ""
'Next i
'End If
'I realize this can be done in a shorter way
'MDIForm_Resize
'End Sub

Sub ResizeBar()

    On Error Resume Next
      SSTab1.Width = pLeft.ScaleWidth - 45
      tvD.Width = SSTab1.Width - 190
      tvW.Width = tvD.Width
      tvS.Width = tvD.Width
      pS.Move pLeft.ScaleWidth - pS.Width, 0, 45, pLeft.ScaleHeight

End Sub ':( On Error Resume still active

Private Sub tvW_NodeClick(ByVal Node As MSComctlLib.Node)

    If Node.Image = 5 Or Node.Image = "closed" Then
        SB.Panels(1).text = "Double-click to view sub-items in '" & Node.text & "'."
      Else
        SB.Panels(1).text = Node.key
    End If

End Sub

Sub LoadCMDLine()

    If Command$() <> "" Then
  Dim lpF As New frmChild ':( Move line to top of current Sub
        Load lpF
        lpF.RTF1.LoadFile Command$(), rtfText
        lpF.Caption = Command$
        lpF.bChanged = False
    End If

End Sub

Function IsInQuotes(SelStart As Long) As Boolean

    On Error Resume Next
    Dim posLT As Long, posGT As Long ':( Move line to top of current Function
      posLT = InStr(SelStart, ActiveForm.RTF1.text, "<")
      If posLT = 0 Then IsInQuotes = True: Exit Function ':( Expand Structure
      posGT = InStr(SelStart, ActiveForm.RTF1.text, ">")
      If posGT = 0 Then IsInQuotes = True: Exit Function ':( Expand Structure
      IsInQuotes = (posLT > posGT)

End Function ':( On Error Resume still active

Sub GoOutsideQuotes(SelStart As Long)

    On Error Resume Next
    Dim iPos As Long ':( Move line to top of current Sub
      iPos = InStr(SelStart, ActiveForm.RTF1.text, ">")
      If iPos = 0 Then ActiveForm.RTF1.SelStart = Len(ActiveForm.RTF1.text): Exit Sub ':( Expand Structure
      ActiveForm.RTF1.SelStart = iPos

End Sub ':( On Error Resume still active

Sub NewTreeView()

  'supposed to subclass TreeView for better display


End Sub

Private Sub FindCD()

  Dim r&, allDrives$, JustOneDrive$, pos%, DriveType& ':( Type Suffixes are obsolete
  Dim CDfound As Integer

    allDrives$ = Space$(64)
    r& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)
    allDrives$ = left$(allDrives$, r&)
    Do
        pos% = InStr(allDrives$, Chr$(0))
        If pos% Then
            JustOneDrive$ = left$(allDrives$, pos%)
            allDrives$ = Mid$(allDrives$, pos% + 1, Len(allDrives$))
            DriveType& = GetDriveType(JustOneDrive$)
            If DriveType& = DRIVE_CDROM Then
                CDfound% = True
                Exit Do '>---> Loop
            End If
        End If
    Loop Until allDrives$ = "" Or DriveType& = DRIVE_CDROM
    If CDfound% Then
        txtCD.text = UCase$(JustOneDrive$)
      Else: MsgBox "No CD-ROM drives were detected on your system." ':( Expand Structure
    End If

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:59 PM) 15 + 1139 = 1154 Lines
