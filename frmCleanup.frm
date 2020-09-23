VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCleanup 
   Caption         =   "Code Cleanup"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdl 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   165
      Width           =   3360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox txtSource 
      Height          =   6225
      Left            =   135
      TabIndex        =   1
      Top             =   495
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   10980
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmCleanup.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "&Format"
      Height          =   375
      Left            =   8460
      TabIndex        =   0
      Top             =   45
      Width           =   1005
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuConfigure 
         Caption         =   "&Configure"
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
   End
End
Attribute VB_Name = "frmCleanup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFormat_Click()
    'txtSource.text = FormatCode(txtSource.text)
End Sub

Private Sub Command1_Click()
cdl.ShowOpen
txtSource.LoadFile (cdl.fileName)
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Resize()
On Error GoTo oops
    With txtSource
        .Width = Me.Width - .left - 200
        .Height = Me.Height - .top - 800
    
        cmdFormat.left = .left + .Width - cmdFormat.Width
    End With
    
oops:
Exit Sub
    
End Sub

Private Sub mnuClear_Click()
    txtSource.text = ""
End Sub

Private Sub mnuConfigure_Click()
    On Error Resume Next
    frmConfigFormat.Show
End Sub

Private Sub mnuCopy_Click()
    Clipboard.SetText txtSource.text
End Sub

Private Sub mnuPaste_Click()
    txtSource = Clipboard.GetText
End Sub

Private Sub txtSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuRightClick
    End If
End Sub
