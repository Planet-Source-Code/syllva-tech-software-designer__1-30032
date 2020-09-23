VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmFont 
   Caption         =   "Font Resizer"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Advanced"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Insert"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   5055
   End
   Begin SHDocVwCtl.WebBrowser wbFont 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   5055
      ExtentX         =   8916
      ExtentY         =   5106
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmFont.frx":0000
      Left            =   3360
      List            =   "frmFont.frx":0019
      TabIndex        =   0
      Top             =   1830
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Choose Font Size --->>"
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   1890
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selected Text:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()

    Text2.text = ""
    Text2.SelText = "<FONT SIZE=" & Chr$(34) & Combo1.text & Chr$(34) & ">" & Text1.text & "</FONT>"
    Render

End Sub

Private Sub Command1_Click()

  'Text2.SelText = "<FONT SIZE=" & Chr$(34) & Combo1.text & Chr$(34) & ">" & Text1.text & "</FONT>"

    Render

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Command3_Click()

    frmMain.ActiveForm.RTF1.SelText = "<FONT SIZE=" & Chr$(34) & Combo1.text & Chr$(34) & ">" & frmMain.ActiveForm.RTF1.SelText & "</FONT>"
    Unload Me

End Sub

Private Sub Form_Load()

    wbFont.Navigate ("about:blank")
    'Combo1.ListIndex = 2
    Render

End Sub

Private Sub Render()

    wbFont.Document.Script.Document.Clear
    wbFont.Document.Script.Document.Write Text2.text
    wbFont.Document.Script.Document.Close

    Exit Sub

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:08 PM) 1 + 49 = 50 Lines
