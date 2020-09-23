VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form colors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Color Scheme"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Build Code"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Text"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   555
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Background"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LINK"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1035
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ALINK"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1515
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "VLINK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1995
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   6015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Insert"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   240
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   615
      Width           =   1335
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   135
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Text Color:"
      Height          =   195
      Left            =   3360
      TabIndex        =   21
      Top             =   645
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Background Color:"
      Height          =   195
      Left            =   3360
      TabIndex        =   20
      Top             =   165
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Link Color:"
      Height          =   195
      Left            =   3360
      TabIndex        =   19
      Top             =   1125
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Active Link Color:"
      Height          =   195
      Left            =   3360
      TabIndex        =   18
      Top             =   1605
      Width           =   1245
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Visited Link Color:"
      Height          =   195
      Left            =   3360
      TabIndex        =   17
      Top             =   2085
      Width           =   1260
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   1095
      Width           =   1335
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   1575
      Width           =   1335
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   2055
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   75
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample Text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   555
      Width           =   1815
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "xyz@domain.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   1035
      Width           =   1815
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "xyz@domain.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   1515
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "xyz@domain.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1995
      Width           =   1815
   End
End
Attribute VB_Name = "colors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RedValue, GreenValue, BlueValue ':( As Variant ?
Dim AColor ':( As Variant ?
Dim ChoosenColor ':( As Variant ?

Private Sub Command1_Click()

    CMDialog1.Flags = &H1& Or &H4&
    CMDialog1.Action = 3
    AColor = CMDialog1.Color
    RedValue = (AColor And &HFF&)
    GreenValue = (AColor And &HFF00&) \ 256
    BlueValue = (AColor And &HFF0000) \ 65536
    ChoosenColor = Format$(Hex$(RedValue) & Hex$(GreenValue) & Hex$(BlueValue), "000000")
    Label3.Caption = ChoosenColor
    Text5.SelText = " BGCOLOR=#" & ChoosenColor
    Label11.BackColor = CMDialog1.Color
    Label12.BackColor = CMDialog1.Color
    Label13.BackColor = CMDialog1.Color
    Label14.BackColor = CMDialog1.Color
    Label15.BackColor = CMDialog1.Color
    Command9.Caption = "&Done"

End Sub

Private Sub Command2_Click()

    CMDialog1.Flags = &H1& Or &H4&
    CMDialog1.Action = 3
    AColor = CMDialog1.Color
    RedValue = (AColor And &HFF&)
    GreenValue = (AColor And &HFF00&) \ 256
    BlueValue = (AColor And &HFF0000) \ 65536
    ChoosenColor = Format$(Hex$(RedValue) & Hex$(GreenValue) & Hex$(BlueValue), "000000")
    Label12.ForeColor = CMDialog1.Color
    Label4.Caption = ChoosenColor
    Text5.SelText = " TEXT=#" & ChoosenColor
    Command9.Caption = "&Done"

End Sub

Private Sub Command4_Click()

    Label11.BackColor = &H80000005
    Label12.BackColor = &H80000005
    Label13.BackColor = &H80000005
    Label14.BackColor = &H80000005
    Label15.BackColor = &H80000005
    Label11.ForeColor = &H80000012
    Label12.ForeColor = &H80000012
    Label13.ForeColor = &H80000012
    Label14.ForeColor = &H80000012
    Label15.ForeColor = &H80000012
    Label3.Caption = ""
    Label4.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    Text5.text = ""

End Sub

Private Sub Command5_Click()

    CMDialog1.Flags = &H1& Or &H4&
    CMDialog1.Action = 3
    AColor = CMDialog1.Color
    RedValue = (AColor And &HFF&)
    GreenValue = (AColor And &HFF00&) \ 256
    BlueValue = (AColor And &HFF0000) \ 65536
    ChoosenColor = Format$(Hex$(RedValue) & Hex$(GreenValue) & Hex$(BlueValue), "000000")
    Label13.ForeColor = CMDialog1.Color
    Label8.Caption = ChoosenColor
    Text5.SelText = " LINK=#" & ChoosenColor
    Command9.Caption = "&Done"

End Sub

Private Sub Command6_Click()

    CMDialog1.Flags = &H1& Or &H4&
    CMDialog1.Action = 3
    AColor = CMDialog1.Color
    RedValue = (AColor And &HFF&)
    GreenValue = (AColor And &HFF00&) \ 256
    BlueValue = (AColor And &HFF0000) \ 65536
    ChoosenColor = Format$(Hex$(RedValue) & Hex$(GreenValue) & Hex$(BlueValue), "000000")
    Label14.ForeColor = CMDialog1.Color
    Label9.Caption = ChoosenColor
    Text5.SelText = " ALINK=#" & ChoosenColor
    Command9.Caption = "&Done"

End Sub

Private Sub Command7_Click()

    CMDialog1.Flags = &H1& Or &H4&
    CMDialog1.Action = 3
    AColor = CMDialog1.Color
    RedValue = (AColor And &HFF&)
    GreenValue = (AColor And &HFF00&) \ 256
    BlueValue = (AColor And &HFF0000) \ 65536
    ChoosenColor = Format$(Hex$(RedValue) & Hex$(GreenValue) & Hex$(BlueValue), "000000")
    Label15.ForeColor = CMDialog1.Color
    Label10.Caption = ChoosenColor
    Text5.SelText = " VLINK=#" & ChoosenColor
    Command9.Caption = "&Done"

End Sub

Private Sub Command8_Click()

    frmMain.ActiveForm.RTF1.Find ("<body")
    frmMain.ActiveForm.RTF1.SelRTF = "<body " & Text5.text
    Unload colors

End Sub

Private Sub Command9_Click()

    Unload Me

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:24 PM) 5 + 121 = 126 Lines
