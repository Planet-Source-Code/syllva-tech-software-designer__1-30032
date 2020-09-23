VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ActiveXobj 
   Caption         =   "Designer - Insert ActiveX Object"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlObject 
      Left            =   10920
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Java Applet (*.class)|*.class"
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtCodeBase 
         Height          =   285
         Left            =   960
         TabIndex        =   47
         Text            =   "http://"
         Top             =   720
         Width           =   5415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "ActiveXobj.frx":0000
         Left            =   4320
         List            =   "ActiveXobj.frx":001F
         TabIndex        =   45
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2520
         TabIndex        =   43
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Insert"
         Height          =   375
         Left            =   3840
         TabIndex        =   42
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5160
         TabIndex        =   41
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtJavaParam 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   2520
         Width           =   6255
      End
      Begin VB.Frame Frame7 
         Caption         =   "Parameters"
         Height          =   1335
         Left            =   2160
         TabIndex        =   35
         Top             =   1080
         Width           =   2055
         Begin VB.TextBox txtnameParam 
            Height          =   285
            Left            =   720
            TabIndex        =   37
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtValueParam 
            Height          =   285
            Left            =   720
            TabIndex        =   36
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "Value:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   840
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Size"
         Height          =   1335
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1815
         Begin VB.TextBox txtWidth 
            Height          =   285
            Left            =   840
            TabIndex        =   32
            Text            =   "120"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtHeight 
            Height          =   285
            Left            =   840
            TabIndex        =   31
            Text            =   "120"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Width"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Height"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.TextBox txtAppletName 
         Height          =   285
         Left            =   3795
         TabIndex        =   29
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   2520
         Picture         =   "ActiveXobj.frx":006E
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   630
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Codebase:  "
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   765
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Alignment: "
         Height          =   195
         Left            =   4320
         TabIndex        =   44
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Name:  "
         Height          =   195
         Left            =   3240
         TabIndex        =   28
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Code:  "
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   285
         Width           =   510
      End
   End
   Begin VB.Frame Frame4 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5160
         TabIndex        =   21
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   2280
         Width           =   6255
      End
      Begin VB.Frame Frame1 
         Caption         =   "Size"
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1815
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   840
            TabIndex        =   16
            Text            =   "120"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   840
            TabIndex        =   15
            Text            =   "120"
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Height"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Width"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Spacing"
         Height          =   1335
         Left            =   2100
         TabIndex        =   9
         Top             =   720
         Width           =   2055
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1080
            TabIndex        =   11
            Text            =   "100"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1080
            TabIndex        =   10
            Text            =   "100"
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Horizontal"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Vertical"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   735
         End
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Text            =   "CLSID:"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Parameters"
         Height          =   1335
         Left            =   4320
         TabIndex        =   1
         Top             =   720
         Width           =   2055
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   720
            TabIndex        =   3
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   720
            TabIndex        =   2
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Value:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Class ID"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   285
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   195
         Left            =   3480
         TabIndex        =   22
         Top             =   285
         Width           =   165
      End
   End
End
Attribute VB_Name = "ActiveXobj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    frmMain.ActiveForm.RTF1.SelText = "<OBJECT CLASSID= """ & Text4.text & """ ID=""" & Text5.text & """ HEIGHT=""" & Text2.text & """ WIDTH=""" & Text3.text & """>" & vbCrLf & Text1.text & "</OBJECT>"

End Sub

Private Sub Command2_Click()

    cdlObject.ShowOpen
    txtCode.text = cdlObject.FileTitle
    txtAppletName.SetFocus

End Sub

Private Sub Command3_Click()

    Text1.text = ""
    Text4.text = "CLSID:"
    Text5.text = ""
    Text2.text = "120"
    Text3.text = "120"
    Text6.text = "100"
    Text7.text = "100"

End Sub

Private Sub Command4_Click()

    Unload Me

End Sub

Private Sub Command5_Click()

    Unload Me

End Sub

Private Sub Command6_Click()

    frmMain.ActiveForm.RTF1.SelText = "<APPLET NAME=" & Chr$(34) & txtAppletName.text & Chr$(34) & " CODE=" & Chr$(34) & txtCode.text & Chr$(34) & " CODEBASE=" & Chr$(34) & txtCodeBase.text & Chr$(34) & " HEIGHT=" & txtHeight.text & " WIDTH=" & Chr$(34) & txtWidth.text & Chr$(34) & ">" & vbCrLf & txtJavaParam.text & vbCrLf & "</APPLET>"
    Unload Me

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Text1.SelText = "     <PARAM NAME= """ & Text8.text & """"
        Text9.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Text1.SelText = " VALUE=""" & Text9.text & """>" & vbCrLf
        Text8.text = ""
        Text9.text = ""
        Text8.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtnameParam_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtJavaParam.SelText = "     <PARAM NAME=" & Chr$(34) & txtnameParam.text & Chr$(34) & " "
        txtValueParam.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtValueParam_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtJavaParam.SelText = " VALUE=" & Chr$(34) & txtValueParam.text & Chr$(34) & ">" & vbCrLf
        txtnameParam.text = ""
        txtValueParam.text = ""
        txtnameParam.SetFocus
        KeyAscii = 0
    End If

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:26 PM) 1 + 91 = 92 Lines
