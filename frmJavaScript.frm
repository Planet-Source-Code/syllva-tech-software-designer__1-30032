VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmJavaScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-Z JavaScript"
   ClientHeight    =   6615
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdl 
      Left            =   1440
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Create Slide Show"
      TabPicture(0)   =   "frmJavaScript.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtSlideFile"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtArray"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Check1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "List1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Create Rollover"
      TabPicture(1)   =   "frmJavaScript.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Another JavaScript Wizard Tab"
      TabPicture(2)   =   "frmJavaScript.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command3 
         Caption         =   "Done"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   11
         Top             =   5880
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   3120
         TabIndex        =   10
         Top             =   2340
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use File name only"
         Height          =   195
         Left            =   5640
         TabIndex        =   9
         Top             =   1890
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin RichTextLib.RichTextBox txtArray 
         Height          =   1095
         Left            =   -3960
         TabIndex        =   8
         Top             =   5880
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmJavaScript.frx":0054
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Keep This Picture"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Preview"
         Height          =   2655
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   2775
         Begin VB.Image imgPreview 
            Height          =   2295
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1815
         TabIndex        =   4
         Text            =   "MyPics"
         Top             =   675
         Width           =   2655
      End
      Begin VB.TextBox txtSlideFile 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1845
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
         Height          =   195
         Left            =   2400
         TabIndex        =   13
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Select picture you want to be the first one in the slide show"
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   5040
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Find pictures needed for the slide show (You can add as many as you need):"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   5400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name of Slide Show:  "
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmJavaScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Select Case Check1.value
Case vbUnchecked
cdl.ShowOpen
txtSlideFile.text = cdl.fileName
imgPreview.Picture = LoadPicture(cdl.fileName)
Case vbChecked
cdl.ShowOpen
txtSlideFile.text = cdl.FileTitle
imgPreview.Picture = LoadPicture(cdl.fileName)
End Select
End Sub

Private Sub Command2_Click()
If txtArray.text <> "" Then
txtArray.SelText = ", " & Chr$(34) & txtSlideFile.text & Chr$(34)
List1.AddItem cdl.FileTitle
Exit Sub
Else
txtArray.SelText = Chr$(34) & txtSlideFile.text & Chr$(34)
List1.AddItem cdl.FileTitle
Exit Sub
End If
End Sub

Private Sub Command3_Click()
frmMain.ActiveForm.RTF1.SelText = "<p><img src=" & Chr$(34) & List1.text & Chr$(34) & " name=" & Chr$(34) & "picSlides" & Chr$(34) & ">" & vbCrLf & _
"<p><a href=" & Chr$(34) & "javascript:prevPicture()" & Chr$(34) & ">" & vbCrLf & _
"&lt;&lt; Previous</a>&nbsp;&nbsp;<a href=" & Chr$(34) & "javascript:nextPicture()" & Chr$(34) & ">" & vbCrLf & _
"Next Picture &gt;&gt;</a>"
PutInHead
End Sub

Private Sub PutInHead()
frmMain.ActiveForm.RTF1.Find ("</head>")
frmMain.ActiveForm.RTF1.SelRTF = "<script language=javascript type=" & Chr$(34) & "text/javascript" & Chr$(34) & ">" & vbCrLf & _
"<!-- Beginning of generated slide show code" & vbCrLf & _
Text1.text & " = new Array(" & txtArray.text & ")" & vbCrLf & _
"thisPic = 0" & vbCrLf & _
"imgCt = " & Text1.text & ".length - 1" & vbCrLf & _
"function prevPicture() {" & vbCrLf & _
"if (document.images && thisPic > 0) {" & vbCrLf & _
"     thisPic--" & vbCrLf & _
"     document.picSlides.src = " & Text1.text & "[thisPic]" & vbCrLf & _
"   }" & vbCrLf & _
"}" & vbCrLf & _
"function nextPicture() {" & vbCrLf & _
"if (document.images && thisPic < imgCt) {" & vbCrLf & _
"     thisPic++" & vbCrLf & _
"     document.picSlides.src = " & Text1.text & "[thisPic]" & vbCrLf & _
"   }" & vbCrLf & _
"}" & vbCrLf & _
"//End of generated slide show code -->" & vbCrLf & _
"</script>" & vbCrLf & "</head>"
End Sub

Private Sub List1_Click()
Command3.Enabled = True
End Sub
