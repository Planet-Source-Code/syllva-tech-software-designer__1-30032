VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLists 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Lists"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlList 
      Left            =   4920
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Gif File (*.gif)|*.gif|JPEG File (*.jpg, *.jpeg|*.jpg;*.jpeg"
   End
   Begin VB.TextBox txtList 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   6735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   6588
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Bulleted List"
      TabPicture(0)   =   "frmLists.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblOpen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCustom"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Definition List"
      TabPicture(1)   =   "frmLists.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtTerm"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtDefinition"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command6"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command7"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Numbered List"
      TabPicture(2)   =   "frmLists.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblNum"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Text2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   -68730
         Picture         =   "frmLists.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   810
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   -69105
         Picture         =   "frmLists.frx":0156
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   810
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   -69480
         Picture         =   "frmLists.frx":0258
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   810
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Convert"
         Height          =   375
         Left            =   -71160
         TabIndex        =   28
         Top             =   3150
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Done"
         Height          =   375
         Left            =   -69720
         TabIndex        =   27
         Top             =   3150
         Width           =   1335
      End
      Begin VB.TextBox txtDefinition 
         Height          =   1215
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   26
         Top             =   1575
         Width           =   6495
      End
      Begin VB.TextBox txtTerm 
         Height          =   285
         Left            =   -74880
         TabIndex        =   24
         Top             =   855
         Width           =   5295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -74880
         TabIndex        =   13
         Top             =   3195
         Width           =   4935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Done"
         Height          =   375
         Left            =   -69720
         TabIndex        =   12
         Top             =   3150
         Width           =   1335
      End
      Begin VB.TextBox txtCustom 
         Height          =   285
         Left            =   480
         TabIndex        =   11
         Top             =   4560
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Done"
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   3150
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   3195
         Width           =   4935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Bullet Style"
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   6495
         Begin VB.OptionButton optCustom 
            Caption         =   "Use Custom Bullet"
            Height          =   195
            Left            =   2040
            TabIndex        =   10
            Top             =   1980
            Width           =   1695
         End
         Begin VB.OptionButton optDisc 
            Caption         =   "Disc Bullet"
            Height          =   195
            Left            =   2040
            TabIndex        =   5
            Top             =   1170
            Width           =   1335
         End
         Begin VB.OptionButton optCircle 
            Caption         =   "Circle Bullet"
            Height          =   195
            Left            =   2040
            TabIndex        =   4
            Top             =   765
            Width           =   1335
         End
         Begin VB.OptionButton optSquare 
            Caption         =   "Square Bullet"
            Height          =   195
            Left            =   2040
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Image imgBullet 
            BorderStyle     =   1  'Fixed Single
            Height          =   1560
            Left            =   360
            Top             =   480
            Width           =   1440
         End
         Begin VB.Shape Shape1 
            Height          =   1815
            Left            =   240
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bullet Style"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   6495
         Begin VB.OptionButton optNumbers 
            Caption         =   "1  (Numbers)"
            Height          =   255
            Left            =   2040
            TabIndex        =   21
            Top             =   1920
            Width           =   2535
         End
         Begin VB.OptionButton optCaps 
            Caption         =   "A  (Capital Letters)"
            Height          =   195
            Left            =   2040
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optLower 
            Caption         =   "a  (Lowercase Letters)"
            Height          =   195
            Left            =   2040
            TabIndex        =   17
            Top             =   750
            Width           =   2655
         End
         Begin VB.OptionButton optRomanCaps 
            Caption         =   "I  (Capital Roman Numerals)"
            Height          =   195
            Left            =   2040
            TabIndex        =   16
            Top             =   1140
            Width           =   2775
         End
         Begin VB.OptionButton optRomanLower 
            Caption         =   "i  (Lowercase Roman Numerals)"
            Height          =   195
            Left            =   2040
            TabIndex        =   15
            Top             =   1530
            Width           =   2775
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3600
            TabIndex        =   19
            Top             =   600
            Width           =   45
         End
         Begin VB.Image imgNum 
            BorderStyle     =   1  'Fixed Single
            Height          =   1560
            Left            =   360
            Top             =   480
            Width           =   1440
         End
         Begin VB.Shape Shape2 
            Height          =   1815
            Left            =   240
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Definition of the term"
         Height          =   255
         Left            =   -74880
         TabIndex        =   25
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Term to define"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -69840
         TabIndex        =   22
         Top             =   3000
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "List Item Data:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   20
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "List Item Data:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Label lblOpen 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   4440
         Visible         =   0   'False
         Width           =   105
      End
   End
End
Attribute VB_Name = "frmLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Select Case optCustom.value
      Case True
        frmMain.ActiveForm.RTF1.SelText = "<ul>" & vbCrLf & txtList.text & "</ul>"
      Case False
        frmMain.ActiveForm.RTF1.SelText = lblOpen.Caption & vbCrLf & txtList.text & "</ul>"
    End Select

End Sub

Private Sub Command2_Click()

    frmMain.ActiveForm.RTF1.SelText = lblNum.Caption & vbCrLf & txtList.text & "</ol>"

End Sub

Private Sub Command3_Click()

    frmMain.ActiveForm.RTF1.SelText = "<dl>" & vbCrLf & txtList.text & "      <dd>" & txtDefinition.text & "</dl>"

End Sub

Private Sub Command5_Click()

    txtTerm.SelText = "<b>" & txtTerm.SelText & "</b>"

End Sub

Private Sub Command6_Click()

    txtTerm.SelText = "<i>" & txtTerm.SelText & "</i>"

End Sub

Private Sub Command7_Click()

    txtTerm.SelText = "<u>" & txtTerm.SelText & "</u>"

End Sub

Private Sub Form_Load()

    imgBullet.Picture = LoadPicture(App.Path & "\Data\square.qcb")

    imgNum.Picture = LoadPicture(App.Path & "\Data\capA.qcb")
    imgBullet.Picture = LoadPicture(App.Path & "\Data\square.qcb")
    imgBullet.Picture = LoadPicture(App.Path & "\Data\square.qcb")

    lblOpen.Caption = "<ul type=" & Chr$(34) & "square" & Chr$(34) & ">"

End Sub

Private Sub optCaps_Click()

    lblNum.Caption = "<ol type=" & Chr$(34) & "A" & Chr$(34) & ">"
    imgNum.Picture = LoadPicture(App.Path & "\Data\capA.qcb")

End Sub

Private Sub optCustom_Click()

    cdlList.ShowOpen
    txtCustom.text = cdlList.FileTitle
    imgBullet.Picture = LoadPicture(cdlList.fileName)

End Sub

Private Sub optLower_Click()

    lblNum.Caption = "<ol type=" & Chr$(34) & "a" & Chr$(34) & ">"
    imgNum.Picture = LoadPicture(App.Path & "\Data\lowA.qcb")

End Sub

Private Sub optNumbers_Click()

    lblNum.Caption = "<ol type=" & Chr$(34) & "1" & Chr$(34) & ">"
    imgNum.Picture = LoadPicture(App.Path & "\Data\num.qcb")

End Sub

Private Sub optRomanCaps_Click()

    lblNum.Caption = "<ol type=" & Chr$(34) & "I" & Chr$(34) & ">"
    imgNum.Picture = LoadPicture(App.Path & "\Data\capRom.qcb")

End Sub

Private Sub optRomanLower_Click()

    lblNum.Caption = "<ol type=" & Chr$(34) & "i" & Chr$(34) & ">"
    imgNum.Picture = LoadPicture(App.Path & "\Data\lowRom.qcb")

End Sub

Private Sub optSquare_Click()

    lblOpen.Caption = "<ul type=" & Chr$(34) & "square" & Chr$(34) & ">"
    imgBullet.Picture = LoadPicture(App.Path & "\Data\square.qcb")

End Sub

Private Sub optCircle_Click()

    lblOpen.Caption = "<ul type=" & Chr$(34) & "circle" & Chr$(34) & ">"
    imgBullet.Picture = LoadPicture(App.Path & "\Data\circle.qcb")

End Sub

Private Sub optDisc_Click()

    lblOpen.Caption = "<ul type=" & Chr$(34) & "disc" & Chr$(34) & ">"
    imgBullet.Picture = LoadPicture(App.Path & "\Data\Disc.qcb")

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    txtList.text = ""

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    Select Case optCustom.value
      Case True
        If KeyAscii = vbKeyReturn Then
            txtList.SelText = "<img src=" & Chr$(34) & txtCustom.text & Chr$(34) & "> " & Text1.text & "<br>" & vbCrLf
            Text1.text = ""
            Text1.SetFocus
            KeyAscii = 0
        End If
      Case False
        If KeyAscii = vbKeyReturn Then
            txtList.SelText = "<li>" & Text1.text & vbCrLf
            Text1.text = ""
            Text1.SetFocus
            KeyAscii = 0
        End If
    End Select

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtList.SelText = "<li>" & Text2.text & vbCrLf
        Text2.text = ""
        Text2.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtDefinition_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtList.SelText = "<dd>" & txtDefinition.text & vbCrLf
        txtDefinition.text = ""
        txtTerm.text = ""
        txtTerm.SetFocus
        KeyAscii = 0
    End If

End Sub

Private Sub txtTerm_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtList.SelText = "<dt>" & txtTerm.text & vbCrLf
        txtDefinition.SetFocus
        KeyAscii = 0
    End If

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:01 PM) 1 + 179 = 180 Lines
