VERSION 5.00
Object = "{C91AB44E-4FFD-4243-ABBC-B9199CE52090}#1.0#0"; "CPVPICSCROLL.OCX"
Begin VB.Form frmImageTwo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Designer - Image browser"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "BMP to JPG"
      Height          =   375
      Left            =   8700
      TabIndex        =   19
      ToolTipText     =   "Convert a bitmap to JPEG"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Move Picture to Images Folder"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2280
      TabIndex        =   18
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Best Fit"
      Height          =   375
      Left            =   8700
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "100%"
      Height          =   375
      Left            =   8700
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Zoom Out"
      Height          =   375
      Left            =   8700
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Zoom In"
      Height          =   375
      Left            =   8700
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Image map"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      ToolTipText     =   "Use Current Image As An Image Map"
      Top             =   4320
      Width           =   1095
   End
   Begin PicScroll.cpvPicScroll imgPic 
      Height          =   3900
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   6879
      BorderStyle     =   1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Do Not Use Full Path"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7590
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   375
      Left            =   8685
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   120
      Pattern         =   "*.jpg;*.gif;*.jpeg"
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   8700
      X2              =   9660
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   8700
      X2              =   9660
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   8760
      TabIndex        =   17
      Top             =   3795
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Picture Size:  "
      Height          =   195
      Left            =   8760
      TabIndex        =   16
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblZoomFactor 
      AutoSize        =   -1  'True
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   8760
      TabIndex        =   15
      Top             =   3195
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Zoom:  "
      Height          =   195
      Left            =   8760
      TabIndex        =   14
      Top             =   3000
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   45
   End
End
Attribute VB_Name = "frmImageTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub Command1_Click()

    Select Case Check1.value
      Case vbUnchecked
        frmMain.ActiveForm.RTF1.SelText = "<IMG SRC=" & Chr$(34) & Label1.Caption & Chr$(34) & ">"
      Case vbChecked
        'frmMain.ActiveForm.RTF1.SelText = "<IMG SRC=" & Chr$(34) & "\images\" & Text1.text & Chr$(34) & ">"
        frmMain.ActiveForm.RTF1.SelText = "<IMG SRC=" & Chr$(34) & Text1.text & Chr$(34) & ">"

    End Select

End Sub

Private Sub Command3_Click()

    imgPic.ZoomIn

End Sub

Private Sub Command4_Click()

    imgPic.ZoomOut

End Sub

Private Sub Command5_Click()

    imgPic.ZoomReal

End Sub

Private Sub Command6_Click()

    imgPic.BestFit

End Sub

Private Sub Dir1_Change()

    File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_DblClick()

    On Error GoTo oops
    Set imgPic.Picture = LoadPicture(File1.Path & "\" & File1.fileName)
    Command6_Click
    Me.Caption = "Designer - " & File1.Path & "\" & File1.fileName
    Label1.Caption = File1.Path & "\" & File1.fileName
    Text1.text = File1.fileName

Exit Sub

oops:
    MsgBox "Picture not found." & vbCrLf & "I'm sorry :-(", vbInformation, "Error"

Exit Sub

End Sub

Private Sub imgPic_PictureSizeChanged()

    lblZoomFactor = imgPic.ZoomPercent & " %"
    lblSize = imgPic.PictureSize.psWidth & " x " & imgPic.PictureSize.psHeight

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:02 PM) 1 + 82 = 83 Lines
