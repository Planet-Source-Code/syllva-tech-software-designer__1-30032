VERSION 5.00
Begin VB.Form frmImage 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmimage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   324
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox pD 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3555
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   2970
      Width           =   240
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   240
      LargeChange     =   15
      Left            =   2205
      SmallChange     =   3
      TabIndex        =   4
      Top             =   4440
      Width           =   2265
   End
   Begin VB.VScrollBar VS 
      Height          =   690
      LargeChange     =   15
      Left            =   4950
      SmallChange     =   3
      TabIndex        =   3
      Top             =   1980
      Width           =   240
   End
   Begin VB.ComboBox cbZoom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   1950
   End
   Begin VB.CommandButton cmCopy 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Copy to clipboard"
      Height          =   315
      Left            =   1935
      TabIndex        =   0
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label lbC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3600
      TabIndex        =   1
      Top             =   68
      Width           =   1125
   End
   Begin VB.Image pB 
      Appearance      =   0  'Flat
      Height          =   1500
      Left            =   90
      Top             =   495
      Width           =   2445
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'######################################
'Designer 1.2 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
Option Explicit
Dim i As Integer

Private Sub cbZoom_Click()

    imlZoom_Click

End Sub

Private Sub cbZoom_GotFocus()

    Form_GotFocus

End Sub

Private Sub cmCopy_Click()

    Clipboard.Clear
    Clipboard.SetData pB.Picture, vbCFBitmap
    MsgBox GetFile(tag) & " was copied to the clipboard.", vbInformation

End Sub

Private Sub cmCopy_GotFocus()

    Form_GotFocus

End Sub

Private Sub Form_GotFocus()

    On Error Resume Next
      If frmMain.tvW.Nodes(tag).Bold = True Then Exit Sub ':( Expand Structure
      For i = 1 To frmMain.tvW.Nodes.Count
          frmMain.tvW.Nodes(i).Bold = False
          If frmMain.tvW.Nodes(i).key = tag Then frmMain.tvW.Nodes(i).Bold = True ':( Expand Structure
      Next i

End Sub ':( On Error Resume still active

Private Sub Form_Load()

    On Error Resume Next
      With cbZoom
          .AddItem "Zoom: 25% (Shrink)"
          .AddItem "Zoom: 50% (Shrink)"
          .AddItem "Zoom: 75% (Shrink)"
          .AddItem "Zoom: 100% (Normal)"
          .AddItem "Zoom: 200% (Stretch)"
          .AddItem "Zoom: 300% (Stretch)"
          .AddItem "Zoom: 400% (Stretch)"
          .ListIndex = 3
      End With 'CBZOOM
      cmCopy.SetFocus
      Form_Resize

End Sub ':( On Error Resume still active

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
      lbC.Caption = GetFile(tag) & " (" & Round(FileLen(tag) / 1024, 2) & " KB)"

End Sub ':( On Error Resume still active

Sub Form_Resize()

    pB.Move 0, cbZoom.Height
    HS.Move 0, ScaleHeight - HS.Height, ScaleWidth - 16, 16
    VS.Move ScaleWidth - VS.Width, 0, 16, ScaleHeight - 16
    HS.Enabled = (pB.Width > ScaleWidth)
    VS.Enabled = (pB.Height > ScaleHeight - 21)
    HS.Max = pB.Width - ScaleWidth
    VS.Max = pB.Height - ScaleHeight
    pD.Move ScaleWidth - 16, ScaleHeight - 16

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
      For i = 1 To frmMain.tvW.Nodes.Count
          frmMain.tvW.Nodes(i).Bold = False
      Next i
      frmMain.ActiveForm.SetFocus
      frmMain.ActiveForm.RTF1.SetFocus

End Sub ':( On Error Resume still active

Private Sub pB_GotFocus()

    For i = 1 To frmMain.tvW.Nodes.Count
        frmMain.tvW.Nodes(i).Bold = False
        If frmMain.tvW.Nodes(i).key = Caption Then frmMain.tvW.Nodes(i).Bold = True ':( Expand Structure
    Next i
    'DisableBar

End Sub

Private Sub imlZoom_Click()

    On Error Resume Next
    Dim i As Long ':( Duplicated Name':( Move line to top of current Sub
      i = ParseInt(cbZoom.text)
      If i = 100 Then pB.Stretch = False: Exit Sub ':( Expand Structure
      pB.Stretch = False
      pB.Stretch = True
      pB.Width = pB.Width * (i / 100)
      pB.Height = pB.Height * (i / 100)
      cmCopy.SetFocus

End Sub ':( On Error Resume still active

Private Sub HS_Change()

    On Error Resume Next
      pB.left = -HS.value
      cmCopy.SetFocus

End Sub ':( On Error Resume still active

Private Sub pB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbC.Caption = X / Screen.TwipsPerPixelX & ", " & Y / Screen.TwipsPerPixelY

End Sub

Private Sub VS_Change()

    On Error Resume Next
      If VS.value = 0 Then pB.top = 22: GoTo n ':( Expand Structure
      pB.top = -VS.value
n:
      cmCopy.SetFocus

End Sub ':( On Error Resume still active

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:03 PM) 9 + 136 = 145 Lines
