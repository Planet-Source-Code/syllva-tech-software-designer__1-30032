VERSION 5.00
Begin VB.Form frmSymbols 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Symbol"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmsymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmOK 
      Cancel          =   -1  'True
      Caption         =   "&Insert"
      Height          =   375
      Left            =   855
      TabIndex        =   3
      Top             =   1305
      Width           =   870
   End
   Begin VB.CommandButton cmNo 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   855
      TabIndex        =   2
      Top             =   1710
      Width           =   870
   End
   Begin VB.ListBox ListSym 
      Height          =   2010
      Left            =   90
      TabIndex        =   1
      Top             =   112
      Width           =   645
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASCII 32"
      Height          =   195
      Left            =   855
      TabIndex        =   4
      Top             =   1035
      Width           =   675
   End
   Begin VB.Label lbSym 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1215
      TabIndex        =   0
      Top             =   405
      UseMnemonic     =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmSymbols"
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

Private Sub cmNo_Click()

    Unload Me

End Sub

Private Sub cmOK_Click()

    frmMain.ActiveForm.RTF1.SelText = "&#" & Asc(lbSym.Caption) & ";"
    Unload Me

End Sub

Private Sub Form_Load()

    SetFont Me
  Dim i As Integer ':( Move line to top of current Sub
    For i = 32 To 255
        ListSym.AddItem Chr$(i)
    Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
      frmMain.ActiveForm.RTF1.SetFocus

End Sub ':( On Error Resume still active

Private Sub ListSym_Click()

    lbSym.Caption = ListSym.text
    lblAsc.Caption = "ASCII " & Asc(ListSym.text)

End Sub

Private Sub ListSym_DblClick()

    cmOK_Click

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:52 PM) 8 + 45 = 53 Lines
