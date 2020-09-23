VERSION 5.00
Begin VB.Form frmFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace..."
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmstart.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmRepAll 
      Caption         =   "Replace &all"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   600
      Width           =   1005
   End
   Begin VB.CommandButton cmRepThis 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   1005
   End
   Begin VB.CommandButton cmFind 
      Caption         =   "&Search..."
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1147
      Width           =   1005
   End
   Begin VB.CheckBox chNoH 
      Caption         =   "Do not Hi&ghlight"
      Height          =   240
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   2025
   End
   Begin VB.CheckBox chWhole 
      Caption         =   "Wh&ole word"
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1185
   End
   Begin VB.CheckBox chCase 
      Caption         =   "Mat&ch case"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1185
   End
   Begin VB.TextBox txR 
      Height          =   315
      Left            =   802
      TabIndex        =   1
      Top             =   427
      Width           =   2985
   End
   Begin VB.TextBox txF 
      Height          =   315
      Left            =   802
      TabIndex        =   0
      Top             =   67
      Width           =   2985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "R&eplace:"
      Height          =   195
      Index           =   1
      Left            =   172
      TabIndex        =   3
      Top             =   487
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Search:"
      Height          =   195
      Index           =   0
      Left            =   172
      TabIndex        =   2
      Top             =   127
      Width           =   555
   End
End
Attribute VB_Name = "frmFind"
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
Dim vCase As Long, vWord As Long, vSelect As Long
Dim lpStart As Long

Private Sub chCase_Click()

    SendMessage chCase.hWnd, WM_KILLFOCUS, 0, 0&
    If chCase.value = 1 Then
        vCase = rtfMatchCase
      Else
        vCase = 0
    End If

End Sub

Private Sub chNoH_Click()

    SendMessage chNoH.hWnd, WM_KILLFOCUS, 0, 0&
    If chNoH.value = 1 Then
        vSelect = rtfNoHighlight
      Else
        vSelect = 0
    End If

End Sub

Private Sub chWhole_Click()

    SendMessage chWhole.hWnd, WM_KILLFOCUS, 0, 0&
    If chWhole.value = 1 Then
        vWord = rtfWholeWord
      Else
        vWord = 0
    End If

End Sub

Private Sub cmFind_Click()

    SendMessage cmFind.hWnd, WM_KILLFOCUS, 0, 0&
  Dim Fin As Long ':( Move line to top of current Sub
    If txF.text = "" Then Exit Sub ':( Expand Structure
    Fin = frmMain.ActiveForm.RTF1.Find(txF.text, lpStart, , vCase + vSelect + vWord)
    If Fin > 0 Then
        lpStart = Fin + 1
      Else
        MsgBox "'" & txF.text & "' cannot be found." & vbNewLine & "0 matches in document.", vbExclamation
    End If

End Sub

Private Sub cmRepAll_Click()

    SendMessage cmRepAll.hWnd, WM_KILLFOCUS, 0, 0&
    If txF.text = "" Or txR.text = "" Then Exit Sub ':( Expand Structure
    frmMain.ActiveForm.RTF1.text = Replace$(frmMain.ActiveForm.RTF1.text, txF.text, txR.text)

End Sub

Private Sub cmRepThis_Click()

    SendMessage cmRepThis.hWnd, WM_KILLFOCUS, 0, 0&
    If txR.text = "" Or txF.text = "" Then Exit Sub ':( Expand Structure
    If frmMain.ActiveForm.RTF1.SelLength = 0 Then Exit Sub ':( Expand Structure
    frmMain.ActiveForm.RTF1.SelText = txR.text
    frmMain.ActiveForm.RTF1.Find txF.text, lpStart, , vCase + vWord

End Sub

Private Sub Form_Load()

    On Error Resume Next
      SetFont Me
      lpStart = 1

End Sub ':( On Error Resume still active

Private Sub txF_Change()

    lpStart = 1

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:09 PM) 10 + 81 = 91 Lines
