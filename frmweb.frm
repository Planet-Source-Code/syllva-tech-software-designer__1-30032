VERSION 5.00
Begin VB.Form frmWeb 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Web Manager"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmweb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Caption         =   "Path:"
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3015
      Begin VB.DriveListBox Drv 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Tag             =   "The DriveList shows the drives on your computer."
         Top             =   360
         Width           =   2715
      End
      Begin VB.DirListBox Folder 
         Height          =   1890
         Left            =   120
         TabIndex        =   5
         Tag             =   "The FolderList displays all folders on your system."
         Top             =   675
         Width           =   2715
      End
   End
   Begin VB.TextBox txName 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   3090
   End
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1995
      TabIndex        =   1
      Tag             =   "Close this dialog without loading or creating the web."
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Continue"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Tag             =   "Continue to create or load the selected web."
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "N&ame of web:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1005
   End
End
Attribute VB_Name = "frmWeb"
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

Private Sub cmdNo_Click()

    ReturnedPath = ""
    Unload Me

End Sub

Private Sub cmdOK_Click()

    If txName.Visible = True And txName.text = "" Then MsgBox "The Name field is required.", vbExclamation: Exit Sub ':( Expand Structure
  Dim StrPS As String ':( Move line to top of current Sub
    If right$(Folder.Path, 1) <> "\" Then StrPS = "\" ':( Expand Structure
    If txName.Visible And Dir$(Folder.Path & StrPS & txName.text, vbDirectory) <> "" Then MsgBox "Directory already exists." & vbNewLine & "Use the open web feature.", vbExclamation: Exit Sub ':( Expand Structure
    ReturnedPath = Folder.Path & StrPS & txName.text
    Unload Me

End Sub

Private Sub Drv_Change()

    On Error GoTo hell
    Folder.Path = Drv.Drive

Exit Sub

hell:
    MsgBox Error, vbExclamation

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
      SetFont Me
      If KeyCode = vbKeyF6 Then
          MsgBox ActiveControl.tag, vbInformation
      End If

End Sub ':( On Error Resume still active

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:48 PM) 8 + 42 = 50 Lines
