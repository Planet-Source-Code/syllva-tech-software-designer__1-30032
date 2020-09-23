VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   2505
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdNextTip 
      BackColor       =   &H80000004&
      Caption         =   "N&ext"
      Height          =   375
      Left            =   2565
      MousePointer    =   99  'Custom
      Picture         =   "frmTip.frx":000C
      TabIndex        =   5
      Top             =   2025
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   135
      ScaleHeight     =   1755
      ScaleWidth      =   3540
      TabIndex        =   2
      Top             =   90
      Width           =   3600
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmTip.frx":0156
         Top             =   127
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Did you know?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   960
         Left            =   90
         TabIndex        =   3
         Top             =   675
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000004&
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3375
      MousePointer    =   99  'Custom
      Picture         =   "frmTip.frx":0898
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2025
      Width           =   375
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   225
      Left            =   135
      TabIndex        =   0
      Top             =   2130
      Width           =   2055
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection
' Name of tips file
Const TIP_FILE = "whtml.tip"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long

Private Sub DoNextTip()

  ' Select a tip at random.

    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

    '  CurrentTip = CurrentTip + 1
    '  If Tips.Count < CurrentTip Then
    '       CurrentTip = 1
    '  End If
    
    ' Show it.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean

  Dim NextTip As String   ' Each tip read in from file.
  Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function '>---> Bottom
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir$(sFile) = "" Then
        LoadTips = False
        Exit Function '>---> Bottom
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()

  ' save whether or not this form should be displayed at startup

    SaveValue "StartupTips", chkLoadTipsAtStartup.Value
    SendMessage chkLoadTipsAtStartup.hWnd, WM_KILLFOCUS, 0, 0&

End Sub

Private Sub cmdNextTip_Click()

    DoNextTip

End Sub

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    On Error Resume Next
      SetFont Me
      ' Set the checkbox, this will force the value to be written back out to the registry
      Me.chkLoadTipsAtStartup.Value = vbChecked
    
      ' Seed Rnd
      Randomize
    
      ' Read in the tips file and display a tip at random.
      If LoadTips(App.Path & "\" & TIP_FILE) = False Then
          MsgBox "Can't find the tips file.", vbExclamation
          Unload Me
      End If

End Sub ':( On Error Resume still active

Public Sub DisplayCurrentTip()

    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:50 PM) 9 + 102 = 111 Lines
