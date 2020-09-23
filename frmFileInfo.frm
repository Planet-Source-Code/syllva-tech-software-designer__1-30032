VERSION 5.00
Begin VB.Form frmFile 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Information"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer TMR 
      Interval        =   100
      Left            =   450
      Top             =   1260
   End
   Begin VB.CommandButton cmOK 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1035
      UseMaskColor    =   -1  'True
      Width           =   870
   End
   Begin VB.TextBox txSize 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   735
      Width           =   3165
   End
   Begin VB.TextBox txDT 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   495
      Width           =   3165
   End
   Begin VB.TextBox txName 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   4200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "File &size (KB):"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   735
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Date / &Time:"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   495
      Width           =   885
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmOK_Click()

    Unload Me

End Sub

Private Sub TMR_Timer()

    On Error Resume Next
      'we needed to wait a while for the form to load, then
      'do the stuff with the FileDateTime and so on.
      txName.text = tag
      txSize.text = Round(FileLen(tag) / 1024, 2) & "(" & FileLen(tag) & " Bytes)"
      txDT.text = FileDateTime(tag)

End Sub ':( On Error Resume still active

Private Sub txDT_Click()

    txDT.SelStart = Len(txDT.text)

End Sub

Private Sub txName_Click()

    txName.SelStart = Len(txName.text)

End Sub

Private Sub txSize_Click()

    txSize.SelStart = Len(txSize.text)

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:10 PM) 1 + 37 = 38 Lines
