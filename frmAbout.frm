VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3720
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6780
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":014A
      Height          =   855
      Left            =   1320
      TabIndex        =   6
      Top             =   2040
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   1200
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 1998 - 2001. All Rights Reserved"
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3330
      Width           =   3225
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   375
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Queen City Software  "
      BeginProperty Font 
         Name            =   "Zephyr"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1110
      Left            =   1320
      TabIndex        =   1
      Top             =   -120
      Width           =   4125
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   3120
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuCredits 
         Caption         =   "Credits"
      End
      Begin VB.Menu mnuAuthors 
         Caption         =   "Authors"
      End
      Begin VB.Menu mnuPopupBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "WebSite"
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Label6.Caption = App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Form_Resize()

    Shape1.Height = Me.Height
    Shape2.Width = Me.Width + 120

End Sub

Private Sub Image1_Click()

  'Put in popup menu stuff here

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:23 PM) 1 + 27 = 28 Lines
