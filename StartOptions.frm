VERSION 5.00
Begin VB.Form StartOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FTP Information"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "URL of your account:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "StartOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    SaveSetting App.Title, "FTP", "URL", Text1.text
    SaveSetting App.Title, "FTP", "Username", Text2.text
    SaveSetting App.Title, "FTP", "Password", Text3.text

End Sub

Private Sub Command3_Click()

    Text1.text = ""
    Text2.text = ""
    Text3.text = ""

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:46 PM) 1 + 24 = 25 Lines
