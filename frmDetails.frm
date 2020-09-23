VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDetails 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   3413
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Names"
      TabPicture(0)   =   "frmDetails.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCompany"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "E-mail"
      TabPicture(1)   =   "frmDetails.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtBusiness"
      Tab(1).Control(1)=   "txtEmail"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Websites"
      TabPicture(2)   =   "frmDetails.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtWebSite"
      Tab(2).Control(1)=   "txtUrl"
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(3)=   "Label5"
      Tab(2).ControlCount=   4
      Begin VB.TextBox txtWebSite 
         Height          =   285
         Left            =   -74880
         TabIndex        =   15
         Top             =   1440
         Width           =   5655
      End
      Begin VB.TextBox txtUrl 
         Height          =   285
         Left            =   -74880
         TabIndex        =   13
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox txtBusiness 
         Height          =   285
         Left            =   -74880
         TabIndex        =   11
         Top             =   1440
         Width           =   3855
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -74880
         TabIndex        =   9
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtCompany 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   5655
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Your business (or your company) website URL:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   14
         Top             =   1200
         Width           =   3300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Your personal website URL:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "What is your business e-mail address?"
         Height          =   195
         Left            =   -74880
         TabIndex        =   10
         Top             =   1200
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "What is your personal e-mail address?"
         Height          =   195
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   2670
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Company name:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Your name:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   810
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Set Information"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   $"frmDetails.frx":0054
      ForeColor       =   &H8000000C&
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   5895
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()

    SaveSetting App.Title, "Details", "PersonalName", txtName
    SaveSetting App.Title, "Details", "BusinessName", txtCompany
    SaveSetting App.Title, "Details", "PersonalEmail", txtEmail
    SaveSetting App.Title, "Details", "BusinessEmail", txtBusiness
    SaveSetting App.Title, "Details", "PersonalURL", txtUrl
    SaveSetting App.Title, "Details", "BusinessURL", txtWebSite

End Sub

Private Sub Command1_Click()

    txtName.text = ""
    txtCompany.text = ""
    txtEmail.text = ""
    txtBusiness = ""
    txtUrl = ""
    txtWebSite = ""

End Sub

Private Sub Command3_Click()

    Unload Me

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:12 PM) 1 + 30 = 31 Lines
