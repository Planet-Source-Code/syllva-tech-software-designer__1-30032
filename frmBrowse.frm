VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBrowse 
   Caption         =   "Open URL"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   810
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfSource 
      Height          =   2295
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4048
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmBrowse.frx":0000
   End
   Begin VB.TextBox txtBrowse 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":00E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":0434
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":0788
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":0ADC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "back"
            Object.ToolTipText     =   "Go Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "forward"
            Object.ToolTipText     =   "Go Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "view"
            Object.ToolTipText     =   "View URL"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "send"
            Object.ToolTipText     =   "Send To Editor"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   725
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   4920
      Y1              =   370
      Y2              =   370
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   5760
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    txtBrowse.text = "http://www."

End Sub

Private Sub Form_Resize()

    wb.Move 0, 725, ScaleWidth, ScaleHeight - 725
    rtfSource.Move 0, 725, ScaleWidth, ScaleHeight - 725
    txtBrowse.Width = Me.Width
    Line1.X2 = Me.Width
    Line2.X2 = Me.Width

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
      Case "back"
        wb.GoBack
      Case "forward"
        wb.GoForward
      Case "view"
        Me.Caption = "Getting Source for " & wb.LocationURL
        rtfSource.text = Inet1.OpenURL(txtBrowse.text)
        rtfSource.Visible = True
        wb.Visible = False
        Me.Caption = "Open URL"
      Case "send"
        frmMain.ActiveForm.RTF1.text = rtfSource.text
    End Select

End Sub

Private Sub txtBrowse_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        wb.Navigate txtBrowse.text
        KeyAscii = 0
    End If

End Sub

Private Sub wb_StatusTextChange(ByVal text As String)

    Me.Caption = "Open URL - " & text

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:22 PM) 1 + 52 = 53 Lines
