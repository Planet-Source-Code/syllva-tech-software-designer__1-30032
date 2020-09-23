VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmCustomFooter 
   Caption         =   "Create a Customized Footer"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   840
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Custom Footers (*.qft)|*.qft"
   End
   Begin SHDocVwCtl.WebBrowser wbFooter 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   2265
      Width           =   6495
      ExtentX         =   11456
      ExtentY         =   3201
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
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   8745
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Child1          =   "cboFonts"
      MinHeight1      =   315
      Width1          =   2370
      NewRow1         =   0   'False
      Child2          =   "Toolbar1"
      MinHeight2      =   330
      Width2          =   1350
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   2565
         TabIndex        =   4
         Top             =   30
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   20
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "new"
               Object.ToolTipText     =   "New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "open"
               Object.ToolTipText     =   "Open"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "save"
               Object.ToolTipText     =   "Save As Template"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "delete"
               Object.ToolTipText     =   "Delete"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cut"
               Object.ToolTipText     =   "Cut"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "copy"
               Object.ToolTipText     =   "Copy"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "paste"
               Object.ToolTipText     =   "Paste"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bold"
               Object.ToolTipText     =   "Bold"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "italic"
               Object.ToolTipText     =   "Italic"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "underline"
               Object.ToolTipText     =   "Underline"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "strike"
               Object.ToolTipText     =   "Strike Thru"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "big"
               Object.ToolTipText     =   "Make Font Bigger"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "small"
               Object.ToolTipText     =   "Make Font Smaller"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "left"
               Object.ToolTipText     =   "Align Left"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "center"
               Object.ToolTipText     =   "Center"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "right"
               Object.ToolTipText     =   "Align Right"
               ImageIndex      =   16
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboFonts 
         Height          =   315
         ItemData        =   "frmCustomFooter.frx":0000
         Left            =   165
         List            =   "frmCustomFooter.frx":0010
         TabIndex        =   3
         Top             =   30
         Width           =   2175
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":014C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":0260
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":0374
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":0488
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":059C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":06B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":07C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":08D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":0B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":0C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":0F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":12BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":13D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":14E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomFooter.frx":15F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfFooter 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3201
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmCustomFooter.frx":194C
   End
End
Attribute VB_Name = "frmCustomFooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    frmMain.ActiveForm.RTF1.Find ("</body>")
    frmMain.ActiveForm.RTF1.SelRTF = "<hr>" & rtfFooter.text & "</body>"

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    wbFooter.Navigate ("about:blank")

End Sub

Private Sub Form_Resize()

    rtfFooter.Move 0, 450, ScaleWidth
    wbFooter.Width = Me.ScaleWidth

End Sub

Private Sub rtfFooter_Change()

    Render

End Sub

Private Sub rtfFooter_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        rtfFooter.SelText = "<br>" & vbCrLf
        KeyAscii = 0
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key

      Case "new"
        rtfFooter.text = ""
        Render

      Case "open"
        cdl.ShowOpen
        rtfFooter.LoadFile cdl.fileName
        Render

      Case "save"

      Case "delete"
        rtfFooter.SelText = ""
        Render

      Case "cut"
        Clipboard.SetText rtfFooter.SelText
        rtfFooter.SelText = ""
        Render

      Case "copy"
        Clipboard.SetText rtfFooter.SelText

      Case "paste"
        rtfFooter.SelText = Clipboard.GetText
        Render

      Case "bold"
        rtfFooter.SelText = "<b>" & rtfFooter.SelText & "</b>"
        Render

      Case "italic"
        rtfFooter.SelText = "<i>" & rtfFooter.SelText & "</i>"
        Render

      Case "underline"
        rtfFooter.SelText = "<u>" & rtfFooter.SelText & "</u>"
        Render

      Case "strike"
        rtfFooter.SelText = "<strike>" & rtfFooter.SelText & "</strike>"
        Render

      Case "big"
        rtfFooter.SelText = "<big>" & rtfFooter.SelText & "</big>"
        Render

      Case "small"
        rtfFooter.SelText = "<small>" & rtfFooter.SelText & "</small>"
        Render

      Case "left"
        rtfFooter.SelText = "<div align=" & Chr$(34) & "left" & Chr$(34) & ">" & rtfFooter.SelText & "</div>"
        Render

      Case "center"
        rtfFooter.SelText = "<center>" & rtfFooter.SelText & "</center>"
        Render

      Case "right"
        rtfFooter.SelText = "<div align=" & Chr$(34) & "right" & Chr$(34) & ">" & rtfFooter.SelText & "</div>"
        Render

    End Select

End Sub

Private Sub Render()

    wbFooter.Document.Script.Document.Clear
    wbFooter.Document.Script.Document.Write rtfFooter.text
    wbFooter.Document.Script.Document.Close

    Exit Sub

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:53:13 PM) 1 + 124 = 125 Lines
