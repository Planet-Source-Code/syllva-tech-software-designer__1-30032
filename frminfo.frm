VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frminfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ListView LV 
      Height          =   1905
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   3360
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlTV"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   467
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTV 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminfo.frx":000C
            Key             =   ""
            Object.Tag             =   "Indicates an HTML or ASP document."
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminfo.frx":0E60
            Key             =   ""
            Object.Tag             =   "Indicates other files such as ZIP archives."
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminfo.frx":1CB4
            Key             =   ""
            Object.Tag             =   "Indicates a closed file folder."
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminfo.frx":2250
            Key             =   ""
            Object.Tag             =   "Indicates an open file folder."
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminfo.frx":27EC
            Key             =   ""
            Object.Tag             =   "Indicates a GIF, JPEG or BMP Image."
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminfo.frx":2D88
            Key             =   ""
            Object.Tag             =   "Indicates the currently open web."
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminfo.frx":4754
            Key             =   ""
            Object.Tag             =   "Specifies an HTML code category."
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminfo.frx":6120
            Key             =   ""
            Object.Tag             =   "Specifies an HTML code element (tag)."
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminfo.frx":7AEC
            Key             =   ""
            Object.Tag             =   "Indicates a Javascript function ID."
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminfo.frx":8088
            Key             =   ""
            Object.Tag             =   "Indicates a Javascript variable ID."
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'######################################
'WonderHTML 1.2 Deluxe Edition: 2001 BETA release
'(C) Sushant S. Pandurangi, [sushant@phreaker.net]
'######################################
'For more software, visit http://sushantshome.tripod.com
'######################################
'Thanks to Andrea Batina for MRU code
Option Explicit

Private Sub Form_Load()
Dim i As Integer
SetFont Me
LV.ColumnHeaders(2).Width = LV.Width - LV.ColumnHeaders(1).Width - 375 'approx
For i = 1 To imlTV.ListImages.Count
LV.ListItems.Add i, "Item" & i, "", , i
LV.ListItems("Item" & i).ListSubItems.Add 1, , imlTV.ListImages(i).Tag
Next i
End Sub

