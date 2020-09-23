VERSION 5.00
Begin VB.UserControl TabCtl 
   Alignable       =   -1  'True
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3840
   ScaleWidth      =   4800
   Begin VB.Image prev 
      Height          =   255
      Left            =   855
      Picture         =   "TabCtl.ctx":0000
      Top             =   2700
      Width           =   885
   End
   Begin VB.Image prev_on 
      Height          =   255
      Left            =   855
      Picture         =   "TabCtl.ctx":03CC
      Top             =   2700
      Width           =   885
   End
   Begin VB.Image source_on 
      Height          =   255
      Left            =   0
      Picture         =   "TabCtl.ctx":0779
      Top             =   2700
      Width           =   885
   End
   Begin VB.Image source 
      Height          =   255
      Left            =   0
      Picture         =   "TabCtl.ctx":0B22
      Top             =   2700
      Width           =   885
   End
End
Attribute VB_Name = "TabCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public TabIndex As Long
Public Event TabChanged(ByVal NewTabIndex As Long)

Private Sub prev_Click()

    prev_on.ZOrder vbBringToFront
    source.ZOrder vbBringToFront
    TabIndex = 2
    RaiseEvent TabChanged(2)

End Sub

Private Sub source_Click()

    source_on.ZOrder vbBringToFront
    prev.ZOrder vbBringToFront
    TabIndex = 1
    RaiseEvent TabChanged(1)

End Sub

Private Sub UserControl_Initialize()

    source_on.ZOrder vbBringToFront
    prev.ZOrder vbBringToFront

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
      source.Move 0, 0
      source_on.Move 0, 0
      prev.Move source.Width, 0
      prev_on.Move source.Width, 0

End Sub ':( On Error Resume still active

Sub CycleTabs()

    If TabIndex = 1 Then
        TabIndex = 2
        prev_on.ZOrder vbBringToFront
        source.ZOrder vbBringToFront
      Else
        source_on.ZOrder vbBringToFront
        prev.ZOrder vbBringToFront
        TabIndex = 1
    End If
    RaiseEvent TabChanged(TabIndex)

End Sub

':) Ulli's VB Code Formatter V2.3.16 (10/30/2001 2:52:31 PM) 3 + 52 = 55 Lines
