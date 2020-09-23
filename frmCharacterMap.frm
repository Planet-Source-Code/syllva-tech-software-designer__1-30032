VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCharacterMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Designer - Character Map"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   720
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Txtcopy 
      Height          =   375
      HideSelection   =   0   'False
      Left            =   4800
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox CboFonts 
      Height          =   315
      Left            =   720
      Sorted          =   -1  'True
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   840
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Txt1 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txt2 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txt3 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   2145
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCharacterMap.frx":0000
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   135
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   10
      Top             =   720
      Width           =   6255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1920
      TabIndex        =   20
      Top             =   3165
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   1920
      TabIndex        =   19
      Top             =   3525
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Hex:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Bin:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Dec:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1920
      TabIndex        =   15
      Top             =   2805
      Width           =   45
   End
   Begin VB.Label Label8 
      Caption         =   "Ch&aracters to copy:"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "&Font:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmCharacterMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'******************************************************************************
'*Character Map recreation -17/jul/2000
'******************************************************************************
Private Type POINTAPI  '  8 Bytes
    x As Long
    y As Long
End Type
Private Declare Function SelectObject& Lib "gdi32" (ByVal hdc As Long, ByVal hObject As _
        Long)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function MoveToEx& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
        ByVal y As Long, lpPoint As POINTAPI)
'Private Declare Function CreateRectRgnIndirect& Lib "gdi32" (lprect As RECT)
Private Declare Function CreateRectRgn& Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As _
        Long, ByVal X2 As Long, ByVal Y2 As Long)
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Private Declare Function LineTo& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal _
        y As Long)
Private Declare Function Rectangle& Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)

Dim asciiList() ' list of character descriptions
Dim sizeX, sizeY, previousX, previousY
Dim mouseDown As Boolean, mouseVisible As Boolean

Private Sub CboFonts_Click()

    drawSquare CboFonts.List(CboFonts.ListIndex)
    Picture2.Font = CboFonts.List(CboFonts.ListIndex)
    Picture2.FontSize = 18
    Txtcopy.Font = CboFonts.List(CboFonts.ListIndex)

    'reselect last square
    drawfocusColour previousX, previousY
    

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    'copy to clipboard
    Clipboard.Clear
    Clipboard.SetText Txtcopy.text, vbCFText

End Sub

Private Sub cmdSelect_Click()
    'Txtcopy.Text = Label7.Caption
    RichTextBox1.text = Label7.Caption
RichTextBox1.Find ("&&")
RichTextBox1.SelRTF = "&"

    inserttext
End Sub
Sub inserttext()
    Dim X1, Y1, char$, lprect As RECT, offsetx, offsety, s
    s = selectedsquare
    Y1 = s \ 32
    X1 = s Mod 32
    char$ = Chr$((Y1 * 32) + (X1 + 1) + 30) '1)
    Txtcopy.SelText = char$

End Sub

Private Sub Command1_Click()
frmMain.ActiveForm.RTF1.SelText = RichTextBox1.text
End Sub

Private Sub Form_Load()
    Dim x, y
    Dim index
    '' Form1.ScaleWidth = 32 * 7
    'Form1.Show
    sizeX = (Picture1.ScaleWidth \ 32) ' + 1 '  32*7=224
    sizeY = (Picture1.ScaleHeight \ 7) ''' + 1 '  32*7=224
    '''MsgBox sizeX & ":" & sizeY
    createAsciiList
    '    Form1.ForeColor = vbBlack
    '    Form1.Picture = LoadPicture()

    ''Form1.Refresh

    ''Form1.AutoRedraw = True
    index = 32
    drawSquare "Times New Roman"



    FillListWithFonts CboFonts 'List1
    CboFonts.ListIndex = 0

    Picture2.Visible = False
    Picture3.Visible = False
    mouseDown = False

    'previousX = 0 'start off with the first square
    'previousY = 0
    'starts off with first square selected
    '''drawfocusColour 0, 0
    '''updateLabel 0, 0
    Picture1_MouseDown 0&, 0&, 0, 0
    Picture1_MouseUp 0&, 0&, 0, 0
    'Picture1.SetFocus
    cmdCopy.Enabled = False

    selectedsquare = 1
End Sub

'''
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub updateLabel(x, y)
    Dim key, k$
    'give keystroke and alt information

    key = (y * 32) + (x + 1) ' + 31
    k$ = "Keystroke: "
    'MsgBox key
    Select Case key
    Case 1
        Label4.Caption = k$ & "Spacebar"
    Case 2 To 95 '
        Label4.Caption = k$ & Chr$(key + 31)
    Case 96 To 97
        Label4.Caption = k$ & "Ctrl+" & (key - 95)
    Case 98 To 224
        Label4.Caption = k$ & "Alt+0" & key + 31
    End Select
    'hex / bin text
    Txt1.text = Hex(key + 31)
    Txt2.text = Bin(key + 31, 8)
    Txt3.text = key + 31
    'Label1.Caption = "Col: " & X1 & " Line: " & Y1 & ":" '& x1 * y1
    Label1.Caption = "Col: " & x & " Line: " & y & " Square:" & (y * 32) + (x + 1) & ": Ascii " & key + 31 ' * (y1 + 1)
    'Debug.Print key
    'asciilist array starts at 0 index
    Select Case key
    Case 1 To 98
        Label7.Caption = asciiList(key - 1)
    Case 99 To 129
        Label7.Caption = asciiList(key - 1)
    Case 130 To 224
        Label7.Caption = asciiList(key - 1)

    End Select
End Sub
Sub createAsciiList()
    ReDim asciiList(250)

    Dim a$, index As Long
    Open App.Path & "\asciiquoteds.txt" For Input As 1
    Do While Not (EOF(1))
        Input #1, a$
        asciiList(index) = a$
        index = index + 1
    Loop


    Close 1
    'For index = 0 To 255 - 31
    'MsgBox asciiList(index)
    'Next index
End Sub

Private Sub Picture1_DblClick()
    RichTextBox1.text = Label7.Caption
RichTextBox1.Find ("&&")
RichTextBox1.SelRTF = "&"

    inserttext
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyDown
        If selectedsquare + 32 < 225 Then
            selectedsquare = selectedsquare + 32
        End If
    Case vbKeyUp
        If selectedsquare - 32 > 0 Then
            selectedsquare = selectedsquare - 32
        End If
    Case vbKeyRight
        If selectedsquare + 1 < 225 Then
            selectedsquare = selectedsquare + 1
        End If
    Case vbKeyLeft
        If selectedsquare - 1 > 0 Then
            selectedsquare = selectedsquare - 1
        End If
    Case Else
        Exit Sub
    End Select
    drawselected (selectedsquare - 1)
    updateLabel (selectedsquare - 1) Mod 32, (selectedsquare - 1) \ 32
    
End Sub
Sub drawselected(s As Long)
    Dim X1, Y1, char$, lprect As RECT, offsetx, offsety
    Y1 = s \ 32
    X1 = s Mod 32
    'erase previous ?
    Picture1.Line (previousX * sizeX + 1, previousY * sizeY + 1)-(previousX * sizeX + (sizeX - 1), previousY * sizeY + (sizeY - 1)), vbWhite, BF
    Picture1.CurrentX = (previousX * sizeX) + 3
    Picture1.CurrentY = (previousY * sizeY)

    Picture1.Print Chr$((previousY * 32) + (previousX + 1) + 31);
    previousX = X1
    previousY = Y1

    char$ = Chr$((Y1 * 32) + (X1 + 1) + 31)
    Picture2.Visible = False: Picture3.Visible = False
    offsetx = (Picture2.ScaleWidth - Picture2.TextWidth(char$)) \ 2
    offsety = (Picture2.ScaleHeight - Picture2.TextHeight(char$)) \ 2
    Picture2.left = (X1 * sizeX - 5) + 10
    Picture2.top = (Y1 * sizeY - 5) + 35
    Picture3.left = Picture2.left + 5
    Picture3.top = Picture2.top + 5
    Picture2.CurrentX = offsetx
    Picture2.CurrentY = offsety '    Chr$((y1 * 32) + (x1 + 1) + 31)
    Picture2.Picture = LoadPicture()
    Picture2.Print Chr$((Y1 * 32) + (X1 + 1) + 31)
    Picture2.Visible = True: Picture3.Visible = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '*******************************************************************************
    '* Name:  Picture1_MouseDown
    '*
    '* Description:
    '*
    '* Date Created:  7/17/00
    '*
    '* Created By: oigres P
    '*
    '* Modified: 7/19/00
    '*
    '*******************************************************************************
    Dim X1, Y1, ret, lprect As RECT
    X1 = x \ sizeX
    Y1 = y \ sizeY
    If Button = vbRightButton Then
        Exit Sub
    End If
    'if in square of picture
    If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then
        ''If x1 <> previousX And y <> previousY Then
        'erase previous focus rectangle
        ''MsgBox IsEmpty(previousX)
        If Not (IsEmpty(previousX) And IsEmpty(previousY)) Then
            lprect.left = X1 * sizeX + 1
            lprect.top = Y1 * sizeY + 1
            lprect.right = X1 * sizeX + (sizeX - 1) + 1 '- 1
            lprect.bottom = Y1 * sizeY + (sizeY - 1) + 1
            ''DrawFocusRect Picture1.hdc, lprect

            Picture1.Line (previousX * sizeX, previousY * sizeY)-(previousX * sizeX + (sizeX), previousY * sizeY + (sizeY)), vbBlack, BF
            Picture1.Line (previousX * sizeX + 1, previousY * sizeY + 1)-(previousX * sizeX + (sizeX - 1), previousY * sizeY + (sizeY - 1)), vbWhite, BF
            Picture1.CurrentX = (previousX * sizeX) + 3
            Picture1.CurrentY = (previousY * sizeY)
            '''
            Picture1.Print Chr$((previousY * 32) + (previousX + 1) + 31);
        End If
        Picture2.Visible = False
        Picture3.Visible = False
        Picture2.left = (X1 * sizeX - 5) + 10
        Picture2.top = (Y1 * sizeY - 5) + 35
        Picture3.left = Picture2.left + 5
        Picture3.top = Picture2.top + 5
        Picture2.Visible = True
        Picture3.Visible = True
        selectedsquare = (Y1 * 32) + (X1 + 1)

        previousX = X1
        previousY = Y1
    End If ' in square

    'draw focus rectangle
    
    Call updateLabel(X1, Y1)
    'hide cursor
    If mouseDown = False Then
        ret = ShowCursor(False)
        'showcursor shows cursor if the return count >=0
        'force it to hide
        While ret >= 0
            ret = ShowCursor(False)
        Wend
        '    If ret >= 0 Then
        '
        '    End If
        mouseVisible = False
        '' Label5.Caption = "showcursor times= " & ret
    End If
    mouseVisible = False
    ''Form1.MousePointer = 15
    Picture2.Visible = True
    Picture3.Visible = True
    mouseDown = True
End Sub

'/******************************************************************************
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim X1, Y1, ret, char$, key
    Dim offsetx, offsety
    Static lastx
    Static lasty

    If mouseDown = True Then

        X1 = x \ sizeX
        Y1 = y \ sizeY
        If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then


            If mouseVisible = True Then
                makeCursorInvisible
            End If
            If lastx = X1 And lasty = Y1 Then Exit Sub
            lastx = X1: lasty = Y1
            key = (Y1 * 32) + (X1 + 1)

            Picture2.Visible = False
            Picture3.Visible = False
            Picture2.left = (X1 * sizeX - 5) + 10
            Picture2.top = (Y1 * sizeY - 5) + 35
            Picture3.left = Picture2.left + 5
            Picture3.top = Picture2.top + 5
            '            Picture2.Visible = True
            '            Picture3.Visible = True
            char$ = Chr$((Y1 * 32) + (X1 + 1) + 31)
            If Picture2.tag = char$ Then
            Else

                '        Picture1.Picture = LoadPicture()
                '        Picture1.CurrentX = 0: Picture1.CurrentY = 0
                '        Picture1.Print Chr$((y1 * 32) + (x1 + 1) + 31)
                previousX = X1
                previousY = Y1
                ''Picture2.Visible = False
                Picture2.tag = char$


                offsetx = (Picture2.ScaleWidth - Picture2.TextWidth(char$)) \ 2
                offsety = (Picture2.ScaleHeight - Picture2.TextHeight(char$)) \ 2
                Picture2.CurrentX = offsetx
                Picture2.CurrentY = offsety '    Chr$((y1 * 32) + (x1 + 1) + 31)
                Picture2.Picture = LoadPicture()
                Picture2.Print Chr$((Y1 * 32) + (X1 + 1) + 31)
                Picture2.Visible = True
                Picture3.Visible = True
            End If 'if tag

            Call updateLabel(X1, Y1)
            previousX = X1
            previousY = Y1
        Else 'not in square
            'showcursor ?
            makeCursorVisible


            Exit Sub
        End If ' x1 >= 0 And x1 <= 31 And y1 >= 0 And y1 <= 6
        ''ShowCursor False
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ret, X1, Y1, lprect As RECT
    X1 = x \ sizeX
    Y1 = y \ sizeY
    If mouseVisible = False Then
        ret = ShowCursor(True)
        While ret < 0
            ret = ShowCursor(True)
        Wend
        mouseVisible = True
    End If
    drawfocusColour previousX, previousY
    If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then

    Else
        If mouseDown = True Then
            Picture2.Visible = False
            Picture3.Visible = False
            'draw red focus rectangle
            drawfocusColour previousX, previousY
        End If

    End If
    Picture2.Visible = False
    Picture3.Visible = False
    mouseDown = False
End Sub
'
Sub makeCursorInvisible()
    Dim ret
    ret = ShowCursor(False)
    While ret >= 0
        ret = ShowCursor(False)
    Wend
    mouseVisible = False
End Sub
Sub makeCursorVisible()
    Dim ret
    ret = ShowCursor(True)
    While ret < 0
        ret = ShowCursor(True)
    Wend
    mouseVisible = True
End Sub
Sub drawSquare(f As String)
    Dim x As Long, y As Long, char$, lppt As POINTAPI
    Dim offsetx, offsety
    Picture1.FontName = f
    Picture1.FontSize = 8
    Picture1.Picture = LoadPicture()
    For x = 0 To 31 '32
        For y = 0 To 6 '7
            ''Picture1.Line (x * sizex, y * sizey)-(x * sizex + (sizey - 1), y * sizex + (sizey - 1)), vbBlack, B
            char$ = Chr$((y * 32) + (x + 1) + 31)
            offsetx = (sizeX - Picture1.TextWidth(char$)) \ 2
            offsety = (sizeY - Picture1.TextHeight(char$)) \ 2
            Picture1.CurrentX = (x * sizeX) + offsetx
            Picture1.CurrentY = (y * sizeY) + offsety
            Picture1.Print char$;

        Next y
    Next x
    For x = 0 To 7
        MoveToEx Picture1.hdc, 0, x * sizeY, lppt
        LineTo Picture1.hdc, sizeX * 32, x * sizeY
    Next x
    For x = 0 To 32
        MoveToEx Picture1.hdc, x * sizeX, 0, lppt
        LineTo Picture1.hdc, x * sizeX, sizeY * 7 'Picture1.ScaleHeight - 1
    Next x

End Sub

Private Sub Txtcopy_Change()
    If Txtcopy.text = "" Then
        cmdCopy.Enabled = False
    Else

        cmdCopy.Enabled = True
    End If
End Sub
Sub drawfocusColour(x, y)
    Dim lprect As RECT
    Picture1.Line (x * sizeX + 1, y * sizeY + 1)-(x * sizeX + (sizeX - 1), _
            y * sizeY + (sizeY - 1)), vbHighlight, BF
    ''Picture1.FillColor = vbHighlight
    'Rectangle Picture1.hdc, x * sizeX + 1, y * sizeY + 1, x * sizeX + (sizeX), y * sizeY + (sizeY)
    ''Picture1.FillColor = vbWhite
    Picture1.CurrentX = (x * sizeX) + 3
    Picture1.CurrentY = (y * sizeY)
    Picture1.ForeColor = vbWhite
    Picture1.Print Chr$((y * 32) + (x + 1) + 31);
    Picture1.ForeColor = vbBlack
    ''previousX = x
    ''previousY = y
    lprect.left = x * sizeX + 1
    lprect.top = y * sizeY + 1
    lprect.right = x * sizeX + (sizeX - 1) + 1 '- 1
    lprect.bottom = y * sizeY + (sizeY - 1) + 1  '- 1

    DrawFocusRect Picture1.hdc, lprect
End Sub


