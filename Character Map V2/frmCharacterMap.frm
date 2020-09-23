VERSION 5.00
Begin VB.Form frmCharacterMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Map"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharacterMap.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6720
      ScaleHeight     =   465
      ScaleWidth      =   345
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   495
      Left            =   6765
      ScaleHeight     =   495
      ScaleWidth      =   375
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2340
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7320
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1845
      Width           =   1215
   End
   Begin VB.CommandButton cmdCut 
      Caption         =   "Cu&t"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1485
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1125
      Width           =   1215
   End
   Begin VB.TextBox txtChars 
      Height          =   285
      Left            =   4800
      TabIndex        =   10
      Top             =   150
      Width           =   1695
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   150
      Width           =   2415
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picCharacterMap 
      AutoRedraw      =   -1  'True
      Height          =   1875
      Left            =   240
      ScaleHeight     =   1815
      ScaleWidth      =   6270
      TabIndex        =   1
      Top             =   840
      Width           =   6330
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C&haracters:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ch&aracters to Copy:"
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   9
      Top             =   195
      Width           =   1470
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Font:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   200
      Width           =   390
   End
End
Attribute VB_Name = "frmCharacterMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI 'Type for holding X & Y co-ordinates
    X As Long
    Y As Long
End Type
Private Const strChars As String = _
    " !""#$%&'()*+,-./0123456789:;<=>?" & _
    "@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_" & _
    "`abcdefghijklmnopqrstuvwxyz{|}~" & _
    "€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ" & _
    " ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿" & _
    "ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞß" & _
    "àáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ" 'All the characters to show
Private Const intCharsPerRow As Integer = 32 'Amount of characters per row
'Button Messages (BM)
Private Const BM_SETSTYLE = &HF4
'Button Styles (BS)
Private Const BS_PUSHBUTTON = &H0&
Private Const BS_USERBUTTON = &H8&
Private intPixelBlockWidth As Integer, _
    intPixelBlockHeight As Integer  'The sizes of the block in pixels
Private Const intMagnification As Integer = 3 'The magnification of the large character
Private Const intShadowOffsetX As Integer = 2, _
    intShadowOffsetY As Integer = 3 'How much to move the shadow over by in pixels
Private intLastOn As Integer 'The last active character
Private sngBlockWidth As Single, sngBlockHeight As Single
Private bolHasFocus As Boolean 'Whether the picture box has focus
Private bolCursorVisible As Boolean 'Whether the cursor is visble or not
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long                      'Finds the cursor's co-ordinates
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long 'API for finding the hWnd of the window under the cursor

Private Const WM_COPY = &H301
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CUT = &H300
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long



'-------------------------------------
'EDIT THESE SUBS TO FIT YOUR OWN NEEDS
Private Sub cmdInsert_Click()
    On Local Error Resume Next
    'ADD INSERT CODE HERE!
End Sub
'END EDITTING
'-------------------------------------


Private Sub DrawCharacter(ByVal Character As String, _
    Optional ByVal Highlighted As Boolean = False, _
    Optional ByVal Focus As Boolean = False)
    On Local Error Resume Next
    With picTemp.Font
        .Bold = False
        .Italic = False
        .Strikethrough = False
        .Underline = False
    End With
    With picTemp
        'Remove old drawings
        .Cls
        'Back/Fore colour = Highlighted or not
        .BackColor = IIf(Highlighted, vbHighlight, vbWindowBackground)
        .ForeColor = IIf(Highlighted, vbHighlightText, vbWindowText)
        'Set the position of the char so that it's centered vertically and horizontally
        .CurrentX = (.ScaleWidth \ 2) - (.TextWidth(Character) \ 2)
        .CurrentY = (.ScaleHeight \ 2) - (.TextHeight(Character) \ 2)
        'Draw the character
        picTemp.Print Character
        'Border
        picTemp.Line (0, 0)-(.ScaleWidth - TwipsX, .ScaleHeight - TwipsY), vbWindowFrame, B
        'Focus rect
        If Focus Then
            'Get the size of the pic box
            Dim rctTemp As RECT
            Call GetClientRect(.hWnd, rctTemp)
            'Move the rect values all in one so we don't end up with a focus rect over the border
            Call InflateRect(rctTemp, -.DrawWidth, -.DrawWidth * 2)
            'Draw the focus
            Call DrawFocusRect(.hDC, rctTemp)
        End If
        'Show changes
        If .AutoRedraw Then .Refresh
    End With
End Sub

Private Sub DrawMap()
    On Local Error Resume Next
    With picCharacterMap
        .Cls
        Dim intLoopCounter As Integer, intRowNumber As Integer, _
            intModulus As Integer
        'Make sure we have the right font
        picTemp.Font.Name = .Font.Name
        'Loop for all chars
        intRowNumber = -1
        For intLoopCounter = 1 To Len(strChars)
            'Get what's left over after dividing by the number of chars per row
            intModulus = (intLoopCounter - 1) Mod intCharsPerRow
            'If it's 0 then it's time to start a new line
            If intModulus = 0 Then intRowNumber = intRowNumber + 1
            'Draw the character to the temp pic box
            Call DrawCharacter(Mid(strChars, intLoopCounter, 1), _
                intLastOn = intLoopCounter, bolHasFocus And intLastOn = intLoopCounter)
            'Now copy it to the correct point in the character map (including the borders)
            Call BitBlt(.hDC, (intModulus * sngBlockWidth) / TwipsX, _
                (intRowNumber * sngBlockHeight) / TwipsY, _
                intPixelBlockWidth + (picTemp.DrawWidth * 2), _
                intPixelBlockHeight + (picTemp.DrawWidth * 2), _
                picTemp.hDC, 0, 0, vbSrcCopy)
        Next intLoopCounter
        '.ScaleLeft = 0
        '.ScaleTop = 0
        '.ScaleWidth = .Width
        '.ScaleHeight = .Height
        '.ScaleWidth = TwipsX(((intModulus * sngBlockWidth) / TwipsX) + intPixelBlockWidth + (picTemp.DrawWidth * 2))
        '.ScaleHeight = TwipsY((((intRowNumber * sngBlockHeight) / TwipsY) + intPixelBlockHeight + (picTemp.DrawWidth * 2)))
        'Show the changes
        If .AutoRedraw Then Call .Refresh
    End With
End Sub

Private Sub HighLightCharacter(ByVal Index As Integer)
    On Local Error Resume Next
    Dim strCharacter As String
    If Index > Len(strChars) Then Index = Len(strChars)
    strCharacter = Mid(strChars, Index, 1)
    With picLarge
        'Remove old drawings
        .Cls
        'Center character
        .CurrentX = (.ScaleWidth \ 2) - (.TextWidth(strCharacter) \ 2)
        .CurrentY = (.ScaleHeight \ 2) - (.TextHeight(strCharacter) \ 2)
        'Draw character
        picLarge.Print strCharacter
        'Show changes
        If .AutoRedraw Then .Refresh
    End With
    
    With picCharacterMap
        Dim sngX As Single, sngY As Single, sngTemp As Single
        sngY = Int(intLastOn / intCharsPerRow) * intPixelBlockHeight
        sngTemp = intLastOn Mod intCharsPerRow
        If sngTemp <> 0 Then
            sngX = (sngTemp - 1) * intPixelBlockWidth
        Else
            sngX = intPixelBlockWidth * (intCharsPerRow - 1)
            sngY = sngY - intPixelBlockHeight
        End If
        
        'Remove the last on character
        If intLastOn >= 0 Then Call DrawCharacter(Mid(strChars, intLastOn, 1))
        Call BitBlt(.hDC, sngX, sngY, intPixelBlockWidth, _
            intPixelBlockHeight, picTemp.hDC, 0, 0, vbSrcCopy)
        
        'Draw the new on character
        Call DrawCharacter(strCharacter, True, bolHasFocus)
        sngY = Int(Index / intCharsPerRow) * intPixelBlockHeight
        sngTemp = Index Mod intCharsPerRow
        If (sngTemp) <> 0 Then
            sngX = (sngTemp - 1) * intPixelBlockWidth
        Else
            sngX = intPixelBlockWidth * (intCharsPerRow - 1)
            sngY = sngY - intPixelBlockHeight
        End If
        Call BitBlt(.hDC, sngX, sngY, _
            intPixelBlockWidth, intPixelBlockHeight, picTemp.hDC, 0, 0, vbSrcCopy)
        If .AutoRedraw Then .Refresh
    End With
    intLastOn = Index
End Sub

Private Sub PositionLargeCharacter(ByVal Index As Integer)
    Dim intRow As Integer, intColumn As Integer
    If Index > Len(strChars) Then Index = Len(strChars)
    intRow = (Index \ (intCharsPerRow)) + 1
    intColumn = Index Mod intCharsPerRow
    If intColumn = 0 Then
        intColumn = intCharsPerRow
        intRow = intRow - 1
    End If
    picLarge.Move picCharacterMap.Left + ((sngBlockWidth * intColumn) - (sngBlockWidth \ 2)) - (picLarge.Width \ 2), _
        picCharacterMap.Top + ((sngBlockHeight * intRow) - (sngBlockHeight \ 2)) - (picLarge.Height \ 2)
    picShadow.Move picLarge.Left + TwipsX(intShadowOffsetX), _
        picLarge.Top + TwipsY(intShadowOffsetY)
    Call SetLargeCharacterVisible
End Sub

Private Sub SetLargeCharacterVisible(Optional ByVal Visible As Boolean = True)
    picLarge.Visible = Visible
    picShadow.Visible = Visible
End Sub

Private Sub cboFont_Click()
    On Local Error Resume Next
    With txtChars.Font
        .Name = cboFont.Text
        .Bold = False
        .Italic = False
        .Strikethrough = False
        .Underline = False
    End With
    With cboFont
        picCharacterMap.FontName = .Text
        picLarge.FontName = .Text
    End With
    Call DrawMap
End Sub

Private Sub cmdClear_Click()
    On Local Error Resume Next
    txtChars.Text = ""
End Sub

Private Sub cmdCopy_Click()
    On Local Error Resume Next
    With txtChars
        'Select all text
        .SelStart = 0
        .SelLength = Len(.Text)
        'Cut to the clipboard
        Call SendMessage(.hWnd, WM_COPY, 0&, 0&)
    End With
End Sub

Private Sub cmdCut_Click()
    On Local Error Resume Next
    With txtChars
        'Select all text
        .SelStart = 0
        .SelLength = Len(.Text)
        'Cut to the clipboard
        Call SendMessage(.hWnd, WM_CUT, 0&, 0&)
    End With
End Sub

Private Sub cmdSelect_Click()
    On Local Error Resume Next
    txtChars.SelText = Mid(strChars, intLastOn, 1)
End Sub

Private Sub picCharacterMap_DblClick()
    Call cmdSelect_Click
End Sub

Private Sub picCharacterMap_GotFocus()
    On Local Error Resume Next
    If Not bolHasFocus Then
        bolHasFocus = True
        Call HighLightCharacter(intLastOn)
    End If
End Sub

Private Sub picCharacterMap_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intTemp As Integer
    Select Case KeyCode
        Case vbKeyLeft
            intTemp = intLastOn - IIf(Shift And vbCtrlMask, 2, 1)
            If intTemp > 0 And intLastOn <> intTemp Then
                Call HighLightCharacter(intTemp)
            Else
                Beep
            End If

        Case vbKeyRight
            intTemp = intLastOn + IIf(Shift And vbCtrlMask, 2, 1)
            If intTemp <= Len(strChars) And intLastOn <> intTemp Then
                Call HighLightCharacter(intTemp)
            Else
                Beep
            End If

        Case vbKeyUp
            If Shift And vbCtrlMask Then
                intTemp = intLastOn Mod intCharsPerRow
                If intTemp = 0 Then intTemp = intCharsPerRow
                If intLastOn <> intTemp And intTemp > 0 Then
                    Call HighLightCharacter(intTemp)
                Else
                    Beep
                End If
            Else
                If intLastOn > intCharsPerRow Then
                    Call HighLightCharacter(intLastOn - intCharsPerRow)
                Else
                    Beep
                End If
            End If
            
        Case vbKeyDown
            If Shift And vbCtrlMask Then
                intTemp = intLastOn Mod intCharsPerRow
                If intTemp = 0 Then intTemp = intCharsPerRow
                intTemp = Len(strChars) - (intCharsPerRow - intTemp)
                If intLastOn <> intTemp And intTemp <= Len(strChars) Then
                    Call HighLightCharacter(intTemp)
                Else
                    Beep
                End If
            Else
                If intLastOn < Len(strChars) - intCharsPerRow + 1 Then
                    Call HighLightCharacter(intLastOn + intCharsPerRow)
                Else
                    Beep
                End If
            End If
        
        Case vbKeyPageUp
            If intLastOn > (intCharsPerRow * 2) Then
                Call HighLightCharacter(intLastOn - (intCharsPerRow * 2))
            Else
                Beep
            End If
            
        Case vbKeyPageDown
            If intLastOn < Len(strChars) - (intCharsPerRow * 2) + 1 Then
                Call HighLightCharacter(intLastOn + (intCharsPerRow * 2))
            Else
                Beep
            End If

        Case vbKeyHome
            If Shift And vbCtrlMask Then
                If intLastOn <> 1 Then
                    Call HighLightCharacter(1)
                Else
                    Beep
                End If
            Else
                intTemp = (intLastOn Mod intCharsPerRow) - 1
                If intTemp = -1 Then intTemp = intCharsPerRow - 1
                If intTemp <> intLastOn And intTemp > 0 Then
                    Call HighLightCharacter(intLastOn - intTemp)
                Else
                    Beep
                End If
            End If
        Case vbKeyEnd
            If Shift And vbCtrlMask Then
                If intLastOn <> Len(strChars) Then
                    Call HighLightCharacter(Len(strChars))
                Else
                    Beep
                End If
            Else
                intTemp = intLastOn Mod intCharsPerRow
                If intLastOn + (intCharsPerRow - intTemp) <> intLastOn And intTemp <> 0 Then
                    Call HighLightCharacter(intLastOn + (intCharsPerRow - intTemp))
                Else
                    Beep
                End If
            End If
        Case Else
            Exit Sub
    End Select
    Call PositionLargeCharacter(intLastOn)
End Sub

Private Sub picCharacterMap_KeyPress(KeyAscii As Integer)
    Dim intCharacterPosition As Integer
    intCharacterPosition = InStr(1, strChars, Chr(KeyAscii))
    If intCharacterPosition > 0 Then
        Call HighLightCharacter(intCharacterPosition)
        Call PositionLargeCharacter(intCharacterPosition)
        Call cmdSelect_Click
    Else
        Call SetLargeCharacterVisible(False)
        Beep
    End If
End Sub

Private Sub picCharacterMap_LostFocus()
    On Local Error Resume Next
    If bolHasFocus Then
        bolHasFocus = False
        Call HighLightCharacter(intLastOn)
        SetLargeCharacterVisible (False)
    End If
End Sub

Private Sub picCharacterMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    Call picCharacterMap_MouseMove(Button, Shift, X, Y)
    If Button And vbLeftButton Then
        Call SetLargeCharacterVisible
    End If
End Sub

Private Sub picCharacterMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    If Button And vbLeftButton Then
        Dim sngTempX As Single, sngTempY As Single
        If X < 0 Then
            sngTempX = 0
        ElseIf X \ sngBlockWidth >= intCharsPerRow Then
            sngTempX = sngBlockWidth * (intCharsPerRow - 1)
        Else
            sngTempX = X
        End If
        sngTempY = (Len(strChars) \ intCharsPerRow)
        If Y < 0 Then
            sngTempY = 0
        ElseIf Y \ sngBlockHeight >= sngTempY Then
            sngTempY = (sngTempY - 1) * sngBlockHeight
            If Len(strChars) Mod intCharsPerRow Then sngTempY = sngTempY + sngBlockHeight
        Else
            sngTempY = Y
        End If
        
        Dim intNewIndex As Integer
        intNewIndex = ((sngTempX \ TwipsX) \ intPixelBlockWidth) + 1 + _
            (((sngTempY \ TwipsY) \ intPixelBlockHeight) * intCharsPerRow)
        If intNewIndex <> intLastOn Then
            Call PositionLargeCharacter(intNewIndex)
            Call HighLightCharacter(intNewIndex)
        End If
        
        Dim rctCharacterMap As RECT
        Call GetWindowRect(picCharacterMap.hWnd, rctCharacterMap)
        If bolCursorVisible And (IsWindowHot(picCharacterMap.hWnd) Or _
            (IsWindowHot(picLarge.hWnd) And IsRECTHot(rctCharacterMap))) Then
            'Hide the cursor
            Call ShowCursor(0)
            bolCursorVisible = False
        ElseIf bolCursorVisible = False And (IsWindowHot(picCharacterMap.hWnd) = False And _
            (IsWindowHot(picLarge.hWnd) = False Or (IsWindowHot(picLarge.hWnd) And IsRECTHot(rctCharacterMap) = False))) Then
            'Show the cursor
            Call ShowCursor(1)
            bolCursorVisible = True
        End If
    End If
End Sub

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Private Sub Form_Load()
    On Local Error Resume Next

    'Load the fonts
    Call LoadFonts(cboFont)
    cboFont.Text = Me.Font.Name
    'Get the correct size (i.e. make the most of the size we have)of the blocks in pixels
    'intPixelBlockWidth = (picCharacterMap.ScaleWidth \ intCharsPerRow) \ TwipsX
    'intPixelBlockHeight = (picCharacterMap.ScaleHeight \ (Len(strChars) \ (intCharsPerRow - 1))) \ TwipsY
    intPixelBlockWidth = (picCharacterMap.ScaleWidth / intCharsPerRow) \ TwipsX
    intPixelBlockHeight = (picCharacterMap.ScaleHeight / (Len(strChars) / intCharsPerRow)) \ TwipsY
    
    'Size of the blocks in twips
    sngBlockWidth = TwipsX(intPixelBlockWidth)
    sngBlockHeight = TwipsY(intPixelBlockHeight)
    'Set the temp pic's size to the size of the block + the width of borders on one _
     side only, as the right/bottom will be covered by the next character
    With picTemp
        .Width = sngBlockWidth + TwipsX(picTemp.DrawWidth)
        .Height = sngBlockHeight + TwipsY(picTemp.DrawWidth)
        'Large/preview pic box size
        picLarge.Width = .Width * intMagnification
        picLarge.Height = .Height * intMagnification
        picLarge.FontSize = .FontSize * intMagnification
    End With
    intLastOn = 0
    'Draw the character map
    Call DrawMap

    'Cursor is visible
    bolCursorVisible = True
    
    'Make sure the large chars are at the front
    Call picShadow.ZOrder(vbBringToFront)
    Call picLarge.ZOrder(vbBringToFront)
    
    'Make buttons 3D
    Call FormatButtons(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    'Get rid of old variables otherwise things don't work properly
    Set frmCharacterMap = Nothing
End Sub

Private Sub picCharacterMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Show the cursor if hidden
    If bolCursorVisible = False Then
        Call ShowCursor(1)
        bolCursorVisible = True
    End If
    If Button = vbLeftButton Then
        Call SetLargeCharacterVisible(False)
    End If
End Sub

Private Sub picLarge_Resize()
    picShadow.Height = picLarge.Height
    picShadow.Width = picLarge.Width
End Sub

Private Function TwipsX(Optional ByVal _
    Amount As Integer = 1) As Single
    On Local Error Resume Next
    'Return the amount of twips in the specified number of pixels
    TwipsX = Amount * Screen.TwipsPerPixelX
End Function

Private Function TwipsY(Optional ByVal _
    Amount As Integer = 1) As Single
    On Local Error Resume Next
    'Return the amount of twips in the specified number of pixels
    TwipsY = Amount * Screen.TwipsPerPixelY
End Function

Private Function IsWindowHot(ByVal hWnd As Long) As Boolean
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    IsWindowHot = WindowFromPoint(CursorPosition.X, CursorPosition.Y) = hWnd 'Return     whether the object is hot
End Function

Private Function IsRECTHot(Area As RECT) As Boolean
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    IsRECTHot = CursorPosition.X >= Area.Left And _
        CursorPosition.X <= Area.Right And _
        CursorPosition.Y >= Area.Top And _
        CursorPosition.Y <= Area.Bottom
End Function

Private Sub FormatButtons(Form As Object)
    On Local Error Resume Next
    'Loop for all controls in form
    Dim lngLoopCounter As Long
    For lngLoopCounter = 0 To Form.Controls.Count - 1
        'If Command Button set style to PushButton
        If TypeOf Form.Controls(lngLoopCounter) Is CommandButton Then _
            Call SendMessage(Form.Controls(lngLoopCounter).hWnd, _
            BM_SETSTYLE, BS_PUSHBUTTON, 0&)
    Next lngLoopCounter
End Sub


