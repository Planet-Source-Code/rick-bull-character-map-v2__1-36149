Attribute VB_Name = "mdlFonts"
Option Explicit
'Font enumeration types
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte

        lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte

        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

' ntmFlags field flags
Private Const NTM_REGULAR = &H40&
Private Const NTM_BOLD = &H20&
Private Const NTM_ITALIC = &H1&

'  tmPitchAndFamily flags
Private Const TMPF_FIXED_PITCH = &H1

Private Const TMPF_VECTOR = &H2
Private Const TMPF_DEVICE = &H8
Private Const TMPF_TRUETYPE = &H4

Private Const ELF_VERSION = 0
Private Const ELF_CULTURE_LATIN = 0

'  EnumFonts Masks
Private Const RASTER_FONTTYPE = &H1
Private Const DEVICE_FONTTYPE = &H2
Private Const TRUETYPE_FONTTYPE = &H4

Private Declare Function EnumFontFamilies Lib "gdi32" Alias _
     "EnumFontFamiliesA" _
     (ByVal hDC As Long, ByVal lpszFamily As String, _
     ByVal lpEnumFontFamProc As Long, lParam As Any) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
     ByVal hDC As Long) As Long

Public Sub LoadFonts(ByVal ComboBox As Object)
    Dim hDC As Long
    ComboBox.Clear
    hDC = GetDC(ComboBox.hWnd)
    Call EnumFontFamilies(hDC, vbNullString, AddressOf EnumFontFamProc, ComboBox)
    Call ReleaseDC(ComboBox.hWnd, hDC)
End Sub

Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
    ByVal FontType As Long, lParam As ComboBox) As Long
    On Local Error Resume Next
    Dim FaceName As String
    Dim FullName As String
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    Call lParam.AddItem(Left$(FaceName, InStr(FaceName, vbNullChar) - 1))
    EnumFontFamProc = 1
End Function


