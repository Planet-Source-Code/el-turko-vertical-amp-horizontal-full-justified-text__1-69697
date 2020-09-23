VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vertical & Horizantal justified text"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   525
      Left            =   7980
      TabIndex        =   3
      Top             =   1440
      Width           =   1275
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vertical"
      ForeColor       =   &H80000008&
      Height          =   465
      Index           =   1
      Left            =   8010
      TabIndex        =   1
      Top             =   870
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Horizantal"
      ForeColor       =   &H80000008&
      Height          =   465
      Index           =   0
      Left            =   8010
      TabIndex        =   0
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Height          =   1635
      Left            =   390
      TabIndex        =   2
      Top             =   6060
      Width           =   6435
   End
   Begin VB.Shape Shape1 
      Height          =   3285
      Left            =   360
      Top             =   270
      Width           =   7365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const c_DTDefFmt = DT_NOPREFIX 'Or DT_SINGLELINE Or DT_VCENTER
Private Const DT_BOTTOM = &H8
Private Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Private Const DT_DISPFILE = 6            '  Display-file
Private Const DT_RASDISPLAY = 1          '  Raster display
Private Const DT_RASPRINTER = 2          '  Raster printer
Private Const DT_PLOTTER = 0             '  Vector plotter
Private Const DT_TOP = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_MODIFYSTRING = 65536
Private Const DT_WORD_ELLIPSIS = 262144
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_RTLREADING = 131072
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Const LF_FACESIZE = 32

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
    lfFaceName As String * LF_FACESIZE
End Type
Private Enum FontWeight 'not used; just FYI
    FW_DONTCARE = 0
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_HEAVY = 900
End Enum
Private Type Size
        cx As Long
        cy As Long
End Type
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const CLIP_LH_ANGLES = &H10
Private Const TRUETYPE_FONTTYPE = &H4
Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const HANGEUL_CHARSET = 129
Const CHINESEBIG5_CHARSET = 136
Const OEM_CHARSET = 255
Const tr_charset = 162
Const OUT_CHARACTER_PRECIS = 2
Const OUT_DEFAULT_PRECIS = 0
Const OUT_DEVICE_PRECIS = 5
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_CHARACTER_PRECIS = 1
Const CLIP_STROKE_PRECIS = 2
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
Const ANTIALIASED_QUALITY = 4
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2
Const OPAQUE = 2
Const TRANSPARENT = 1
Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
Private Const PHYSICALWIDTH = 110
Private Const PHYSICALHEIGHT = 111
Private Const PHYSICALOFFSETX = 112
Private Const PHYSICALOFFSETY = 113
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal U As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextJustification Lib "gdi32" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Sub Command1_Click()
  Dim F As LOGFONT, hPrevFont As Long, FontName As String
  Dim FONTSIZE As Integer

Dim i As Integer
Dim tzk As Size: Dim std As RECT: Dim ddd, tabsabit1, tabsabit2 As Long: Dim Ltxt() As String: Dim sze As String
Dim hFont As Long: Dim sonuc As Long
If Option1(0).Value = True Then
Form1.Cls
  FontName = "Verdana" & Chr$(0)
  F.lfCharSet = 1
  F.lfFaceName = FontName
  F.lfHeight = (9 * -20) / Screen.TwipsPerPixelY
  F.lfEscapement = 0
  F.lfOrientation = 0
  F.lfQuality = ANTIALIASED_QUALITY
  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(Me.hdc, hFont)

Ltxt = SplitLines2(Label1.Caption, Form1, CSng(Shape1.Width - 2) - CSng(Shape1.Left + 1))
std.Top = Shape1.Top: std.Left = Shape1.Left + 1: std.Right = Shape1.Left + Shape1.Width - 2: std.Bottom = std.Top

For i = 1 To UBound(Ltxt)
ret = GetTextExtentPoint32(Me.hdc, Ltxt(i), Len(Ltxt(i)), tzk)
std.Bottom = std.Bottom + tzk.cy
sonuc = std.Right - std.Left - tzk.cx
retval = SetTextJustification(Me.hdc, Abs(sonuc), CountStrings(CStr(Ltxt(i)), " "))
dvmct:
DrawText Me.hdc, CStr(Ltxt(i)), -1, std, DT_LEFT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_EXPANDTABS
std.Top = std.Bottom
retval = SetTextJustification(Me.hdc, 0, 0)
Next
  hFont = SelectObject(Me.hdc, hPrevFont)
  DeleteObject hFont

Else

Form1.Cls
  FontName = "Verdana" & Chr$(0)
  F.lfCharSet = 1
  F.lfFaceName = FontName
  F.lfHeight = (9 * -20) / Screen.TwipsPerPixelY
  F.lfEscapement = 900
  F.lfOrientation = 900
  F.lfQuality = ANTIALIASED_QUALITY
  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(Me.hdc, hFont)

Ltxt = SplitLines2(Label1.Caption, Form1, CSng(Shape1.Height + Shape1.Top) - CSng(Shape1.Top))
std.Left = Shape1.Left
std.Top = Shape1.Height + Shape1.Top - 2
std.Right = std.Left
std.Bottom = -std.Top
 

For i = 1 To UBound(Ltxt)
ret = GetTextExtentPoint32(Me.hdc, Ltxt(i), Len(Ltxt(i)), tzk)
std.Right = std.Right + tzk.cy
sonuc = Shape1.Height - tzk.cx - 4
retval = SetTextJustification(Me.hdc, sonuc, CountStrings(CStr(Ltxt(i)), " "))

dvmct2:
DrawText Me.hdc, CStr(Ltxt(i)), -1, std, DT_LEFT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_EXPANDTABS
std.Left = std.Right
retval = SetTextJustification(Me.hdc, 0, 0)
Next
  hFont = SelectObject(Me.hdc, hPrevFont)
  DeleteObject hFont
End If

End Sub

Private Function SplitLines2(Txt As String, P As Object, W As Single) As String()
    Dim Lines() As String, CurrW As Single, CurrWord As String
    Dim L As Integer, i As Integer, WCnt As Integer
    CurrW = 0
    L = Len(Txt)


    If (P.TextWidth(Txt) > W) Or (InStr(Txt, vbCr) > 0) Then
            i = 1
            WCnt = 1
            ReDim Lines(WCnt) As String
            Do Until i > L
                CurrWord = ""
                Do Until i > L Or Mid(Txt, i, 1) <= " "
                    CurrWord = CurrWord & Mid(Txt, i, 1)
                    i = i + 1
                Loop
                If CurrW + P.TextWidth(CurrWord) > W Then
                    WCnt = WCnt + 1
                    ReDim Preserve Lines(WCnt) As String
                    CurrW = 0
                End If
                Lines(WCnt) = Lines(WCnt) + CurrWord
                CurrW = P.TextWidth(Lines(WCnt))
                Do Until i > L Or Mid(Txt, i, 1) > " "
                    Select Case Mid(Txt, i, 1)
                    Case " "
                        Lines(WCnt) = Lines(WCnt) + " "
                        CurrW = P.TextWidth(Lines(WCnt))
                    Case vbLf
                    Case vbCr
                        WCnt = WCnt + 1
                        ReDim Preserve Lines(WCnt) As String
                        CurrW = 0
                    Case Chr(9)
                        Lines(WCnt) = Lines(WCnt) + " "
                        CurrW = P.TextWidth(Lines(WCnt))
                    End Select
                    i = i + 1
                Loop
            Loop
    Else
            ReDim Lines(1) As String
            Lines(1) = Txt
    End If


    For i = 1 To WCnt
        Lines(i) = LTrim(RTrim(Lines(i)))
    Next i
    SplitLines2 = Lines
End Function

Private Function CountStrings(ByVal where As String, ByVal what As String) As Long
Dim pos As Long, count As Long
    pos = 1
    Do While InStr(pos, where, what)
        pos = InStr(pos, where, what) + 1
        count = count + 1
    Loop
    CountStrings = count
End Function

Private Sub Form_Load()
Label1.Caption = "Visual Basic (VB) is a RAD (Rapid Application Development) tool, that allows programmers to create Windows applications in very little time.   It is the most popular programming language in the world, and has more programmers and lines of code than any of its nearest competitors."
End Sub
