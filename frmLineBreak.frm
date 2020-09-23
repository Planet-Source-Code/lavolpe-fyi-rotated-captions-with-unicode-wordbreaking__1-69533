VERSION 5.00
Begin VB.Form frmLineBreak 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WordBreak & Rotated Text Example"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFlags 
      Caption         =   "Single Line"
      Height          =   240
      Index           =   4
      Left            =   5190
      TabIndex        =   21
      Top             =   1485
      Width           =   1650
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Clip Text to Rectangle"
      Height          =   240
      Index           =   3
      Left            =   3180
      TabIndex        =   20
      Top             =   1485
      Width           =   1950
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Hide Prefix"
      Height          =   240
      Index           =   2
      Left            =   1965
      TabIndex        =   19
      Top             =   1485
      Width           =   1275
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Process As Caption"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   18
      Top             =   1485
      Value           =   1  'Checked
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   360
      Left            =   120
      TabIndex        =   17
      Top             =   4800
      Width           =   1680
   End
   Begin VB.PictureBox picUnicodeHolder 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   3525
      ScaleHeight     =   990
      ScaleWidth      =   3345
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   135
      Width           =   3345
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "If you are seeing this during run-time, then your system does not support unicode windows."
         Height          =   645
         Left            =   120
         TabIndex        =   16
         Top             =   165
         Width           =   2730
      End
   End
   Begin VB.OptionButton optUnicode 
      Caption         =   "Use this textbox to display results"
      Height          =   240
      Index           =   1
      Left            =   3525
      TabIndex        =   14
      Top             =   1155
      Width           =   3150
   End
   Begin VB.OptionButton optUnicode 
      Caption         =   "Use this textbox to display results"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   1155
      Value           =   -1  'True
      Width           =   3240
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   105
      ScaleHeight     =   190
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   444
      TabIndex        =   1
      Top             =   1845
      Width           =   6720
      Begin VB.CheckBox chkFlags 
         Caption         =   "Make right to left (not valid for all fonts)"
         Height          =   330
         Index           =   0
         Left            =   3300
         TabIndex        =   12
         Top             =   2490
         Width           =   3315
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   1755
         TabIndex        =   5
         Top             =   1980
         Width           =   1350
         Begin VB.OptionButton optJustify 
            Caption         =   "Center Align"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   8
            Top             =   555
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton optJustify 
            Caption         =   "Bottom Align"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   7
            Top             =   285
            Width           =   1500
         End
         Begin VB.OptionButton optJustify 
            Caption         =   "Top Align"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   6
            Top             =   45
            Width           =   1500
         End
      End
      Begin VB.OptionButton optJustify 
         Caption         =   "Center Justify"
         Height          =   255
         Index           =   2
         Left            =   345
         TabIndex        =   4
         Top             =   2535
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton optJustify 
         Caption         =   "Right Justify"
         Height          =   255
         Index           =   1
         Left            =   345
         TabIndex        =   3
         Top             =   2265
         Width           =   1500
      End
      Begin VB.OptionButton optJustify 
         Caption         =   "Left Justify"
         Height          =   255
         Index           =   0
         Left            =   345
         TabIndex        =   2
         Top             =   2025
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Rotated 270 degrees"
         Height          =   240
         Index           =   2
         Left            =   4935
         TabIndex        =   11
         Top             =   60
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Rotated 90 degrees"
         Height          =   240
         Index           =   1
         Left            =   3150
         TabIndex        =   10
         Top             =   60
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Non-rotated (zero degrees)"
         Height          =   240
         Index           =   0
         Left            =   345
         TabIndex        =   9
         Top             =   60
         Width           =   2070
      End
      Begin VB.Shape Bounds270 
         BorderColor     =   &H00C0C0C0&
         Height          =   2145
         Left            =   4890
         Top             =   300
         Width           =   1635
      End
      Begin VB.Shape Bounds90 
         BorderColor     =   &H00C0C0C0&
         Height          =   2145
         Left            =   3135
         Top             =   315
         Width           =   1635
      End
      Begin VB.Shape BoundsH 
         BorderColor     =   &H00C0C0C0&
         Height          =   1635
         Left            =   330
         Top             =   315
         Width           =   2145
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmLineBreak.frx":0000
      Top             =   135
      Width           =   3345
   End
End
Attribute VB_Name = "frmLineBreak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' SO WHAT IS THIS ?

' Have you ever needed to rotate text at 90 or 270 degrees?  Pretty easy really, but
' have you ever needed to draw at that same degree while word breaking?  Very difficult
' because there is no API that does it for you. Sure you could build a simple little
' word break routine, but will it word break for unicode if your work is to be used
' in other countries? Didn't think about that one, did you?

' This project has two purposes. Its primary target is to word wrap relatively short
' text like those found in button and window captions.

' 1. Word break using Windows wordbreak algorithms. This is accomplished by creating
'    an API window and formatting it just right so that it does the wordbreaking, and
'    all we do is simply get the character positions of the words that start new lines.
'    Clever, huh :)
'    Ah, but you say, DrawText does that for us. True, but it won't do it for rotated
'    fonts.  Also, no longer such a big deal, but DrawTextW is not compatible with Win9x
'    however, these routines use TextOutW which exists on Win9x systems and NT too.
'
'    So why is word breaking necessary? Rotated fonts only. No standard API, except
'    maybe GDI+, will render rotated, multiline text. But if we can word break rotated
'    fonts, then we can render a line at a time. Hey, "Why not just word break horizontal
'    fonts and then rotate the measurements?" Because many, if not most, rotated fonts
'    do not draw at the same exact pixels as non-rotated fonts, therefore, your
'    measurements will be off.
'
'    Ok enough about word breaking and rotated fonts. There supposedly is another
'    workaround using SetWorldTransform; however, I haven't gotten to it yet to test it out

' 2. Correctly positions 90 & 270 degree rotated text to any rectangle.

' Notes (to do):
' create a class version

' APIs used to create memory API edit control (textbox)
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function CreateWindowExW Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Const ES_MULTILINE As Long = &H4&
Private Const ES_CENTER As Long = &H1&
Private Const ES_LEFT As Long = &H0&
Private Const ES_RIGHT As Long = &H2&
Private Const WS_EX_RTLREADING As Long = &H2000&

' Unicode & ANSI declarations for setting up the API textbox
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageA Lib "user32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETTEXT As Long = &HC
Private Const WM_GETTEXT As Long = &HD
Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_LINELENGTH As Long = &HC1
Private Const EM_SETRECT As Long = &HB3
Private Const EM_POSFROMCHAR As Long = &HD6
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long
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
    lfFaceName As String * 32
End Type
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Const OBJ_FONT As Long = 6

Private Declare Function TextOutW Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsA" (ByVal hdc As Long, ByRef lpMetrics As TEXTMETRIC) As Long
Private Type TEXTMETRIC
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
End Type
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetTextExtentPointW Lib "gdi32.dll" (ByVal hdc As Long, ByVal lpszString As Long, ByVal cbString As Long, ByRef lpSize As POINTAPI) As Long
Private Declare Function GetTextExtentPointA Lib "gdi32.dll" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, ByRef lpSize As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

' DC Text alignment & justification settings
Private Declare Function GetClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Const TA_LEFT As Long = 0
Private Const TA_RIGHT As Long = 2
Private Const TA_CENTER As Long = 6
Private Const TA_TOP As Long = 0
Private Const TA_RTLREADING As Long = 256

' Enumerations used for the PrintText routine
Private Enum eFormatFlags       ' same values as DT_Flags
    ptxTOP = &H0                ' top aligned
    ptxLEFT = &H0               ' left justified text
    ptxCENTER = &H1             ' center justified
    ptxRIGHT = &H2              ' right justified
    ptxVCENTER = &H4            ' center aligned
    ptxBOTTOM = &H8             ' bottom aligned
    ptxSINGLELINE = &H20        ' single line, else multiline
    ptxNOCLIP = &H100           ' do not clip text else clipped
    ptxNOPREFIX = &H800         ' do not underscore accelerator key
    ptxRTLREADING = &H20000     ' Right-to-Left else Left-to-Right
    ptxHIDEPREFIX = &H100000    ' Process as plain text else as Caption
End Enum
Private Enum eTextAngles
    Angle0 = 0
    Angle90 = 90
    Angle270 = 270
End Enum

Private tUnicodeHWND As Long    ' API generated unicode text box displayed on screen

Private Sub PrintText(ByVal inText As String, ByVal destDC As Long, _
                        ByVal CanvasX As Long, ByVal CanvasY As Long, _
                        ByVal CanvasCx As Long, ByVal CanvasCy As Long, _
                        Optional ByVal Angle As eTextAngles = Angle0, _
                        Optional ByVal Formatting As eFormatFlags = 0&)


    ' Prints text to destination DC. Uses the font currently selected into the DC
    ' and will rotate that font if needed.
    
    ' Parameters
    '   inText :: the ansi or unicode string to print
    '   destDC :: the DC to print to; current font selected into it is used
    '   CanvasX :: the left coordinate of the rectangle to print to
    '   CanvasY :: the top coordinate of the rectangle to print to
    '   CanvasCx :: the width of the rectangle to print to
    '   CanvasCy :: the height of the rectangle to print to
    '   Angle :: either 0, 90, or 270. Any other value will default to 0 degrees
    '   Formatting :: various values that determine formatting
    
    Dim hFont As Long, hOldFont As Long ' to get font from DC
    Dim uLFont As LOGFONT               ' used to create rotated font
    Dim uTextMetric As TEXTMETRIC       ' used to get font height
    Dim fRect As RECT                   ' formatting rect, used for clipping if needed
    
    Dim X As Long, Y As Long            ' positioning: CurrentX, CurrentY
    Dim txtOffset As POINTAPI           ' per line offsets
    Dim BreakOffsets() As Long          ' where each line begins
    Dim maxLines As Long                ' nr lines to fit in fRect
    Dim LineNr As Long                  ' looping counter
    
    Dim tHwnd As Long                   ' hwnd to our API textbox
    Dim tAlign As Long                  ' window & dc alignment flags
    Dim bIsUnicode As Boolean           ' whether api window is unicode or not
    Dim extPT As POINTAPI               ' width/height of accelerator key if used
    Dim accelKeyPos As Long             ' position in string for accelerator key
    Dim hRgn As Long, cRgn As Long      ' clipping region handles if needed
    Dim lineFirst As Long               ' first line to be displayed in passed canvas
    
    ' setup the api window alignment & style
    If (Formatting And ptxCENTER) = ptxCENTER Then
        tAlign = ES_CENTER
    ElseIf (Formatting And ptxRIGHT) = ptxRIGHT Then
        tAlign = ES_RIGHT
    Else
        tAlign = ES_LEFT
    End If
    If (Formatting And ptxRTLREADING) Then tAlign = tAlign Or WS_EX_RTLREADING
    If (Formatting And ptxSINGLELINE) = 0& Then tAlign = tAlign Or ES_MULTILINE
    
    On Error Resume Next
    ' create our API window, memory only -- it won't be visible
    tHwnd = CreateWindowExW(0&, StrPtr("Edit"), 0&, tAlign, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
    If tHwnd = 0& Then
        If Err Then Err.Clear
        tHwnd = CreateWindowExA(0&, "Edit", vbNullString, tAlign, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
    Else
        bIsUnicode = True
    End If
    If tHwnd = 0& Then
        MsgBox "Error: " & Err.Description, vbExclamation + vbOKOnly, "Oops"
        If Err Then Err.Clear
    End If
    
    ' process acclerator keys as needed
    If (Formatting And ptxNOPREFIX) = 0& Then ' process as caption
        If Not InStr(inText, "&") = 0& Then
            ' rules used here are:
            ' 1. "&&" = "&"
            ' 2. "& " = " "
            ' 3. & and other character makes the other character a hotkey
            ' 4. Last hotkey in string is recognized if multiple are set
            txtOffset.X = Len(inText)         ' adjusted length of text
            maxLines = txtOffset.X            ' original length of text
            For X = 1 To maxLines - 1
                If Mid$(inText, X, 1) = "&" Then
                    ' shift string one character left
                    Mid$(inText, X, maxLines - X) = Mid$(inText, X + 1, maxLines - X)
                    txtOffset.X = txtOffset.X - 1 ' reduce string length
                    ' set hotkey if valid
                    If Not Mid$(inText, X, 1) = "&" Then
                        If Not Mid$(inText, X, 1) = " " Then accelKeyPos = X
                    End If
                End If
            Next
            ' test for & on last character - invalid
            If Right$(inText, txtOffset.X) = "&" Then
                accelKeyPos = 0                 ' if set then no hotkey is valid
                txtOffset.X = txtOffset.X - 1   ' reduce length
            End If
            ' fill and remaining characters after shifting with null characters
            For X = txtOffset.X + 1 To maxLines
                Mid$(inText, X, 1) = vbNullChar
            Next
            ' does user want the hotkey underlined
            If (Formatting And ptxHIDEPREFIX) Then
                accelKeyPos = 0&
            Else
                ' get the width of the hotkey
                If accelKeyPos Then
                    If bIsUnicode Then
                        GetTextExtentPointW destDC, StrPtr(inText) + (accelKeyPos - 1) * 2, 1, extPT
                    Else
                        GetTextExtentPointA destDC, Mid$(inText, accelKeyPos, 1), 1, extPT
                    End If
                End If
            End If
        End If
    End If
    
    ' get the font from the DC
    hOldFont = GetCurrentObject(destDC, OBJ_FONT)
    GetGDIObject hOldFont, Len(uLFont), uLFont
    
    Select Case Angle       ' validate angle is valid
        Case 0, 90, 270
        Case Else: Angle = Angle0
    End Select
    ' use DC font for template, rotate as needed & create
    uLFont.lfOrientation = Angle * 10
    uLFont.lfEscapement = uLFont.lfOrientation
    uLFont.lfOutPrecision = 2
    hFont = CreateFontIndirect(uLFont)
    ' select our font into target DC
    SelectObject destDC, hFont
    
    ' calculate the drawing rectangle for the API text box
    If Angle = Angle0 Then
        fRect.Right = CanvasCx: fRect.Bottom = CanvasCy
    Else    ' flip rectangle 90/270 degrees
        fRect.Right = CanvasCy: fRect.Bottom = CanvasCx
    End If
    
    ' ok set the API textbox properties, using Ansi or Unicode messages
    If bIsUnicode Then
        SendMessageW tHwnd, WM_SETFONT, hFont, ByVal 0&
        SendMessageW tHwnd, EM_SETRECT, 0&, ByVal VarPtr(fRect)
        SendMessageW tHwnd, WM_SETTEXT, 0&, ByVal StrPtr(inText)
    Else
        SendMessageA tHwnd, WM_SETFONT, hFont, ByVal 0&
        SendMessageA tHwnd, EM_SETRECT, 0&, ByVal VarPtr(fRect)
        SendMessageA tHwnd, WM_SETTEXT, 0&, ByVal inText
    End If
    
    ' the textbox already wordbroke the text as needed, we now
    ' will ask it where the linebreaks are...
    LineNr = GetLineBreaks(tHwnd, BreakOffsets(), bIsUnicode, accelKeyPos)
    DestroyWindow tHwnd
    
    ' determine max number of lines that can be displayed w/o clipping
    GetTextMetrics destDC, uTextMetric
    If (Formatting And ptxNOCLIP) = ptxNOCLIP Then
        maxLines = LineNr   ' no clipping, may look like garbage
    Else
        maxLines = (fRect.Bottom \ uTextMetric.tmHeight)
        If (maxLines * uTextMetric.tmHeight) < fRect.Bottom Then maxLines = maxLines + 1
        If maxLines > LineNr Then maxLines = LineNr
    End If
    
    ' now for calculating offsets
    If Angle = Angle90 Then
        txtOffset.Y = 0&
        txtOffset.X = uTextMetric.tmHeight              ' distance to move each line of text
        If (Formatting And ptxBOTTOM) = ptxBOTTOM Then  ' start on right edge of rect
            X = (CanvasCx - maxLines * txtOffset.X) + CanvasX - 1
            lineFirst = LineNr - maxLines + 1           ' first line to render
        ElseIf (Formatting And ptxVCENTER) = ptxVCENTER Then ' start on center of rect
            X = (CanvasCx - LineNr * txtOffset.X) \ 2 + CanvasX - 1
            lineFirst = (LineNr - maxLines) \ 2 + 1
        Else                                            ' start on left edge of rect
            X = CanvasX
            lineFirst = 1
        End If
        Y = CanvasCy + CanvasY
    
    ElseIf Angle = Angle270 Then
        txtOffset.Y = 0&
        txtOffset.X = -uTextMetric.tmHeight ' distance to move each line of text
        If (Formatting And ptxBOTTOM) = ptxBOTTOM Then ' start on left edge of rect
            X = CanvasCx - (CanvasCx + maxLines * txtOffset.X) + CanvasX - 1
            lineFirst = LineNr - maxLines + 1           ' first line to render
        ElseIf (Formatting And ptxVCENTER) = ptxVCENTER Then ' start on center of rect
            X = CanvasCx - (CanvasCx + LineNr * txtOffset.X) \ 2 + CanvasX - 1
            lineFirst = (LineNr - maxLines) \ 2 + 1
        Else                                ' start on right edge of rect
            X = CanvasX + CanvasCx
            lineFirst = 1
        End If
        Y = CanvasY
        
    Else ' Angle0 is default
        txtOffset.X = 0&
        txtOffset.Y = uTextMetric.tmHeight  ' distance to move each line of text vertically
        If (Formatting And ptxBOTTOM) = ptxBOTTOM Then ' start on bottom edge of rect
            Y = (CanvasCy - maxLines * txtOffset.Y) + CanvasY
            lineFirst = LineNr - maxLines + 1           ' first line to render
        ElseIf (Formatting And ptxVCENTER) = ptxVCENTER Then ' start on center of rect
            lineFirst = (LineNr - maxLines) \ 2 + 1
            Y = (CanvasCy - LineNr * txtOffset.Y) \ 2 + CanvasY
        Else                                ' start on top edge of rect
            Y = CanvasY
            lineFirst = 1
        End If
        X = CanvasX
    End If
    
    ' set DC properties & keep return value
    tAlign = TA_LEFT Or TA_TOP
    If (Formatting And ptxRTLREADING) = ptxRTLREADING Then tAlign = tAlign Or TA_RTLREADING
    tAlign = SetTextAlign(destDC, tAlign)
    
    ' setup clipping rectangle if requested
    If (Formatting And ptxNOCLIP) = 0& Then
        hRgn = CreateRectRgn(CanvasX, CanvasY, CanvasCx + CanvasX, CanvasCy + CanvasY)
        If GetClipRgn(destDC, hRgn) = 0& Then
            cRgn = hRgn
            hRgn = 0&
        Else
            cRgn = CreateRectRgn(CanvasX, CanvasY, CanvasCx + CanvasX, CanvasCy + CanvasY)
        End If
        SelectClipRgn destDC, cRgn
        DeleteObject cRgn
    End If
    
    For LineNr = lineFirst To lineFirst + maxLines ' loop thru the line breaks
        ' note that we are using StrPtr & VB stores all strings as Unicode (2bytes per char)
        ' so when linebreaks were calculated, the character position was pre-multiplied
        ' by 2, therefore we don't need to do it here
        ' Also, we are using StrPtr so we don't have use Mid$() which creates another string
        If Angle = Angle0 Then
            TextOutW destDC, X + BreakOffsets(3, LineNr), Y, StrPtr(inText) + BreakOffsets(2, LineNr), BreakOffsets(1, LineNr)
            If LineNr = BreakOffsets(2, 0) Then
                MoveToEx destDC, X + BreakOffsets(1, 0), Y + extPT.Y - 1, ByVal 0&
                LineTo destDC, X + BreakOffsets(1, 0) + extPT.X, Y + extPT.Y - 1
            End If
        ElseIf Angle = Angle90 Then
            TextOutW destDC, X, Y - BreakOffsets(3, LineNr), StrPtr(inText) + BreakOffsets(2, LineNr), BreakOffsets(1, LineNr)
            If LineNr = BreakOffsets(2, 0) Then
                MoveToEx destDC, X + extPT.Y, Y - BreakOffsets(1, 0), ByVal 0&
                LineTo destDC, X + extPT.Y, Y - BreakOffsets(1, 0) - extPT.X
            End If
        Else
            TextOutW destDC, X, Y + BreakOffsets(3, LineNr), StrPtr(inText) + BreakOffsets(2, LineNr), BreakOffsets(1, LineNr)
            If LineNr = BreakOffsets(2, 0) Then
                MoveToEx destDC, X - extPT.Y, Y + BreakOffsets(1, 0), ByVal 0&
                LineTo destDC, X - extPT.Y, Y + BreakOffsets(1, 0) + extPT.X
            End If
        End If
        X = X + txtOffset.X
        Y = Y + txtOffset.Y
    Next
    ' remove clipping rectangle if used & replace previous clipping region if any
    If (Formatting And ptxNOCLIP) = 0& Then
        SelectClipRgn destDC, hRgn
        If hRgn Then DeleteObject hRgn
    End If
    SetTextAlign destDC, tAlign             ' reset to previous justification settings
    
    DeleteObject SelectObject(destDC, hOldFont) ' replace old font, destroy our font

End Sub


Private Function GetLineBreaks(ctrlHwnd As Long, arrayBreaks() As Long, isUnicode As Boolean, Optional ByVal accelKeyPos As Long) As Long

    ' Helper function to determine where linebreaks occur
    ' See PrintText
    
    Dim nrLines As Long
    Dim Index As Long
    
    ' The main difference here is use of Ansi or Unicode messages, otherwise
    ' the logic is the same. The linebreak positions are pre-multiplied by
    ' two because VB stores its strings as Unicode and the PrintText references
    ' the string via StrPtr. So, for example, character position 5 is the 9th & 10th byte
    
    ' The returned array contents are:
    ' (1,0) is accelerator line number if used, (2,0) is acclerator X coordinate
    ' (1,LineNr) is length of the line
    ' (2,LineNr) is character pos*2 which starts new line
    ' (3,LineNr) is the X coordinate of that character
    
    If isUnicode Then
        nrLines = SendMessageW(ctrlHwnd, EM_GETLINECOUNT, 0, ByVal 0&)
        ReDim arrayBreaks(1 To 3, 0 To nrLines)
        For Index = 1 To nrLines
            arrayBreaks(2, Index) = SendMessageW(ctrlHwnd, EM_LINEINDEX, Index - 1, ByVal 0&)
            arrayBreaks(1, Index) = SendMessageW(ctrlHwnd, EM_LINELENGTH, arrayBreaks(2, Index), ByVal 0&)
            arrayBreaks(3, Index) = SendMessageW(ctrlHwnd, EM_POSFROMCHAR, arrayBreaks(2, Index), ByVal 0&)
            If Not arrayBreaks(3, Index) = -1& Then arrayBreaks(3, Index) = (arrayBreaks(3, Index) And &H7FFF&)
            arrayBreaks(2, Index) = arrayBreaks(2, Index) * 2
        Next
        If accelKeyPos Then
            arrayBreaks(2, 0) = SendMessageW(ctrlHwnd, &HD6, accelKeyPos - 1, ByVal 0&)
        Else
            arrayBreaks(2, 0) = -1
        End If
    Else
        nrLines = SendMessageA(ctrlHwnd, EM_GETLINECOUNT, 0, ByVal 0&)
        ReDim arrayBreaks(1 To 3, 0 To nrLines)
        For Index = 1 To nrLines
            arrayBreaks(2, Index) = SendMessageA(ctrlHwnd, EM_LINEINDEX, Index - 1, ByVal 0&)
            arrayBreaks(1, Index) = SendMessageA(ctrlHwnd, EM_LINELENGTH, arrayBreaks(2, Index), ByVal 0&)
            arrayBreaks(3, Index) = SendMessageA(ctrlHwnd, EM_POSFROMCHAR, arrayBreaks(2, Index), ByVal 0&)
            If Not arrayBreaks(3, Index) = -1& Then arrayBreaks(3, Index) = (arrayBreaks(3, Index) And &H7FFF&)
            arrayBreaks(2, Index) = arrayBreaks(2, Index) * 2
        Next
        If accelKeyPos Then
            arrayBreaks(2, 0) = SendMessageA(ctrlHwnd, &HD6, accelKeyPos - 1, ByVal 0&)
        Else
            arrayBreaks(2, 0) = -1
        End If
    End If
    If arrayBreaks(2, 0) > -1 Then
        arrayBreaks(1, 0) = (arrayBreaks(2, 0) And &H7FFF)
        arrayBreaks(2, 0) = SendMessageA(ctrlHwnd, EM_LINEFROMCHAR, accelKeyPos, ByVal 0&) + 1
    End If
    GetLineBreaks = nrLines
    
End Function

Private Sub DoSample()

    Dim tJust As Long, tAlign As Long
    Dim lValue As Long
    Dim sText As String
    Dim lFlags As Long
    
    Select Case True    ' justification
    Case optJustify(0): lFlags = ptxLEFT
    Case optJustify(1): lFlags = ptxRIGHT
    Case Else: lFlags = ptxCENTER
    End Select
    
    Select Case True    ' alignment
    Case optJustify(3): lFlags = lFlags Or ptxTOP
    Case optJustify(4): lFlags = lFlags Or ptxBOTTOM
    Case Else: lFlags = lFlags Or ptxVCENTER
    End Select
    
    ' right to left?
    If chkFlags(0) Then lFlags = lFlags Or ptxRTLREADING
    ' process as caption (hot keys, double &&, etc) or not?
    If chkFlags(1) = 0 Then lFlags = lFlags Or ptxNOPREFIX
    ' if processing as caption, hide underscore on hotkey?
    If chkFlags(2) Then lFlags = lFlags Or ptxHIDEPREFIX
    ' see if clipping is requested
    If chkFlags(3) = 0 Then lFlags = lFlags Or ptxNOCLIP
    ' see if single line is requested
    If chkFlags(4) Then lFlags = lFlags Or ptxSINGLELINE
    
    If optUnicode(1) = 0 Then   ' use VB's text box: ANSI strings
        sText = Text1.Text
    Else                        ' use API's text box: UNICODE
        lValue = SendMessageW(tUnicodeHWND, WM_GETTEXTLENGTH, 0, ByVal 0&)
        sText = String$(lValue, 0)
        SendMessageW tUnicodeHWND, WM_GETTEXT, lValue + 1, ByVal StrPtr(sText)
    End If
    
    ' The gray bounding rects you see in the picture box are shape controls
    ' The offsets below are to prevent the text from printing under the
    ' shape control borders
        
    Picture1.Cls
    PrintText sText, Picture1.hdc, BoundsH.Left + 1, BoundsH.Top + 1, BoundsH.Width - 2, BoundsH.Height - 2, Angle0, lFlags
    PrintText sText, Picture1.hdc, Bounds90.Left + 1, Bounds90.Top + 1, Bounds90.Width - 2, Bounds90.Height - 2, Angle90, lFlags
    PrintText sText, Picture1.hdc, Bounds270.Left + 1, Bounds270.Top + 1, Bounds270.Width - 2, Bounds270.Height - 2, Angle270, lFlags
    Picture1.Refresh
    
End Sub

Private Sub Form_Load()

    Dim Index As Long
    
    ' setup text1
    Text1.Text = "Now is the time for all good men to come to the aid of their &country."
    ' make all of our controls same backcolor as form
    Me.BackColor = Picture1.BackColor
    For Index = optJustify.LBound To optJustify.UBound
        optJustify(Index).BackColor = Me.BackColor
    Next
    For Index = Label1.LBound To Label1.UBound
        Label1(Index).BackColor = Me.BackColor
    Next
    For Index = chkFlags.LBound To chkFlags.UBound
        chkFlags(Index).BackColor = Me.BackColor
    Next
    For Index = optUnicode.LBound To optUnicode.UBound
        optUnicode(Index).BackColor = Me.BackColor
    Next
    Frame1.BackColor = Me.BackColor

    ' here we will attempt to create a Unicode Textbox
    Const ES_SUNKEN As Long = &H4000
    Const ES_AUTOVSCROLL As Long = &H40&
    Const WS_CHILD As Long = &H40000000
    Const WS_VISIBLE As Long = &H10000000
    Const WS_EX_RIGHTSCROLLBAR As Long = &H0&
    Const WS_EX_CLIENTEDGE As Long = &H200&
    Const WS_VSCROLL As Long = &H200000
    Const WM_GETFONT As Long = &H31

    ' to set API textbox backcolor, we need to give it a parent that has the
    ' backcolor set. The edit control gets its backcolor from the host container,
    ' otherwise we'd have to subclass it to change bkg color
    picUnicodeHolder.BackColor = Text1.BackColor
    
    ' create it & test for success
    tUnicodeHWND = CreateWindowExW(WS_EX_RIGHTSCROLLBAR Or WS_EX_CLIENTEDGE, StrPtr("Edit"), 0, _
        ES_MULTILINE Or WS_CHILD Or WS_VISIBLE Or ES_AUTOVSCROLL Or WS_VSCROLL, _
        0, 0, Text1.Width, Text1.Height, picUnicodeHolder.hWnd, 0, App.hInstance, ByVal 0)
    
    If tUnicodeHWND = 0 Then            ' failure
        optUnicode(1).Enabled = False
    Else
        Label2.Visible = False          ' hide unicode compatibility message
        ' now set its font to text1's font & give it a default text
        SendMessageW tUnicodeHWND, WM_SETFONT, SendMessageW(Text1.hWnd, WM_GETFONT, 0, ByVal 0&), ByVal 0&
        SendMessageW tUnicodeHWND, WM_SETTEXT, 0&, ByVal StrPtr("Unicode Text Box - Copy && Paste as needed")
    End If
    Set Picture1.Font = Text1.Font  ' set picture1's font to text1's font
    
    Show
    Call DoSample               ' show initial results
End Sub

Private Sub Command1_Click()    ' refresh
    DoSample
End Sub

Private Sub optJustify_Click(Index As Integer)
    Call DoSample
End Sub

Private Sub optUnicode_Click(Index As Integer)
    Call DoSample
End Sub

Private Sub chkFlags_Click(Index As Integer)
    DoSample
End Sub

