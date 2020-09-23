Attribute VB_Name = "Module1"
'Make A Gradient Title Bar

'You can download our freeware Gradient Title Bar OCX Control at ActiveX Section.
'Add a module to your project.
'Insert the following code to your module:

Public GradForceColors As Boolean
Public GradVerticalGradient As Boolean
Public GradForcedText As Long, GradForcedTextA As Long
Public GradForcedFirst As Long, GradForcedSecond As Long
Public GradForcedFirstA As Long, GradForcedSecondA As Long

Dim GradhWnd As Long, GradIcon As Long
Dim DrawDC As Long, tmpDC As Long
Dim hRgn As Long
Dim tmpGradFont As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" _
(ByVal hWnd As Long, ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32" Alias "SetPropA" _
(ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" _
(ByVal hWnd As Long, ByVal lpString As String) As Long

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
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Declare Function SystemParametersInfo Lib "user32" Alias _
"SystemParametersInfoA" (ByVal uAction As Long, _
ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const SPI_GETNONCLIENTMETRICS = 41

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

Dim CaptionFont As LOGFONT

Private Declare Function CreateFontIndirect Lib "gdi32" Alias _
"CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) _
As Long

Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) _
As Long

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" _
(ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Private Const IDC_SIZENS = 32645&
Private Const IDC_SIZEWE = 32644&
Private Const IDC_SIZENWSE = 32642&
Private Const IDC_SIZENESW = 32643&

Private Declare Function GetWindowText Lib "user32" Alias _
"GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, _
ByVal cch As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
lpRect As RECT) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000
Private Const WS_DLGFRAME = &H400000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_POPUP = &H80000000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_VISIBLE = &H10000000

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
ByVal hDC As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function OffsetClipRgn Lib "gdi32" (ByVal hDC As Long, _
ByVal x As Long, ByVal Y As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
ByVal x As Long, ByVal Y As Long) As Long

Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, _
ByVal cxWidth As Long, ByVal cyWidth As Long, _
ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
ByVal diFlags As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
(ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, _
lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_END_ELLIPSIS = &H8000&

Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, _
ByVal hRgn As Long) As Long

Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, _
ByVal hRgn As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
ByVal x As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) _
As Long

Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, _
ByVal nBkMode As Long) As Long

Private Const TRANSPARENT = 1

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, _
ByVal crColor As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" _
(ByVal nIndex As Long) As Long

Private Const SM_CXBORDER = 5
Private Const SM_CXDLGFRAME = 7
Private Const SM_CXFRAME = 32
Private Const SM_CXICON = 11
Private Const SM_CXSMSIZE = 30
Private Const SM_CYBORDER = 6
Private Const SM_CYCAPTION = 4
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYFRAME = 33
Private Const SM_CYICON = 12
Private Const SM_CYMENU = 15
Private Const SM_CYSMSIZE = 31

Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, _
lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) _
As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
(ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function CreateRectRgnIndirect Lib "gdi32" _
(lpRect As RECT) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, _
ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
ByVal Y2 As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" _
(ByVal crColor As Long) As Long

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, _
lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Const DFC_CAPTION = 1
Private Const DFCS_CAPTIONRESTORE = &H3
Private Const DFCS_CAPTIONMIN = &H1
Private Const DFCS_CAPTIONMAX = &H2
Private Const DFCS_CAPTIONHELP = &H4
Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_INACTIVE = &H100
Private Const WM_SIZE = &H5
Private Const WM_SETCURSOR = &H20
Private Const WM_GETICON = &H7F
Private Const WM_SETICON = &H80
Private Const WM_NCACTIVATE = &H86
Private Const WM_MDIACTIVATE = &H222
Private Const WM_KILLFOCUS = &H8
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_MDIGETACTIVE = &H229
Private Const MA_ACTIVATE = 1
Private Const WM_SETTEXT = &HC
Private Const WM_NCPAINT = &H85
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCRBUTTONDOWN = &HA4
Private Const WM_SYSCOMMAND = &H112
Private Const WM_INITMENUPOPUP = &H117
Private Const SC_MOUSEMENU = &HF090&
Private Const SC_MOVE = &HF010&
Private Const HTCAPTION = 2
Private Const HTSYSMENU = 3
Private Const HTLEFT = 10
Private Const HTRIGHT = 11
Private Const HTTOP = 12
Private Const HTTOPLEFT = 13
Private Const HTTOPRIGHT = 14
Private Const HTBOTTOM = 15
Private Const HTBOTTOMLEFT = 16
Private Const HTBOTTOMRIGHT = 17

Private Function LoWord(LongIn As Long) As Integer
    If (LongIn And &HFFFF&) > &H7FFF Then
        LoWord = (LongIn And &HFFFF&) - &H10000
    Else
        LoWord = LongIn And &HFFFF&
    End If
End Function

Private Sub GetColors(IsActive As Boolean, LColor As Long, RColor As Long)
    If IsActive Then
        If GradForceColors Then
            LColor = GradForcedFirst
            RColor = GradForcedSecond
        Else
            LColor = vbBlack
            RColor = GetSysColor(COLOR_ACTIVECAPTION)
        End If
    Else
        If GradForceColors Then
            LColor = GradForcedFirstA
            RColor = GradForcedSecondA
        Else
            LColor = vbBlack
            RColor = GetSysColor(COLOR_INACTIVECAPTION)
        End If
    End If
End Sub

Public Sub GradientGetCapsFont()
    Dim NCM As NONCLIENTMETRICS
    Dim lfNew As LOGFONT
    NCM.cbSize = Len(NCM)
    Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
    CaptionFont = NCM.lfCaptionFont
End Sub

Private Sub GetCaptionRect(hWnd As Long, rct As RECT)
    Dim XBorder As Long
    Dim fStyle As Long
    Dim YHeight As Long
    YHeight = GetSystemMetrics(SM_CYCAPTION)
    fStyle = GetWindowLong(hWnd, GWL_STYLE)
    Select Case fStyle And &H80
        Case &H80
            XBorder = GetSystemMetrics(SM_CXDLGFRAME)
        Case Else
            XBorder = GetSystemMetrics(SM_CXFRAME)
    End Select
    rct.Left = XBorder
    rct.Right = XBorder
    rct.Top = XBorder
    rct.Bottom = rct.Top + YHeight - 1
End Sub

Private Sub GradateColors(Colors() As Long, ByVal Color1 As Long, ByVal Color2 As Long)
    Dim i As Long
    Dim dblR As Double, dblG As Double, dblB As Double
    Dim addR As Double, addG As Double, addB As Double
    Dim bckR As Double, bckG As Double, bckB As Double
    dblR = CDbl(Color1 And &HFF)
    dblG = CDbl(Color1 And &HFF00&) / 255
    dblB = CDbl(Color1 And &HFF0000) / &HFF00&
    bckR = CDbl(Color2 And &HFF&)
    bckG = CDbl(Color2 And &HFF00&) / 255
    bckB = CDbl(Color2 And &HFF0000) / &HFF00&
    addR = (bckR - dblR) / UBound(Colors)
    addG = (bckG - dblG) / UBound(Colors)
    addB = (bckB - dblB) / UBound(Colors)
    For i = 0 To UBound(Colors)
        dblR = dblR + addR
        dblG = dblG + addG
        dblB = dblB + addB
        If dblR > 255 Then dblR = 255
        If dblG > 255 Then dblG = 255
        If dblB > 255 Then dblB = 255
        If dblR < 0 Then dblR = 0
        If dblG < 0 Then dblG = 0
        If dblG < 0 Then dblB = 0
        Colors(i) = RGB(dblR, dblG, dblB)
    Next
End Sub

Private Function DrawGradient(ByVal Color1 As Long, ByVal Color2 As Long) As Long
    Dim i As Long
    Dim DestWidth As Long, DestHeight As Long
    Dim StartPnt As Long, EndPnt As Long
    Dim PixelStep As Long, XBorder As Long
    Dim WndRect As RECT
    Dim OldFont As Long
    Dim fStyle As Long, fText As String
    Dim SMSize As Long, SMSizeY As Long
    On Error Resume Next
    SMSize = GetSystemMetrics(SM_CXSMSIZE)
    SMSizeY = GetSystemMetrics(SM_CYSMSIZE)
    GetWindowRect GradhWnd, WndRect
    With WndRect
        DestWidth = .Right - .Left
    End With
    DestHeight = GetSystemMetrics(SM_CYCAPTION)
    fText = Space$(255)
    Call GetWindowText(GradhWnd, fText, 255)
    fText = Trim$(fText)
    fStyle = GetWindowLong(GradhWnd, GWL_STYLE)
    Select Case fStyle And &H80
        Case &H80
            XBorder = GetSystemMetrics(SM_CXDLGFRAME)
            DestWidth = (DestWidth - XBorder)
        Case Else
            XBorder = GetSystemMetrics(SM_CXFRAME)
            DestWidth = DestWidth - XBorder
    End Select
    StartPnt = XBorder
    EndPnt = XBorder + DestWidth - 4
    Dim rct As RECT
    Dim hBr As Long
    With rct
        If Not GradVerticalGradient Then
            PixelStep = DestWidth \ 8
            ReDim Colors(PixelStep) As Long
            GradateColors Colors(), Color1, Color2
            .Top = XBorder
            .Left = XBorder
            .Right = XBorder + (DestWidth \ PixelStep)
            .Bottom = (XBorder + DestHeight - 1)
            If (fStyle And &H80) = &H80 Then EndPnt = EndPnt + 1
            For i = 0 To PixelStep - 1
                hBr = CreateSolidBrush(Colors(i))
                FillRect DrawDC, rct, hBr
                DeleteObject hBr
                OffsetRect rct, (DestWidth \ PixelStep), 0
                If i = PixelStep - 2 Then .Right = EndPnt
            Next
        Else
            PixelStep = DestHeight \ 1
            ReDim Colors(PixelStep) As Long
            GradateColors Colors(), Color2, Color1
            .Top = XBorder
            .Left = XBorder
            If (fStyle And &H80) = &H80 Then
                .Right = (XBorder * 2) + DestWidth + 2
            Else
                .Right = (XBorder * 2) + DestWidth
            End If
            .Bottom = XBorder + (DestHeight \ PixelStep)
            For i = 0 To PixelStep - 1
                hBr = CreateSolidBrush(Colors(i))
                FillRect DrawDC, rct, hBr
                DeleteObject hBr
                OffsetRect rct, 0, (DestHeight \ PixelStep)
                If i = PixelStep - 2 Then .Bottom = XBorder + (DestHeight - 1)
                .Bottom = XBorder + (DestHeight - 1)
            Next
        End If
        .Top = XBorder
        If GradIcon <> 0 Then
            .Left = XBorder + SMSize + 2
            DrawIconEx DrawDC, XBorder + 1, XBorder + 1, GradIcon, _
            SMSize - 2, SMSize - 2, ByVal 0&, ByVal 0&, 2
        Else
            .Left = XBorder
        End If
        tmpGradFont = CreateFontIndirect(CaptionFont)
        OldFont = SelectObject(DrawDC, tmpGradFont)
        SetBkMode DrawDC, TRANSPARENT
        If GradForceColors Then
            If Color1 = GradForcedFirst Then
                SetTextColor DrawDC, GradForcedText
            Else
                SetTextColor DrawDC, GradForcedTextA
            End If
        Else
            If Color2 = GetSysColor(COLOR_ACTIVECAPTION) Then
                SetTextColor DrawDC, GetSysColor(COLOR_CAPTIONTEXT)
            Else
                SetTextColor DrawDC, GetSysColor(COLOR_INACTIVECAPTIONTEXT)
            End If
        End If
        .Left = .Left + 2
        .Right = .Right - 10
        DrawText DrawDC, fText, Len(fText) - 1, rct, _
        DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER
        SelectObject DrawDC, OldFont
        DeleteObject tmpGradFont
        tmpGradFont = 0
        Dim frct As RECT
        If (fStyle And WS_SYSMENU) = WS_SYSMENU Then
            Dim CurMaxPic As Long
            If IsZoomed(GradhWnd) Then
                CurMaxPic = DFCS_CAPTIONRESTORE
            Else
                CurMaxPic = DFCS_CAPTIONMAX
            End If
            With frct
                .Right = DestWidth - 2
                .Left = .Right - SMSize + 2
                .Top = XBorder + 2
                .Bottom = .Top + (DestHeight - 5)
            End With
            DrawFrameControl DrawDC, frct, DFC_CAPTION, DFCS_CAPTIONCLOSE
            OffsetRect frct, -(SMSize), 0
            If (fStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX And _
            (fStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then
                DrawFrameControl DrawDC, frct, DFC_CAPTION, CurMaxPic
                OffsetRect frct, -(SMSize) + 2, 0
                DrawFrameControl DrawDC, frct, DFC_CAPTION, DFCS_CAPTIONMIN
            ElseIf (fStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX Then
                DrawFrameControl DrawDC, frct, DFC_CAPTION, CurMaxPic
                OffsetRect frct, -(SMSize) + 2, 0
                DrawFrameControl DrawDC, frct, DFC_CAPTION, DFCS_CAPTIONMIN _
                Or DFCS_INACTIVE
            ElseIf (fStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then
                DrawFrameControl DrawDC, frct, DFC_CAPTION, CurMaxPic Or DFCS_INACTIVE
                OffsetRect frct, -(SMSize) + 2, 0
                DrawFrameControl DrawDC, frct, DFC_CAPTION, DFCS_CAPTIONMIN
            End If
        End If
        .Left = XBorder
        .Right = .Right + 12
        If tmpDC <> 0 Then
            BitBlt tmpDC, .Left, .Top, .Right - .Left - 10, .Bottom - .Top, _
            DrawDC, .Left, .Top, vbSrcCopy
            ExcludeClipRect tmpDC, XBorder, XBorder, DestWidth, _
            XBorder + (DestHeight - 1)
        End If
        '.Right - .Left - 8
    End With
End Function

Public Function GradientCallback(ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

    Dim OldGradProc As Long
    Dim OldBMP As Long, NewBMP As Long
    Dim rcWnd As RECT
    Dim tmpFrm As Form
    Dim tmpCol1 As Long, tmpCol2 As Long
    Static GettingIcon As Boolean
    GradhWnd = hWnd
    OldGradProc = GetProp(GradhWnd, "OldMeProc")
    
    If Not GettingIcon Then
        GettingIcon = True
        GradIcon = SendMessage(hWnd, WM_GETICON, 0, ByVal 0&)
        GettingIcon = False
    End If
    
    Select Case wMsg
        Case WM_NCACTIVATE, WM_MDIACTIVATE, WM_KILLFOCUS, WM_MOUSEACTIVATE
            GetWindowRect GradhWnd, rcWnd
            tmpDC = GetWindowDC(GradhWnd)
            DrawDC = CreateCompatibleDC(tmpDC)
            NewBMP = CreateCompatibleBitmap(tmpDC, rcWnd.Right - rcWnd.Left, 50)
            OldBMP = SelectObject(DrawDC, NewBMP)
            With rcWnd
                hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
                SelectClipRgn tmpDC, hRgn
                OffsetClipRgn tmpDC, -.Left, -.Top
            End With
            If wMsg = WM_KILLFOCUS And GetParent(GradhWnd) <> 0 Then
                GetColors False, tmpCol1, tmpCol2
            ElseIf wMsg = WM_NCACTIVATE And wParam And _
                (GetParent(GradhWnd) = 0) Then
                GetColors True, tmpCol1, tmpCol2
            ElseIf wMsg = WM_NCACTIVATE And wParam = 0 And _
                (GetParent(GradhWnd) = 0) Then
                GetColors False, tmpCol1, tmpCol2
            ElseIf wParam = GradhWnd And GetParent(GradhWnd) <> 0 Then
                GetColors False, tmpCol1, tmpCol2
            ElseIf SendMessage(GetParent(GradhWnd), WM_MDIGETACTIVE, _
                0, 0) = GradhWnd Then
                GetColors True, tmpCol1, tmpCol2
            ElseIf GetActiveWindow() = GradhWnd Then
                GetColors True, tmpCol1, tmpCol2
            Else
                GetColors False, tmpCol1, tmpCol2
            End If
            DrawGradient tmpCol1, tmpCol2
            SelectObject DrawDC, OldBMP
            DeleteObject NewBMP
            DeleteDC DrawDC
            OffsetClipRgn tmpDC, rcWnd.Left, rcWnd.Top
            GetClipRgn tmpDC, hRgn
            If wMsg = WM_MOUSEACTIVATE Then
                GradientCallback = MA_ACTIVATE
            Else
                GradientCallback = 1
            End If
            ReleaseDC GradhWnd, tmpDC
            DeleteObject hRgn
            tmpDC = 0
            Exit Function
        
        Case WM_SETTEXT, WM_NCPAINT, WM_NCLBUTTONDOWN, _
            WM_NCRBUTTONDOWN, WM_SYSCOMMAND, WM_INITMENUPOPUP
            GetWindowRect GradhWnd, rcWnd
            tmpDC = GetWindowDC(GradhWnd)
            DrawDC = CreateCompatibleDC(tmpDC)
            NewBMP = CreateCompatibleBitmap(tmpDC, rcWnd.Right - rcWnd.Left, 50)
            OldBMP = SelectObject(DrawDC, NewBMP)
            With rcWnd
                hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
                SelectClipRgn tmpDC, hRgn
                OffsetClipRgn tmpDC, -.Left, -.Top
            End With
            If (GetActiveWindow() = GradhWnd) Then
                GetColors True, tmpCol1, tmpCol2
            ElseIf SendMessage(GetParent(GradhWnd), WM_MDIGETACTIVE, _
                0, 0) = GradhWnd Then
                GetColors True, tmpCol1, tmpCol2
            Else
                GetColors False, tmpCol1, tmpCol2
            End If
            DrawGradient tmpCol1, tmpCol2
            SelectObject DrawDC, OldBMP
            DeleteObject NewBMP
            DeleteDC DrawDC
            OffsetClipRgn tmpDC, rcWnd.Left, rcWnd.Top
            GetClipRgn tmpDC, hRgn
            GradientCallback = CallWindowProc(OldGradProc, hWnd, WM_NCPAINT, _
            hRgn, lParam)
            ReleaseDC GradhWnd, tmpDC
            DeleteObject hRgn
            tmpDC = 0
            If wMsg = (WM_NCLBUTTONDOWN And wParam <> HTSYSMENU And _
                wParam <> HTCAPTION) Or wMsg = _
                (WM_SYSCOMMAND And Not (wParam = SC_MOUSEMENU)) Then
                GetCaptionRect GradhWnd, rcWnd
                ExcludeClipRect tmpDC, rcWnd.Left, rcWnd.Top, _
                rcWnd.Right, rcWnd.Bottom
            ElseIf wMsg = WM_NCLBUTTONDOWN And wParam = HTCAPTION Then
                If IsZoomed(GradhWnd) = 0 Then
                    GradientCallback = SendMessage(GradhWnd, WM_SYSCOMMAND, _
                    SC_MOVE, ByVal 0&)
                End If
                Exit Function
            Else
                Exit Function
            End If
        Case WM_SIZE
            If hWnd = GradhWnd Then
                SendMessage GradhWnd, WM_NCPAINT, 0, 0
            End If
        Case WM_SETCURSOR
        Select Case LoWord(lParam)
            Case HTTOP, HTBOTTOM
                SetCursor LoadCursor(ByVal 0&, IDC_SIZENS)
            Case HTLEFT, HTRIGHT
                SetCursor LoadCursor(ByVal 0&, IDC_SIZEWE)
            Case HTTOPLEFT, HTBOTTOMRIGHT
                SetCursor LoadCursor(ByVal 0&, IDC_SIZENWSE)
            Case HTTOPRIGHT, HTBOTTOMLEFT
                SetCursor LoadCursor(ByVal 0&, IDC_SIZENESW)
            Case Else
                GoTo JustCallBack
        End Select
        GradientCallback = 1
        Exit Function
    End Select
JustCallBack:
    GradientCallback = CallWindowProc(OldGradProc, hWnd, wMsg, wParam, lParam)
End Function

Public Sub GradientForm(frm As Form)
    SetBarColours
    Dim tmpProc As Long
    tmpProc = SetWindowLong(frm.hWnd, GWL_WNDPROC, _
    AddressOf GradientCallback)
    SetProp frm.hWnd, "OldMeProc", tmpProc
End Sub

Public Sub GradientReleaseForm(frm As Form)
    Dim tmpProc As Long
    tmpProc = GetProp(frm.hWnd, "OldMeProc")
    RemoveProp frm.hWnd, "OldMeProc"
    If tmpProc = 0 Then Exit Sub
    SetWindowLong frm.hWnd, GWL_WNDPROC, tmpProc
End Sub

Public Sub SetBarColours()
    GradForceColors = True
    'Replace the 'True' below with 'False' if you want that the gradient will
    'be drawn  horizonally.
    GradVerticalGradient = True
    'Set colors for active caption
    GradForcedText = GetSysColor(COLOR_CAPTIONTEXT)
    'Replace the two color values below to change the active title bar color
    GradForcedFirst = GetSysColor(COLOR_ACTIVECAPTION)
    GradForcedSecond = GetSysColor(COLOR_SCROLLBAR)
    'Set colors for Inactive caption
    GradForcedTextA = GetSysColor(COLOR_INACTIVECAPTIONTEXT)
    'Replace the two color values below to change the inactive title bar color
    GradForcedFirstA = GetSysColor(COLOR_INACTIVECAPTION)
    GradForcedSecondA = GetSysColor(COLOR_SCROLLBAR)
    GradientGetCapsFont
End Sub
