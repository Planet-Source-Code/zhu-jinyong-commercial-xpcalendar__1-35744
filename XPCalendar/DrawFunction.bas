Attribute VB_Name = "mdrawFunction"
'Email Me: zhujy@samling.com.my
'Welcome to visit our company WebSite at:
'www.Samling.com.my
'www.Samling.com.cn
'Samling Group---One of a leading and largest WoodBased Industry company in Asia
'Copyright © 2001-2002 by Zhu JinYong from People Republic of China
'Thanks to Abdul Gafoor.GK ,BadSoft and Carles.P.V.

Option Explicit
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal fuFlags As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function InflateRect Lib "user32" _
                          (lpRect As RECT, _
                          ByVal dx As Long, ByVal dy As Long) As Long
Public Declare Function RoundRect Lib "gdi32" _
                        (ByVal hdc As Long, _
                        ByVal X1 As Long, ByVal Y1 As Long, _
                        ByVal X2 As Long, ByVal Y2 As Long, _
                        ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SelectObject Lib "gdi32" _
                        (ByVal hdc As Long, _
                        ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SetTextColor Lib "gdi32" _
                        (ByVal hdc As Long, _
                        ByVal crColor As Long) As Long

Public Declare Function CreatePen Lib "gdi32" _
                        (ByVal nPenStyle As Long, _
                        ByVal nWidth As Long, _
                        ByVal crColor As Long) As Long

Public Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" _
                        (ByVal hdc As Long, _
                        pVertex As TRIVERTEX, _
                        ByVal dwNumVertex As Long, _
                        pMesh As GRADIENT_RECT, _
                        ByVal dwNumMesh As Long, _
                        ByVal dwMode As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" _
                          (ByVal hdc As Long, _
                          ByVal nStretchMode As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, _
                          ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, _
                          ByVal cxWidth As Long, ByVal cyWidth As Long, _
                          ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
                          ByVal diFlags As Long) As Long
Public Type TRIVERTEX
    X As Long
    Y As Long
    R As Integer
    G As Integer
    B As Integer
    Alpha As Integer
End Type

Public Type RGB
    R As Integer
    G As Integer
    B As Integer
End Type

Public Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long  '
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Const PS_SOLID = 0
Const COLOR_WINDOW = 5

'/* Bitmap Manipulate */
Const SRCAND = &H8800C6
Const SRCCOPY = &HCC0020
Const SRCERASE = &H440328
Const SRCINVERT = &H660046
Const SRCPAINT = &HEE0086

' /* State type */
Const DSS_NORMAL = &H0
Const DSS_UNION = &H10
Const DSS_DISABLED = &H20
Const DSS_MONO = &H80
Const DSS_RIGHT = &H8000

' Standard GDI draw icon function:
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = &H3
Const DI_COMPAT = &H4
Const DI_DEFAULTSIZE = &H8

'
'## Paint item back area (Standard)
'
Public Sub DrawBack(ByVal hdc As Long, _
                    R As RECT, _
                    ByVal Color As Long)

  Dim lBrush As Long, lOldBr As Long
  
    lBrush = CreateSolidBrush(Color)
    lOldBr = SelectObject(hdc, lBrush)
    Call FillRect(hdc, R, lBrush)
    SelectObject hdc, lOldBr
    DeleteObject lBrush
    DeleteObject lOldBr
    
End Sub

Public Sub SetColors()

  '## Item back color [Normal]

    cBackNrm = GetLngColor(m_BackNormal)

    '## Item back color [Selected]
    cBackSel = GetLngColor(m_BackSelected)

    '## Item back color 1 [Selected (Gradient style)]
    cGrad1 = GetRGBColors(GetLngColor(m_BackSelectedG1))

    '## Item back color 2 [Selected (Gradient style)]
    cGrad2 = GetRGBColors(GetLngColor(m_BackSelectedG2))

    '## Item box border color [Selected (Box style)]
    cBox = GetLngColor(m_BoxBorder)

    '## Item font color [Normal]
    cFontNrm = GetLngColor(m_FontNormal)

    '## Item font color [Selected]
    cFontSel = GetLngColor(m_FontSelected)

End Sub

'## GetLngColor ------------------------------------------------------------------
Public Function GetLngColor(Color As Long) As Long

    If Color And &H80000000 Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
      Else
        GetLngColor = Color
    End If

End Function

'## GetRGBColors -----------------------------------------------------------------
Public Function GetRGBColors(Color As Long) As RGB

  Dim HexColor As String

    HexColor = String(6 - Len(Hex(Color)), "0") & Hex(Color)
    GetRGBColors.R = "&H" & Mid(HexColor, 5, 2) & "00"
    GetRGBColors.G = "&H" & Mid(HexColor, 3, 2) & "00"
    GetRGBColors.B = "&H" & Mid(HexColor, 1, 2) & "00"

End Function

Private Function COLOR_UniColor(ColorVal As Long) As Long

  ' Returns Color as long, accepts SystemColorConstants

    COLOR_UniColor = ColorVal
    If ColorVal > &HFFFFFF Or ColorVal < 0 Then COLOR_UniColor = GetSysColor(ColorVal And &HFFFFFF)

End Function

'
'## Paint item back area (Gradient);Thank Carlves.P.v.
'
Public Sub DrawBackGrad(ByVal hdc As Long, _
                        R As RECT, _
                        Color1 As RGB, _
                        Color2 As RGB, _
                        Direction As Long)

  Dim V(1) As TRIVERTEX
  Dim GRct As GRADIENT_RECT

  Dim Ret As Long

    '# from
    With V(0)
        .X = R.Left
        .Y = R.Top
        .R = Color1.R
        .G = Color1.G
        .B = Color1.B
        .Alpha = 0
    End With
    '# to
    With V(1)
        .X = R.Right
        .Y = R.Bottom
        .R = Color2.R
        .G = Color2.G
        .B = Color2.B
        .Alpha = 0
    End With

    GRct.UpperLeft = 0
    GRct.LowerRight = 1

    Ret = GradientFillRect(hdc, V(0), 2, GRct, 1, Direction)

End Sub

Public Sub DrawBitmap(ByVal hdc As Long, Picture As StdPicture, _
                      ByVal PicSizeWidth As Long, _
                      ByVal PicSizeHeight As Long, _
                      ByVal OriginalPicSizeW As Long, _
                      ByVal OriginalPicSizeH As Long, _
                      EnabledPic As Byte, _
                      CurPictRECT As RECT, _
                      Optional Mode As Integer = 1, _
                      Optional MaskColor As OLE_COLOR, _
                      Optional AsShadow As Byte = 0)

  Dim DC1 As Long
  Dim BM1 As Long
  Dim DC2 As Long
  Dim BM2 As Long
  Dim UZUN1 As Long
  Dim UZUN2 As Long
  Dim hBrush As Long
  Dim DC3 As Long
  Dim BM3 As Long
    DC1 = CreateCompatibleDC(hdc)
    DC2 = CreateCompatibleDC(hdc)
    BM1 = CreateCompatibleBitmap(hdc, OriginalPicSizeW, OriginalPicSizeH)
    BM2 = CreateCompatibleBitmap(hdc, PicSizeWidth, PicSizeHeight)
    UZUN1 = SelectObject(DC1, BM1)
    UZUN2 = SelectObject(DC2, BM2)

    If EnabledPic = 0 Then 'DISABLED BITMAP
  
        DC3 = CreateCompatibleDC(hdc)
        BM3 = SelectObject(DC3, Picture.Handle)

        SetBkColor DC1, GetSysColor(COLOR_WINDOW) '&HFFFFFF

        DRAWRECT DC1, 0, 0, _
                 OriginalPicSizeW, OriginalPicSizeH, GetSysColor(COLOR_WINDOW), 1

        TransParentPic DC1, DC1, DC3, 0, 0, _
                       OriginalPicSizeW, OriginalPicSizeH, 0, 0, MaskColor

        On Error Resume Next
          If (PicSizeWidth < OriginalPicSizeW Or PicSizeHeight < OriginalPicSizeH) Then
              SetStretchBltMode DC2, &H3
            Else
              SetStretchBltMode DC2, &H0
          End If

          Select Case Mode
            Case 0
              BitBlt DC2, 0, 0, _
                     PicSizeWidth, _
                     PicSizeHeight, _
                     DC1, 0, 0, SRCCOPY
            Case 1
              StretchBlt DC2, 0, 0, _
                         PicSizeWidth, _
                         PicSizeHeight, _
                         DC1, 0, 0, OriginalPicSizeW, OriginalPicSizeH, SRCCOPY

            Case 2
              Call TilePic2DC(DC2, 0, 0, CurPictRECT.Right - CurPictRECT.Left, _
                   CurPictRECT.Bottom - CurPictRECT.Top, _
                   DC1, OriginalPicSizeW, _
                   OriginalPicSizeH)

          End Select

          SelectObject DC2, UZUN2

          If AsShadow = 1 Then 'IF WE ARE TO DRAW SHADOW, DRAW GRAY:
              hBrush = CreateSolidBrush(RGB(176, 176, 176))
              Call DrawState(hdc, hBrush, 0, BM2, 0, CurPictRECT.Left, _
                   CurPictRECT.Top, 0, 0, DSS_DISABLED Or DI_COMPAT)
              DeleteObject hBrush
            Else 'WE ARE DRAWING THE DISABLED APPEARANCE OF THE BITMAP
              Call DrawState(hdc, 0, 0, BM2, 0, CurPictRECT.Left, _
                   CurPictRECT.Top, 0, 0, DSS_DISABLED Or DI_COMPAT)
          End If

          DeleteObject BM3
          DeleteDC DC3
         
        Else 'ENABLED BITMAP

          Call DrawState(DC1, 0, 0, Picture, 0, 0, 0, 0, 0, _
               DSS_NORMAL Or DI_COMPAT)

          If (PicSizeWidth < OriginalPicSizeW Or PicSizeHeight < OriginalPicSizeH) Then
              SetStretchBltMode DC2, &H3
            Else
              SetStretchBltMode DC2, &H0
          End If

          Select Case Mode
            Case 0
              BitBlt DC2, 0, 0, _
                     PicSizeWidth, _
                     PicSizeHeight, _
                     DC1, 0, 0, SRCCOPY

            Case 1
              StretchBlt DC2, 0, 0, _
                         PicSizeWidth, _
                         PicSizeHeight, _
                         DC1, 0, 0, OriginalPicSizeW, OriginalPicSizeH, SRCCOPY

            Case 2 ''~~ Tile BMP inside RECT without transparency
              Call TilePic2DC(DC2, 0, 0, CurPictRECT.Right - CurPictRECT.Left, _
                   CurPictRECT.Bottom - CurPictRECT.Top, _
                   DC1, OriginalPicSizeW, _
                   OriginalPicSizeH)

          End Select
          TransParentPic hdc, hdc, DC2, 0, 0, _
                         PicSizeWidth, PicSizeHeight, _
                         CurPictRECT.Left, CurPictRECT.Top, MaskColor

      End If

      Call DeleteObject(SelectObject(DC1, UZUN1))
      Call DeleteObject(SelectObject(DC2, UZUN2))
      DeleteDC DC1
      DeleteDC DC2
      Call ReleaseDC(0&, hdc)

End Sub

Public Sub DrawPIcon(ByVal hdc As Long, PicIcon As Picture, ByVal PicSizeWidth As Long, ByVal PicSizeHeight As Long, ByVal OriginalPicSizeW As Long, ByVal OriginalPicSizeH As Long, EnabledPic As Byte, CurPictRECT As RECT, _
                     Optional AsShadow As Byte = 0)

  'disabled icons by means of converting them to bitmap first.

    If EnabledPic = 0 Then 'DISABLED ICON
  Dim DC1 As Long
  Dim BM1 As Long
  Dim DC2 As Long
  Dim BM2 As Long
  Dim UZUN1 As Long
  Dim UZUN2 As Long
  Dim hBrush As Long

        DC1 = CreateCompatibleDC(hdc)
        BM1 = CreateCompatibleBitmap(hdc, OriginalPicSizeW, OriginalPicSizeH)

        DC2 = CreateCompatibleDC(hdc)
        BM2 = CreateCompatibleBitmap(hdc, PicSizeWidth, PicSizeHeight)

        UZUN1 = SelectObject(DC1, BM1)
        UZUN2 = SelectObject(DC2, BM2)

        If AsShadow = 1 Then 'IF WE ARE TO DRAW SHADOW, DRAW GRAY:
            hBrush = CreateSolidBrush(RGB(176, 176, 176))
            Call DrawState(DC1, hBrush, 0, PicIcon, 0, 0, 0, 0, 0, _
                 DSS_MONO Or DI_NORMAL)
            DeleteObject hBrush
          Else 'WE ARE DRAWING THE DISABLED APPEARANCE OF THE ICON
            Call DrawState(DC1, 0, 0, PicIcon, 0, 0, 0, 0, 0, _
                 DSS_DISABLED Or DI_NORMAL)
        End If

        If ((CurPictRECT.Right - CurPictRECT.Left) < OriginalPicSizeW Or (CurPictRECT.Bottom - CurPictRECT.Top) < OriginalPicSizeH) Then
            SetStretchBltMode DC2, &H3
          Else
            SetStretchBltMode DC2, &H0
        End If

        StretchBlt DC2, 0, 0, _
                   CurPictRECT.Right - CurPictRECT.Left, _
                   CurPictRECT.Bottom - CurPictRECT.Top, _
                   DC1, 0, 0, OriginalPicSizeW, OriginalPicSizeH, SRCCOPY

        TransParentPic hdc, hdc, DC2, 0, 0, _
                       PicSizeWidth, PicSizeHeight, _
                       CurPictRECT.Left, CurPictRECT.Top, &HFFFFFF

        Call DeleteObject(SelectObject(DC1, UZUN1))
        Call DeleteObject(SelectObject(DC2, UZUN2))
        DeleteDC DC1
        DeleteDC DC2
        Call ReleaseDC(0&, hdc)

      Else 'ENABLED ICON
        DrawIconEx hdc, CurPictRECT.Left, CurPictRECT.Top, _
                   PicIcon.Handle, _
                   CurPictRECT.Right - CurPictRECT.Left, _
                   CurPictRECT.Bottom - CurPictRECT.Top, 0, 0, DI_NORMAL
    End If

End Sub

Public Sub TilePic2DC( _
                      ByVal hDCDestin As Long, _
                      ByVal X As Long, _
                      ByVal Y As Long, _
                      ByVal Width As Long, _
                      ByVal Height As Long, _
                      ByVal hDCSource As Long, _
                      ByVal SrcWidth As Long, _
                      ByVal SrcHeight As Long _
                      )

  Dim lSrcX As Long
  Dim lSrcY As Long
  Dim lSrcStartX As Long
  Dim lSrcStartY As Long
  Dim lSrcStartWidth As Long
  Dim lSrcStartHeight As Long
  Dim lDstX As Long
  Dim lDstY As Long
  Dim lDstWidth As Long
  Dim lDstHeight As Long

    lSrcStartX = (X Mod SrcWidth)
    lSrcStartY = (Y Mod SrcHeight)
    lSrcStartWidth = (SrcWidth - lSrcStartX)
    lSrcStartHeight = (SrcHeight - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY

    lDstY = Y
    lDstHeight = lSrcStartHeight

    Do While lDstY < (Y + Height)
        If (lDstY + lDstHeight) > (Y + Height) Then
            lDstHeight = Y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = X
        lSrcX = lSrcStartX
        Do While lDstX < (X + Width)
            If (lDstX + lDstWidth) > (X + Width) Then
                lDstWidth = X + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            BitBlt hDCDestin, lDstX, lDstY, lDstWidth, lDstHeight, hDCSource, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = SrcWidth
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = SrcHeight
    Loop

End Sub

Private Sub TransParentPic(DestDC As Long, _
                           DestDCTrans As Long, _
                           SrcDC As Long, _
                           SrcRectLeft As Long, SrcRectTop As Long, _
                           SrcRectRight As Long, SrcRectBottom As Long, _
                           DstX As Long, _
                           DstY As Long, _
                           MaskColor As Long)

  Dim nRet As Long, W As Integer, H As Integer
  Dim MonoMaskDC As Long, hMonoMask As Long
  Dim MonoInvDC As Long, hMonoInv As Long
  Dim ResultDstDC As Long, hResultDst As Long
  Dim ResultSrcDC As Long, hResultSrc As Long
  Dim hPrevMask As Long, hPrevInv As Long
  Dim hPrevSrc As Long, hPrevDst As Long
  Dim SrcRect As RECT

    With SrcRect
        .Left = SrcRectLeft
        .Top = SrcRectTop
        .Right = SrcRectRight
        .Bottom = SrcRectBottom
    End With

    W = SrcRectRight - SrcRectLeft '+ 1
    H = SrcRectBottom - SrcRectTop '+ 1

    'create monochrome mask and inverse masks
    MonoMaskDC = CreateCompatibleDC(DestDCTrans)
    MonoInvDC = CreateCompatibleDC(DestDCTrans)
    hMonoMask = CreateBitmap(W, H, 1, 1, ByVal 0&)
    hMonoInv = CreateBitmap(W, H, 1, 1, ByVal 0&)
    hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
    hPrevInv = SelectObject(MonoInvDC, hMonoInv)

    'create keeper DCs and bitmaps
    ResultDstDC = CreateCompatibleDC(DestDCTrans)
    ResultSrcDC = CreateCompatibleDC(DestDCTrans)
    hResultDst = CreateCompatibleBitmap(DestDCTrans, W, H)
    hResultSrc = CreateCompatibleBitmap(DestDCTrans, W, H)
    hPrevDst = SelectObject(ResultDstDC, hResultDst)
    hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)

    'copy src to monochrome mask
  Dim OldBC As Long
    OldBC = SetBkColor(SrcDC, MaskColor)
    nRet = BitBlt(MonoMaskDC, 0, 0, W, H, SrcDC, _
           SrcRect.Left, SrcRect.Top, SRCCOPY)
    MaskColor = SetBkColor(SrcDC, OldBC)

    'create inverse of mask
    nRet = BitBlt(MonoInvDC, 0, 0, W, H, _
           MonoMaskDC, 0, 0, &H330008)

    'get background
    nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
           DestDCTrans, DstX, DstY, SRCCOPY)

    'AND with Monochrome mask
    nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
           MonoMaskDC, 0, 0, SRCAND)

    'get overlapper
    nRet = BitBlt(ResultSrcDC, 0, 0, W, H, SrcDC, _
           SrcRect.Left, SrcRect.Top, SRCCOPY)

    'AND with inverse monochrome mask
    nRet = BitBlt(ResultSrcDC, 0, 0, W, H, _
           MonoInvDC, 0, 0, SRCAND)

    'XOR these two
    nRet = BitBlt(ResultDstDC, 0, 0, W, H, _
           ResultSrcDC, 0, 0, SRCINVERT)

    'output results
    nRet = BitBlt(DestDC, DstX, DstY, W, H, _
           ResultDstDC, 0, 0, SRCCOPY)

    'clean up
    hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
    DeleteObject hMonoMask

    hMonoInv = SelectObject(MonoInvDC, hPrevInv)
    DeleteObject hMonoInv

    hResultDst = SelectObject(ResultDstDC, hPrevDst)
    DeleteObject hResultDst

    hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
    DeleteObject hResultSrc

    DeleteDC MonoMaskDC
    DeleteDC MonoInvDC
    DeleteDC ResultDstDC
    DeleteDC ResultSrcDC

End Sub

Public Sub SetBorderStyle(frm As Form)

  Dim lPen As Long

    Select Case m_CalendarBorderStyle
      Case [Flat]
        Call DrawFlatRect(frm.hdc, 0, 0, _
             frm.ScaleWidth, frm.ScaleHeight, _
             m_CalendarBdFlatBorderColour)

      Case [Raised Thin]
        Call DrawRaisedThinRect(frm.hdc, 0, 0, _
             frm.ScaleWidth, frm.ScaleHeight, _
             m_CalendarBdHighlightColour, _
             m_CalendarBdShadowColour)

      Case [Sunken Thin]
        Call DrawSunkenThinRect(frm.hdc, 0, 0, _
             frm.ScaleWidth, frm.ScaleHeight, _
             m_CalendarBdShadowColour, _
             m_CalendarBdHighlightColour)

      Case [Raised 3D]
        Call DrawRaised3DRect(frm.hdc, 0, 0, _
             frm.ScaleWidth, frm.ScaleHeight, _
             m_CalendarBdHighlightColour, _
             m_CalendarBdShadowDKColour, _
             m_CalendarBdHighlightDKColour, _
             m_CalendarBdShadowColour)

      Case [Sunken 3D]
        Call DrawSunken3DRect(frm.hdc, 0, 0, _
             frm.ScaleWidth, frm.ScaleHeight, _
             m_CalendarBdShadowDKColour, _
             m_CalendarBdHighlightColour, _
             m_CalendarBdShadowColour, _
             m_CalendarBdHighlightDKColour)

      Case [Etched]
        Call DrawEtchedRect(frm.hdc, 0, 0, _
             frm.ScaleWidth, frm.ScaleHeight, _
             m_CalendarBdHighlightColour, _
             m_CalendarBdShadowColour)

      Case [Bump]
        Call DrawBumpRect(frm.hdc, 0, 0, _
             frm.ScaleWidth, frm.ScaleHeight, _
             m_CalendarBdHighlightColour, _
             m_CalendarBdShadowColour)
    End Select

End Sub

Private Function TranslateColour(lColour As Long) As Long

    TranslateColor lColour, 0, TranslateColour

End Function

Public Sub DrawLine(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)

  Dim CurPos As POINTAPI
  Dim Obj As Long, lPen As Long

    lPen = CreatePen(PS_SOLID, 1, Color)
    Obj = SelectObject(hdc, lPen)
    MoveToEx hdc, X1, Y1, CurPos '0 would cause OUT Of MEMORY!!
    LineTo hdc, X2, Y2
    SelectObject hdc, Obj
    DeleteObject lPen

End Sub

Private Sub DRAWRECT(DestHDC As Long, ByVal RectLEFT As Long, _
                     ByVal RectTOP As Long, _
                     ByVal RectRIGHT As Long, ByVal RectBOTTOM As Long, _
                     ByVal COLOR_WINDOW As Long, _
                     Optional FillRectWithColor As Byte = 0)

  Dim MyRect As RECT, lBrush As Long, lOldBr As Long

    lBrush = CreateSolidBrush(COLOR_UniColor(COLOR_WINDOW))
    lOldBr = SelectObject(DestHDC, lBrush)
    With MyRect
        .Left = RectLEFT
        .Top = RectTOP
        .Right = RectRIGHT
        .Bottom = RectBOTTOM
    End With
    If FillRectWithColor = 1 Then FillRect DestHDC, MyRect, lBrush Else FrameRect DestHDC, MyRect, lBrush
    'DeleteObject lBrush
    SelectObject DestHDC, lOldBr
    DeleteObject lBrush
    DeleteObject lOldBr
    
End Sub

Public Function DrawBumpRect(ByVal hdc As Long, _
                             ByVal X1 As Long, ByVal Y1 As Long, _
                             ByVal X2 As Long, ByVal Y2 As Long, _
                             BdHighlightColor As OLE_COLOR, _
                             BdShadowColor As OLE_COLOR)

  Dim CurPos As POINTAPI
  Dim lPen As Long
  Dim lPenOld As Long
  
    MoveToEx hdc, X2, Y1, CurPos
    lPen = CreatePen(0, 0, TranslateColour(BdHighlightColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X1, Y1
    LineTo hdc, X1, Y2 - 1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    lPen = CreatePen(0, 0, TranslateColour(BdShadowColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X2 - 1, Y2 - 1
    LineTo hdc, X2 - 1, Y1 - 1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    MoveToEx hdc, X2 - 2, Y1 + 1, CurPos
    lPen = CreatePen(0, 0, TranslateColour(BdShadowColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X1 + 1, Y1 + 1
    LineTo hdc, X1 + 1, Y2 - 2
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    lPen = CreatePen(0, 0, TranslateColour(BdHighlightColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X2 - 2, Y2 - 2
    LineTo hdc, X2 - 2, Y1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld

End Function

Public Function DrawEtchedRect(ByVal hdc As Long, _
                               ByVal X1 As Long, ByVal Y1 As Long, _
                               ByVal X2 As Long, ByVal Y2 As Long, _
                               BdHighlightColor As OLE_COLOR, _
                               BdShadowColor As OLE_COLOR)

  Dim CurPos As POINTAPI
  Dim lPen As Long
  Dim lPenOld As Long
  
    MoveToEx hdc, X2, Y1, CurPos
    lPen = CreatePen(0, 0, TranslateColour(BdShadowColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X1, Y1
    LineTo hdc, X1, Y2 - 1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    lPen = CreatePen(0, 0, TranslateColour(BdHighlightColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X2 - 1, Y2 - 1
    LineTo hdc, X2 - 1, Y1 - 1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    MoveToEx hdc, X2 - 2, Y1 + 1, CurPos
    lPen = CreatePen(0, 0, TranslateColour(BdHighlightColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X1 + 1, Y1 + 1
    LineTo hdc, X1 + 1, Y2 - 2
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    lPen = CreatePen(0, 0, TranslateColour(BdShadowColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X2 - 2, Y2 - 2
    LineTo hdc, X2 - 2, Y1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
End Function

Public Function DrawSunken3DRect(ByVal hdc As Long, _
                                 ByVal X1 As Long, ByVal Y1 As Long, _
                                 ByVal X2 As Long, ByVal Y2 As Long, _
                                 BdShadowDKColour As OLE_COLOR, _
                                 BdHighlightColor As OLE_COLOR, _
                                 BdShadowColour As OLE_COLOR, _
                                 BdHighlightDKColor As OLE_COLOR)

  Dim CurPos As POINTAPI
  Dim lPen As Long
  Dim lPenOld As Long
  
    MoveToEx hdc, X2, Y1, CurPos
    lPen = CreatePen(0, 0, TranslateColour(BdShadowDKColour))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X1, Y1
    LineTo hdc, X1, Y2 - 1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    lPen = CreatePen(0, 0, TranslateColour(BdHighlightColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X2 - 1, Y2 - 1
    LineTo hdc, X2 - 1, Y1 - 1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    MoveToEx hdc, X2 - 2, Y1 + 1, CurPos
    lPen = CreatePen(0, 0, TranslateColour(BdShadowColour))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X1 + 1, Y1 + 1
    LineTo hdc, X1 + 1, Y2 - 2
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    lPen = CreatePen(0, 0, TranslateColour(BdHighlightDKColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X2 - 2, Y2 - 2
    LineTo hdc, X2 - 2, Y1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld

End Function

Public Function DrawRaised3DRect(ByVal hdc As Long, _
                                 ByVal X1 As Long, ByVal Y1 As Long, _
                                 ByVal X2 As Long, ByVal Y2 As Long, _
                                 BdHighlightColor As OLE_COLOR, _
                                 BdShadowDKColour As OLE_COLOR, _
                                 BdHighlightDKColor As OLE_COLOR, _
                                 BdShadowColour As OLE_COLOR)

  Dim CurPos As POINTAPI
  Dim lPen As Long
  Dim lPenOld As Long
  
    MoveToEx hdc, X2, Y1, CurPos
    lPen = CreatePen(0, 0, TranslateColour(BdHighlightColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X1, Y1
    LineTo hdc, X1, Y2 - 1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    lPen = CreatePen(0, 0, TranslateColour(BdShadowDKColour))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X2 - 1, Y2 - 1
    LineTo hdc, X2 - 1, Y1 - 1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    MoveToEx hdc, X2 - 2, Y1 + 1, CurPos
    lPen = CreatePen(0, 0, TranslateColour(BdHighlightDKColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X1 + 1, Y1 + 1
    LineTo hdc, X1 + 1, Y2 - 2
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    lPen = CreatePen(0, 0, TranslateColour(BdShadowColour))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X2 - 2, Y2 - 2
    LineTo hdc, X2 - 2, Y1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    

End Function

Public Function DrawSunkenThinRect(ByVal hdc As Long, _
                                   ByVal X1 As Long, ByVal Y1 As Long, _
                                   ByVal X2 As Long, ByVal Y2 As Long, _
                                   BdShadowColour As OLE_COLOR, _
                                   BdHighlightColor As OLE_COLOR)

  Dim CurPos As POINTAPI
  Dim lPen As Long
  Dim lPenOld As Long
  
    MoveToEx hdc, X2, Y1, CurPos
    lPen = CreatePen(0, 0, TranslateColour(BdShadowColour))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X1, Y1
    LineTo hdc, X1, Y2 - 1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    lPen = CreatePen(0, 0, TranslateColour(BdHighlightColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X2 - 1, Y2 - 1
    LineTo hdc, X2 - 1, Y1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    

End Function

'Raised Thin

Public Function DrawRaisedThinRect(ByVal hdc As Long, _
                                   ByVal X1 As Long, ByVal Y1 As Long, _
                                   ByVal X2 As Long, ByVal Y2 As Long, _
                                   BdHighlightColor As OLE_COLOR, _
                                   BdShadowColour As OLE_COLOR)

  Dim CurPos As POINTAPI
  Dim lPen As Long
  Dim lPenOld As Long
  
    MoveToEx hdc, X2, Y1, CurPos
    lPen = CreatePen(0, 0, TranslateColour(BdHighlightColor))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X1, Y1
    LineTo hdc, X1, Y2 - 1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
    lPen = CreatePen(0, 0, TranslateColour(BdShadowColour))
    lPenOld = SelectObject(hdc, lPen)
    LineTo hdc, X2 - 1, Y2 - 1
    LineTo hdc, X2 - 1, Y1
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
End Function

Public Function DrawFlatRect(ByVal hdc As Long, _
                             ByVal X1 As Long, ByVal Y1 As Long, _
                             ByVal X2 As Long, ByVal Y2 As Long, _
                             BdFlatBorderColour As OLE_COLOR)

  Dim lPen As Long
  Dim lPenOld As Long
  
    lPen = CreatePen(0, 0, TranslateColour(BdFlatBorderColour))
    lPenOld = SelectObject(hdc, lPen)
    Rectangle hdc, X1, Y1, X2, Y2
    'DeleteObject lPen
    SelectObject hdc, lPenOld
    DeleteObject lPen
    DeleteObject lPenOld
    
End Function

':) Ulli's VB Code Formatter V2.12.7 (7/17/02 9:14:56 AM) 121 + 765 = 886 Lines
