VERSION 5.00
Begin VB.Form frmCalendar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Email Me: zhujy@samling.com.my
'Welcome to visit our company WebSite at:
'www.Samling.com.my
'www.Samling.com.cn
'Samling Group---One of a leading and largest WoodBased Industry company in Asia
'Copyright © 2001-2002 by Zhu JinYong from People Republic of China
'Thanks to Abdul Gafoor.GK ,BadSoft and Carles.P.V.
'Updated on 20/06/2002  bugs fixed,Wrong WeekNumber with different "Regional Settings".
Option Explicit

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
' Window style bit functions:
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                          (ByVal hwnd As Long, ByVal nIndex As Long, _
                          ByVal dwNewLong As Long _
                          ) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                          (ByVal hwnd As Long, ByVal nIndex As Long _
                          ) As Long
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_BORDER = &H800000
Private Const WS_THICKFRAME = &H40000
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

'Module specific variable declarations
Private Type cpCalendarInformation
    'Clr As OLE_COLOR
    Rct As RECT
    'Tip As String
End Type

'Grid dimensions for days
Const GRID_ROWS = 6
Const GRID_COLS = 7

'Private variables

Const m_Margin_Offset = 4
Const m_nInitialT = 0
Const m_nButLabelHeight = 20
Const AlignLeft_Offset = 4
Const AlignRight_Offset = 4

Const BDR_RAISEDINNER = &H4
Const BDR_SUNKENINNER = &H8
Const BDR_RAISEDOUTER = &H1
Const BDR_SUNKENOUTER = &H2
Const BF_BOTTOM = &H8
Const BF_FLAT = &H4000      ' For flat rather than 3D borders
Const BF_LEFT = &H1
Const BF_MONO = &H8000      ' For monochrome borders.
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const EDGE_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
Const EDGE_SUNKEN = BDR_SUNKENOUTER Or BDR_SUNKENINNER
Const EDGE_XPFlat = BDR_RAISEDINNER Or BDR_SUNKENOUTER
Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Const DT_CENTER = &H1
Const DT_VCENTER = &H4
Const DT_SINGLELINE = &H20
Const DT_CALCRECT = &H400
Const DT_BOTTOM = &H8
Const DT_LEFT = &H0
Const DT_NOCLIP = &H100
Const DT_TOP = &H0
Const GRADIENT_FILL_RECT_H  As Long = &H0
Const GRADIENT_FILL_RECT_V  As Long = &H1
Const PS_SOLID = 0

Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const HWND_NOTOPMOST = -2

Dim IsFocus(2)               As Boolean
Dim IsMouseDown(2)           As Boolean
Dim IsOutsideCalendarStamp   As Boolean
Dim IsInsideCalendarStamp    As Boolean
Dim Clrs(60)                 As cpCalendarInformation
Dim intfrmWidth              As Integer
Dim intfrmHeight             As Integer
Dim FocusRec                 As RECT
Dim MouseButId               As Integer
Dim MouseDownButId           As Integer
Dim temMouseButIDFocus       As Integer
Dim iNumDaysPerWeek          As Integer
Dim strDayCaption(7)         As String
Dim FrmBkClr                 As Long
Dim ButID                    As Integer
Dim temButID                 As Integer
Dim pl, pt                   As Integer
Dim MonthStart               As Date
Dim MonthEnd                 As Date
Dim NextMonth                As Date
Dim TimeNow                  As Date
Dim PrevMonth                As Date
Dim PrevDays                 As Integer
Dim temSelectedDate          As String
Dim g_Button                 As Integer
Dim g_Shift                  As Integer
Dim g_X                      As Single
Dim g_Y                      As Single
Dim Pt1                      As Integer
Dim Pt2                      As Integer
Dim Pt3                      As Integer
Dim intHighlightDay          As Integer
Dim intHoverDay              As Integer
Dim LPos                     As Long
Dim TPos                     As Long

Public CurrDate             As Date
Public SelectedDate         As String
Public IsSelected           As Boolean
Private WithEvents tHoverSelection As XTimer
Attribute tHoverSelection.VB_VarHelpID = -1

' /* State type */
Const DSS_NORMAL = &H0
Const DSS_UNION = &H10
Const DSS_DISABLED = &H20
Const DSS_MONO = &H80
Const DSS_RIGHT = &H8000
Const DST_COMPLEX = &H0
Const DST_TEXT = &H1
Const DST_PREFIXTEXT = &H2
Const DST_ICON = &H3
Const DST_BITMAP = &H4

Private Sub Form_DblClick()

    SetCapture hwnd 'Preserve hWnd on DblClick
    Form_MouseDown g_Button, g_Shift, g_X, g_Y
    Call ReleaseCapture

    If m_CalendarClickBehivor = [dtDblClickHide] Then

        If IsOutsideCalendarStamp Then
            IsOutsideCalendarStamp = False
            IsInsideCalendarStamp = False
            Call SetCapture(Me.hwnd)

          ElseIf IsInsideCalendarStamp Then
            IsInsideCalendarStamp = False
            IsOutsideCalendarStamp = False
            If MouseButId = -1 Then
                'IsSelected = True
                SelectedDate = temSelectedDate
                Unload Me
                Set frmCalendar = Nothing
                Exit Sub
            End If
            Call ReleaseCapture
            Call Form_KeyDown(vbKeyEscape, 0)

        End If

    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then IsSelected = True: SelectedDate = temSelectedDate: Unload Me: Set frmCalendar = Nothing
    If (KeyCode = vbKeyEscape) Then
        ReleaseCapture
        Unload Me
    End If

End Sub

Private Sub Form_Load()

  Dim R As RECT

    Me.KeyPreview = True
    Me.ScaleMode = vbPixels
    Me.Font.Name = "Arial"
    Call SetCapture(hwnd)
    Pt3 = 0
    MouseButId = -1
    MouseDownButId = -1
    IsSelected = False
    IsOutsideCalendarStamp = False
    IsInsideCalendarStamp = False

    CurrDate = CDate(IIf((Len(CStr(m_Value)) > 0), m_Value, Date))

    '~~ Initialize Cell Loaction on the Form
    Call InitialRECT

    '~~ Draw Background
    Call DrawBck

    '~~ Draw Button
    Call DrawButton

    '~~ Attach Board Style
    If Not m_CalendarBorderStyle Then
        Call SetBorderStyle(Me)
    End If

    '~~ Calendar
    Set tHoverSelection = New XTimer
    Call DrawWeekLabel
    Call DrawWeekline
    Call SetColors
    Call DrawCalendar

End Sub

Private Sub InitialRECT()

  Dim i As Integer
  Dim j As Integer

    pt = m_nInitialT

    '## Button RECT Defination
    LPos = m_Margin_Offset
    TPos = m_Margin_Offset
    Call SetRect(Clrs(1).Rct, LPos, TPos, LPos + m_nGridWidth, TPos + m_nButLabelHeight)

    LPos = IIf(m_ShowWeek, m_WeekColumnWidth + 2 * m_Margin_Offset + m_nGridWidth * GRID_ROWS, m_Margin_Offset + m_nGridWidth * GRID_ROWS)
    TPos = m_Margin_Offset
    Call SetRect(Clrs(2).Rct, LPos, TPos, LPos + m_nGridWidth, TPos + m_nButLabelHeight)

    pt = (pt + m_Margin_Offset) + m_nButLabelHeight

    '## Draw line I
    Pt1 = pt
    pt = pt + 2 + 1

    '## 7 RECTs for Week label
    For i = 3 To 9
        LPos = IIf(m_ShowWeek, m_WeekColumnWidth + 2 * m_Margin_Offset + (i - 3) * m_nGridWidth, m_Margin_Offset + (i - 3) * m_nGridWidth)
        TPos = (pt + m_Margin_Offset)
        Call SetRect(Clrs(i).Rct, LPos, TPos, LPos + m_nGridWidth, TPos + m_nButLabelHeight)
    Next i

    pt = pt + (1 * m_nButLabelHeight)

    '## Draw line II
    Pt2 = pt
    pt = pt + 2 + 1
    Pt3 = pt

    '## No.10 to 51 RECT to draw Calendar,Total RECTs GRID_COLS*GRID_ROWS=42
  Dim n As Integer
    For i = 1 To GRID_ROWS
        n = (i - 1) * 7 + 10
        For j = 1 To GRID_COLS
            LPos = IIf(m_ShowWeek, m_WeekColumnWidth + 2 * m_Margin_Offset + (j - 1) * m_nGridWidth, m_Margin_Offset + (j - 1) * m_nGridWidth)
            TPos = (pt + m_Margin_Offset) + m_nGridHeight * (i - 1)
            Call SetRect(Clrs(n + (j - 1)).Rct, LPos, TPos, LPos + m_nGridWidth, TPos + m_nGridHeight)
        Next

    Next

    '## No.52 RECT For Displaying Date
    LPos = Clrs(1).Rct.Right + 2
    TPos = m_Margin_Offset
    Call SetRect(Clrs(52).Rct, LPos, TPos, _
         LPos + Clrs(2).Rct.Left - Clrs(1).Rct.Right - 2 * 2, _
         TPos + m_nButLabelHeight)
    '## No.53 to 53+GRID_ROWS-1 for Week Number
    If m_ShowWeek Then
        For i = 53 To 53 + GRID_ROWS - 1
            LPos = m_Margin_Offset
            TPos = Clrs(10 + (i - 53) * GRID_COLS).Rct.Top
            Call SetRect(Clrs(i).Rct, LPos, TPos, LPos + m_WeekColumnWidth, TPos + m_nGridHeight)

        Next
    End If

    intfrmWidth = Clrs(51).Rct.Right + m_Margin_Offset + 2
    intfrmHeight = Clrs(51).Rct.Bottom + m_Margin_Offset + 2

End Sub

Private Sub DrawBck()

  Dim Clr As Long, lBrush As Long, lOldBr As Long
  Dim Rct As RECT

    Width = intfrmWidth * Screen.TwipsPerPixelX

    Height = intfrmHeight * Screen.TwipsPerPixelY

    Call SetRect(Rct, 0, 0, ScaleWidth, ScaleHeight)

    Call OleTranslateColor(m_BackNormal, ByVal 0&, Clr)
    lBrush = CreateSolidBrush(Clr)
    lOldBr = SelectObject(hdc, lBrush)
    Call FillRect(hdc, Rct, lBrush)
    'DeleteObject lBrush
    SelectObject hdc, lOldBr
    DeleteObject lBrush
    DeleteObject lOldBr

    Call DrawEdge(hdc, Rct, EDGE_XPFlat, BF_RECT)
    Call SetRectEmpty(Rct)

End Sub

Private Sub DrawButton()

  Dim Rct As RECT

    For ButID = 1 To 2
        Call SetRect(Rct, Clrs(ButID).Rct.Left, Clrs(ButID).Rct.Top, Clrs(ButID).Rct.Right, Clrs(ButID).Rct.Bottom)

  Dim CurFontName As String
  Dim CurFontSize As Integer
        CurFontName = Font.Name
        CurFontSize = Font.Size
        FrmBkClr = Me.ForeColor
        Me.ForeColor = &H80000012

        If ButID = 1 Then
            Font.Name = "Marlett"
            Font.Size = 12
            Call DrawText(hdc, "3", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
          ElseIf ButID = 2 Then Font.Name = "Marlett"
            Font.Size = 12
            Call DrawText(hdc, "4", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
        End If
        Font.Name = CurFontName
        Font.Size = CurFontSize
        Me.ForeColor = FrmBkClr
        Call DrawButEdge(ButID, 0)
        Call SetRectEmpty(Rct)

    Next ButID

End Sub

Private Sub DrawWeekLabel()

  Dim buffer As String
  Dim intFirstDayOfWeek As Integer

    FrmBkClr = Me.ForeColor
    '## draw line
    Me.ForeColor = vb3DShadow
    
    CurrentX = m_Margin_Offset
    CurrentY = Clrs(1).Rct.Bottom + 2
    Line -(Clrs(2).Rct.Right, CurrentY)
    Me.ForeColor = vb3DHighlight
    
    CurrentX = m_Margin_Offset
    CurrentY = Clrs(1).Rct.Bottom + 2 + 1
    Line -(Clrs(2).Rct.Right, CurrentY)
    Me.ForeColor = FrmBkClr

    intFirstDayOfWeek = CInt(m_CalendarFirstDayOfWeek)

    For iNumDaysPerWeek = 1 To GRID_COLS
        Select Case m_CalendarDayHeaderFormat
          Case [dtSingleLetter]
            strDayCaption(iNumDaysPerWeek) = Left$(Format$(intFirstDayOfWeek, "Ddd"), 1)
          Case [dtMedium]
            strDayCaption(iNumDaysPerWeek) = Format$(intFirstDayOfWeek, "Ddd")
          Case [dtFullName]
            strDayCaption(iNumDaysPerWeek) = Format$(intFirstDayOfWeek, "Dddd")
        End Select

        ButID = 2 + iNumDaysPerWeek

        buffer$ = strDayCaption(iNumDaysPerWeek)
        Call DrawText(hdc, buffer$, Len(buffer$), Clrs(ButID).Rct, DT_CENTER Or DT_SINGLELINE) 'DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
        intFirstDayOfWeek = intFirstDayOfWeek + 1
        If (intFirstDayOfWeek > 7) Then intFirstDayOfWeek = 1

    Next iNumDaysPerWeek

    FrmBkClr = Me.ForeColor
    '## draw line m_Margin_Offset
    Me.ForeColor = vb3DShadow
    
    CurrentX = m_Margin_Offset
    CurrentY = Pt2 + 2
    Line -(Clrs(2).Rct.Right, CurrentY)
    Me.ForeColor = vb3DHighlight
    
    CurrentX = m_Margin_Offset
    CurrentY = Pt2 + 2 + 1
    Line -(Clrs(2).Rct.Right, CurrentY)
    Me.ForeColor = FrmBkClr

End Sub

Private Sub DrawWeekNumber()

  Dim dtWeekDate    As Date
  Dim intWeek       As Integer
  Dim RctWeekText   As RECT
  Dim intWeekRECT   As Integer
  Dim buffer        As String

    If m_ShowWeek Then

        dtWeekDate = CurrDate
        dtWeekDate = CDate(Month(dtWeekDate) & "/" & "01" & "/" & Year(dtWeekDate))
        dtWeekDate = VBA.Format$(dtWeekDate, "mm/dd/yyyy")
        intWeek = WeekNumber(dtWeekDate)
        For intWeekRECT = 53 To 53 + GRID_ROWS - 1
            If (intWeek + intWeekRECT - 53) >= 54 Then
                intWeek = 1
                buffer = ""
                Exit For
              Else
                buffer = CStr(intWeek + intWeekRECT - 53) & " "
            End If
            SetRect RctWeekText, _
                    Clrs(intWeekRECT).Rct.Left, _
                    Clrs(intWeekRECT).Rct.Top + m_DateYOffset, _
                    Clrs(intWeekRECT).Rct.Right, _
                    Clrs(intWeekRECT).Rct.Bottom
            '## Week Number always at Center RECT,you can replace
            '   AlignCenter with m_DateAlign to get 3 different
            '   kinds of Alignment
            Call DrawText(hdc, buffer$, Len(buffer), RctWeekText, AlignRight Or Choose(m_DatePosn + 1, DT_TOP, DT_VCENTER, DT_BOTTOM) Or DT_SINGLELINE)
        Next intWeekRECT
        Call SetRectEmpty(RctWeekText)
    End If

End Sub

Private Sub DrawWeekline()

    Me.ForeColor = vb3DShadow

    If m_ShowWeek Then
        Me.ForeColor = vb3DShadow
        Line (Clrs(53).Rct.Right, Clrs(53).Rct.Top)-(Clrs(53 + GRID_ROWS - 1).Rct.Right, Clrs(53 + GRID_ROWS - 1).Rct.Bottom)
        Me.ForeColor = vb3DHighlight
        Line (Clrs(53).Rct.Right + 1 + 2, Clrs(53).Rct.Top)-(Clrs(53 + GRID_ROWS - 1).Rct.Right + 1 + 2, Clrs(53 + GRID_ROWS - 1).Rct.Bottom)
    End If

    Me.ForeColor = FrmBkClr

End Sub

Private Sub DrawGridLine()

  '~~ Draw GridLine

  Dim s As Integer, t As Integer, m As Integer

    Select Case m_Gridline
      Case 3
        For s = 1 To GRID_ROWS
            t = (s - 1) * 7 + 10
            For m = 1 To GRID_COLS
                Me.Line (Clrs(t + (m - 1)).Rct.Left, Clrs(t + (m - 1)).Rct.Top)- _
                        (Clrs(t + (m - 1)).Rct.Right, Clrs(t + (m - 1)).Rct.Bottom), _
                        m_GridLineColor, B
            Next m
        Next s
      Case 4
        For s = 1 To GRID_ROWS
            t = (s - 1) * 7 + 10
            For m = 1 To GRID_COLS
                Call DrawLine(hdc, Clrs(t + (m - 1)).Rct.Left, Clrs(t + (m - 1)).Rct.Top, _
                     Clrs(t + (m - 1)).Rct.Right, Clrs(t + (m - 1)).Rct.Top, _
                     m_GridLineColor)
            Next m
        Next s
        Call DrawLine(hdc, Clrs(10 + (GRID_ROWS - 1) * GRID_COLS).Rct.Left, _
             Clrs(10 + (GRID_ROWS - 1) * GRID_COLS).Rct.Bottom, _
             Clrs(10 + GRID_ROWS * GRID_COLS - 1).Rct.Right, _
             Clrs(10 + GRID_ROWS * GRID_COLS - 1).Rct.Bottom, _
             m_GridLineColor)
      Case 5
        For s = 1 To GRID_ROWS
            t = (s - 1) * 7 + 10
            For m = 1 To GRID_COLS
                Call DrawLine(hdc, Clrs(t + (m - 1)).Rct.Right, Clrs(t + (m - 1)).Rct.Top, _
                     Clrs(t + (m - 1)).Rct.Right, Clrs(t + (m - 1)).Rct.Bottom, _
                     m_GridLineColor)
            Next m
        Next s
        Call DrawLine(hdc, Clrs(10).Rct.Left, Clrs(10).Rct.Top, _
             Clrs(10 + (GRID_ROWS - 1) * GRID_COLS).Rct.Left, Clrs(10 + (GRID_ROWS - 1) * GRID_COLS).Rct.Bottom, _
             m_GridLineColor)

    End Select

End Sub

Private Sub DrawCalendar()

  Dim NumDays As Integer, CurrPos As Integer, bCurrMonth As Boolean
  Dim buffer As String
  Dim intNrDaysInMonth As Integer, intFirstWeekDay As Integer
  Dim intHelpFlag As Integer
  Dim DateTextRct As RECT
  Dim CurFontName As String
  Dim CurFontSize As Integer

    temButID = 0
    ButID = 0
    intNrDaysInMonth = 0
    intFirstWeekDay = 0

    Call ClearDayCell
    Call ClsCalendar
    HiTest True

    '## Paint day text

    'Determine if this month is today's month
    If Month(CurrDate) = Month(Date) And Year(CurrDate) = Year(Date) Then
        bCurrMonth = True
    End If
    'Get first date in the month
    MonthStart = DateSerial(Year(CurrDate), Month(CurrDate), 1)
    'Get previous date in the month
    PrevMonth = DateAdd("m", -1, MonthStart)
    'Get Last day in current month
    MonthEnd = DateAdd("d", -1, DateAdd("m", 1, MonthStart))
    'Get first day in next month
    NextMonth = DateAdd("d", -1, DateAdd("m", 2, MonthStart))

    FrmBkClr = Me.ForeColor
    Me.ForeColor = &H80000012
    Call DrawText(hdc, Format(MonthStart, "mmmm yyyy"), Len(Format(MonthStart, "mmmm yyyy")), Clrs(52).Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
    Me.ForeColor = FrmBkClr
    'Number of days in the month
    NumDays = DateDiff("d", MonthStart, DateAdd("m", 1, MonthStart))
    PrevDays = DateDiff("d", PrevMonth, DateAdd("m", 1, PrevMonth))

    'Get first weekday in the month (0 - based)
    intFirstWeekDay = Weekday(MonthStart, m_CalendarFirstDayOfWeek) - 1

    'Tweak for 1-based For/Next index,just for easy calculation
    intFirstWeekDay = intFirstWeekDay - 1

    If intFirstWeekDay >= 0 Then '## j>0 it means first Row has previous month days
        For intNrDaysInMonth = 0 To intFirstWeekDay 'Add previous month days (greyed)
            buffer$ = CStr(PrevDays - intFirstWeekDay + intNrDaysInMonth)
            Me.ForeColor = &H808080
            Me.Font.Bold = False
            Me.Font.Size = 8
            ButID = 10 + intNrDaysInMonth
            SetRect DateTextRct, _
                    Clrs(ButID).Rct.Left + Choose(m_DateAlign + 1, AlignLeft_Offset + m_DateXOffset, m_DateXOffset, m_DateXOffset), _
                    Clrs(ButID).Rct.Top + m_DateYOffset, _
                    Clrs(ButID).Rct.Right + Choose(m_DateAlign + 1, AlignLeft_Offset + m_DateXOffset, m_DateXOffset, m_DateXOffset), _
                    Clrs(ButID).Rct.Bottom

            Call DrawText(hdc, buffer$, Len(buffer), DateTextRct, m_DateAlign Or Choose(m_DatePosn + 1, DT_TOP, DT_VCENTER, DT_BOTTOM) Or DT_SINGLELINE)

            SetRect DateTextRct, _
                    Clrs(ButID).Rct.Left, Clrs(ButID).Rct.Top, _
                    Clrs(ButID).Rct.Right, Clrs(ButID).Rct.Bottom
        Next intNrDaysInMonth

      ElseIf intFirstWeekDay = -1 Then ButID = 9

    End If

    temButID = ButID

    'Display dates for current month as black color
    'Display Prev/Next month days as greyed

    For intNrDaysInMonth = 1 To 41 - intFirstWeekDay         'NumDays

        If intNrDaysInMonth <= NumDays Then
            'Show date as bold if date is today
            If bCurrMonth And intNrDaysInMonth = Day(Date) Then
                Me.Font.Size = m_TodayFont.Size
                Me.Font.Bold = m_TodayFont.Bold
                Me.Font.Italic = m_TodayFont.Italic
                Me.Font.Strikethrough = m_TodayFont.Strikethrough
                Me.Font.Underline = m_TodayFont.Underline

              Else
                Me.Font.Bold = False
                Me.FontUnderline = False
                Me.Font.Size = 8
            End If
            Me.ForeColor = 0

            For intHelpFlag = 1 To 6 '## Maximum number days in a month is less than 6*7=42days
                If m_CalendarFirstDayOfWeek <> 1 Then
                    If intNrDaysInMonth = intHelpFlag * GRID_COLS - m_CalendarFirstDayOfWeek - intFirstWeekDay Then
                        Me.ForeColor = &HFF0000
                      ElseIf intNrDaysInMonth = intHelpFlag * GRID_COLS - (m_CalendarFirstDayOfWeek - 1) - intFirstWeekDay Then
                        Me.ForeColor = &HFF&
                    End If
                  Else
                    If intNrDaysInMonth = intHelpFlag * GRID_COLS - 1 - intFirstWeekDay Then
                        Me.ForeColor = &HFF0000
                      ElseIf intNrDaysInMonth = intHelpFlag * GRID_COLS - 7 - intFirstWeekDay Then
                        Me.ForeColor = &HFF&
                    End If

                End If
            Next intHelpFlag

            '~~ HotTracking ForeColor
            If intHighlightDay = intNrDaysInMonth And m_HotTracking = True Then
                Me.ForeColor = m_HoverColor
                Me.Font.Size = 10
            End If

            buffer$ = CStr(intNrDaysInMonth)
            ButID = temButID + intNrDaysInMonth
            SetRect DateTextRct, _
                    Clrs(ButID).Rct.Left + Choose(m_DateAlign + 1, AlignLeft_Offset + m_DateXOffset, m_DateXOffset, m_DateXOffset), _
                    Clrs(ButID).Rct.Top + m_DateYOffset, _
                    Clrs(ButID).Rct.Right + Choose(m_DateAlign + 1, AlignLeft_Offset + m_DateXOffset, m_DateXOffset, m_DateXOffset), _
                    Clrs(ButID).Rct.Bottom
            Call DrawText(hdc, buffer$, Len(buffer), DateTextRct, m_DateAlign Or Choose(m_DatePosn + 1, DT_TOP, DT_VCENTER, DT_BOTTOM) Or DT_SINGLELINE)

            '~~ Save cell infomation into array
            With DayCell(ButID - 10)
                .Day = intNrDaysInMonth
                .X1 = Clrs(ButID).Rct.Left
                .X2 = Clrs(ButID).Rct.Right
                .Y1 = Clrs(ButID).Rct.Top
                .Y2 = Clrs(ButID).Rct.Bottom
                Select Case m_Gridline
                  Case 0
                    Me.Line (.X1, .Y1)-(.X2, .Y2), m_GridLineColor, B
                  Case 1
                    Call DrawLine(hdc, .X1, .Y1, .X2, .Y1, m_GridLineColor)
                    Call DrawLine(hdc, .X2, .Y2, .X1, .Y2, m_GridLineColor)
                  Case 2
                    Call DrawLine(hdc, .X2, .Y1, .X2, .Y2, m_GridLineColor)
                    Call DrawLine(hdc, .X1, .Y2, .X1, .Y1, m_GridLineColor)

                End Select
            End With

          Else  'Add next month as 1 to ... (greyed)
            buffer$ = CStr(intNrDaysInMonth - NumDays)
            Me.ForeColor = &H808080
            Me.Font.Size = 8
            Me.Font.Bold = False
            ButID = temButID + intNrDaysInMonth

            SetRect DateTextRct, _
                    Clrs(ButID).Rct.Left + Choose(m_DateAlign + 1, AlignLeft_Offset + m_DateXOffset, m_DateXOffset, m_DateXOffset), _
                    Clrs(ButID).Rct.Top + m_DateYOffset, _
                    Clrs(ButID).Rct.Right + Choose(m_DateAlign + 1, AlignLeft_Offset + m_DateXOffset, m_DateXOffset, m_DateXOffset), _
                    Clrs(ButID).Rct.Bottom

            Call DrawText(hdc, buffer$, Len(buffer), DateTextRct, m_DateAlign Or Choose(m_DatePosn + 1, DT_TOP, DT_VCENTER, DT_BOTTOM) Or DT_SINGLELINE)

        End If

    Next intNrDaysInMonth
    SetRectEmpty DateTextRct

    Call DrawWeekNumber
    Call DrawWeekline
    Call DrawGridLine

    '~~ Refresh,update painting
    Refresh

End Sub

Private Sub Form_LostFocus()

    Call ReleaseCapture
    Call Form_KeyDown(vbKeyEscape, 0)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not IsActiveWindow Then Exit Sub
  Dim i As Integer

    For i = 1 To 2
        If IsMouseOnBut(i) Then
            MouseButId = i
            Call DrawButEdge(i, 1)
            Exit For
          Else
            MouseButId = -1
            Call DrawButEdge(i, 0)
        End If
    Next i
    Refresh

    '~~ Real HotTracking and HoverSelection Functions
    If (m_HoverSelection Or m_HotTracking) = True And MouseButId = -1 Then

  Static intHighlightday_Tag As Integer 'anti-repainting tag
  Static intFoundDay_Tag As Integer     'anti-repainting tag

  Dim bHelpTag As Boolean

        For i = LBound(DayCell) To UBound(DayCell)

            If X > DayCell(i).X1 And X < DayCell(i).X2 And Y > DayCell(i).Y1 And Y < DayCell(i).Y2 Then

                intHighlightDay = DayCell(i).Day
                intFoundDay_Tag = 1

                '~~ If it's a different day then reset the timer
                If intHoverDay <> DayCell(i).Day And m_HoverSelection = True Then

                    intHoverDay = DayCell(i).Day
                    tHoverSelection.Interval = 1000  '## 1 second
                    tHoverSelection.Enabled = True

                End If

                If intHighlightday_Tag <> intHighlightDay And m_HotTracking = True Then
                    intHighlightday_Tag = intHighlightDay
                    DrawCalendar
                End If
                'intHighlightDay = 0
                bHelpTag = True
                Exit For
            End If
        Next i

        'If mouse is moving out of  day Map of active month,only do drawCalendar
        'once to clear Hottracking color and restore normal Forecolor
        'with Tag(intFoundDay_Tag) help.
        If Not bHelpTag And intFoundDay_Tag = 1 And m_HotTracking = True Then
            intFoundDay_Tag = intFoundDay_Tag + 1
            intHoverDay = 0
            intHighlightday_Tag = 0
            tHoverSelection.Enabled = False
            intHighlightDay = 0
            DrawCalendar
        End If
    End If

    If (GetCapture() <> Me.hwnd) Then
        Call SetCapture(Me.hwnd)

    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    g_Button = Button
    g_Shift = Shift
    g_X = X
    g_Y = Y

    Select Case Button
      Case vbRightButton
        '...
      Case vbLeftButton
        If Not (MouseButId = -1) Then
            On Error Resume Next
              IsMouseDown(MouseButId) = True
              MouseDownButId = MouseButId
              Call SetRect(FocusRec, Clrs(MouseButId).Rct.Left, Clrs(MouseButId).Rct.Top, Clrs(MouseButId).Rct.Right, Clrs(MouseButId).Rct.Bottom)
              If (X >= FocusRec.Left And X <= FocusRec.Right) And (Y >= FocusRec.Top And Y <= FocusRec.Bottom) Then
                  IsFocus(MouseButId) = True
                  Call DrawButFocus
                  Call DrawButton
                  Call DrawButEdge(MouseButId, 2)
                  Refresh  'Refresh Button Down Edge
                  temMouseButIDFocus = MouseButId
                Else
                  IsFocus(MouseButId) = False
              End If
              Call SetRectEmpty(FocusRec)
              MouseDownButId = -1
          End If
      End Select

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim IsMouseOver As Boolean
  Dim MaxNrDay      As Integer

    IsMouseOver = X >= 0 And Y >= 0 And X <= ScaleWidth And Y <= ScaleHeight
    MaxNrDay = Day(DateAdd("d", -1, DateSerial(Year(CurrDate), Month(CurrDate) + 1, 1)))
    Select Case Button
      Case vbRightButton
        '...
      Case vbLeftButton
        Call ClickCalendar
  Dim i As Integer
        For i = LBound(DayCell) To UBound(DayCell)

            If X > DayCell(i).X1 And X < DayCell(i).X2 And Y > DayCell(i).Y1 And Y < DayCell(i).Y2 Then

                If DayCell(i).Day >= 1 And DayCell(i).Day <= MaxNrDay Then
                    IsInsideCalendarStamp = True
                    IsSelected = True
                    SetNewDate DateSerial(Year(CurrDate), Month(CurrDate), DayCell(i).Day)
                End If

                Exit For
            End If
        Next i

    End Select

    If m_CalendarClickBehivor = [dtDblClickHide] Then
        If IsMouseOver Then

            '## since any mouse click will release capture, set mouse capture again
            Call SetCapture(Me.hwnd)
          Else

            Call ReleaseCapture
            Call Form_KeyDown(vbKeyEscape, 0)
            Exit Sub

        End If

      ElseIf m_CalendarClickBehivor = [dtOneClickHide] Then
        If IsMouseOver Then
            If IsOutsideCalendarStamp Then
                IsOutsideCalendarStamp = False
                IsInsideCalendarStamp = False
                Call SetCapture(Me.hwnd)

              ElseIf IsInsideCalendarStamp Then
                IsInsideCalendarStamp = False
                IsOutsideCalendarStamp = False
                If MouseButId < 1 Then
                    'IsSelected = True
                    SelectedDate = temSelectedDate
                    Unload Me
                    Set frmCalendar = Nothing
                    Exit Sub
                End If
                Call ReleaseCapture
                Call Form_KeyDown(vbKeyEscape, 0)

            End If
            '## since any mouse click will release capture, set mouse capture again
            Call SetCapture(Me.hwnd)
          Else
            '## outside of the form.  release mouse capture and unload it
            Call ReleaseCapture
            Call Form_KeyDown(vbKeyEscape, 0)
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  '## No key movility if Calendar visible

    If KeyCode <= 37 And KeyCode >= 40 Then Exit Sub

    '## Erase Hottracking
    intHighlightDay = 0

    Select Case KeyCode

        '## Left
      Case 37  '## Previous Day
        SetNewDate DateAdd("d", -1, CurrDate)

        '## Right
      Case 39  '## Next Day
        SetNewDate DateAdd("d", 1, CurrDate)

        '## Up
      Case 38
        SetNewDate DateAdd("d", -7, CurrDate)

        '## Down
      Case 40
        SetNewDate DateAdd("d", 7, CurrDate)

        '## Previous Month
      Case 33
        If Day(DateAdd("d", (-1), MonthStart)) >= Day(CurrDate) Then
            SetNewDate DateAdd("d", (-1) * Day(DateAdd("d", (-1), MonthStart)), CurrDate)
          Else
            SetNewDate DateAdd("d", (-1), MonthStart)
        End If

        '## Next Month
      Case 34
        If Day(NextMonth) >= Day(CurrDate) Then
            SetNewDate DateAdd("d", Day(MonthEnd), CurrDate)
          Else
            SetNewDate NextMonth
        End If

        '## Home
      Case 36
        SetNewDate MonthStart

        '## End
      Case 35
        SetNewDate MonthEnd

        '## Tab
      Case 9
        '## Next Day
        SetNewDate DateAdd("d", 1, CurrDate)

        '## Escape
      Case 27
        ReleaseCapture
        Unload Me
        Set frmCalendar = Nothing

    End Select

End Sub

Private Sub DrawButFocus()

  Dim Rct As RECT
  Dim Clr As Long, lBrush As Long, lOldBr As Long

    If MouseButId <> -1 Then

        If IsFocus(MouseButId) = True Then

            If MouseDownButId = 1 Or 2 Then
                Call SetRect(Rct, Clrs(temMouseButIDFocus).Rct.Left + 1, Clrs(temMouseButIDFocus).Rct.Top + 1, Clrs(temMouseButIDFocus).Rct.Right - 1, Clrs(temMouseButIDFocus).Rct.Bottom - 1)
                Call OleTranslateColor(m_BackNormal, ByVal 0&, Clr)
                lBrush = CreateSolidBrush(Clr)
                lOldBr = SelectObject(hdc, lBrush)
                Call FillRect(hdc, Rct, lBrush)
                'DeleteObject lBrush
                SelectObject hdc, lOldBr
                DeleteObject lBrush
                DeleteObject lOldBr
                Call SetRectEmpty(Rct)
            End If

            Call SetRect(FocusRec, Clrs(MouseButId).Rct.Left + 2, Clrs(MouseButId).Rct.Top + 2, Clrs(MouseButId).Rct.Right - 2, Clrs(MouseButId).Rct.Bottom - 2)
            Call DrawFocusRect(hdc, FocusRec)
            Call SetRectEmpty(FocusRec)
            IsFocus(MouseButId) = False
            Call DrawButEdge(MouseButId, 1)
        End If

    End If

End Sub

Private Sub ClsCalendar()

  Dim Rct As RECT
  Dim i As Integer
  Dim Clr As Long, lBrush As Long, lOldBr As Long

    'Erase Last Calendar Map
    Call SetRect(Rct, m_Margin_Offset, Clrs(10).Rct.Top, ScaleWidth - 2, ScaleHeight - 2)
    Call OleTranslateColor(m_BackNormal, ByVal 0&, Clr)
    lBrush = CreateSolidBrush(Clr)
    lOldBr = SelectObject(hdc, lBrush)
    Call FillRect(hdc, Rct, lBrush)
    'DeleteObject lBrush
    SelectObject hdc, lOldBr
    DeleteObject lBrush
    DeleteObject lOldBr
   
    'Erase Last Date Caption
    Call SetRect(Rct, Clrs(52).Rct.Left, Clrs(52).Rct.Top, Clrs(52).Rct.Right, Clrs(52).Rct.Bottom)
    Call OleTranslateColor(m_BackNormal, ByVal 0&, Clr)
    lBrush = CreateSolidBrush(Clr)
    lOldBr = SelectObject(hdc, lBrush)
    Call FillRect(hdc, Rct, lBrush)
    'DeleteObject lBrush
    SelectObject hdc, lOldBr
    DeleteObject lBrush
    DeleteObject lOldBr
    
    Call SetRectEmpty(Rct)

End Sub

Private Sub DrawButEdge(ClrId As Integer, EdgeStyle As Integer)

    Select Case EdgeStyle
      Case 0
        Call DrawEdge(hdc, Clrs(ClrId).Rct, BDR_RAISEDINNER, BF_RECT Or BF_FLAT)
      Case 1
        Call DrawEdge(hdc, Clrs(ClrId).Rct, BDR_RAISEDINNER, BF_RECT)
      Case 2
        Call DrawEdge(hdc, Clrs(ClrId).Rct, BDR_SUNKENOUTER, BF_RECT)
    End Select

    'Refresh

End Sub

Private Function IsMouseOnBut(ButID As Integer) As Boolean

  Dim pti As POINTAPI

    Call GetCursorPos(pti)
    Call ScreenToClient(Me.hwnd, pti)

    IsMouseOnBut = (pti.X >= Clrs(ButID).Rct.Left And pti.X <= Clrs(ButID).Rct.Right) And _
                   (pti.Y >= Clrs(ButID).Rct.Top And pti.Y <= Clrs(ButID).Rct.Bottom)

End Function

Private Sub SetNewDate(NewDate As Date)

    If Month(CurrDate) = Month(NewDate) And Year(CurrDate) = Year(NewDate) Then

        CurrDate = NewDate

        Call DrawCalendar
      Else

        CurrDate = NewDate

        Call DrawCalendar
    End If

End Sub

Private Sub ClickCalendar()

    Select Case MouseButId
      Case 1 'Prev
        SetNewDate DateAdd("m", -1, CurrDate)
      Case 2 'Next
        SetNewDate DateAdd("m", 1, CurrDate)
    End Select

End Sub

Private Sub HiTest(bSelected As Boolean)

  Dim intHiTest As Integer, intTodayRECT As Integer
  Dim rctVerSign As RECT
  Dim intWeekNo As Integer
  Dim intWeekFocusRct As Integer
  Dim TodayRct As RECT
  
    'Get HiTest WeekDay from FirstDayOfWeek to LastDayofWeek (1 to 7)
    'Weekday(CurrDate, m_CalendarFirstDayOfWeek)

    'Compute location for current date
    intHiTest = Weekday(DateSerial(Year(CurrDate), Month(CurrDate), vbSunday), m_CalendarFirstDayOfWeek) - 1
    intHiTest = intHiTest + (Day(CurrDate) - 1)

    '## Calendar Output Text Option
    Select Case m_CalendarOption
      Case 0  '[Default-Only Date]
        temSelectedDate = CStr(CurrDate)
    End Select
    TimeNow = TimeSerial(Hour(Time), Minute(Time), Second(Time))

    On Error Resume Next
      '## Selected Style
      Select Case m_SelectModeStyle
        Case 0 '[Standard]
          DrawBack hdc, Clrs(intHiTest + 10).Rct, cBackSel

        Case 1 '[Gradient-V]
          DrawBackGrad hdc, Clrs(intHiTest + 10).Rct, cGrad1, cGrad2, GRADIENT_FILL_RECT_V

        Case 2 '[Gradient-H]
          DrawBackGrad hdc, Clrs(intHiTest + 10).Rct, cGrad1, cGrad2, GRADIENT_FILL_RECT_H

        Case 3 '[ByPicture]

          If Not m_SelectionPicture Is Nothing Then

              If m_SelectionPicture.Type = vbPicTypeBitmap Then

                  Call DrawBitmap(hdc, m_SelectionPicture, _
                       m_nGridWidth, _
                       m_nGridHeight, _
                       ScaleX(m_SelectionPicture.Width, 8, ScaleMode), _
                       ScaleY(m_SelectionPicture.Height, 8, ScaleMode), _
                       1, Clrs(intHiTest + 10).Rct, m_CalendarSelectPicDrawMode + 1, m_SelectedBMPMaskColor)
                Else
                  Call DrawPIcon(hdc, m_SelectionPicture, _
                       m_nGridWidth, _
                       m_nGridHeight, _
                       ScaleX(m_SelectionPicture.Width, 8, ScaleMode), _
                       ScaleY(m_SelectionPicture.Height, 8, ScaleMode), _
                       1, Clrs(intHiTest + 10).Rct)
              End If
            Else
              DrawBack hdc, Clrs(intHiTest + 10).Rct, cBackSel
          End If

      End Select

      If m_ShowWeek = True And m_ShowWeekSignPicture Then
    
          intWeekNo = WeekNumber(CurrDate)
          intWeekFocusRct = Format(Int(intHiTest / 7), 0) + 53
          If Not (m_WeekSignPicture Is Nothing) Then
              SetRect rctVerSign, _
                      IIf(Val(m_WeekSignPicXOffset) <= 0, Clrs(intWeekFocusRct).Rct.Left, Clrs(intWeekFocusRct).Rct.Left + m_WeekSignPicXOffset), _
                      Clrs(intWeekFocusRct).Rct.Top + _
                      Choose(m_DatePosn + 1, m_WeekSignPicYOffset + m_DateYOffset, (m_nGridHeight + m_WeekSignPicYOffset + m_DateYOffset - m_WeekSignPictureHeight) / 2, m_nGridHeight - m_WeekSignPictureHeight), _
                      Clrs(intWeekFocusRct).Rct.Left + m_WeekSignPicXOffset + m_WeekSignPictureWidth, _
                      Clrs(intWeekFocusRct).Rct.Top + _
                      Choose(m_DatePosn + 1, m_WeekSignPicYOffset + m_DateYOffset, (m_nGridHeight + m_WeekSignPicYOffset + m_DateYOffset - m_WeekSignPictureHeight) / 2, m_nGridHeight - m_WeekSignPictureHeight)

              If m_WeekSignPicture.Type = vbPicTypeBitmap Then

                  Call DrawBitmap(hdc, m_WeekSignPicture, _
                       m_WeekSignPictureWidth, _
                       m_WeekSignPictureHeight, _
                       m_WeekSignOriginalPicSizeW, _
                       m_WeekSignOriginalPicSizeH, _
                       1, rctVerSign, _
                       1, m_WeekSignPicMaskColor)

                Else

                  Call DrawPIcon(hdc, m_WeekSignPicture, _
                       m_WeekSignPictureWidth, _
                       m_WeekSignPictureHeight, _
                       m_WeekSignOriginalPicSizeW, _
                       m_WeekSignOriginalPicSizeH, _
                       1, rctVerSign)

              End If
            Else

              SetRect rctVerSign, _
                      IIf(Val(m_WeekSignPicXOffset) < 0, Clrs(intWeekFocusRct).Rct.Left, Clrs(intWeekFocusRct).Rct.Left + m_WeekSignPicXOffset), _
                      Clrs(intWeekFocusRct).Rct.Top + _
                      Choose(m_DatePosn + 1, m_WeekSignPicYOffset + m_DateYOffset, (m_nGridHeight + m_WeekSignPicYOffset + m_DateYOffset - ScaleY(LoadResPicture("WeekSignBMP", vbResBitmap).Height, 8, ScaleMode)) / 2, m_nGridHeight - ScaleY(LoadResPicture("WeekSignBMP", vbResBitmap).Height, 8, ScaleMode)), _
                      Clrs(intWeekFocusRct).Rct.Left + m_WeekSignPicXOffset + ScaleX(LoadResPicture("WeekSignBMP", vbResBitmap).Width, 8, ScaleMode), _
                      Clrs(intWeekFocusRct).Rct.Top + _
                      Choose(m_DatePosn + 1, m_WeekSignPicYOffset + m_DateYOffset, (m_nGridHeight + m_WeekSignPicYOffset + m_DateYOffset - ScaleY(LoadResPicture("WeekSignBMP", vbResBitmap).Height, 8, ScaleMode)) / 2, m_nGridHeight - ScaleY(LoadResPicture("WeekSignBMP", vbResBitmap).Height, 8, ScaleMode))

              '~~ LoadResPicture("TODAYBMP", vbResBitmap) like StdPicture,has all StdPicture properties.
              Call DrawBitmap(hdc, LoadResPicture("WeekSignBMP", vbResBitmap), _
                   ScaleX(LoadResPicture("WeekSignBMP", vbResBitmap).Width, 8, ScaleMode), _
                   ScaleY(LoadResPicture("WeekSignBMP", vbResBitmap).Height, 8, ScaleMode), _
                   ScaleX(LoadResPicture("WeekSignBMP", vbResBitmap).Width, 8, ScaleMode), _
                   ScaleY(LoadResPicture("WeekSignBMP", vbResBitmap).Height, 8, ScaleMode), _
                   1, rctVerSign, _
                   1, &H0) 'MaskColor is &H0 based on Resource bitmap file

          End If
          Call SetRectEmpty(rctVerSign)
      End If
      '~~ Tracing the location of texture of Today to draw BMP or Icon inside Cell

      intTodayRECT = Weekday(DateSerial(Year(Date), Month(Date), 1), m_CalendarFirstDayOfWeek) - 1
      intTodayRECT = intTodayRECT + (Day(Date) - 1)

      If Month(CurrDate) = Month(Date) And Year(CurrDate) = Year(Date) Then

          SetRect TodayRct, Clrs(intTodayRECT + 10).Rct.Left + _
                  Choose(m_DateAlign + 1, m_DateXOffset + m_TodayXOffset, (m_nGridWidth + m_DateXOffset + m_TodayXOffset - m_TodayPictureWidth) / 2, m_nGridWidth - m_TodayPictureWidth), _
                  Clrs(intTodayRECT + 10).Rct.Top + _
                  Choose(m_DatePosn + 1, m_DateYOffset + m_TodayYOffset, (m_nGridHeight + m_DateYOffset + m_TodayYOffset - m_TodayPictureHeight) / 2, m_nGridHeight - m_TodayPictureHeight), _
                  Clrs(intTodayRECT + 10).Rct.Left + _
                  Choose(m_DateAlign + 1, m_DateXOffset + m_TodayXOffset, (m_nGridWidth + m_DateXOffset + m_TodayXOffset - m_TodayPictureWidth) / 2, m_nGridWidth - m_TodayPictureWidth) + m_TodayPictureWidth, _
                  Clrs(intTodayRECT + 10).Rct.Top + _
                  Choose(m_DatePosn + 1, m_DateYOffset + m_TodayYOffset, (m_nGridHeight + m_DateYOffset + m_TodayYOffset - m_TodayPictureHeight) / 2, m_nGridHeight)
          If Not (m_TodayPicture Is Nothing) Then
              If m_TodayPicture.Type = vbPicTypeBitmap Then

                  Call DrawBitmap(hdc, m_TodayPicture, _
                       m_TodayPictureWidth, _
                       m_TodayPictureHeight, _
                       m_TodayOriginalPicSizeW, _
                       m_TodayOriginalPicSizeH, _
                       1, TodayRct, _
                       1, m_TodayMaskColor)
                Else

                  Call DrawPIcon(hdc, m_TodayPicture, _
                       m_TodayPictureWidth, _
                       m_TodayPictureHeight, _
                       m_TodayOriginalPicSizeW, _
                       m_TodayOriginalPicSizeH, _
                       1, TodayRct)

              End If
            Else

              SetRect TodayRct, Clrs(intTodayRECT + 10).Rct.Left + _
                      Choose(m_DateAlign + 1, m_DateXOffset + m_TodayXOffset, (m_nGridWidth + m_DateXOffset + m_TodayXOffset - ScaleX(LoadResPicture("TODAYBMP", vbResBitmap).Width, 8, ScaleMode)) / 2, m_nGridWidth - ScaleX(LoadResPicture("TODAYBMP", vbResBitmap).Width, 8, ScaleMode)), _
                      Clrs(intTodayRECT + 10).Rct.Top + _
                      Choose(m_DatePosn + 1, m_DateYOffset + m_TodayYOffset, (m_nGridHeight + m_DateYOffset + m_TodayYOffset - ScaleY(LoadResPicture("TODAYBMP", vbResBitmap).Height, 8, ScaleMode)) / 2, m_nGridHeight - ScaleY(LoadResPicture("TODAYBMP", vbResBitmap).Height, 8, ScaleMode)), _
                      Clrs(intTodayRECT + 10).Rct.Left + _
                      Choose(m_DateAlign + 1, m_DateXOffset + m_TodayXOffset, (m_nGridWidth + m_DateXOffset + m_TodayXOffset - ScaleX(LoadResPicture("TODAYBMP", vbResBitmap).Width, 8, ScaleMode)) / 2, m_nGridWidth - ScaleX(LoadResPicture("TODAYBMP", vbResBitmap).Width, 8, ScaleMode)) + ScaleX(LoadResPicture("TODAYBMP", vbResBitmap).Width, 8, ScaleMode), _
                      Clrs(intTodayRECT + 10).Rct.Top + _
                      Choose(m_DatePosn + 1, m_DateYOffset + m_TodayYOffset, (m_nGridHeight + m_DateYOffset + m_TodayYOffset - ScaleY(LoadResPicture("TODAYBMP", vbResBitmap).Height, 8, ScaleMode)) / 2, m_nGridHeight - ScaleY(LoadResPicture("TODAYBMP", vbResBitmap).Height, 8, ScaleMode))
              '~~ LoadResPicture("TODAYBMP", vbResBitmap) like StdPicture,has all StdPicture properties.
              Call DrawBitmap(hdc, LoadResPicture("TODAYBMP", vbResBitmap), _
                   ScaleX(LoadResPicture("TODAYBMP", vbResBitmap).Width, 8, ScaleMode), _
                   ScaleY(LoadResPicture("TODAYBMP", vbResBitmap).Height, 8, ScaleMode), _
                   ScaleX(LoadResPicture("TODAYBMP", vbResBitmap).Width, 8, ScaleMode), _
                   ScaleY(LoadResPicture("TODAYBMP", vbResBitmap).Height, 8, ScaleMode), _
                   1, TodayRct, _
                   1, &HFFFFFF) 'MaskColor is &HFFFFFF based on Resource bitmap file

          End If
          SetRectEmpty TodayRct
      End If

End Sub

Private Sub Form_Terminate()

  Dim i As Integer

    For i = 0 To 60
        SetRectEmpty Clrs(i).Rct
    Next i
    Call ClearDayCell
    If Not (tHoverSelection Is Nothing) Then Set tHoverSelection = Nothing
    Call ReleaseCapture
    Unload Me
    Set frmCalendar = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call ReleaseCapture
    Unload Me
    Set frmCalendar = Nothing

End Sub

Private Sub ClearDayCell()

  Dim i As Integer

    For i = LBound(DayCell) To UBound(DayCell)
        With DayCell(i)
            .Day = 0
            Set .Img = Nothing
            .X1 = 0
            .X2 = 0
            .Y1 = 0
            .Y2 = 0
        End With
    Next i

End Sub

Private Sub tHoverSelection_Tick()

    tHoverSelection.Enabled = False
    If Not IsActiveWindow Then Exit Sub

    On Error Resume Next
      If m_HoverSelection And intHoverDay > 0 Then
          CurrDate = DateSerial(Year(CurrDate), Month(CurrDate), intHoverDay)
          DrawCalendar
      End If

End Sub

Private Function IsActiveWindow() As Boolean

    On Error Resume Next
      If GetActiveWindow() <> Me.hwnd Then
          tHoverSelection.Enabled = False
          IsActiveWindow = False
        Else
          IsActiveWindow = True
      End If
      DoEvents

End Function

':) Ulli's VB Code Formatter V2.12.7 (7/29/02 11:00:37 AM) 164 + 1135 = 1299 Lines
