VERSION 5.00
Begin VB.UserControl XPCalendar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   BeginProperty Font 
      Name            =   "Marlett"
      Size            =   9.75
      Charset         =   2
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   117
   ToolboxBitmap   =   "XPCalendar.ctx":0000
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1290
   End
End
Attribute VB_Name = "XPCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Email Me: zhujy@samling.com.my
'Welcome to visit our company WebSite at:
'www.Samling.com.my
'www.Samling.com.cn
'Samling Group---One of a leading and largest WoodBased Industry company in Asia
'Copyright © 2001-2002 by Zhu JinYong from People Republic of China
'Thanks to Abdul Gafoor.GK ,BadSoft and Carles.P.V.
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal fuFlags As Long) As Long

Public Enum pbcStyle
    pbXP = 0
    'pbsmart = 1
End Enum

Const DST_COMPLEX = &H0
Const DST_TEXT = &H1
Const DST_PREFIXTEXT = &H2
Const DST_ICON = &H3
Const DST_BITMAP = &H4

Const DSS_NORMAL = &H0
Const DSS_UNION = &H10
Const DSS_DISABLED = &H20
Const DSS_MONO = &H80
Const DSS_RIGHT = &H8000

Const SM_CXHTHUMB = 10
Const DT_BOTTOM = &H8
Const DT_CENTER = &H1
Const DT_LEFT = &H0
Const DT_NOCLIP = &H100
Const DT_NOPREFIX = &H800
Const DT_RIGHT = &H2
Const DT_SINGLELINE = &H20
Const DT_TOP = &H0
Const DT_VCENTER = &H4
Const DT_WORDBREAK = &H10
Const m_def_IconSizeWidth = 16
Const m_def_IconSizeHeight = 16
Const m_def_Enabled = True
Const m_def_ShowIcon = True
Const m_def_Style = pbXP
Const m_def_FocusColor = &HC00000
Const m_def_OutputText = ""
Const m_def_BorderColor = &HFF8080
Const m_def_BorderColorOver = &H80FF&
Const m_def_BorderColorDown = &HFF&
Const m_def_BgColor = &HFFFFFF
Const m_def_BgColorOver = &HFFFFFF
Const m_def_BgColorDown = &HFFFFFF
Const m_def_ButtonBgColor = &HFFC0C0
Const m_def_ButtonBgColorOver = &HD0B0B0
Const m_def_ButtonBgColorDown = &HFFC0C0
Const m_def_BackNormal = &HC8FFFF
Const m_def_BackSelected = &HC000&
Const m_def_BackSelectedG1 = &HC000&
Const m_def_BackSelectedG2 = &HE0E0E0
Const m_def_HoverColor = &HFF0000
Const m_def_GridLineColor = &H80FF&
Const m_def_nGridWidth = 30
Const m_def_nGridHeight = 30
Const m_def_SelectMode = 0
Const m_def_SelectModeStyle = 0
Const m_def_SelectControlType = 0
Const m_def_CalendarDateOption = 0
Const m_def_ShowWeek = True
'Const m_def_CustomDateFormat = "MMMM D,YYYY"

Const m_def_CalendarOption = 0

Const m_def_CalendarFirstDayOfWeek = 1
Const m_def_DayHeaderFormat = 1
Const m_def_ClickBehivor = 0
Const m_def_HoverSelection = True
Const m_def_HotTracking = True
Const m_def_nBorderStyle = 0
Const m_def_CalendarBorderStyle = 0

Const m_def_CalendarBdHighlightColour = &H80000014 'vb3DHighlight
Const m_def_CalendarBdHighlightDKColour = &H80000016 'vb3DLight
Const m_def_CalendarBdShadowColour = &H80000010 'vb3DShadow
Const m_def_CalendarBdShadowDKColour = &HFF&      'vb3DDKShadow
Const m_def_CalendarBdFlatBorderColour = vbBlack

Const DEF_ACTIVE_DAY_FORECOLOR = &H200FF

Dim m_MaxLength           As Integer
Dim m_BorderColor         As OLE_COLOR
Dim m_BorderColorOver     As OLE_COLOR
Dim m_BorderColorDown     As OLE_COLOR
Dim m_BgColor             As OLE_COLOR
Dim m_BgColorOver         As OLE_COLOR
Dim m_BgColorDown         As OLE_COLOR
Dim m_ButtonBgColor       As OLE_COLOR
Dim m_ButtonBgColorOver   As OLE_COLOR
Dim m_ButtonBgColorDown   As OLE_COLOR
Dim m_ButtonCount         As Long
Dim m_OutputText                As String
Dim mstrText              As String
Dim m_oStartColor         As OLE_COLOR
Dim m_oEndColor           As OLE_COLOR
Dim m_FocusColor          As OLE_COLOR

Dim UsrRect               As RECT
Dim ButtRect              As RECT
Dim ButtRectUp            As RECT
Dim ButtRectDown          As RECT
Dim Ret                   As Long
Dim CrlRet                As Long
Dim IsMOver               As Boolean
Dim IsMDown               As Boolean
Dim IsButtDown            As Boolean
Dim IsCrlOver             As Boolean
Dim Clicked               As Boolean
Dim InFocus               As Boolean
Dim m_DropListEnabled     As Boolean
Dim m_Enabled             As Boolean
Dim m_Icon                As StdPicture
Dim M_IconOK              As Boolean
Dim m_ShowIcon            As Boolean
Dim m_IconSizeWidth       As Long
Dim m_IconSizeHeight      As Long
Dim m_SelectControl       As UserControlType
Dim CurrTime              As Date
Dim temtxt                As String
Dim strFormatWhenEdit     As String
Dim m_DateFormatWhenEdit As dtFormatWhenEditConstants
Dim CurrentSection        As dtSectionConstants
Dim strCustomDateFormat   As String
Dim strMask               As String
Dim strSeperator          As String
Dim strPlaceHolder        As String
Dim bAllowNull            As Boolean
Dim varOldDate            As Variant

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Overall Usercontrol Declare----------------------------------------------------------------------------
Public Enum UserControlType
    [XPCalendar] = 0
End Enum

'## Calendar Declare------------------------------------------------------------------------------------
Public Enum CalendarSelectModeStyle
    [Standard]
    [Gradient_V]
    [Gradient_H]
    [byPicture]
End Enum

Public Enum CalendarSelectPictureDrawModeType
    [PictureStretched] = 0
    [PictureTiled] = 1
End Enum

Public Enum dtDaysOfTheWeek
    [dtSunday] = 1
    [dtMonday] = 2
    [dtTuesday] = 3
    [dtWednesday] = 4
    [dtThursday] = 5
    [dtFriday] = 6
    [dtSaturday] = 7
End Enum
'-----------------------------
Public Enum dtFormatWhenEditConstants
    [dd/mm/yyyy] = 1
    [mm/dd/yyyy] = 2
    [dd/yyyy/mm] = 3
    [mm/yyyy/dd] = 4
    [yyyy/dd/mm] = 5
    [yyyy/mm/dd] = 6
End Enum
Public Enum dtSectionConstants
    dtDaySection
    dtMonthSection
    dtYearSection
    dtInvalid
End Enum
Public Enum dtDateValidationConstants
    dtValidDate
    dtInvalidDate
    dtNullDate
End Enum

Private Const SEPERATORS = "/-\:"
Private Const PLACE_HOLDERS = "_#- "
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
' date format
Private Const LOCALE_SLONGDATE = &H20        'Long date format string
Private Const LOCALE_SSHORTDATE = &H1F       'Short date format string
Private Const m_def_MinDate As Date = "01/01/1901"
Private Const m_def_MaxDate As Date = "31/12/2099"
Private Const m_def_CustomDateFormat = "dd/MM/yyyy"
Private Const m_def_DateFormat = 2
Private Const m_def_AllowNull = True
Private Const m_def_DateFormatWhenEdit = dtFormatWhenEditConstants.[dd/mm/yyyy]
Private Const m_def_Seperator = "/"
Private Const m_def_PlaceHolder = "_"

Public Enum dtFormatConstants
    [dtLongDate] = 1
    [dtShortDate] = 2
    [dtCustom] = 3
End Enum

Public Enum CalendarDateSettingType
    [Default(Now)] = 0
    [UserInput] = 1
End Enum

Public Enum CalendarDateTimeOption
    [DateOnly(Default)] = 0
End Enum

Public Enum dtDayHeaderFormats
    [dtSingleLetter] = 0
    [dtMedium] = 1
    [dtFullName] = 2
End Enum

Enum CalendarBorderStyleType
    [None] = 0
    [Flat] = 1
    [Raised Thin] = 2
    [Raised 3D] = 3
    [Sunken Thin] = 4
    [Sunken 3D] = 5
    [Etched] = 6
    [Bump] = 7
End Enum

Public Enum dtClickBehivor
    [dtOneClickHide] = 0
    [dtDblClickHide] = 1
End Enum

Public Enum PictureSize
    [size16x16] = 0
    [size32x32] = 1
    [SizeDefault] = 2
    [SizeCustom] = 3
    [size64x64] = 4
End Enum

Public Enum DatePositionType
    [PosnTop] = 0
    [PosnCenter] = 1
    [PosnBottom] = 2
End Enum

Public Enum DateAlignType
    [AlignLeft] = 0
    [AlignCenter] = 1
    [AlignRight] = 2
End Enum

Public Enum GridLineType
    [HVLineActiveMonth] = 0
    [HLineActiveMonth] = 1
    [VLineActiveMonth] = 2
    [HVMonthLine] = 3
    [HMonthLine] = 4
    [VMonthLine] = 5
    [None] = 6
End Enum

Event Click()
Event MouseOver()
Event MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event Change()
Event CalendarChoose(Text As String)
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

Public Sub DrawControl(ByVal StatusType As Long)

  Dim CurFontName As String
  Dim Brsh As Long, Clr As Long
  Dim lx As Long, ty As Long
  Dim rx As Long, by As Long
  Dim xx As Long
  Dim sh As Long
  Dim textline As Long
  Dim align As Long
  Dim lr As Long

    On Error Resume Next
      lx = ScaleLeft
      ty = ScaleTop
      rx = ScaleWidth
      by = ScaleHeight

      On Error Resume Next
      SetRect UsrRect, 0, 0, rx, by
      If Not m_Enabled Then StatusType = 0
      Cls
      Select Case m_SelectControl
        Case 0
          Select Case StatusType

              '## Draw Button Normal--No Focus,No Mouse Event
            Case 0

              Call SetRect(UsrRect, 0, 0, rx - by, by)
              OleTranslateColor m_BgColor, ByVal 0&, Clr
              Brsh = CreateSolidBrush(Clr)
              FillRect hdc, UsrRect, Brsh
              DeleteObject Brsh

              Call SetRect(ButtRect, rx - by, 0, rx, by)
              OleTranslateColor m_ButtonBgColor, ByVal 0&, Clr
              Brsh = CreateSolidBrush(Clr)
              FillRect hdc, ButtRect, Brsh
              DeleteObject Brsh

              Call SetRect(ButtRect, rx - by, 0, rx, by)
              If InFocus Then
                  OleTranslateColor m_FocusColor, ByVal 0&, Clr
                Else
                  OleTranslateColor m_BorderColor, ByVal 0&, Clr
              End If
              Brsh = CreateSolidBrush(Clr)
              FrameRect hdc, UsrRect, Brsh
              DeleteObject Brsh

              SetRect UsrRect, 0, 0, rx, by
              If InFocus Then
                  OleTranslateColor m_FocusColor, ByVal 0&, Clr
                Else
                  OleTranslateColor m_BorderColor, ByVal 0&, Clr
              End If
              Brsh = CreateSolidBrush(Clr)
              FrameRect hdc, UsrRect, Brsh
              DeleteObject Brsh

              If m_ShowIcon Then
                  If Not m_Enabled Then
                      lr = DrawState(hdc, 0, 0, m_Icon, 0, rx - (by + m_IconSizeWidth) / 2, (by / 2) - (m_IconSizeHeight / 2), m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_DISABLED)
                    Else
                      lr = DrawState(hdc, 0, 0, m_Icon, 0, rx - (by + m_IconSizeWidth) / 2, (by / 2) - (m_IconSizeHeight / 2), m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_NORMAL)
                  End If

              End If

              '## Draw Button Over
            Case 1

              SetRect UsrRect, 0, 0, rx, by

              OleTranslateColor m_BgColorOver, ByVal 0&, Clr
              Brsh = CreateSolidBrush(Clr)
              FillRect hdc, UsrRect, Brsh
              DeleteObject Brsh

              Call SetRect(ButtRect, rx - by, 0, rx, by)
              OleTranslateColor m_ButtonBgColorOver, ByVal 0&, Clr
              Brsh = CreateSolidBrush(Clr)
              FillRect hdc, ButtRect, Brsh
              DeleteObject Brsh

              Call SetRect(ButtRect, rx - by, 0, rx, by)
              OleTranslateColor m_BorderColorOver, ByVal 0&, Clr
              Brsh = CreateSolidBrush(Clr)
              FrameRect hdc, ButtRect, Brsh
              DeleteObject Brsh

              SetRect UsrRect, 0, 0, rx, by
              OleTranslateColor m_BorderColorOver, ByVal 0&, Clr
              Brsh = CreateSolidBrush(Clr)
              FrameRect hdc, UsrRect, Brsh
              DeleteObject Brsh
              If m_ShowIcon Then
                  Brsh = CreateSolidBrush(RGB(136, 141, 157))
                  lr = DrawState(hdc, Brsh, 0, m_Icon, 0, rx - (by + m_IconSizeWidth) / 2 + 1.5, (by / 2) - (m_IconSizeHeight / 2) + 1.5, m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_MONO)
                  DeleteObject Brsh
                  lr = DrawState(hdc, 0, 0, m_Icon, 0, rx - (by + m_IconSizeWidth) / 2 - 1.5, (by / 2) - (m_IconSizeHeight / 2) - 1.5, m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_NORMAL)
              End If
              'Draw Button Down
            Case 2

              SetRect UsrRect, 0, 0, rx, by

              OleTranslateColor m_BgColorDown, ByVal 0&, Clr
              Brsh = CreateSolidBrush(Clr)
              FillRect hdc, UsrRect, Brsh
              DeleteObject Brsh

              Call SetRect(ButtRect, rx - by, 0, rx, by)
              OleTranslateColor m_ButtonBgColorDown, ByVal 0&, Clr
              Brsh = CreateSolidBrush(Clr)
              FillRect hdc, ButtRect, Brsh
              DeleteObject Brsh
              SetRectEmpty ButtRect

              Call SetRect(ButtRect, rx - by, 0, rx, by)

              OleTranslateColor m_BorderColorDown, ByVal 0&, Clr

              Brsh = CreateSolidBrush(Clr)
              FrameRect hdc, ButtRect, Brsh
              DeleteObject Brsh

              SetRect UsrRect, 0, 0, rx, by
              OleTranslateColor m_BorderColorDown, ByVal 0&, Clr
              Brsh = CreateSolidBrush(Clr)
              FrameRect hdc, UsrRect, Brsh
              DeleteObject Brsh
              If m_ShowIcon Then
                  lr = DrawState(hdc, 0, 0, m_Icon, 0, rx - (by + m_IconSizeWidth) / 2, (by / 2) - (m_IconSizeHeight / 2), m_IconSizeWidth, m_IconSizeHeight, DST_ICON Or DSS_NORMAL)
              End If

          End Select
          If Not m_ShowIcon Then
              CurFontName = Font.Name
              Font.Name = "Marlett"
              If by <> 0 Then
                  Call DrawText(hdc, "6", 1&, ButtRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
              End If
              Font.Name = CurFontName
          End If

      End Select
      DeleteObject m_Icon

End Sub

Public Sub RefreshControl()

    If IsCrlOver Then
        Call DrawControl(2)
      Else
        Call DrawControl(0)
    End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)

    If (KeyCode = vbKeyF4) Or ((KeyCode = vbKeyDown) And (Shift = vbAltMask)) Then
        If m_Enabled = True Then
            RaiseEvent Click
            Call XPCalendarshow(1)
        End If
      ElseIf (KeyCode = vbKeyDelete) Or (KeyCode = vbKeyBack) Then
        Call DeleteNumber(KeyCode)
      ElseIf (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Then
        Call MoveInsertionPoint(KeyCode)
      ElseIf (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Then
        Call ChangeSectionValue(KeyCode)
      ElseIf (KeyCode = vbKeyHome) Then
        Text1.SelStart = 0
        Call MakeSelection(GetCurrentSection())
      ElseIf (KeyCode = vbKeyEnd) Then
        Text1.SelStart = Text1.MaxLength
        Call MakeSelection(GetCurrentSection())
    End If

    KeyCode = 0

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

    If (KeyAscii >= vbKeySpace) Then
        If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
            Call InsertNumber(Chr$(KeyAscii))
          ElseIf (KeyAscii = Asc(strSeperator)) Then
            Call MoveToNextSection
        End If
    End If

    KeyAscii = 0

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)

    Select Case KeyCode
      Case vbKeyInsert
        Value = VBA.Date
        Call Text1_GotFocus
    End Select

End Sub

Private Sub Text1_OLECompleteDrag(Effect As Long)

    RaiseEvent OLECompleteDrag(Effect)

End Sub

Public Sub OLEDrag()

    Text1.OLEDrag

End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)

End Sub

Private Sub Text1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)

    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)

End Sub

Private Sub Text1_OLESetData(Data As DataObject, DataFormat As Integer)

    RaiseEvent OLESetData(Data, DataFormat)

End Sub

Private Sub Text1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)

    RaiseEvent OLEStartDrag(Data, AllowedEffects)

End Sub

Private Sub Text1_Change()

  ' m_Text = Text1.text

    RaiseEvent Change

End Sub

Private Sub Text1_GotFocus()

    If m_Enabled Then
        InFocus = True

  Dim dtDate As Date

        With Text1
            .MaxLength = Len(strFormatWhenEdit)

            If (Len(CStr(Value)) > 0) Then
                Call GetDate(dtDate)
                .Text = VBA.Format$(dtDate, strFormatWhenEdit)
                varOldDate = dtDate
              Else
                If bAllowNull Then
                    .Text = strMask
                    varOldDate = ""
                  Else
                    .Text = VBA.Format$(VBA.Date, strFormatWhenEdit)
                    varOldDate = .Text
                End If
            End If
            mstrText = .Text

        End With
        Call MakeSelection(GetCurrentSection())
    End If

End Sub

Private Sub Text1_LostFocus()

    InFocus = False

  Dim strDate As String
  Dim dtDate As Date
  Dim enuDtType As dtDateValidationConstants

    With Text1
        .MaxLength = 0

        Select Case GetDate(dtDate)
          Case dtValidDate
            Value = dtDate
          Case dtInvalidDate
            Value = varOldDate
          Case dtNullDate
            Value = IIf(bAllowNull, "", varOldDate)
        End Select
    End With

    Call RefreshControl

End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)
    If (GetCurrentSection() <> CurrentSection) Then Call MakeSelection

End Sub

Private Sub UserControl_Initialize()

    Call DrawControl(0)

End Sub

Private Sub UserControl_LostFocus()

    InFocus = False
    Call RefreshControl

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_Enabled Then
        RaiseEvent MouseMove(Button, Shift, X, Y)
        If Not IsButtDown Then
            UserControl_MouseOut Button, Shift, X, Y
        End If
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_Enabled Then
        IsButtDown = False
        'If m_DropListEnabled = True Then

        Call DrawControl(0)
        Select Case m_SelectControl
          Case 0

            If (X >= ButtRect.Left And X <= ButtRect.Right) And (Y >= ButtRect.Top And Y <= ButtRect.Bottom) Then

                Select Case m_SelectControl
                  Case 0
                    RaiseEvent Click
                    Call XPCalendarshow(1)
                End Select
            End If

        End Select

    End If

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next

      '## Single Line Textbox
      Text1.Move 2, (ScaleHeight / 2) - (Text1.Height / 2), ScaleWidth - ScaleHeight - 4

      Call RefreshControl

End Sub

Function UserControl_MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)

  '## Also can use SubClass to Trace Mouse Movement

    IsCrlOver = False
    IsMOver = False
    Select Case m_SelectControl
      Case 0 '## Case XPCalendar
        If (X >= ButtRect.Left And X <= ButtRect.Right) And (Y >= ButtRect.Top And Y <= ButtRect.Bottom) Then
            If IsMOver = False Then

                IsMOver = True
                Ret = SetCapture(UserControl.hwnd)

                RaiseEvent MouseOver
                Call DrawControl(1)
            End If
          Else
            IsMOver = False
            Ret = ReleaseCapture()

        End If

        If (X >= 0 And X <= ScaleWidth) And (Y >= 0 And Y <= ScaleHeight) Then

            If IsCrlOver = False Then

                IsCrlOver = True
                CrlRet = SetCapture(UserControl.hwnd)

                RaiseEvent MouseOver
                Call DrawControl(1)
            End If
          Else
            IsMOver = False
            IsCrlOver = False
            CrlRet = ReleaseCapture()
            RaiseEvent MouseOut(Button, Shift, X, Y)
            Call DrawControl(0)
        End If

    End Select

End Function

Public Property Get ShowWeek() As Boolean

    ShowWeek = m_ShowWeek

End Property

Public Property Let ShowWeek(ByVal New_ShowWeek As Boolean)

    m_ShowWeek = New_ShowWeek
    PropertyChanged "ShowWeek"

End Property

Public Property Get Enabled() As Boolean

    Enabled = m_Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    Call RefreshControl

End Property

Public Property Get ShowIcon() As Boolean

    ShowIcon = m_ShowIcon

End Property

Public Property Let ShowIcon(ByVal New_ShowIcon As Boolean)

    m_ShowIcon = New_ShowIcon
    PropertyChanged "ShowIcon"
    RefreshControl

End Property

Public Property Get Icon() As StdPicture

    Set Icon = m_Icon

End Property

Public Property Set Icon(ByVal New_Icon As StdPicture)

    Set m_Icon = New_Icon

    PropertyChanged "Icon"
    Call RefreshControl

End Property

Public Property Get IconSizeWidth() As Long

    IconSizeWidth = m_IconSizeWidth

End Property

Public Property Let IconSizeWidth(ByVal New_IconSizeWidth As Long)

    m_IconSizeWidth = New_IconSizeWidth
    PropertyChanged "IconSizeWidth"
    Call RefreshControl

End Property

Public Property Get IconSizeHeight() As Long

    IconSizeHeight = m_IconSizeHeight

End Property

Public Property Let IconSizeHeight(ByVal New_IconSizeHeight As Long)

    m_IconSizeHeight = New_IconSizeHeight
    PropertyChanged "IconSizeHeight"
    Call RefreshControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColor() As OLE_COLOR

    BorderColor = m_BorderColor

End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)

    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    Call RefreshControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColorOver() As OLE_COLOR

    BorderColorOver = m_BorderColorOver

End Property

Public Property Let BorderColorOver(ByVal New_BorderColorOver As OLE_COLOR)

    m_BorderColorOver = New_BorderColorOver
    PropertyChanged "BorderColorOver"
    Call RefreshControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColorDown() As OLE_COLOR

    BorderColorDown = m_BorderColorDown

End Property

Public Property Let BorderColorDown(ByVal New_BorderColorDown As OLE_COLOR)

    m_BorderColorDown = New_BorderColorDown
    PropertyChanged "BorderColorDown"
    Call RefreshControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BgColor() As OLE_COLOR

    BgColor = m_BgColor

End Property

Public Property Let BgColor(ByVal New_BgColor As OLE_COLOR)

    m_BgColor = New_BgColor
    PropertyChanged "BgColor"
    Call RefreshControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BgColorOver() As OLE_COLOR

    BgColorOver = m_BgColorOver

End Property

Public Property Let BgColorOver(ByVal New_BgColorOver As OLE_COLOR)

    m_BgColorOver = New_BgColorOver
    PropertyChanged "BgColorOver"
    Call RefreshControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BgColorDown() As OLE_COLOR

    BgColorDown = m_BgColorDown

End Property

Public Property Let BgColorDown(ByVal New_BgColorDown As OLE_COLOR)

    m_BgColorDown = New_BgColorDown
    PropertyChanged "BgColorDown"
    Call RefreshControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonBgColor() As OLE_COLOR

    ButtonBgColor = m_ButtonBgColor

End Property

Public Property Let ButtonBgColor(ByVal New_ButtonBgColor As OLE_COLOR)

    m_ButtonBgColor = New_ButtonBgColor
    PropertyChanged "ButtonBgColor"
    Call RefreshControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonBgColorOver() As OLE_COLOR

    ButtonBgColorOver = m_ButtonBgColorOver

End Property

Public Property Let ButtonBgColorOver(ByVal New_ButtonBgColorOver As OLE_COLOR)

    m_ButtonBgColorOver = New_ButtonBgColorOver
    PropertyChanged "ButtonBgColorOver"
    Call RefreshControl

End Property

Public Property Get ButtonBgColorDown() As OLE_COLOR

    ButtonBgColorDown = m_ButtonBgColorDown

End Property

Public Property Let ButtonBgColorDown(ByVal New_ButtonBgColorDown As OLE_COLOR)

    m_ButtonBgColorDown = New_ButtonBgColorDown
    PropertyChanged "ButtonBgColorDown"
    Call RefreshControl

End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_Enabled = m_def_Enabled
    m_ShowIcon = m_def_ShowIcon
    m_IconSizeWidth = m_def_IconSizeWidth
    m_IconSizeHeight = m_def_IconSizeHeight
    m_BorderColor = m_def_BorderColor
    m_BorderColorOver = m_def_BorderColorOver
    m_BorderColorDown = m_def_BorderColorDown
    m_BgColor = m_def_BgColor
    m_BgColorOver = m_def_BgColorOver
    m_BgColorDown = m_def_BgColorDown
    m_ButtonBgColor = m_def_ButtonBgColor
    m_ButtonBgColorOver = m_def_ButtonBgColorOver
    m_ButtonBgColorDown = m_def_ButtonBgColorDown
    m_CalendarBdHighlightColour = m_def_CalendarBdHighlightColour
    m_CalendarBdHighlightDKColour = m_def_CalendarBdHighlightDKColour
    m_CalendarBdShadowColour = m_def_CalendarBdShadowColour
    m_CalendarBdShadowDKColour = m_def_CalendarBdShadowDKColour
    m_CalendarBdFlatBorderColour = m_def_CalendarBdFlatBorderColour
    m_GridLineColor = m_def_GridLineColor
    m_OutputText = m_def_OutputText

    m_oStartColor = vbWhite
    m_oEndColor = vbButtonFace

    m_FocusColor = m_def_FocusColor

    m_DropListEnabled = True

    '## Calendar initial Properties

    Set m_SelectionPicture = Nothing

    m_BackNormal = m_def_BackNormal
    m_BackSelected = m_def_BackSelected
    m_BackSelectedG1 = m_def_BackSelectedG1
    m_BackSelectedG2 = m_def_BackSelectedG2
    m_SelectControl = m_def_SelectControlType
    m_SelectModeStyle = m_def_SelectModeStyle
    m_CalendarFirstDayOfWeek = m_def_CalendarFirstDayOfWeek
    m_CalendarDayHeaderFormat = m_def_DayHeaderFormat

    m_DateFormat = m_def_DateFormat
    'mstrCustomFormat = m_def_CustomFormat

    m_CalendarOption = m_def_CalendarOption
    m_CalendarClickBehivor = m_def_ClickBehivor

    m_nGridWidth = m_def_nGridWidth
    m_nGridHeight = m_def_nGridHeight
    m_HoverSelection = m_def_HoverSelection
    m_HotTracking = m_def_HotTracking
    m_HoverColor = m_def_HoverColor

    With m_TodayFont
        .Name = "Tahoma"
        .Size = 8
        .Bold = True
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With

    m_SelectedBMPMaskColor = 0
    m_TodayMaskColor = 0
    m_WeekSignPicMaskColor = 0
    '---------------------------------------
    m_ShowWeek = m_def_ShowWeek
    m_dtMaxDate = m_def_MaxDate
    m_dtMinDate = m_def_MinDate
    strSeperator = m_def_Seperator
    strPlaceHolder = m_def_PlaceHolder
    m_DateFormat = m_def_DateFormat
    strCustomDateFormat = m_def_CustomDateFormat
    bAllowNull = m_def_AllowNull
    m_DateFormatWhenEdit = m_def_DateFormatWhenEdit
    strDateFormat = GetFormat()
    strFormatWhenEdit = GetEditFormatString()
    Value = Date

    Call RefreshControl

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    m_ShowIcon = PropBag.ReadProperty("ShowIcon", m_def_ShowIcon)
    m_IconSizeWidth = PropBag.ReadProperty("IconSizeWidth", m_def_IconSizeWidth)
    m_IconSizeHeight = PropBag.ReadProperty("IconSizeHeight", m_def_IconSizeHeight)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderColorOver = PropBag.ReadProperty("BorderColorOver", m_def_BorderColorOver)
    m_BorderColorDown = PropBag.ReadProperty("BorderColorDown", m_def_BorderColorDown)
    m_BgColor = PropBag.ReadProperty("BgColor", m_def_BgColor)
    m_BgColorOver = PropBag.ReadProperty("BgColorOver", m_def_BgColorOver)
    m_BgColorDown = PropBag.ReadProperty("BgColorDown", m_def_BgColorDown)
    m_ButtonBgColor = PropBag.ReadProperty("ButtonBgColor", m_def_ButtonBgColor)
    m_ButtonBgColorOver = PropBag.ReadProperty("ButtonBgColorOver", m_def_ButtonBgColorOver)
    m_ButtonBgColorDown = PropBag.ReadProperty("ButtonBgColorDown", m_def_ButtonBgColorDown)

    m_OutputText = PropBag.ReadProperty("OutputText", m_def_OutputText)

    m_FocusColor = PropBag.ReadProperty("FocusColor", m_def_FocusColor)
    m_MaxLength = PropBag.ReadProperty("TextMaxLength", 30)
    Text1.MaxLength = PropBag.ReadProperty("TextMaxLength", 0)
    'Text1.ForeColor = PropBag.ReadProperty("Text_ForeColor", &H80000008)
    Text1.Enabled = PropBag.ReadProperty("Text_Enabled", True)
    Text1.Locked = PropBag.ReadProperty("Text_Locked", False)
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Text1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Text1.FontBold = PropBag.ReadProperty("FontBold", Ambient.Font.Bold)
    Text1.FontItalic = PropBag.ReadProperty("FontItalic", Ambient.Font.Italic)
    Text1.FontName = PropBag.ReadProperty("FontName", Ambient.Font.Name)
    Text1.FontSize = PropBag.ReadProperty("FontSize", Ambient.Font.Size)
    Text1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", Ambient.Font.Strikethrough)
    Text1.FontUnderline = PropBag.ReadProperty("FontUnderline", Ambient.Font.Underline)

    m_DropListEnabled = PropBag.ReadProperty("DropListEnabled", True)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)

    '## Calendar Read Properties
    m_CalendarSelectPicDrawMode = PropBag.ReadProperty("CalendarSelectPicDrawMode", 0)
    m_SelectModeStyle = PropBag.ReadProperty("CalendarSelectStyle", m_def_SelectModeStyle)

    m_SelectedBMPMaskColor = PropBag.ReadProperty("CalendarSelectePicMaskColor", 0)
    Set m_SelectionPicture = PropBag.ReadProperty("CalendarSelectePicture", Nothing)
    m_BackNormal = PropBag.ReadProperty("CalendarBackNormal", m_def_BackNormal)

    m_BackSelected = PropBag.ReadProperty("CalendarBackSelected", m_def_BackSelected)
    m_BackSelectedG1 = PropBag.ReadProperty("CalendarBackSelectedG1", m_def_BackSelectedG1)
    m_BackSelectedG2 = PropBag.ReadProperty("CalendarBackSelectedG2", m_def_BackSelectedG2)
    '----------------------------------------------
    m_dtMaxDate = PropBag.ReadProperty("MaxDate", m_def_MaxDate)
    m_dtMinDate = PropBag.ReadProperty("MinDate", m_def_MinDate)
    Text1.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    strCustomDateFormat = PropBag.ReadProperty("CustomDateFormat", m_def_CustomDateFormat)
    bAllowNull = PropBag.ReadProperty("AllowNull", m_def_AllowNull)
    strSeperator = PropBag.ReadProperty("Seperator", m_def_Seperator)
    strPlaceHolder = PropBag.ReadProperty("PlaceHolder", m_def_PlaceHolder)
    Text1.OLEDragMode = PropBag.ReadProperty("OLEDragMode", OLEDragConstants.vbOLEDragManual)
    Text1.OLEDropMode = PropBag.ReadProperty("OLEDropMode", OLEDropConstants.vbOLEDropNone)
    Text1.MousePointer = PropBag.ReadProperty("MousePointer", MousePointerConstants.vbDefault)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_DateFormat = PropBag.ReadProperty("DateFormat", m_def_DateFormat)
    strDateFormat = GetFormat()
    m_DateFormatWhenEdit = PropBag.ReadProperty("DateFormatWhenEdit", m_def_DateFormatWhenEdit)
    strFormatWhenEdit = GetEditFormatString()
    strMask = GetMaskString()

    m_Value = PropBag.ReadProperty("Value", Date)
    If (Len(CStr(m_Value)) = 0) Then
        If bAllowNull Then
            Text1.Text = IIf(Ambient.UserMode, GetMaskString(), "")
          Else
            m_Value = VBA.Date
            Text1.Text = VBA.Format$(m_Value, strDateFormat)
            mstrText = VBA.Format$(m_Value, strFormatWhenEdit)
        End If
      Else
        Text1.Text = VBA.Format$(m_Value, strDateFormat)
        mstrText = VBA.Format$(m_Value, strFormatWhenEdit)
    End If

    m_CalendarFirstDayOfWeek = PropBag.ReadProperty("CalendarFirstDayOfWeek", m_def_CalendarFirstDayOfWeek)
    m_CalendarDayHeaderFormat = PropBag.ReadProperty("CalendarDayHeaderFormat", m_def_DayHeaderFormat)
    m_CalendarOption = PropBag.ReadProperty("CalendarOption", m_def_CalendarOption)
    m_CalendarClickBehivor = PropBag.ReadProperty("CalendarClickBehivor", m_def_ClickBehivor)
    m_SelectControl = PropBag.ReadProperty("SelectControl", m_def_SelectControlType)
    m_nGridWidth = PropBag.ReadProperty("CalendarGridWidth", m_def_nGridWidth)
    m_nGridHeight = PropBag.ReadProperty("CalendarGridHeight", m_def_nGridHeight)
    m_HoverSelection = PropBag.ReadProperty("HoverSelection", m_def_HoverSelection)
    m_HotTracking = PropBag.ReadProperty("HotTracking", m_def_HotTracking)
    m_HoverColor = PropBag.ReadProperty("HoverColor", m_def_HoverColor)
    m_CalendarBorderStyle = PropBag.ReadProperty("CalendarBorderStyle", m_def_CalendarBorderStyle)

    'Fonts
    m_TodayForeColor = PropBag.ReadProperty("TodayForeColor", DEF_ACTIVE_DAY_FORECOLOR)
    '~~ Active Day font
    With m_TodayFont
        .Name = PropBag.ReadProperty("TodayFontName", "Tahoma")
        .Size = PropBag.ReadProperty("TodayFontSize", 8)
        .Bold = PropBag.ReadProperty("TodayFontBold", True)
        .Italic = PropBag.ReadProperty("TodayFontItalic", True)
        .Underline = PropBag.ReadProperty("TodayFontUnderline", False)
        .Strikethrough = PropBag.ReadProperty("TodayFontStrikethrough", False)
    End With

    m_TodayMaskColor = PropBag.ReadProperty("TodayMaskColor", 0)
    Set m_TodayPicture = PropBag.ReadProperty("TodayPicture", Nothing)
    m_TodayPictureWidth = PropBag.ReadProperty("TodayPictureWidth", 32)
    m_TodayPictureHeight = PropBag.ReadProperty("TodayPictureHeight", 32)
    m_TodayPictureSize = PropBag.ReadProperty("TodayPictureSize", 1)
    m_TodayOriginalPicSizeW = PropBag.ReadProperty("TodayOriginalPicSizeW", 32)
    m_TodayOriginalPicSizeH = PropBag.ReadProperty("TodayOriginalPicSizeH", 32)
    m_DateAlign = PropBag.ReadProperty("DateAlign", 0)
    m_DatePosn = PropBag.ReadProperty("DatePosn", 0)
    m_DateXOffset = PropBag.ReadProperty("DateXOffset", 0)
    m_DateYOffset = PropBag.ReadProperty("DateYOffset", 0)
    m_TodayXOffset = PropBag.ReadProperty("TodayXOffset", 0)
    m_TodayYOffset = PropBag.ReadProperty("TodayYOffset", 0)
    m_GridLineColor = PropBag.ReadProperty("GridLineColor", m_def_GridLineColor)
    m_Gridline = PropBag.ReadProperty("Gridline", 0)
    m_CalendarBdHighlightColour = PropBag.ReadProperty("CalendarBdHighlightColour", m_def_CalendarBdHighlightColour)
    m_CalendarBdHighlightDKColour = PropBag.ReadProperty("CalendarBdHighlightDKColour", m_def_CalendarBdHighlightDKColour)
    m_CalendarBdShadowColour = PropBag.ReadProperty("CalendarBdShadowColour", m_def_CalendarBdShadowColour)
    m_CalendarBdShadowDKColour = PropBag.ReadProperty("CalendarBdShadowDKColour", m_def_CalendarBdShadowDKColour)
    m_CalendarBdFlatBorderColour = PropBag.ReadProperty("CalendarBdFlatBorderColour", m_def_CalendarBdFlatBorderColour)

    m_ShowWeek = PropBag.ReadProperty("ShowWeek", m_def_ShowWeek)
    m_WeekColumnWidth = PropBag.ReadProperty("WeekColumnWidth", m_nGridWidth)
    m_ShowWeekSignPicture = PropBag.ReadProperty("ShowWeekSignPicture", True)
    m_WeekSignPicXOffset = PropBag.ReadProperty("WeekSignPicXOffset", 0)
    m_WeekSignPicYOffset = PropBag.ReadProperty("WeekSignPicYOffset", 0)
    m_WeekSignPicMaskColor = PropBag.ReadProperty("WeekSignPicMaskColor", 0)
    Set m_WeekSignPicture = PropBag.ReadProperty("WeekSignPicture", Nothing)
    m_WeekSignPictureWidth = PropBag.ReadProperty("WeekSignPictureWidth", 32)
    m_WeekSignPictureHeight = PropBag.ReadProperty("WeekSignPictureHeight", 32)
    m_WeekSignPictureSize = PropBag.ReadProperty("WeekSignPictureSize", 1)
    m_WeekSignOriginalPicSizeW = PropBag.ReadProperty("WeekSignPictureOriginalPicSizeW", 32)
    m_WeekSignOriginalPicSizeH = PropBag.ReadProperty("WeekSignPictureOriginalPicSizeH", 32)

    Call RefreshControl

End Sub

Private Sub UserControl_Terminate()

    Set m_TodayFont = Nothing
    Set m_Icon = Nothing
    Set m_SelectionPicture = Nothing
    Set m_WeekSignPicture = Nothing
    Set m_TodayPicture = Nothing
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("ShowIcon", m_ShowIcon, m_def_ShowIcon)
    Call PropBag.WriteProperty("IconSizeWidth", m_IconSizeWidth, m_def_IconSizeWidth)
    Call PropBag.WriteProperty("IconSizeHeight", m_IconSizeHeight, m_def_IconSizeHeight)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderColorOver", m_BorderColorOver, m_def_BorderColorOver)
    Call PropBag.WriteProperty("BorderColorDown", m_BorderColorDown, m_def_BorderColorDown)
    Call PropBag.WriteProperty("BgColor", m_BgColor, m_def_BgColor)
    Call PropBag.WriteProperty("BgColorOver", m_BgColorOver, m_def_BgColorOver)
    Call PropBag.WriteProperty("BgColorDown", m_BgColorDown, m_def_BgColorDown)
    Call PropBag.WriteProperty("ButtonBgColor", m_ButtonBgColor, m_def_ButtonBgColor)
    Call PropBag.WriteProperty("ButtonBgColorOver", m_ButtonBgColorOver, m_def_ButtonBgColorOver)
    Call PropBag.WriteProperty("ButtonBgColorDown", m_ButtonBgColorDown, m_def_ButtonBgColorDown)
    Call PropBag.WriteProperty("OutputText", m_OutputText, m_def_OutputText)
    'Call PropBag.WriteProperty("TextMaxLength", Text1.MaxLength, 0)
    Call PropBag.WriteProperty("TextMaxLength", m_MaxLength, 30)
    Call PropBag.WriteProperty("FocusColor", m_FocusColor, m_def_FocusColor)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BackColor", Text1.BackColor, &H80000005)
    Call PropBag.WriteProperty("FontBold", Text1.FontBold, Ambient.Font.Bold)
    Call PropBag.WriteProperty("FontItalic", Text1.FontItalic, Ambient.Font.Italic)
    Call PropBag.WriteProperty("FontName", Text1.FontName, Ambient.Font.Name)
    Call PropBag.WriteProperty("FontSize", Text1.FontSize, Ambient.Font.Size)
    Call PropBag.WriteProperty("FontStrikethru", Text1.FontStrikethru, Ambient.Font.Strikethrough)
    Call PropBag.WriteProperty("FontUnderline", Text1.FontUnderline, Ambient.Font.Underline)
    Call PropBag.WriteProperty("Text_Enabled", Text1.Enabled, True)
    Call PropBag.WriteProperty("Text_Locked", Text1.Locked, False)
    Call PropBag.WriteProperty("DropListEnabled", m_DropListEnabled, True)

    '## Calendar Write Properties

    Call PropBag.WriteProperty("CalendarSelectPicDrawMode", m_CalendarSelectPicDrawMode, 0)
    Call PropBag.WriteProperty("CalendarSelectStyle", m_SelectModeStyle, m_def_SelectModeStyle)
    Call PropBag.WriteProperty("CalendarSelectePicMaskColor", m_SelectedBMPMaskColor, 0)
    Call PropBag.WriteProperty("CalendarSelectePicture", m_SelectionPicture, Nothing)
    Call PropBag.WriteProperty("CalendarBackNormal", m_BackNormal, m_def_BackNormal)
    Call PropBag.WriteProperty("CalendarBackSelected", m_BackSelected, m_def_BackSelected)
    Call PropBag.WriteProperty("CalendarBackSelectedG1", m_BackSelectedG1, m_def_BackSelectedG1)
    Call PropBag.WriteProperty("CalendarBackSelectedG2", m_BackSelectedG2, m_def_BackSelectedG2)
    '-------------------------------------------------------------
    Call PropBag.WriteProperty("DateFormat", m_DateFormat, m_def_DateFormat)
    Call PropBag.WriteProperty("MaxDate", m_dtMaxDate, m_def_MaxDate)
    Call PropBag.WriteProperty("MinDate", m_dtMinDate, m_def_MinDate)
    Call PropBag.WriteProperty("RightToLeft", Text1.RightToLeft, False)
    Call PropBag.WriteProperty("CustomDateFormat", strCustomDateFormat, m_def_CustomDateFormat)
    Call PropBag.WriteProperty("AllowNull", bAllowNull, m_def_AllowNull)
    Call PropBag.WriteProperty("DateFormatWhenEdit", m_DateFormatWhenEdit, m_def_DateFormatWhenEdit)
    Call PropBag.WriteProperty("Seperator", strSeperator, m_def_Seperator)
    Call PropBag.WriteProperty("PlaceHolder", strPlaceHolder, m_def_PlaceHolder)
    Call PropBag.WriteProperty("Value", m_Value, Date)
    Call PropBag.WriteProperty("OLEDragMode", Text1.OLEDragMode, OLEDragConstants.vbOLEDragManual)
    Call PropBag.WriteProperty("OLEDropMode", Text1.OLEDropMode, OLEDropConstants.vbOLEDropNone)
    Call PropBag.WriteProperty("MousePointer", Text1.MousePointer, MousePointerConstants.vbDefault)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("CalendarOption", m_CalendarOption, m_def_CalendarOption)
    Call PropBag.WriteProperty("CalendarFirstDayOfWeek", m_CalendarFirstDayOfWeek, m_def_CalendarFirstDayOfWeek)
    Call PropBag.WriteProperty("CalendarDayHeaderFormat", m_CalendarDayHeaderFormat, m_def_DayHeaderFormat)
    Call PropBag.WriteProperty("CalendarClickBehivor", m_CalendarClickBehivor, m_def_ClickBehivor)
    Call PropBag.WriteProperty("SelectControl", m_SelectControl, m_def_SelectControlType)
    Call PropBag.WriteProperty("CalendarGridWidth", m_nGridWidth, m_def_nGridWidth)
    Call PropBag.WriteProperty("CalendarGridHeight", m_nGridHeight, m_def_nGridHeight)
    Call PropBag.WriteProperty("HoverSelection", m_HoverSelection, m_def_HoverSelection)
    Call PropBag.WriteProperty("HotTracking", m_HotTracking, m_def_HotTracking)
    Call PropBag.WriteProperty("HoverColor", m_HoverColor, m_def_HoverColor)
    Call PropBag.WriteProperty("CalendarBorderStyle", m_CalendarBorderStyle, m_def_CalendarBorderStyle)
    Call PropBag.WriteProperty("TodayForeColor", m_TodayForeColor, DEF_ACTIVE_DAY_FORECOLOR)
    '~~ Active day font
    With m_TodayFont
        Call PropBag.WriteProperty("TodayFontName", .Name, "Tahoma")
        Call PropBag.WriteProperty("TodayFontSize", .Size, 8)
        Call PropBag.WriteProperty("TodayFontBold", .Bold, True)
        Call PropBag.WriteProperty("TodayFontItalic", .Italic, True)
        Call PropBag.WriteProperty("TodayFontUnderline", .Underline, False)
        Call PropBag.WriteProperty("TodayFontStrikethrough", .Strikethrough, False)
    End With

    Call PropBag.WriteProperty("TodayMaskColor", m_TodayMaskColor, 0)
    Call PropBag.WriteProperty("TodayPicture", m_TodayPicture, Nothing)
    Call PropBag.WriteProperty("TodayPictureWidth", m_TodayPictureWidth, 32)
    Call PropBag.WriteProperty("TodayPictureHeight", m_TodayPictureHeight, 32)
    Call PropBag.WriteProperty("TodayPictureSize", m_TodayPictureSize, 1)
    Call PropBag.WriteProperty("TodayOriginalPicSizeW", m_TodayOriginalPicSizeW, 32)
    Call PropBag.WriteProperty("TodayOriginalPicSizeH", m_TodayOriginalPicSizeH, 32)
    Call PropBag.WriteProperty("DateAlign", m_DateAlign, 0)
    Call PropBag.WriteProperty("DatePosn", m_DatePosn, 0)
    Call PropBag.WriteProperty("DateXOffset", m_DateXOffset, 0)
    Call PropBag.WriteProperty("DateYOffset", m_DateYOffset, 0)
    Call PropBag.WriteProperty("TodayXOffset", m_TodayXOffset, 0)
    Call PropBag.WriteProperty("TodayYOffset", m_TodayYOffset, 0)
    Call PropBag.WriteProperty("GridLineColor", m_GridLineColor, m_def_GridLineColor)
    Call PropBag.WriteProperty("Gridline", m_Gridline, 0)
    Call PropBag.WriteProperty("CalendarBdHighlightColour", m_CalendarBdHighlightColour, m_def_CalendarBdHighlightColour)
    Call PropBag.WriteProperty("CalendarBdHighlightDKColour", m_CalendarBdHighlightDKColour, m_def_CalendarBdHighlightDKColour)
    Call PropBag.WriteProperty("CalendarBdShadowColour", m_CalendarBdShadowColour, m_def_CalendarBdShadowColour)
    Call PropBag.WriteProperty("CalendarBdShadowDKColour", m_CalendarBdShadowDKColour, m_def_CalendarBdShadowDKColour)
    Call PropBag.WriteProperty("CalendarBdFlatBorderColour", m_CalendarBdFlatBorderColour, m_def_CalendarBdFlatBorderColour)

    Call PropBag.WriteProperty("ShowWeekSignPicture", m_ShowWeekSignPicture, True)
    Call PropBag.WriteProperty("WeekColumnWidth", m_WeekColumnWidth, m_nGridWidth)
    Call PropBag.WriteProperty("ShowWeek", m_ShowWeek, m_def_ShowWeek)
    Call PropBag.WriteProperty("WeekSignPicXOffset", m_WeekSignPicXOffset, 0)
    Call PropBag.WriteProperty("WeekSignPicYOffset", m_WeekSignPicYOffset, 0)
    Call PropBag.WriteProperty("WeekSignPicMaskColor", m_WeekSignPicMaskColor, 0)
    Call PropBag.WriteProperty("WeekSignPicture", m_WeekSignPicture, Nothing)
    Call PropBag.WriteProperty("WeekSignPictureWidth", m_WeekSignPictureWidth, 32)
    Call PropBag.WriteProperty("WeekSignPictureHeight", m_WeekSignPictureHeight, 32)
    Call PropBag.WriteProperty("WeekSignPictureSize", m_WeekSignPictureSize, 1)
    Call PropBag.WriteProperty("WeekSignPictureOriginalPicSizeW", m_WeekSignOriginalPicSizeW, 32)
    Call PropBag.WriteProperty("WeekSignPictureOriginalPicSizeH", m_WeekSignOriginalPicSizeH, 32)

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If m_Enabled Then
        RaiseEvent KeyDown(KeyCode, Shift)
        Call DrawControl(1)
    End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    If m_Enabled Then

        RaiseEvent KeyPress(KeyAscii)
    End If

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    If m_Enabled Then
        RaiseEvent KeyUp(KeyCode, Shift)
        Call DrawControl(0)
    End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)
    Select Case m_SelectControl
      Case 0
        If (X >= ButtRect.Left And X <= ButtRect.Right) And (Y >= ButtRect.Top And Y <= ButtRect.Bottom) Then
            IsButtDown = True
            Call DrawControl(2)
          Else
            IsButtDown = False
        End If

    End Select

End Sub

Public Sub XPCalendarshow(Show As Integer)

    If Show = 0 And IsWindowVisible(frmCalendar.hwnd) = 0 Then
        GoTo ShowDropDown_Exit
      ElseIf Show = 1 And (IsWindowVisible(frmCalendar.hwnd) <> 0 Or m_DropListEnabled = False) Then
        GoTo ShowDropDown_Exit
    End If

    If Show Then

  Dim ClrPos      As RECT
  Dim crx         As Long
  Dim varValue    As Variant

        '~~ Call 'LostFocus' event of textbox for updating Value of Date if Text1.text has any valid changes
        '   Then this Value is given to frmCalendar.CurrDate to Highlight Date Cell
        Call Text1_LostFocus
        '~~ Save old date
        varValue = Me.Value

        Call GetWindowRect(hwnd, ClrPos)
        Call RefreshControl

        Load frmCalendar
        With frmCalendar
            .Left = ClrPos.Left * Screen.TwipsPerPixelX
            .Top = ClrPos.Bottom * Screen.TwipsPerPixelY
            .IsSelected = False

            'Modal Mode
            .Show 1
            If Not IsNull(CStr(.SelectedDate)) Then
                If .IsSelected And IsInRange(.SelectedDate) = True Then
                    Me.Value = .SelectedDate
                    RaiseEvent CalendarChoose(.SelectedDate)
                  Else
                    Me.Value = varValue
                End If
              Else
                Me.Value = varValue
            End If
            Unload frmCalendar
            Set frmCalendar = Nothing
            Call SetRectEmpty(ClrPos)

        End With
        IsMDown = False
        IsCrlOver = False
        'Call DrawControl(2)
        Call RefreshControl

    End If

ShowDropDown_Exit:

    Exit Sub

End Sub

'## ControlType Select-------------------------------------------------------------
Public Property Get SelectControl() As UserControlType

    SelectControl = m_SelectControl

End Property

Public Property Let SelectControl(ByVal New_SelectControl As UserControlType)

    m_SelectControl = New_SelectControl

    PropertyChanged "SelectControl"

End Property

'## Calendar Select Picture Draw Type --------------------------------------------------------------
Public Property Get CalendarSelectPicDrawMode() As CalendarSelectPictureDrawModeType

    CalendarSelectPicDrawMode = m_CalendarSelectPicDrawMode

End Property

Public Property Let CalendarSelectPicDrawMode(ByVal New_CalendarSelectPicDrawMode As CalendarSelectPictureDrawModeType)

    m_CalendarSelectPicDrawMode = New_CalendarSelectPicDrawMode

    PropertyChanged "CalendarSelectPicDrawMode"

End Property

'## Calendar SelectModeStyle --------------------------------------------------------------
Public Property Get CalendarSelectStyle() As CalendarSelectModeStyle

    CalendarSelectStyle = m_SelectModeStyle

End Property

Public Property Let CalendarSelectStyle(ByVal New_CalendarSelectStyle As CalendarSelectModeStyle)

    m_SelectModeStyle = New_CalendarSelectStyle

    PropertyChanged "CalendarSelectStyle"

End Property

Public Property Get CalendarSelectePicMaskColor() As OLE_COLOR

    CalendarSelectePicMaskColor = m_SelectedBMPMaskColor

End Property

Public Property Let CalendarSelectePicMaskColor(ByVal New_CalendarSelectePicMaskColorr As OLE_COLOR)

    m_SelectedBMPMaskColor = New_CalendarSelectePicMaskColorr
    PropertyChanged "CalendarSelectePicMaskColor"

End Property

'## Calendar SelectionPicture -------------------------------------------------------------
Public Property Get CalendarSelectePicture() As Picture

    Set CalendarSelectePicture = m_SelectionPicture

End Property

Public Property Set CalendarSelectePicture(ByVal New_CalendarSelectePicture As Picture)

    Set m_SelectionPicture = New_CalendarSelectePicture

    PropertyChanged "CalendarSelectePicture"

End Property

'## BackNormal -------------------------------------------------------------------
Public Property Get CalendarBackNormal() As OLE_COLOR

    CalendarBackNormal = m_BackNormal

End Property

Public Property Let CalendarBackNormal(ByVal New_CalendarBackNormal As OLE_COLOR)

    m_BackNormal = New_CalendarBackNormal
    cBackNrm = GetLngColor(m_BackNormal)
    PropertyChanged "CalendarBackNormal"

End Property

'## BackSelected -----------------------------------------------------------------
Public Property Get CalendarBackSelected() As OLE_COLOR

    CalendarBackSelected = m_BackSelected

End Property

Public Property Let CalendarBackSelected(ByVal New_CalendarBackSelected As OLE_COLOR)

    m_BackSelected = New_CalendarBackSelected
    cBackSel = GetLngColor(m_BackSelected)
    PropertyChanged "CalendarBackSelected"

End Property

'## BackSelectedG1 ---------------------------------------------------------------
Public Property Get CalendarBackSelectedG1() As OLE_COLOR

    CalendarBackSelectedG1 = m_BackSelectedG1

End Property

Public Property Let CalendarBackSelectedG1(ByVal New_CalendarBackSelectedG1 As OLE_COLOR)

    m_BackSelectedG1 = New_CalendarBackSelectedG1
    cGrad1 = GetRGBColors(GetLngColor(m_BackSelectedG1))
    PropertyChanged "CalendarBackSelectedG1"

End Property

'## BackSelectedG2 ---------------------------------------------------------------
Public Property Get CalendarBackSelectedG2() As OLE_COLOR

    CalendarBackSelectedG2 = m_BackSelectedG2

End Property

Public Property Let CalendarBackSelectedG2(ByVal New_CalendarBackSelectedG2 As OLE_COLOR)

    m_BackSelectedG2 = New_CalendarBackSelectedG2
    cGrad2 = GetRGBColors(GetLngColor(m_BackSelectedG2))
    PropertyChanged "CalendarBackSelectedG2"

End Property

Public Property Get CalendarFirstDayOfWeek() As dtDaysOfTheWeek

    CalendarFirstDayOfWeek = m_CalendarFirstDayOfWeek

End Property

Public Property Let CalendarFirstDayOfWeek(ByVal New_CalendarFirstDayOfWeek As dtDaysOfTheWeek)

    If (New_CalendarFirstDayOfWeek >= [dtSunday]) And (New_CalendarFirstDayOfWeek <= [dtSaturday]) Then
        m_CalendarFirstDayOfWeek = New_CalendarFirstDayOfWeek
        PropertyChanged "CalendarFirstDayOfWeek"

    End If

End Property

Public Property Get CalendarClickBehivor() As dtClickBehivor

    CalendarClickBehivor = m_CalendarClickBehivor

End Property

Public Property Let CalendarClickBehivor(ByVal New_CalendarClickBehivor As dtClickBehivor)

    m_CalendarClickBehivor = New_CalendarClickBehivor
    PropertyChanged "CalendarClickBehivor"

End Property

Public Property Get CalendarDayHeaderFormat() As dtDayHeaderFormats

    CalendarDayHeaderFormat = m_CalendarDayHeaderFormat

End Property

Public Property Let CalendarDayHeaderFormat(ByVal New_CalendarDayHeaderFormat As dtDayHeaderFormats)

    m_CalendarDayHeaderFormat = New_CalendarDayHeaderFormat
    PropertyChanged "CalendarDayHeaderFormat"

End Property

'## Calendar Date setting Selection --------------------------------------------------------------------
Public Property Get CalendarOption() As CalendarDateTimeOption

    CalendarOption = m_CalendarOption

End Property

Public Property Let CalendarOption(ByVal New_CalendarOption As CalendarDateTimeOption)

    m_CalendarOption = New_CalendarOption
    PropertyChanged "CalendarOption"

End Property

'## Calendar Grid Width--------------------------------------------------------------------
Public Property Get CalendarGridWidth() As Integer

    CalendarGridWidth = m_nGridWidth

End Property

Public Property Let CalendarGridWidth(ByVal New_CalendarGridWidth As Integer)

    If New_CalendarGridWidth >= 20 Then
        m_nGridWidth = New_CalendarGridWidth
      Else
        m_nGridWidth = m_def_nGridWidth
    End If
    PropertyChanged "CalendarGridWidth"

End Property

'## Calendar Grid Height--------------------------------------------------------------------
Public Property Get CalendarGridHeight() As Integer

    CalendarGridHeight = m_nGridHeight

End Property

Public Property Let CalendarGridHeight(ByVal New_CalendarGridHeight As Integer)

    If New_CalendarGridHeight >= 20 Then
        m_nGridHeight = New_CalendarGridHeight
      Else
        m_nGridHeight = m_def_nGridHeight
    End If
    PropertyChanged "CalendarGridHeight"

End Property

'~~ HoverSelection ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Property Get HoverSelection() As Boolean

    HoverSelection = m_HoverSelection

End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)

    m_HoverSelection = New_HoverSelection
    PropertyChanged "HoverSelection"

End Property

Public Property Get HotTracking() As Boolean

    HotTracking = m_HotTracking

End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)

    m_HotTracking = New_HotTracking

    PropertyChanged HotTracking

End Property

Public Property Get HoverColor() As OLE_COLOR

    HoverColor = m_HoverColor

End Property

Public Property Let HoverColor(ByVal New_HoverColor As OLE_COLOR)

    m_HoverColor = New_HoverColor

    PropertyChanged HoverColor

End Property

' Font

Public Property Get Font() As Font

    Set Font = Text1.Font

End Property

Public Property Set Font(ByVal vData As Font)

    Set Text1.Font = vData
    PropertyChanged "Font"
    Call RefreshControl

End Property

Public Property Get FontBold() As Boolean

    FontBold = Text1.FontBold

End Property

Public Property Let FontBold(ByVal vData As Boolean)

    Text1.FontBold() = vData
    PropertyChanged "FontBold"

End Property

Public Property Get FontItalic() As Boolean

    FontItalic = Text1.FontItalic

End Property

Public Property Let FontItalic(ByVal vData As Boolean)

    Text1.FontItalic() = vData
    PropertyChanged "FontItalic"

End Property

Public Property Get FontName() As String

    FontName = Text1.FontName

End Property

Public Property Let FontName(ByVal vData As String)

    Text1.FontName() = vData
    PropertyChanged "FontName"

End Property

Public Property Get FontSize() As Single

    FontSize = Text1.FontSize

End Property

Public Property Let FontSize(ByVal vData As Single)

    Text1.FontSize() = vData
    PropertyChanged "FontSize"

End Property

Public Property Get FontStrikethru() As Boolean

    FontStrikethru = Text1.FontStrikethru

End Property

Public Property Let FontStrikethru(ByVal vData As Boolean)

    Text1.FontStrikethru() = vData
    PropertyChanged "FontStrikethru"

End Property

Public Property Get FontUnderline() As Boolean

    FontUnderline = Text1.FontUnderline

End Property

Public Property Let FontUnderline(ByVal vData As Boolean)

    Text1.FontUnderline() = vData
    PropertyChanged "FontUnderline"

End Property

Public Property Get TextMaxLength() As Integer

    TextMaxLength = m_MaxLength

End Property

Public Property Let TextMaxLength(ByVal New_TextMaxLength As Integer)

    m_MaxLength = New_TextMaxLength

    Text1.MaxLength = New_TextMaxLength

    PropertyChanged "TextMaxLength"

End Property

Public Property Get OutputText() As String

    OutputText = VBA.Format$(m_Value, strDateFormat)

End Property

Public Property Let OutputText(ByVal New_OutputText As String)

  'm_OutputText = New_OutputText
  'PropertyChanged "OutputText"


End Property

Public Property Get Text_Enabled() As Boolean

    Text_Enabled = Text1.Enabled

End Property

Public Property Let Text_Enabled(ByVal Text_New_Enabled As Boolean)

    Text1.Enabled() = Text_New_Enabled
    PropertyChanged "Text_Enabled"

End Property

Public Property Get Text_Locked() As Boolean

    Text_Locked = Text1.Locked

End Property

Public Property Let Text_Locked(ByVal Text_New_Locked As Boolean)

    Text1.Locked() = Text_New_Locked
    PropertyChanged "Text_Locked"

End Property

Public Property Get DropListEnabled() As Boolean

    DropListEnabled = m_DropListEnabled

End Property

Public Property Let DropListEnabled(ByVal New_DropListEnabled As Boolean)

    m_DropListEnabled = New_DropListEnabled

    PropertyChanged "DropListEnabled"

End Property

Public Property Get FocusColor() As OLE_COLOR

    FocusColor = m_FocusColor

End Property

Public Property Let FocusColor(ByVal New_FocusColor As OLE_COLOR)

    m_FocusColor = New_FocusColor
    PropertyChanged "FocusColor"
    Call RefreshControl

End Property

Public Property Get CalendarBorderStyle() As CalendarBorderStyleType

    CalendarBorderStyle = m_CalendarBorderStyle

End Property

Public Property Let CalendarBorderStyle(ByVal New_CalendarBorderStyle As CalendarBorderStyleType)

    m_CalendarBorderStyle = New_CalendarBorderStyle

    PropertyChanged "CalendarBorderStyle"

End Property

Public Property Get TodayForeColor() As OLE_COLOR

    TodayForeColor = m_TodayForeColor

End Property

Public Property Let TodayForeColor(ByVal New_TodayForeColor As OLE_COLOR)

    m_TodayForeColor = New_TodayForeColor
    PropertyChanged "TodayForeColor"

End Property

Public Property Get TodayFont() As StdFont

    Set TodayFont = m_TodayFont

End Property

Public Property Set TodayFont(ByVal New_TodayFont As StdFont)

    Set m_TodayFont = New_TodayFont
    PropertyChanged "TodayFont"

End Property

Public Property Get TodayFontBold() As Boolean

    TodayFontBold = m_TodayFont.Bold

End Property

'Public Property Let TodayFontBold(ByVal New_TodayFontBold As Boolean)
'm_TodayFont.Bold = New_TodayFontBold
'PropertyChanged "TodayFontBold"
'End Property

Public Property Get TodayFontItalic() As Boolean

    TodayFontItalic = m_TodayFont.Italic

End Property

'Public Property Let TodayFontItalic(ByVal New_TodayFontItalic As Boolean)
' m_TodayFont.Italic = New_TodayFontItalic
'PropertyChanged "TodayFontItalic"
'End Property

Public Property Get TodayFontName() As String

    TodayFontName = m_TodayFont.Name

End Property

'Public Property Let TodayFontName(ByVal New_TodayFontName As String)
'm_TodayFont.Name = New_TodayFontName
'PropertyChanged "TodayFontName"
'End Property

Public Property Get TodayFontSize() As Long

    TodayFontSize = m_TodayFont.Size

End Property

'Public Property Let TodayFontSize(ByVal New_TodayFontSize As Long)
'm_TodayFont.Size = New_TodayFontSize
'PropertyChanged "TodayFontSize"
'End Property
'TodayFontStrikethrough
Public Property Get TodayFontStrikethrough() As Boolean

    TodayFontStrikethrough = m_TodayFont.Strikethrough

End Property

'Public Property Let TodayFontStrikethrough(ByVal New_TodayFontStrikethrough As Boolean)
' m_TodayFont.Strikethrough = New_TodayFontStrikethrough
' PropertyChanged "TodayFontStrikethrough"
'End Property

Public Property Get TodayMaskColor() As OLE_COLOR

    TodayMaskColor = m_TodayMaskColor

End Property

Public Property Let TodayMaskColor(ByVal New_TodayMaskColor As OLE_COLOR)

    m_TodayMaskColor = New_TodayMaskColor
    PropertyChanged "TodayMaskColor"

End Property

Public Property Get TodayPicture() As Picture

    Set TodayPicture = m_TodayPicture

End Property

Public Property Set TodayPicture(ByVal New_TodayPicture As Picture)

    Set m_TodayPicture = New_TodayPicture

    If m_TodayPicture Is Nothing Then
        m_TodayOriginalPicSizeW = 32
        m_TodayOriginalPicSizeH = 32
      Else
        m_TodayOriginalPicSizeW = UserControl.ScaleX(m_TodayPicture.Width, 8, UserControl.ScaleMode)
        m_TodayOriginalPicSizeH = UserControl.ScaleY(m_TodayPicture.Height, 8, UserControl.ScaleMode)
    End If
    PropertyChanged "TodayPicture"
    If Not (m_TodayPicture Is Nothing) And m_TodayPictureSize = [SizeDefault] Then
        m_TodayPictureWidth = UserControl.ScaleX(m_TodayPicture.Width, 8, UserControl.ScaleMode)
        m_TodayPictureHeight = UserControl.ScaleY(m_TodayPicture.Height, 8, UserControl.ScaleMode)
    End If

End Property

Public Property Get TodayPictureWidth() As Long

    TodayPictureWidth = m_TodayPictureWidth

End Property

Public Property Let TodayPictureWidth(ByVal New_TodayPictureWidth As Long)

    m_TodayPictureWidth = New_TodayPictureWidth
    PropertyChanged "TodayPictureWidth"

End Property

Public Property Get TodayPictureHeight() As Long

    TodayPictureHeight = m_TodayPictureHeight

End Property

Public Property Let TodayPictureHeight(ByVal New_TodayPictureHeight As Long)

    m_TodayPictureHeight = New_TodayPictureHeight
    PropertyChanged "TodayPictureHeight"

End Property

Public Property Get TodayPictureSize() As PictureSize

    TodayPictureSize = m_TodayPictureSize

End Property

Public Property Let TodayPictureSize(ByVal New_TodayPictureSize As PictureSize)

    m_TodayPictureSize = New_TodayPictureSize
    PropertyChanged "TodayPictureSize"

    If New_TodayPictureSize = size16x16 Then
        m_TodayPictureWidth = 16
        m_TodayPictureHeight = 16
      ElseIf New_TodayPictureSize = size32x32 Then
        m_TodayPictureWidth = 32
        m_TodayPictureHeight = 32
      ElseIf New_TodayPictureSize = size64x64 Then
        m_TodayPictureWidth = 64
        m_TodayPictureHeight = 64
      ElseIf New_TodayPictureSize = SizeDefault Then
        If Not (m_TodayPicture Is Nothing) Then
            m_TodayPictureWidth = m_TodayOriginalPicSizeW
            m_TodayPictureHeight = m_TodayOriginalPicSizeH
          Else
            m_TodayPictureWidth = 32
            m_TodayPictureHeight = 32
        End If
    End If

End Property

Public Property Get WeekColumnWidth() As Long

    WeekColumnWidth = m_WeekColumnWidth

End Property

Public Property Let WeekColumnWidth(ByVal New_WeekColumnWidth As Long)

    m_WeekColumnWidth = New_WeekColumnWidth

    PropertyChanged "WeekColumnWidth"

End Property

Public Property Get ShowWeekSignPicture() As Boolean

    ShowWeekSignPicture = m_ShowWeekSignPicture

End Property

Public Property Let ShowWeekSignPicture(ByVal New_ShowWeekSignPicture As Boolean)

    m_ShowWeekSignPicture = New_ShowWeekSignPicture
    PropertyChanged "ShowWeekSignPicture"

End Property

Public Property Get WeekSignPicMaskColor() As OLE_COLOR

    WeekSignPicMaskColor = m_WeekSignPicMaskColor

End Property

Public Property Let WeekSignPicMaskColor(ByVal New_WeekSignPicMaskColor As OLE_COLOR)

    m_WeekSignPicMaskColor = New_WeekSignPicMaskColor
    PropertyChanged "WeekSignPicMaskColor"

End Property

Public Property Get WeekSignPictureWidth() As Long

    WeekSignPictureWidth = m_WeekSignPictureWidth

End Property

Public Property Let WeekSignPictureWidth(ByVal New_WeekSignPictureWidth As Long)

    m_WeekSignPictureWidth = New_WeekSignPictureWidth
    PropertyChanged "WeekSignPictureWidth"

End Property

Public Property Get WeekSignPictureHeight() As Long

    WeekSignPictureHeight = m_WeekSignPictureHeight

End Property

Public Property Let WeekSignPictureHeight(ByVal New_WeekSignPictureHeight As Long)

    m_WeekSignPictureHeight = New_WeekSignPictureHeight
    PropertyChanged "WeekSignPictureHeight"

End Property

Public Property Get WeekSignPictureSize() As PictureSize

    WeekSignPictureSize = m_WeekSignPictureSize

End Property

Public Property Let WeekSignPictureSize(ByVal New_WeekSignPictureSize As PictureSize)

    m_WeekSignPictureSize = New_WeekSignPictureSize
    PropertyChanged "WeekSignPictureSize"

    If New_WeekSignPictureSize = size16x16 Then
        m_WeekSignPictureWidth = 16
        m_WeekSignPictureHeight = 16
      ElseIf New_WeekSignPictureSize = size32x32 Then
        m_WeekSignPictureWidth = 32
        m_WeekSignPictureHeight = 32
      ElseIf New_WeekSignPictureSize = size64x64 Then
        m_WeekSignPictureWidth = 64
        m_WeekSignPictureHeight = 64
      ElseIf New_WeekSignPictureSize = SizeDefault Then
        If Not (m_WeekSignPicture Is Nothing) Then
            m_WeekSignPictureWidth = m_WeekSignOriginalPicSizeW
            m_WeekSignPictureHeight = m_WeekSignOriginalPicSizeH
          Else
            m_WeekSignPictureWidth = 32
            m_WeekSignPictureHeight = 32
        End If
    End If

End Property

Public Property Get DateAlign() As DateAlignType

    DateAlign = m_DateAlign

End Property

Public Property Let DateAlign(ByVal New_DateAlign As DateAlignType)

    m_DateAlign = New_DateAlign
    PropertyChanged "DateAlign"

End Property

Public Property Get DatePosn() As DatePositionType

    DatePosn = m_DatePosn

End Property

Public Property Let DatePosn(ByVal New_DatePosn As DatePositionType)

    m_DatePosn = New_DatePosn
    PropertyChanged "DatePosn"

End Property

Public Property Get DateXOffset() As Long

    DateXOffset = m_DateXOffset

End Property

Public Property Let DateXOffset(ByVal New_DateXOffset As Long)

    m_DateXOffset = New_DateXOffset
    PropertyChanged "DateXOffset"

End Property

Public Property Get DateYOffset() As Long

    DateYOffset = m_DateYOffset

End Property

Public Property Let DateYOffset(ByVal New_DateYOffset As Long)

    m_DateYOffset = New_DateYOffset
    PropertyChanged "DateYOffset"

End Property

Public Property Get TodayXOffset() As Long

    TodayXOffset = m_TodayXOffset

End Property

Public Property Let TodayXOffset(ByVal New_TodayXOffset As Long)

    m_TodayXOffset = New_TodayXOffset
    PropertyChanged "TodayXOffset"

End Property

Public Property Get TodayYOffset() As Long

    TodayYOffset = m_TodayYOffset

End Property

Public Property Let TodayYOffset(ByVal New_TodayYOffset As Long)

    m_TodayYOffset = New_TodayYOffset
    PropertyChanged "TodayYOffset"

End Property

Public Property Get GridLineColor() As OLE_COLOR

    GridLineColor = m_GridLineColor

End Property

Public Property Let GridLineColor(ByVal New_GridLineColor As OLE_COLOR)

    m_GridLineColor = New_GridLineColor
    PropertyChanged "GridLineColor"

End Property

Public Property Get GridLine() As GridLineType

    GridLine = m_Gridline

End Property

Public Property Let GridLine(ByVal New_GridLine As GridLineType)

    m_Gridline = New_GridLine
    PropertyChanged "Gridline"

End Property

Public Property Get CalendarBdHighlightColour() As OLE_COLOR

    CalendarBdHighlightColour = m_CalendarBdHighlightColour

End Property

Public Property Let CalendarBdHighlightColour(ByVal New_CalendarBdHighlightColour As OLE_COLOR)

    m_CalendarBdHighlightColour = New_CalendarBdHighlightColour
    PropertyChanged "CalendarBdHighlightColour"

End Property

Public Property Get CalendarBdHighlightDKColour() As OLE_COLOR

    CalendarBdHighlightDKColour = m_CalendarBdHighlightDKColour

End Property

Public Property Let CalendarBdHighlightDKColour(ByVal New_CalendarBdHighlightDKColour As OLE_COLOR)

    m_CalendarBdHighlightDKColour = New_CalendarBdHighlightDKColour
    PropertyChanged "CalendarBdHighlightDKColour"

End Property

Public Property Get CalendarBdShadowColour() As OLE_COLOR

    CalendarBdShadowColour = m_CalendarBdShadowColour

End Property

Public Property Let CalendarBdShadowColour(ByVal New_CalendarBdShadowColour As OLE_COLOR)

    m_CalendarBdShadowColour = New_CalendarBdShadowColour
    PropertyChanged "CalendarBdShadowColour"

End Property

Public Property Get CalendarBdShadowDKColour() As OLE_COLOR

    CalendarBdShadowDKColour = m_CalendarBdShadowDKColour

End Property

Public Property Let CalendarBdShadowDKColour(ByVal New_CalendarBdShadowDKColour As OLE_COLOR)

    m_CalendarBdShadowDKColour = New_CalendarBdShadowDKColour
    PropertyChanged "CalendarBdShadowDKColour"

End Property

Public Property Get CalendarBdFlatBorderColour() As OLE_COLOR

    CalendarBdFlatBorderColour = m_CalendarBdFlatBorderColour

End Property

Public Property Let CalendarBdFlatBorderColour(ByVal New_CalendarBdFlatBorderColour As OLE_COLOR)

    m_CalendarBdFlatBorderColour = New_CalendarBdFlatBorderColour
    PropertyChanged "CalendarBdFlatBorderColour"

End Property

Public Property Get WeekSignPicXOffset() As Long

    WeekSignPicXOffset = m_WeekSignPicXOffset

End Property

Public Property Let WeekSignPicXOffset(ByVal New_WeekSignPicXOffset As Long)

    m_WeekSignPicXOffset = New_WeekSignPicXOffset
    PropertyChanged "WeekSignPicXOffset"

End Property

Public Property Get WeekSignPicYOffset() As Long

    WeekSignPicYOffset = m_WeekSignPicYOffset

End Property

Public Property Let WeekSignPicYOffset(ByVal New_WeekSignPicYOffset As Long)

    m_WeekSignPicYOffset = New_WeekSignPicYOffset
    PropertyChanged "WeekSignPicYOffset"

End Property

Public Property Get WeekSignPicture() As Picture

    Set WeekSignPicture = m_WeekSignPicture

End Property

Public Property Set WeekSignPicture(ByVal New_WeekSignPicture As Picture)

    Set m_WeekSignPicture = New_WeekSignPicture
    PropertyChanged "WeekSignPicture"

    Set m_WeekSignPicture = New_WeekSignPicture

    If m_WeekSignPicture Is Nothing Then
        m_WeekSignOriginalPicSizeW = 32
        m_WeekSignOriginalPicSizeH = 32
      Else
        m_WeekSignOriginalPicSizeW = UserControl.ScaleX(m_WeekSignPicture.Width, 8, UserControl.ScaleMode)
        m_WeekSignOriginalPicSizeH = UserControl.ScaleY(m_WeekSignPicture.Height, 8, UserControl.ScaleMode)

    End If
    PropertyChanged "WeekSignPicture"
    If Not (m_WeekSignPicture Is Nothing) And m_WeekSignPictureSize = [SizeDefault] Then
        m_WeekSignPictureWidth = UserControl.ScaleX(m_WeekSignPicture.Width, 8, UserControl.ScaleMode)
        m_WeekSignPictureHeight = UserControl.ScaleY(m_WeekSignPicture.Height, 8, UserControl.ScaleMode)
    End If

End Property

Public Property Get CustomDateFormat() As String

    CustomDateFormat = strCustomDateFormat

End Property

Public Property Let CustomDateFormat(ByVal New_CustomDateFormat As String)

    strCustomDateFormat = New_CustomDateFormat
    strDateFormat = GetFormat()
    PropertyChanged "DateCustomFormat"

End Property

Public Property Get CalendarDateFormat() As dtFormatConstants

    CalendarDateFormat = m_DateFormat

End Property

Public Property Let CalendarDateFormat(ByVal New_CalendarDateFormat As dtFormatConstants)

    m_DateFormat = New_CalendarDateFormat
    strDateFormat = GetFormat()
    PropertyChanged "CalendarDateFormat"

End Property

Public Property Get AllowNull() As Boolean

    AllowNull = bAllowNull

End Property

Public Property Let AllowNull(ByVal vData As Boolean)

    bAllowNull = vData
    PropertyChanged "AllowNull"

    If Not bAllowNull Then
        If (Value = "") Then
            Value = Date
        End If
    End If

End Property

Public Property Get DateFormatWhenEdit() As dtFormatWhenEditConstants

    DateFormatWhenEdit = m_DateFormatWhenEdit

End Property

Public Property Let DateFormatWhenEdit(ByVal vData As dtFormatWhenEditConstants)

    m_DateFormatWhenEdit = vData
    strMask = GetMaskString()
    strFormatWhenEdit = GetEditFormatString()
    mstrText = VBA.Format$(m_Value, strFormatWhenEdit)
    PropertyChanged "DateFormatWhenEdit"

End Property

Public Property Get Seperator() As String

    Seperator = strSeperator

End Property

Public Property Let Seperator(ByVal vData As String)

    If Ambient.UserMode Then Err.Raise 382      '// In run-time raise an error
    If (Len(vData) = 1) Then
        If (vData <> strPlaceHolder) Then
            If InStr(1, SEPERATORS, vData) Then
                strSeperator = vData
                PropertyChanged "Seperator"
            End If
        End If
    End If

End Property

Public Property Get PlaceHolder() As String

    PlaceHolder = strPlaceHolder

End Property

Public Property Let PlaceHolder(ByVal vData As String)

    If Ambient.UserMode Then Err.Raise 382      '// In run-time raise an error
    If (Len(vData) = 1) Then
        If (vData <> strPlaceHolder) Then
            If InStr(1, PLACE_HOLDERS, vData) Then
                strPlaceHolder = vData
                PropertyChanged "PlaceHolder"
            End If
        End If
    End If

End Property

Public Property Get MaxDate() As Date

    MaxDate = m_dtMaxDate

End Property

Public Property Let MaxDate(ByVal vData As Date)

    If (vData > m_dtMinDate) Then
        m_dtMaxDate = vData
        PropertyChanged "MaxDate"
    End If

End Property

Public Property Get MinDate() As Date

    MinDate = m_dtMinDate

End Property

Public Property Let MinDate(ByVal vData As Date)

    If (vData < m_dtMaxDate) Then
        m_dtMinDate = vData
        PropertyChanged "MinDate"
    End If

End Property

Public Property Get Value() As Variant

    Value = m_Value

End Property

Public Property Let Value(ByVal vData As Variant)

    If IsDate(vData) Then
        m_Value = CDate(vData)
      ElseIf (Len(CStr(vData)) = 0) Then
        m_Value = IIf(bAllowNull, vData, Date)
      Else
        Exit Property
    End If

    If (Len(CStr(m_Value)) > 0) Then
        Text1.Text = VBA.Format$(m_Value, strDateFormat)
        mstrText = VBA.Format$(m_Value, strFormatWhenEdit)
      Else
        mstrText = GetMaskString()
        Text1.Text = IIf(Ambient.UserMode, mstrText, "")
    End If

    PropertyChanged "Value"

End Property

Public Property Get OLEDragMode() As OLEDragConstants

    OLEDragMode = Text1.OLEDragMode

End Property

Public Property Let OLEDragMode(ByVal vData As OLEDragConstants)

    Text1.OLEDragMode() = vData
    PropertyChanged "OLEDragMode"

End Property

Public Property Get OLEDropMode() As OLEDropConstants

    OLEDropMode = Text1.OLEDropMode

End Property

Public Property Let OLEDropMode(ByVal vData As OLEDropConstants)

    Text1.OLEDropMode() = vData
    PropertyChanged "OLEDropMode"

End Property

Public Property Get MousePointer() As MousePointerConstants

    MousePointer = Text1.MousePointer

End Property

Public Property Let MousePointer(ByVal vData As MousePointerConstants)

    Text1.MousePointer() = vData
    PropertyChanged "MousePointer"

End Property

Public Property Get MouseIcon() As Picture

    Set MouseIcon = Text1.MouseIcon

End Property

Public Property Set MouseIcon(ByVal vData As Picture)

    Set Text1.MouseIcon = vData
    PropertyChanged "MouseIcon"

End Property

Public Property Get RightToLeft() As Boolean

    RightToLeft = Text1.RightToLeft

End Property

Public Property Let RightToLeft(ByVal vData As Boolean)

    Text1.RightToLeft() = vData
    PropertyChanged "RightToLeft"

End Property

Public Function IsNull(DateIn As String) As Boolean

    IsNull = IsEmpty(DateIn) Or (DateIn = "")

End Function

Private Function GetMaskString() As String

  Dim i()         As Integer
  Dim strMask     As String

    ReDim i(1 To 3)
    Select Case m_DateFormatWhenEdit
      Case [dd/mm/yyyy], [mm/dd/yyyy]
        i(1) = 2
        i(2) = 2
        i(3) = 4
      Case [dd/yyyy/mm], [mm/yyyy/dd]
        i(1) = 2
        i(2) = 4
        i(3) = 2
      Case [yyyy/dd/mm], [yyyy/mm/dd]
        i(1) = 4
        i(2) = 2
        i(3) = 2
    End Select

    GetMaskString = String$(i(1), strPlaceHolder) & strSeperator & _
                    String$(i(2), strPlaceHolder) & strSeperator & _
                    String$(i(3), strPlaceHolder)

End Function

Private Sub InsertNumber(ByVal sChar As String)

  Dim intStart    As Integer
  Dim intEnd      As Integer
  Dim strText     As String
  Dim strNewText  As String
  Dim intPos      As Integer
  Dim i           As Integer

    With Text1
        '// If insertion point is at maximum length, exit the procedure
        If (.SelStart = .MaxLength) Then Exit Sub
        '// If some text has been selected, delete it first
        If (.SelLength > 0) Then Call DeleteSelection
        '// Get current section information
        Call GetSectionInfo(GetCurrentSection(), intStart, intEnd)
        '// Get position of insertion point from the beginning of current section
        intPos = .SelStart - intStart
        '// Get section text
        strText = Mid$(mstrText, intStart + 1, (intEnd - intStart))
        '// Parse the section text to form new text
        strNewText = Mid$(strText, 1, intPos) & sChar & _
                     Replace(strText, strPlaceHolder, "", intPos + 1, 1)
        strNewText = Left$(strNewText, Len(strText))
        '// Apply the parsed string in main date string
        Mid$(mstrText, intStart + 1, Len(strNewText)) = strNewText
        '// Update text box
        .Text = mstrText
        .SelStart = intPos + intStart + 1
    End With

    Call MoveInsertionPoint(0)

End Sub

Private Sub DeleteNumber(ByVal iDeleteMode As Integer)

  Dim intStart    As Integer
  Dim intEnd      As Integer
  Dim strText     As String
  Dim strNewText  As String
  Dim intPos      As Integer

    With Text1
        '// If a selection has been made, delete it
        If (.SelLength > 0) Then
            Call DeleteSelection
            Exit Sub
        End If

        If (iDeleteMode = vbKeyDelete) Then
            '// If insertion point is at the end, exit the function
            If (.SelStart = .MaxLength) Then Exit Sub
            '// If the next letter is a seperator, move insertion point one step forward
            If (Mid$(mstrText, .SelStart + 1, 1) = strSeperator) Then
                .SelStart = .SelStart + 1
                CurrentSection = GetCurrentSection()
            End If
            '// Get current section information
            Call GetSectionInfo(GetCurrentSection(), intStart, intEnd)
            '// Get position of insertion point from the beginning of current section
            intPos = .SelStart - intStart
            '// Get section text
            strText = Mid$(mstrText, intStart + 1, (intEnd - intStart))
            '// Parse the section text to form new text
            strNewText = Mid$(strText, 1, intPos) & _
                         Mid$(strText, intPos + 2, (intEnd - intStart) - intPos)
            strNewText = strNewText & strPlaceHolder
            '// Apply the parsed string in main date string
            Mid$(mstrText, intStart + 1, Len(strNewText)) = strNewText
          Else
            '// If insertion point is at the end, exit the function
            If (.SelStart = 0) Then Exit Sub
            '// If the letter just before is a seperator, move insertion point one step backward
            If (Mid$(mstrText, .SelStart, 1) = strSeperator) Then
                .SelStart = .SelStart - 1
                CurrentSection = GetCurrentSection()
            End If
            '// Get current section information
            Call GetSectionInfo(GetCurrentSection(), intStart, intEnd)
            '// Get position of insertion point from the beginning of current section
            intPos = .SelStart - intStart
            '// Get section text
            strText = Mid$(mstrText, intStart + 1, (intEnd - intStart))
            '// Parse the section text to form new text
            strNewText = Mid$(strText, 1, intPos - 1) & _
                         Mid$(strText, intPos + 1, (intEnd - intStart) - intPos)
            strNewText = strNewText & strPlaceHolder
            '// Apply the parsed string in main date string
            Mid$(mstrText, intStart + 1, Len(strNewText)) = strNewText
        End If

        .Text = mstrText
        .SelStart = intPos + intStart - IIf((iDeleteMode = vbKeyBack), 1, 0)
    End With

End Sub

Private Sub DeleteSelection()

  Dim strDel      As String
  Dim strChar     As String
  Dim intPos      As Integer
  Dim intLen      As Integer
  Dim i           As Integer

    With Text1
        intPos = .SelStart
        intLen = .SelLength

        For i = (intPos + 1) To (intPos + intLen)
            strChar = Mid$(mstrText, i, 1)
            strDel = strDel & IIf((strChar = strSeperator), strSeperator, strPlaceHolder)
        Next i
        Mid$(mstrText, intPos + 1, intLen) = strDel

        .Text = mstrText
        .SelStart = intPos
    End With

End Sub

Private Sub MakeSelection( _
                          Optional eSection As dtSectionConstants = dtInvalid, _
                          Optional ByVal iStart As Integer = -1, _
                          Optional ByVal iEnd As Integer = -1)

  Dim intStart    As Integer
  Dim intEnd      As Integer

    If (iStart = -1) Or (iEnd = -1) Then
        If (eSection = dtInvalid) Then
            Call GetSectionInfo(GetCurrentSection(), intStart, intEnd)
          Else
            Call GetSectionInfo(eSection, intStart, intEnd)
        End If
      Else
        intStart = iStart
        intEnd = iEnd
    End If

    Text1.SelStart = intStart
    Text1.SelLength = (intEnd - intStart)

    CurrentSection = GetCurrentSection()

End Sub

Private Function GetCurrentSection() As dtSectionConstants

  Dim strCurPosChar   As String
  Dim intPos          As Integer

    With Text1
        If (.SelLength > 0) Then
            If (InStr(1, .SelText, strSeperator) > 0) Then
                GetCurrentSection = dtInvalid
                Exit Function
            End If
        End If

        intPos = .SelStart '+ IIf((.SelStart = .MaxLength), 0, 1)
        If (.SelStart <> .MaxLength) Or (.SelStart = 0) Then
            intPos = intPos + 1
        End If
        strCurPosChar = Mid$(strFormatWhenEdit, intPos, 1)
        If (strCurPosChar = strSeperator) Then
            strCurPosChar = Mid$(strFormatWhenEdit, intPos - 1, 1)
        End If
    End With

    Select Case strCurPosChar
      Case "d"
        GetCurrentSection = dtDaySection
      Case "m"
        GetCurrentSection = dtMonthSection
      Case "y"
        GetCurrentSection = dtYearSection
    End Select

End Function

Private Sub MoveInsertionPoint(ByVal iMoveMode As Integer)

  Dim intPos          As Integer
  Dim intDirection    As Integer
  Dim enuSection      As dtSectionConstants

    With Text1
        '// Find the current position of insertion point
        intPos = .SelStart
        '// Get the direction to move the insertion point
        intDirection = IIf((iMoveMode = vbKeyRight) Or (iMoveMode = 0), 1, -1)
        '// If insertion point has to be moved, validate & move
        If (iMoveMode <> 0) Then
            If (.SelLength = 0) Then
                .SelStart = .SelStart + intDirection
              Else
                If (iMoveMode = vbKeyRight) Then
                    If ((.SelStart + .SelLength) <> .MaxLength) Then
                        .SelStart = (.SelStart + .SelLength) + intDirection
                    End If
                  Else
                    If (.SelStart <> 0) Then
                        .SelStart = .SelStart + intDirection
                    End If
                End If
            End If
        End If
        '// If insertion point is at the end of current section,
        '// move the point to next section
        If (Mid$(mstrText, .SelStart + 1, 1) = strSeperator) And (.SelLength = 0) Then
            .SelStart = .SelStart + intDirection
        End If
        '// If current section differs from old section,
        '// select the current section.
        enuSection = GetCurrentSection()
        If (enuSection = dtInvalid) Then Exit Sub
        If (enuSection <> CurrentSection) Then
            If (.SelStart <> 0) And (.SelStart <> .MaxLength) Then
                Call MakeSelection(enuSection)
            End If
        End If
    End With

End Sub

Private Sub ChangeSectionValue(ByVal iChangeMode As Integer)

  Dim enuSection      As dtSectionConstants
  Dim intDirection    As Integer
  Dim intStart        As Integer
  Dim intEnd          As Integer
  Dim intValue        As Integer

    enuSection = GetCurrentSection()
    If (enuSection = dtInvalid) Then Exit Sub

    intDirection = IIf((iChangeMode = vbKeyUp), 1, -1)
    Select Case enuSection
      Case dtDaySection
        Call GetSectionInfo(dtDaySection, intStart, intEnd, intValue)
        intValue = IIf((intValue = 0), Day(Date), intValue + intDirection)
        If (intValue > 31) Then
            intValue = 1
          ElseIf (intValue < 1) Then
            intValue = 31
        End If
        Mid$(mstrText, intStart + 1, (intEnd - intStart)) = VBA.Format$(intValue, "00")
      Case dtMonthSection
        Call GetSectionInfo(dtMonthSection, intStart, intEnd, intValue)
        intValue = IIf((intValue = 0), Month(Date), intValue + intDirection)
        If (intValue > 12) Then
            intValue = 1
          ElseIf (intValue < 1) Then
            intValue = 12
        End If
        Mid$(mstrText, intStart + 1, (intEnd - intStart)) = VBA.Format$(intValue, "00")
      Case dtYearSection
        Call GetSectionInfo(dtYearSection, intStart, intEnd, intValue)
        intValue = IIf((intValue = 0), Year(Date), intValue + intDirection)
        If (intValue > Year(m_dtMaxDate)) Then
            intValue = Year(m_dtMinDate)
          ElseIf (intValue < Year(m_dtMinDate)) Then
            intValue = Year(m_dtMaxDate)
        End If
        Mid$(mstrText, intStart + 1, (intEnd - intStart)) = VBA.Format$(intValue, "0000")
    End Select

    Text1.Text = mstrText
    Call MakeSelection(, intStart, intEnd)

End Sub

Private Sub GetSectionInfo( _
                           eSection As dtSectionConstants, _
                           Optional iStart As Integer, _
                           Optional iEnd As Integer, _
                           Optional iValue As Integer)

  Dim strVal          As String
  Dim strCurPosChar   As String
  Dim intPos          As Integer

    Select Case eSection
      Case dtDaySection
        iStart = InStr(1, strFormatWhenEdit, "d") - 1
        iEnd = iStart + 2
      Case dtMonthSection
        iStart = InStr(1, strFormatWhenEdit, "m") - 1
        iEnd = iStart + 2
      Case dtYearSection
        iStart = InStr(1, strFormatWhenEdit, "y") - 1
        iEnd = iStart + 4
    End Select

    strVal = Mid$(mstrText, iStart + 1, (iEnd - iStart))
    strVal = Replace(strVal, strPlaceHolder, "")
    iValue = CInt(Val(Trim(strVal)))

End Sub

Private Sub MoveToNextSection()

  Dim intPos  As Integer

    With Text1
        intPos = InStr(.SelStart + 1, mstrText, strSeperator)
        If (intPos > 0) Then
            .SelStart = intPos + 1
            Call MakeSelection(GetCurrentSection())
        End If
    End With

End Sub

Private Function GetEditFormatString() As String

  Dim strTmp          As String

    Select Case m_DateFormatWhenEdit
      Case [dd/mm/yyyy]
        strTmp = "dd" & strSeperator & "mm" & strSeperator & "yyyy"
      Case [mm/dd/yyyy]
        strTmp = "mm" & strSeperator & "dd" & strSeperator & "yyyy"
      Case [dd/yyyy/mm]
        strTmp = "dd" & strSeperator & "yyyy" & strSeperator & "mm"
      Case [mm/yyyy/dd]
        strTmp = "mm" & strSeperator & "yyyy" & strSeperator & "dd"
      Case [yyyy/dd/mm]
        strTmp = "yyyy" & strSeperator & "dd" & strSeperator & "mm"
      Case [yyyy/mm/dd]
        strTmp = "yyyy" & strSeperator & "mm" & strSeperator & "dd"
    End Select
    GetEditFormatString = strTmp

End Function

Private Function GetDate(dDate As Date) As dtDateValidationConstants

  Dim intDay          As Integer
  Dim intMonth        As Integer
  Dim intYear         As Integer
  Dim intLastDay      As Integer
  Dim intNewMonth     As Integer

    Call GetSectionInfo(dtDaySection, , , intDay)
    Call GetSectionInfo(dtMonthSection, , , intMonth)
    Call GetSectionInfo(dtYearSection, , , intYear)

    If (intDay = 0) And (intMonth = 0) And (intYear = 0) Then
        GetDate = dtNullDate
      ElseIf (intDay >= 1) And ((intMonth >= 1) And (intMonth <= 12)) Then
        '// If the year is not valid, take current year
        If (intYear = 0) Then intYear = CInt(Year(Date))
        '// Find the last day of entered month
        intNewMonth = CInt(Month(DateSerial(intYear, intMonth, intDay)))
        intLastDay = CInt(Day(DateAdd("d", -1, DateAdd("m", 1, DateSerial(intYear, intMonth, 1)))))
        '// Generate date
        If (intDay <= intLastDay) And (intMonth = intNewMonth) Then
            dDate = DateSerial(intYear, intMonth, intDay)
            If IsInRange(dDate) Then
                GetDate = dtValidDate
              Else
                GetDate = dtInvalidDate
            End If
          Else
            GetDate = dtInvalidDate
        End If
      Else
        GetDate = dtInvalidDate
    End If

End Function

Private Function IsInRange(dDate As Date) As Boolean

    IsInRange = (dDate >= m_dtMinDate) And (dDate <= m_dtMaxDate)

End Function

Private Function GetFormat() As String

  Dim strFormat As String

    If (m_DateFormat = 1) Or (m_DateFormat = 2) Then
        strFormat = Space$(128)
        Call GetLocaleInfo(GetSystemDefaultLCID(), _
             IIf((m_DateFormat = dtShortDate), LOCALE_SSHORTDATE, LOCALE_SLONGDATE), _
             strFormat, 128&)
        strFormat = Left$(strFormat, InStr(1, strFormat, vbNullChar) - 1)
      Else
        strFormat = strCustomDateFormat
    End If
    GetFormat = strFormat

End Function

':) Ulli's VB Code Formatter V2.12.7 (7/17/02 9:16:54 AM) 309 + 2734 = 3043 Lines
