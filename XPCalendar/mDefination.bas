Attribute VB_Name = "mDefination"
'Email Me: zhujy@samling.com.my
'Welcome to visit our company WebSite at:
  'www.Samling.com.my
  'www.Samling.com.cn
'Samling Group---One of a leading and largest WoodBased Industry company in Asia
'Copyright © 2001-2002 by Zhu JinYong from People Republic of China

Option Explicit

Public m_SelectModeStyle                As CalendarSelectModeStyle
Public m_CalendarSelectPicDrawMode      As CalendarSelectPictureDrawModeType
Public m_CalendarFirstDayOfWeek         As dtDaysOfTheWeek
Public m_DateFormat                     As dtFormatConstants
Public m_CalendarOption                 As CalendarDateTimeOption
Public m_CalendarDayHeaderFormat        As dtDayHeaderFormats
Public m_CalendarClickBehivor           As dtClickBehivor
Public m_CalendarBorderStyle            As CalendarBorderStyleType
Public m_DateAlign                      As DateAlignType
Public m_DatePosn                       As DatePositionType
Public m_Gridline                       As GridLineType
Public m_TodayPictureSize               As PictureSize
Public m_TodayFont                      As New StdFont
Public m_SelectionPicture               As StdPicture
Public m_TodayPicture                   As StdPicture
Public m_WeekSignPicture                As StdPicture
Public m_WeekSignPictureSize            As PictureSize
Public m_HoverSelection                 As Boolean
Public m_HotTracking                    As Boolean
Public m_ShowWeek                       As Boolean
Public m_ShowWeekSignPicture            As Boolean
Public m_TodayPictureWidth              As Long
Public m_TodayPictureHeight             As Long
Public m_TodayOriginalPicSizeW          As Long
Public m_TodayOriginalPicSizeH          As Long
Public m_WeekSignPictureWidth           As Long
Public m_WeekSignPictureHeight          As Long
Public m_WeekSignOriginalPicSizeW       As Long
Public m_WeekSignOriginalPicSizeH       As Long
Public m_WeekColumnWidth                As Long
Public m_TodayForeColor                 As OLE_COLOR
Public m_SelectedBMPMaskColor           As OLE_COLOR
Public m_TodayMaskColor                 As OLE_COLOR
Public m_BackNormal                     As OLE_COLOR
Public m_BackSelected                   As OLE_COLOR
Public m_BackSelectedG1                 As OLE_COLOR
Public m_BackSelectedG2                 As OLE_COLOR
Public m_BoxBorder                      As OLE_COLOR
Public m_FontNormal                     As OLE_COLOR
Public m_FontSelected                   As OLE_COLOR
Public m_HoverColor                     As OLE_COLOR
Public m_GridLineColor                  As OLE_COLOR
Public m_CalendarBdHighlightColour      As OLE_COLOR
Public m_CalendarBdHighlightDKColour    As OLE_COLOR
Public m_CalendarBdShadowColour         As OLE_COLOR
Public m_CalendarBdShadowDKColour       As OLE_COLOR
Public m_CalendarBdFlatBorderColour     As OLE_COLOR
Public m_WeekSignPicMaskColor           As OLE_COLOR
Public m_DateXOffset                    As Long
Public m_DateYOffset                    As Long
Public m_TodayXOffset                   As Long
Public m_TodayYOffset                   As Long
Public m_nGridWidth                     As Long
Public m_nGridHeight                    As Long
Public m_WeekSignPicXOffset             As Long
Public m_WeekSignPicYOffset             As Long
'-------------------------------------------------
Public strDateFormat                    As String
Public m_dtMaxDate           As Date
Public m_dtMinDate           As Date
'---------------------------------------------------

Public cBackNrm                         As Long                '# Back color [Normal]
Public cBackSel                         As Long                '# Back color [Selected]
Public cFontNrm                         As Long                '# Font color [Normal]
Public cBox                             As Long                '# Box border color
Public cFontSel                         As Long                '# Font color [Selected]
Public cGrad1                           As RGB                 '# Gradient color from [Selected]
Public cGrad2                           As RGB                 '# Gradient color  to  [Selected]
Public m_Value As Variant

Public Type DayCellType
    Day As Integer
    X1 As Single
    X2 As Single
    Y1 As Single
    Y2 As Single
    Img As Picture
    Tips As String
End Type

'Storage for daily information
Public DayCell(53)              As DayCellType




