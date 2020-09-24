VERSION 5.00
Object = "{B641EEDB-8451-11D6-BD5F-0010A4F59E39}#1.1#0"; "XPCALENDAR.OCX"
Begin VB.Form Form1 
   Caption         =   "Zhu's Drawdown XPCalendar"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   ScaleHeight     =   8355
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin XP_Calendar.XPCalendar XPCalendar1 
      Height          =   375
      Left            =   240
      TabIndex        =   48
      Top             =   1680
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      Icon            =   "Form1.frx":0000
      FontSize        =   9.75
      CalendarSelectPicDrawMode=   1
      CalendarSelectePicMaskColor=   16777215
      CalendarSelectePicture=   "Form1.frx":059A
      CalendarBorderStyle=   3
      TodayForeColor  =   0
      TodayFontName   =   "MS Serif"
      TodayFontSize   =   8.25
      TodayFontBold   =   0   'False
      TodayFontItalic =   0   'False
      TodayMaskColor  =   16777215
      TodayPicture    =   "Form1.frx":08EC
      TodayPictureWidth=   26
      TodayPictureHeight=   15
      TodayPictureSize=   2
      TodayOriginalPicSizeW=   26
      TodayOriginalPicSizeH=   15
      DateAlign       =   1
      DatePosn        =   1
      WeekColumnWidth =   28
      WeekSignPicXOffset=   2
      WeekSignPictureWidth=   11
      WeekSignPictureHeight=   10
      WeekSignPictureSize=   2
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   41
      Top             =   480
      Width           =   3255
   End
   Begin VB.ComboBox comSelectedStyle 
      Height          =   315
      ItemData        =   "Form1.frx":0DEE
      Left            =   5760
      List            =   "Form1.frx":0DFE
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   600
      Width           =   2445
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   8175
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      Begin VB.VScrollBar VSWeekColumnWidth 
         Height          =   375
         Left            =   3960
         TabIndex        =   58
         Top             =   6840
         Width           =   255
      End
      Begin VB.TextBox txtWeekColumnWidth 
         Height          =   375
         Left            =   2520
         TabIndex        =   57
         Top             =   6840
         Width           =   1695
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Show Week Sign Picture"
         Height          =   255
         Left            =   2280
         TabIndex        =   56
         Top             =   6360
         Width           =   2295
      End
      Begin VB.VScrollBar VSWeekSignPicYOffset 
         Height          =   375
         Left            =   3960
         Max             =   50
         Min             =   -50
         TabIndex        =   53
         Top             =   7560
         Width           =   255
      End
      Begin VB.TextBox txtWeekSignPicYOffset 
         Height          =   375
         Left            =   2520
         TabIndex        =   52
         Top             =   7560
         Width           =   1695
      End
      Begin VB.VScrollBar VSWeekSignPicXOffset 
         Height          =   375
         Left            =   1800
         Max             =   50
         Min             =   -50
         TabIndex        =   51
         Top             =   7560
         Width           =   255
      End
      Begin VB.TextBox txtWeekSignPicXOffset 
         Height          =   405
         Left            =   360
         TabIndex        =   50
         Top             =   7560
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Week Number Visible"
         Height          =   255
         Left            =   2280
         TabIndex        =   49
         Top             =   6000
         Width           =   1935
      End
      Begin VB.ComboBox comEditDateFormat 
         Height          =   315
         ItemData        =   "Form1.frx":0E3F
         Left            =   2280
         List            =   "Form1.frx":0E55
         TabIndex        =   47
         Text            =   "Edit Date Format"
         Top             =   5640
         Width           =   2055
      End
      Begin VB.TextBox txtdtCustomDateFormat 
         Height          =   375
         Left            =   2280
         TabIndex        =   44
         Text            =   "yyyy-mm-dd"
         Top             =   4920
         Width           =   1935
      End
      Begin VB.ComboBox comDateFormat 
         Height          =   315
         ItemData        =   "Form1.frx":0EBF
         Left            =   2280
         List            =   "Form1.frx":0ECC
         TabIndex        =   42
         Text            =   "Calendar Date Format"
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox comGridLine 
         Height          =   315
         ItemData        =   "Form1.frx":0F03
         Left            =   360
         List            =   "Form1.frx":0F1C
         TabIndex        =   39
         Top             =   6480
         Width           =   1695
      End
      Begin VB.VScrollBar VSTodayYOffset 
         Height          =   375
         Left            =   1800
         Max             =   20
         Min             =   -20
         TabIndex        =   38
         Top             =   5760
         Width           =   255
      End
      Begin VB.TextBox txtTodayYOffset 
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   5760
         Width           =   1695
      End
      Begin VB.VScrollBar VSTodayXOffset 
         Height          =   375
         Left            =   1800
         Max             =   20
         Min             =   -20
         TabIndex        =   35
         Top             =   5040
         Width           =   255
      End
      Begin VB.TextBox txtTodayXOffset 
         Height          =   375
         Left            =   360
         TabIndex        =   34
         Top             =   5040
         Width           =   1695
      End
      Begin VB.VScrollBar VSDateYOffset 
         Height          =   375
         Left            =   1800
         Max             =   20
         Min             =   -20
         TabIndex        =   32
         Top             =   4320
         Width           =   255
      End
      Begin VB.TextBox txtDateYOffset 
         Height          =   405
         Left            =   360
         TabIndex        =   31
         Text            =   "DateYOffset"
         Top             =   4320
         Width           =   1695
      End
      Begin VB.VScrollBar VSDateXOffset 
         Height          =   375
         Left            =   1800
         Max             =   50
         Min             =   -50
         TabIndex        =   29
         Top             =   3600
         Width           =   255
      End
      Begin VB.TextBox txtDateXOffset 
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Text            =   "DateXOffset"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.ComboBox comDatePosn 
         Height          =   315
         ItemData        =   "Form1.frx":0FAB
         Left            =   360
         List            =   "Form1.frx":0FB8
         TabIndex        =   26
         Text            =   "Date Postion"
         Top             =   3000
         Width           =   1695
      End
      Begin VB.ComboBox comDateAlign 
         Height          =   315
         ItemData        =   "Form1.frx":0FE8
         Left            =   360
         List            =   "Form1.frx":0FF5
         TabIndex        =   24
         Text            =   "DateAlignment"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.ComboBox comBoardStyle 
         Height          =   315
         ItemData        =   "Form1.frx":1029
         Left            =   360
         List            =   "Form1.frx":1045
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         Caption         =   "HotTracking=True"
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Caption         =   "HoverSelection =True"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   3480
         Width           =   1935
      End
      Begin VB.ComboBox comClickBehivor 
         Height          =   315
         ItemData        =   "Form1.frx":10BB
         Left            =   2160
         List            =   "Form1.frx":10C5
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.ComboBox comDayCaption 
         Height          =   315
         ItemData        =   "Form1.frx":10F1
         Left            =   2160
         List            =   "Form1.frx":10FE
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   2520
         Width           =   2415
      End
      Begin VB.ComboBox comFirstDayofWeek 
         Height          =   315
         ItemData        =   "Form1.frx":1134
         Left            =   2160
         List            =   "Form1.frx":114D
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   1800
         Width           =   2415
      End
      Begin VB.ComboBox comCalendarOption 
         Height          =   315
         ItemData        =   "Form1.frx":11BB
         Left            =   2160
         List            =   "Form1.frx":11CE
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1080
         Width           =   2415
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   480
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Week Number Column Width:"
         Height          =   255
         Left            =   2160
         TabIndex        =   59
         Top             =   6600
         Width           =   2535
      End
      Begin VB.Label Label22 
         Caption         =   "Week Sign Picture YOffset"
         Height          =   255
         Left            =   2400
         TabIndex        =   55
         Top             =   7320
         Width           =   2055
      End
      Begin VB.Label Label21 
         Caption         =   "Week Sign Picture XOffset"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   7320
         Width           =   2055
      End
      Begin VB.Label Label20 
         Caption         =   "Calendar Edit Date Format :"
         Height          =   255
         Left            =   2280
         TabIndex        =   46
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label19 
         Caption         =   "Calendar Custom Date Format :"
         Height          =   255
         Left            =   2160
         TabIndex        =   45
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Calendar Date Format :"
         Height          =   255
         Left            =   2160
         TabIndex        =   43
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Grid Line:"
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "TodayYOffset :"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "TodayXOffset :"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "DateYOffset :"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "DateXOffset :"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Date Text Position :"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "DateAlignment:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Calendar Board Style"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Click/Double Click to Hide :"
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Day Caption :"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "FirstDayofWeek:"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Calendar Selected Highlight Style :"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "CalendarOption"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Calendar Height :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Calendar Width :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "XPCalendar1_CalendarChoose(    )"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Email Me: zhujy@samling.com.my
'Welcome to visit our company WebSite at:
  'www.Samling.com.my
  'www.Samling.com.cn
'Samling Group---One of a leading and largest WoodBased Industry company in Asia
'Copyright only credit to The Author

Option Explicit

Private Sub Check1_Click()
If Check1.Value = 1 Then

     XPCalendar1.ShowWeek = True
  Else
     XPCalendar1.ShowWeek = False
  End If
End Sub

Private Sub Check2_Click()
  If Check2.Value = 1 Then

     XPCalendar1.HoverSelection = True
  Else
     XPCalendar1.HoverSelection = False
  End If
End Sub
Private Sub Check3_Click()
  If Check3.Value = 1 Then

     XPCalendar1.HotTracking = True
  Else
     XPCalendar1.HotTracking = False
  End If
End Sub

Private Sub Check4_Click()
  If Check4.Value = 1 Then

     XPCalendar1.ShowWeekSignPicture = True
  Else
     XPCalendar1.ShowWeekSignPicture = False
  End If
End Sub

Private Sub comBoardStyle_Click()
XPCalendar1.CalendarBorderStyle = comBoardStyle.ListIndex
End Sub

Private Sub comCalendarOption_Click()
    XPCalendar1.CalendarOption = comCalendarOption.ListIndex
End Sub



Private Sub comDateAlign_Click()
XPCalendar1.DateAlign = comDateAlign.ListIndex
comDateAlign.ListIndex = XPCalendar1.DateAlign
End Sub


Private Sub comDateFormat_Click()
XPCalendar1.CalendarDateFormat = comDateFormat.ListIndex + 1
If XPCalendar1.CalendarDateFormat = dtCustom Then
   XPCalendar1.CustomDateFormat = txtdtCustomDateFormat.Text
End If
End Sub

Private Sub comDatePosn_Click()
XPCalendar1.DatePosn = comDatePosn.ListIndex
    comDatePosn.ListIndex = XPCalendar1.DatePosn
End Sub



Private Sub comEditDateFormat_Click()
XPCalendar1.DateFormatWhenEdit = comEditDateFormat.ListIndex + 1
End Sub

Private Sub comGridLine_Click()
XPCalendar1.GridLine = comGridLine.ListIndex
    comGridLine.ListIndex = XPCalendar1.GridLine
End Sub
Private Sub comSelectedStyle_Click()
    XPCalendar1.CalendarSelectStyle = comSelectedStyle.ListIndex
End Sub
Private Sub comDayCaption_Click()
    XPCalendar1.CalendarDayHeaderFormat = comDayCaption.ListIndex
End Sub
Private Sub comFirstDayofWeek_Click()
    XPCalendar1.CalendarFirstDayOfWeek = comFirstDayofWeek.ListIndex + 1
End Sub
Private Sub comclickbehivor_Click()
    XPCalendar1.CalendarClickBehivor = comClickBehivor.ListIndex
End Sub



Private Sub Form_Load()

'initialize Visual Features
    XPCalendar1.ShowWeek = True
    XPCalendar1.CalendarGridHeight = 35
    XPCalendar1.CalendarGridWidth = 30
    XPCalendar1.CalendarBorderStyle = [Raised 3D]
    XPCalendar1.CalendarSelectStyle = byPicture
    XPCalendar1.DateAlign = AlignCenter
    XPCalendar1.DatePosn = PosnCenter
    XPCalendar1.GridLine = 6 'None GridLine
    XPCalendar1.CalendarSelectPicDrawMode = PictureTiled
    XPCalendar1.HotTracking = True
    XPCalendar1.HoverSelection = True
    XPCalendar1.FontSize = 10
   
'---------------------------------------------
      Text1.Text = XPCalendar1.Value
      Text2 = XPCalendar1.CalendarGridHeight
      VScroll1 = XPCalendar1.CalendarGridHeight
      Text3 = XPCalendar1.CalendarGridWidth
      VScroll2 = XPCalendar1.CalendarGridWidth
      txtDateXOffset = XPCalendar1.DateXOffset
      txtDateYOffset = XPCalendar1.DateYOffset
      VSDateXOffset = XPCalendar1.DateXOffset
      VSDateYOffset = XPCalendar1.DateYOffset
      
      txtTodayXOffset = XPCalendar1.TodayXOffset
      txtTodayYOffset = XPCalendar1.TodayYOffset
      VSTodayXOffset = XPCalendar1.TodayXOffset
      VSTodayYOffset = XPCalendar1.TodayYOffset
      VSWeekColumnWidth = XPCalendar1.WeekColumnWidth
      txtWeekColumnWidth = XPCalendar1.WeekColumnWidth
      txtWeekSignPicXOffset = XPCalendar1.WeekSignPicXOffset
      txtWeekSignPicYOffset = XPCalendar1.WeekSignPicYOffset
      VSWeekSignPicXOffset = XPCalendar1.WeekSignPicXOffset
      VSWeekSignPicYOffset = XPCalendar1.WeekSignPicYOffset
      
      comBoardStyle.ListIndex = XPCalendar1.CalendarBorderStyle
      comSelectedStyle.ListIndex = XPCalendar1.CalendarSelectStyle
      comCalendarOption.ListIndex = XPCalendar1.CalendarOption
      comDayCaption.ListIndex = XPCalendar1.CalendarDayHeaderFormat
      comFirstDayofWeek.ListIndex = XPCalendar1.CalendarFirstDayOfWeek - 1
      comClickBehivor.ListIndex = XPCalendar1.CalendarClickBehivor
      comDateFormat.ListIndex = XPCalendar1.CalendarDateFormat - 1
      comDateAlign.ListIndex = XPCalendar1.DateAlign
      comDatePosn.ListIndex = XPCalendar1.DatePosn
      comGridLine.ListIndex = XPCalendar1.GridLine
      comEditDateFormat.ListIndex = XPCalendar1.DateFormatWhenEdit - 1
     
    If XPCalendar1.ShowWeek = True Then
         Check1.Value = 1
    Else
         Check1.Value = 0
    End If
    If XPCalendar1.HoverSelection = True Then
         Check2.Value = 1
    Else
         Check2.Value = 0
    End If
     If XPCalendar1.HotTracking = True Then
         Check3.Value = 1
    Else
         Check3.Value = 0
    End If
    
    If XPCalendar1.ShowWeekSignPicture = True Then
         Check4.Value = 1
    Else
         Check4.Value = 0
    End If
End Sub
Private Sub Text2_Change()
      XPCalendar1.CalendarGridHeight = Val(Text2)
End Sub
Private Sub Text3_Change()
      XPCalendar1.CalendarGridWidth = Val(Text3)
End Sub

Private Sub txtDateXOffset_Change()
XPCalendar1.DateXOffset = Val(txtDateXOffset)
End Sub
Private Sub txtDateYOffset_Change()
XPCalendar1.DateYOffset = Val(txtDateYOffset)
End Sub

Private Sub txtdtCustomDateFormat_Change()
XPCalendar1.CustomDateFormat = txtdtCustomDateFormat.Text
End Sub

Private Sub txtTodayXOffset_Change()
XPCalendar1.TodayXOffset = Val(txtTodayXOffset)
End Sub
Private Sub txtTodayYOffset_Change()
XPCalendar1.TodayYOffset = Val(txtTodayYOffset)
End Sub
Private Sub txtWeekSignPicXOffset_Change()
XPCalendar1.WeekSignPicXOffset = Val(txtWeekSignPicXOffset)
End Sub
Private Sub txtWeekSignPicYOffset_Change()
XPCalendar1.WeekSignPicYOffset = Val(txtWeekSignPicYOffset)
End Sub
Private Sub txtWeekcolumnWidth_Change()
XPCalendar1.WeekColumnWidth = Val(txtWeekColumnWidth)
End Sub
Private Sub VScroll1_Change()
      Text2 = VScroll1
End Sub
Private Sub VScroll2_Change()
      Text3 = VScroll2
End Sub

Private Sub VSDateXOffset_Change()
txtDateXOffset = VSDateXOffset
End Sub

Private Sub VSDateYOffset_Change()
txtDateYOffset = VSDateYOffset
End Sub
Private Sub VSTodayXOffset_Change()
txtTodayXOffset = VSTodayXOffset
End Sub
Private Sub VSTodayYOffset_Change()
txtTodayYOffset = VSTodayYOffset
End Sub

Private Sub VSWeekColumnWidth_Change()
txtWeekColumnWidth = VSWeekColumnWidth
End Sub

Private Sub VSWeekSignPicXOffset_Change()
txtWeekSignPicXOffset = VSWeekSignPicXOffset
End Sub
Private Sub VSWeekSignPicYOffset_Change()
txtWeekSignPicYOffset = VSWeekSignPicYOffset
End Sub
Private Sub XPCalendar1_CalendarChoose(Text As String)
      Text1.Text = XPCalendar1.OutputText
      Text4.Text = XPCalendar1.Value
End Sub




