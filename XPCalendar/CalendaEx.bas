Attribute VB_Name = "mCalendarEx"
'Email Me: zhujy@samling.com.my
'Welcome to visit our company WebSite at:
  'www.Samling.com.my
  'www.Samling.com.cn
'Samling Group---One of a leading and largest WoodBased Industry company in Asia
'Copyright © 2001-2002 by Zhu JinYong from People Republic of China

Option Explicit

Public Function WeekNumber(DateIn As Date) As Integer
' Gets the weeknumber for the currently selected date
    WeekNumber = DatePart("ww", DateIn, 0)
End Function

Public Function Quarter(DateIn As Date) As Integer
' Gets the weeknumber for the currently selected date
    Quarter = DatePart("q", DateIn)
End Function
