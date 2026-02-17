Attribute VB_Name = "DateTimeHelper"
Option Explicit

' DateTimeHelper
' Common helpers for date/time manipulation and formatting.

Public Function CombineDateTime(ByVal d As Date, ByVal t As Date) As Date
    CombineDateTime = DateSerial(Year(d), Month(d), Day(d)) + _
                      TimeSerial(Hour(t), Minute(t), Second(t))
End Function

Public Function StartOfDay(ByVal d As Date) As Date
    StartOfDay = DateSerial(Year(d), Month(d), Day(d))
End Function

Public Function EndOfDay(ByVal d As Date) As Date
    EndOfDay = StartOfDay(d) + TimeSerial(23, 59, 59)
End Function

