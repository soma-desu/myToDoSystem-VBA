Attribute VB_Name = "HolidayService"
Option Explicit

' HolidayService
' Pure helper for holiday-related date calculations.

Public Function GetNextWorkDay(ByVal today As Date, _
                               ByVal holidays As Collection) As Date
    ' TODO: implement skipping holidays and weekends as needed.
    GetNextWorkDay = today
End Function

