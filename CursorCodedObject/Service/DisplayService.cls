VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DisplayService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' DisplayService
' Builds DTOs for various views from domain entities.

Public Function BuildTodayToDoList(ByVal allToDos As Collection, _
                                   ByVal today As Date) As Collection
    ' Returns Collection of ToDoRowDTO
    Dim result As New Collection
    ' TODO: filter and transform BasicToDo/RepeatToDoInstance into ToDoRowDTO
    Set BuildTodayToDoList = result
End Function

Public Function BuildTomorrowToDoList(ByVal allToDos As Collection, _
                                      ByVal tomorrow As Date, _
                                      ByVal holidays As Collection) As Collection
    Dim result As New Collection
    ' TODO: compute next work day, filter and transform
    Set BuildTomorrowToDoList = result
End Function

Public Function BuildSpecificDayToDoList(ByVal allToDos As Collection, _
                                         ByVal targetDate As Date) As Collection
    Dim result As New Collection
    ' TODO: filter by date and transform
    Set BuildSpecificDayToDoList = result
End Function

Public Function BuildTwoWeekScheduleCalendar(ByVal allToDos As Collection, _
                                             ByVal startDate As Date) As Collection
    Dim result As New Collection
    ' Returns Collection of DayScheduleDTO
    ' TODO: build 14-day schedule
    Set BuildTwoWeekScheduleCalendar = result
End Function

Public Function BuildDailyExecutedTaskTime(ByVal allToDos As Collection, _
                                           ByVal yearMonth As Date, _
                                           ByVal granularity As String) As Collection
    Dim result As New Collection
    ' Returns Collection of ExecutedTaskRowDTO
    ' TODO: aggregate executed time per day and grouping granularity
    Set BuildDailyExecutedTaskTime = result
End Function

