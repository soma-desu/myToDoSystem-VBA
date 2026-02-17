VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' SheetPresenter
' Responsible only for writing DTOs to Worksheet/ListObject.

Public Sub RenderToDoList(ByVal sheet As Worksheet, _
                          ByVal rows As Collection)
    ' rows: Collection of ToDoRowDTO
    ' TODO: write rows into target ListObject
End Sub

Public Sub RenderTwoWeekCalendar(ByVal sheet As Worksheet, _
                                 ByVal schedule As Collection)
    ' schedule: Collection of DayScheduleDTO
    ' TODO: render calendar view to sheet
End Sub

Public Sub RenderExecutedTaskSummary(ByVal sheet As Worksheet, _
                                     ByVal summary As Collection)
    ' summary: Collection of ExecutedTaskRowDTO
    ' TODO: render aggregation view to sheet
End Sub

