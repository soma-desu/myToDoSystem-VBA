VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToDoService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ToDoService
' Pure orchestration of BasicToDo / RepeatToDo use cases.

Private m_basicToDoRepository As IToDoRepository
Private m_repeatToDoRepository As IRepeatToDoRepository
Private m_holidayRepository As IHolidayRepository

Public Property Set BasicToDoRepository(ByVal repo As IToDoRepository)
    Set m_basicToDoRepository = repo
End Property

Public Property Set RepeatToDoRepository(ByVal repo As IRepeatToDoRepository)
    Set m_repeatToDoRepository = repo
End Property

Public Property Set HolidayRepository(ByVal repo As IHolidayRepository)
    Set m_holidayRepository = repo
End Property

' --- BasicToDo CRUD ---

Public Function AddBasicToDo(ByVal input As BasicToDoInputDTO) As ServiceResult
    Dim result As New ServiceResult
    ' TODO: validation + entity construction + repository save
    Set AddBasicToDo = result
End Function

Public Function EditBasicToDo(ByVal toDoId As String, ByVal input As BasicToDoInputDTO) As ServiceResult
    Dim result As New ServiceResult
    ' TODO: load existing, apply changes immutably, validate, save
    Set EditBasicToDo = result
End Function

Public Function DoneBasicToDo(ByVal toDoId As String, ByVal finishTime As Date) As ServiceResult
    Dim result As New ServiceResult
    ' TODO: mark as done (set actual finish/beginning), save
    Set DoneBasicToDo = result
End Function

Public Function DuplicateBasicToDo(ByVal toDoId As String) As ServiceResult
    Dim result As New ServiceResult
    ' TODO: load, clone with new id, save
    Set DuplicateBasicToDo = result
End Function

Public Function DeleteBasicToDo(ByVal toDoId As String) As ServiceResult
    Dim result As New ServiceResult
    ' TODO: delete via repository
    Set DeleteBasicToDo = result
End Function

' --- RepeatToDo CRUD ---

Public Function AddRepeatToDo(ByVal input As RepeatToDoInputDTO) As ServiceResult
    Dim result As New ServiceResult
    ' TODO: create prototype and/or first instance
    Set AddRepeatToDo = result
End Function

Public Function EditRepeatToDo(ByVal instanceId As String, ByVal input As RepeatToDoInputDTO) As ServiceResult
    Dim result As New ServiceResult
    ' TODO: edit repeat instance or prototype depending on isUnique
    Set EditRepeatToDo = result
End Function

Public Function DoneRepeatToDo(ByVal instanceId As String, ByVal finishTime As Date) As ServiceResult
    Dim result As New ServiceResult
    ' TODO: mark repeat instance as done
    Set DoneRepeatToDo = result
End Function

Public Function DuplicateRepeatToDo(ByVal instanceId As String) As ServiceResult
    Dim result As New ServiceResult
    ' TODO: duplicate instance or create new series
    Set DuplicateRepeatToDo = result
End Function

Public Function DeleteRepeatToDo(ByVal instanceId As String) As ServiceResult
    Dim result As New ServiceResult
    ' TODO: delete repeat instance / prototype
    Set DeleteRepeatToDo = result
End Function

