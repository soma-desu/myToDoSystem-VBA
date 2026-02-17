VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelMasterRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ExcelMasterRepository
' Implements IMasterRepository using Excel as storage.

Implements IMasterRepository

Public Function IMasterRepository_FindAllClients() As Collection
    Dim result As New Collection
    ' TODO: map rows to Client entities
    Set IMasterRepository_FindAllClients = result
End Function

Public Function IMasterRepository_FindAllCampaigns() As Collection
    Dim result As New Collection
    ' TODO: map rows to Campaign entities
    Set IMasterRepository_FindAllCampaigns = result
End Function

Public Function IMasterRepository_FindAllMilestones() As Collection
    Dim result As New Collection
    ' TODO: map rows to Milestone entities
    Set IMasterRepository_FindAllMilestones = result
End Function

Public Function IMasterRepository_FindAllToDoTypes() As Collection
    Dim result As New Collection
    ' TODO: map rows to ToDoType entities
    Set IMasterRepository_FindAllToDoTypes = result
End Function

Public Function IMasterRepository_FindAllPriorities() As Collection
    Dim result As New Collection
    ' TODO: map rows to Priority entities
    Set IMasterRepository_FindAllPriorities = result
End Function

Public Function IMasterRepository_FindAllEventTypes() As Collection
    Dim result As New Collection
    ' TODO: map rows to EventType entities
    Set IMasterRepository_FindAllEventTypes = result
End Function

