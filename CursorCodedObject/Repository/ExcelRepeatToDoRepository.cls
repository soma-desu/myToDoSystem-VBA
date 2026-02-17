VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelRepeatToDoRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ExcelRepeatToDoRepository
' Implements IRepeatToDoRepository using Excel ListObjects as storage.

Implements IRepeatToDoRepository

Public Function IRepeatToDoRepository_FindAllInstances() As Collection
    Dim result As New Collection
    ' TODO: read all RepeatToDoInstance
    Set IRepeatToDoRepository_FindAllInstances = result
End Function

Public Function IRepeatToDoRepository_FindInstanceById(ByVal id As String) As RepeatToDoInstance
    Dim entity As RepeatToDoInstance
    ' TODO: map row to RepeatToDoInstance
    Set IRepeatToDoRepository_FindInstanceById = entity
End Function

Public Function IRepeatToDoRepository_SaveInstance(ByVal entity As RepeatToDoInstance) As Boolean
    ' TODO: insert or update
    IRepeatToDoRepository_SaveInstance = True
End Function

Public Function IRepeatToDoRepository_DeleteInstance(ByVal id As String) As Boolean
    ' TODO: delete instance
    IRepeatToDoRepository_DeleteInstance = True
End Function

Public Function IRepeatToDoRepository_GenerateNewInstanceId() As String
    ' TODO: generate RI-xxxxxx
    IRepeatToDoRepository_GenerateNewInstanceId = ""
End Function

Public Function IRepeatToDoRepository_FindPrototypeById(ByVal id As String) As RepeatToDoPrototype
    Dim entity As RepeatToDoPrototype
    ' TODO: map row to RepeatToDoPrototype
    Set IRepeatToDoRepository_FindPrototypeById = entity
End Function

Public Function IRepeatToDoRepository_SavePrototype(ByVal entity As RepeatToDoPrototype) As Boolean
    ' TODO: insert or update prototype
    IRepeatToDoRepository_SavePrototype = True
End Function

Public Function IRepeatToDoRepository_DeletePrototype(ByVal id As String) As Boolean
    ' TODO: delete prototype
    IRepeatToDoRepository_DeletePrototype = True
End Function

Public Function IRepeatToDoRepository_GenerateNewPrototypeId() As String
    ' TODO: generate RP-xxxxxx
    IRepeatToDoRepository_GenerateNewPrototypeId = ""
End Function

