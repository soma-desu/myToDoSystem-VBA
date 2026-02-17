VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelToDoRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ExcelToDoRepository
' Implements IToDoRepository using Excel ListObject as storage.

Implements IToDoRepository

Public Function IToDoRepository_FindAll() As Collection
    Dim result As New Collection
    ' TODO: read all BasicToDo records from ListObject
    Set IToDoRepository_FindAll = result
End Function

Public Function IToDoRepository_FindById(ByVal id As String) As BasicToDo
    Dim entity As BasicToDo
    ' TODO: find row by id and map to BasicToDo
    Set IToDoRepository_FindById = entity
End Function

Public Function IToDoRepository_Save(ByVal entity As BasicToDo) As Boolean
    ' TODO: insert or update row in ListObject
    IToDoRepository_Save = True
End Function

Public Function IToDoRepository_Delete(ByVal id As String) As Boolean
    ' TODO: delete row by id
    IToDoRepository_Delete = True
End Function

Public Function IToDoRepository_GenerateNewId() As String
    ' TODO: generate new BT-xxxxxx id
    IToDoRepository_GenerateNewId = ""
End Function

