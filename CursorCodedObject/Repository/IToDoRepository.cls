VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IToDoRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Interface IToDoRepository

Public Function FindAll() As Collection
End Function

Public Function FindById(ByVal id As String) As BasicToDo
End Function

Public Function Save(ByVal entity As BasicToDo) As Boolean
End Function

Public Function Delete(ByVal id As String) As Boolean
End Function

Public Function GenerateNewId() As String
End Function

