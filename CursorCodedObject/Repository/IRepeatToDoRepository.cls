VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRepeatToDoRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Interface IRepeatToDoRepository

Public Function FindAllInstances() As Collection
End Function

Public Function FindInstanceById(ByVal id As String) As RepeatToDoInstance
End Function

Public Function SaveInstance(ByVal entity As RepeatToDoInstance) As Boolean
End Function

Public Function DeleteInstance(ByVal id As String) As Boolean
End Function

Public Function GenerateNewInstanceId() As String
End Function

Public Function FindPrototypeById(ByVal id As String) As RepeatToDoPrototype
End Function

Public Function SavePrototype(ByVal entity As RepeatToDoPrototype) As Boolean
End Function

Public Function DeletePrototype(ByVal id As String) As Boolean
End Function

Public Function GenerateNewPrototypeId() As String
End Function

