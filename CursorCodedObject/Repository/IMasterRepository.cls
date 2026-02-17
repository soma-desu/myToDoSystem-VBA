VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMasterRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Interface IMasterRepository

Public Function FindAllClients() As Collection
End Function

Public Function FindAllCampaigns() As Collection
End Function

Public Function FindAllMilestones() As Collection
End Function

Public Function FindAllToDoTypes() As Collection
End Function

Public Function FindAllPriorities() As Collection
End Function

Public Function FindAllEventTypes() As Collection
End Function

