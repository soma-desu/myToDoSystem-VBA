VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Priority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Priority Entity

Private m_priorityId As String
Private m_priorityTag As String
Private m_description As String

Public Property Get PriorityId() As String
    PriorityId = m_priorityId
End Property

Public Property Get PriorityTag() As String
    PriorityTag = m_priorityTag
End Property

Public Property Get Description() As String
    Description = m_description
End Property

Public Sub Init(ByVal priorityId As String, _
                ByVal priorityTag As String, _
                ByVal description As String)
    m_priorityId = priorityId
    m_priorityTag = priorityTag
    m_description = description
End Sub

