VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' EventType Entity

Private m_eventTypeId As String
Private m_eventName As String
Private m_description As String

Public Property Get EventTypeId() As String
    EventTypeId = m_eventTypeId
End Property

Public Property Get EventName() As String
    EventName = m_eventName
End Property

Public Property Get Description() As String
    Description = m_description
End Property

Public Sub Init(ByVal eventTypeId As String, _
                ByVal eventName As String, _
                ByVal description As String)
    m_eventTypeId = eventTypeId
    m_eventName = eventName
    m_description = description
End Sub

