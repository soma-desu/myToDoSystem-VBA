VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RawEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' RawEvent Entity
' Represents an input event from Excel (sheet, listobject, target range, etc.)

Private m_eventId As String
Private m_eventTypeId As String
Private m_sheetName As String
Private m_listObjectName As String
Private m_targetAddress As String
Private m_generationDateAndTime As Date

Public Property Get EventId() As String
    EventId = m_eventId
End Property

Public Property Get EventTypeId() As String
    EventTypeId = m_eventTypeId
End Property

Public Property Get SheetName() As String
    SheetName = m_sheetName
End Property

Public Property Get ListObjectName() As String
    ListObjectName = m_listObjectName
End Property

Public Property Get TargetAddress() As String
    TargetAddress = m_targetAddress
End Property

Public Property Get GenerationDateAndTime() As Date
    GenerationDateAndTime = m_generationDateAndTime
End Property

Public Sub Init( _
    ByVal eventId As String, _
    ByVal eventTypeId As String, _
    ByVal sheetName As String, _
    ByVal listObjectName As String, _
    ByVal targetAddress As String, _
    ByVal generationDateAndTime As Date)

    m_eventId = eventId
    m_eventTypeId = eventTypeId
    m_sheetName = sheetName
    m_listObjectName = listObjectName
    m_targetAddress = targetAddress
    m_generationDateAndTime = generationDateAndTime
End Sub

