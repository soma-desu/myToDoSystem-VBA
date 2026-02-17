VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToDoRowDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ToDoRowDTO
' Output DTO for ToDo table row (matches IOTableDesign columns).

Private m_toDoId As String
Private m_clientName As String
Private m_campaignName As String
Private m_milestoneName As String
Private m_toDoName As String
Private m_toDoTypeName As String
Private m_scheduledExecutionTime As String
Private m_estimatedExecutionalDuration As String
Private m_scheduledResttimeDuration As String
Private m_deadlineDateAndTime As String
Private m_actualBeginningTime As String
Private m_actualFinishTime As String
Private m_priority As String
Private m_scheduledExecutionDate As String

Public Property Get ToDoId() As String
    ToDoId = m_toDoId
End Property
Public Property Let ToDoId(ByVal value As String)
    m_toDoId = value
End Property

Public Property Get ClientName() As String
    ClientName = m_clientName
End Property
Public Property Let ClientName(ByVal value As String)
    m_clientName = value
End Property

Public Property Get CampaignName() As String
    CampaignName = m_campaignName
End Property
Public Property Let CampaignName(ByVal value As String)
    m_campaignName = value
End Property

Public Property Get MilestoneName() As String
    MilestoneName = m_milestoneName
End Property
Public Property Let MilestoneName(ByVal value As String)
    m_milestoneName = value
End Property

Public Property Get ToDoName() As String
    ToDoName = m_toDoName
End Property
Public Property Let ToDoName(ByVal value As String)
    m_toDoName = value
End Property

Public Property Get ToDoTypeName() As String
    ToDoTypeName = m_toDoTypeName
End Property
Public Property Let ToDoTypeName(ByVal value As String)
    m_toDoTypeName = value
End Property

Public Property Get ScheduledExecutionTime() As String
    ScheduledExecutionTime = m_scheduledExecutionTime
End Property
Public Property Let ScheduledExecutionTime(ByVal value As String)
    m_scheduledExecutionTime = value
End Property

Public Property Get EstimatedExecutionalDuration() As String
    EstimatedExecutionalDuration = m_estimatedExecutionalDuration
End Property
Public Property Let EstimatedExecutionalDuration(ByVal value As String)
    m_estimatedExecutionalDuration = value
End Property

Public Property Get ScheduledResttimeDuration() As String
    ScheduledResttimeDuration = m_scheduledResttimeDuration
End Property
Public Property Let ScheduledResttimeDuration(ByVal value As String)
    m_scheduledResttimeDuration = value
End Property

Public Property Get DeadlineDateAndTime() As String
    DeadlineDateAndTime = m_deadlineDateAndTime
End Property
Public Property Let DeadlineDateAndTime(ByVal value As String)
    m_deadlineDateAndTime = value
End Property

Public Property Get ActualBeginningTime() As String
    ActualBeginningTime = m_actualBeginningTime
End Property
Public Property Let ActualBeginningTime(ByVal value As String)
    m_actualBeginningTime = value
End Property

Public Property Get ActualFinishTime() As String
    ActualFinishTime = m_actualFinishTime
End Property
Public Property Let ActualFinishTime(ByVal value As String)
    m_actualFinishTime = value
End Property

Public Property Get Priority() As String
    Priority = m_priority
End Property
Public Property Let Priority(ByVal value As String)
    m_priority = value
End Property

Public Property Get ScheduledExecutionDate() As String
    ScheduledExecutionDate = m_scheduledExecutionDate
End Property
Public Property Let ScheduledExecutionDate(ByVal value As String)
    m_scheduledExecutionDate = value
End Property

