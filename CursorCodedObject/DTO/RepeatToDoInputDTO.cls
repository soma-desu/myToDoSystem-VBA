VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepeatToDoInputDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' RepeatToDoInputDTO
' DTO for user input for repeat ToDo prototype/instance.

Private m_campaignName As String
Private m_milestoneName As String
Private m_toDoName As String
Private m_toDoTypeName As String
Private m_scheduledExecutionTime As Date
Private m_spanUntilDeadline As Long
Private m_estimatedExecutionalDuration As Date
Private m_scheduledResttimeDuration As Date
Private m_priority As String

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

Public Property Get ScheduledExecutionTime() As Date
    ScheduledExecutionTime = m_scheduledExecutionTime
End Property
Public Property Let ScheduledExecutionTime(ByVal value As Date)
    m_scheduledExecutionTime = value
End Property

Public Property Get SpanUntilDeadline() As Long
    SpanUntilDeadline = m_spanUntilDeadline
End Property
Public Property Let SpanUntilDeadline(ByVal value As Long)
    m_spanUntilDeadline = value
End Property

Public Property Get EstimatedExecutionalDuration() As Date
    EstimatedExecutionalDuration = m_estimatedExecutionalDuration
End Property
Public Property Let EstimatedExecutionalDuration(ByVal value As Date)
    m_estimatedExecutionalDuration = value
End Property

Public Property Get ScheduledResttimeDuration() As Date
    ScheduledResttimeDuration = m_scheduledResttimeDuration
End Property
Public Property Let ScheduledResttimeDuration(ByVal value As Date)
    m_scheduledResttimeDuration = value
End Property

Public Property Get Priority() As String
    Priority = m_priority
End Property
Public Property Let Priority(ByVal value As String)
    m_priority = value
End Property

