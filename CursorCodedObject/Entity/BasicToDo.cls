VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BasicToDo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' BasicToDo Entity

Private m_basicToDoId As String
Private m_toDoName As String
Private m_scheduledExecutionDate As Date
Private m_scheduledExecutionTime As Date
Private m_deadlineDateAndTime As Date
Private m_estimatedExecutionalDuration As Date
Private m_scheduledResttimeDuration As Date
Private m_actualBeginningTime As Date
Private m_actualFinishTime As Date
Private m_generationDateAndTime As Date

Private m_milestoneId As String
Private m_toDoTypeId As String
Private m_priorityId As String

Public Property Get BasicToDoId() As String
    BasicToDoId = m_basicToDoId
End Property

Public Property Get ToDoName() As String
    ToDoName = m_toDoName
End Property

Public Property Get ScheduledExecutionDate() As Date
    ScheduledExecutionDate = m_scheduledExecutionDate
End Property

Public Property Get ScheduledExecutionTime() As Date
    ScheduledExecutionTime = m_scheduledExecutionTime
End Property

Public Property Get DeadlineDateAndTime() As Date
    DeadlineDateAndTime = m_deadlineDateAndTime
End Property

Public Property Get EstimatedExecutionalDuration() As Date
    EstimatedExecutionalDuration = m_estimatedExecutionalDuration
End Property

Public Property Get ScheduledResttimeDuration() As Date
    ScheduledResttimeDuration = m_scheduledResttimeDuration
End Property

Public Property Get ActualBeginningTime() As Date
    ActualBeginningTime = m_actualBeginningTime
End Property

Public Property Get ActualFinishTime() As Date
    ActualFinishTime = m_actualFinishTime
End Property

Public Property Get GenerationDateAndTime() As Date
    GenerationDateAndTime = m_generationDateAndTime
End Property

Public Property Get MilestoneId() As String
    MilestoneId = m_milestoneId
End Property

Public Property Get ToDoTypeId() As String
    ToDoTypeId = m_toDoTypeId
End Property

Public Property Get PriorityId() As String
    PriorityId = m_priorityId
End Property

Public Sub Init( _
    ByVal basicToDoId As String, _
    ByVal toDoName As String, _
    ByVal scheduledExecutionDate As Date, _
    ByVal scheduledExecutionTime As Date, _
    ByVal deadlineDateAndTime As Date, _
    ByVal estimatedExecutionalDuration As Date, _
    ByVal scheduledResttimeDuration As Date, _
    ByVal actualBeginningTime As Date, _
    ByVal actualFinishTime As Date, _
    ByVal generationDateAndTime As Date, _
    ByVal milestoneId As String, _
    ByVal toDoTypeId As String, _
    ByVal priorityId As String)

    m_basicToDoId = basicToDoId
    m_toDoName = toDoName
    m_scheduledExecutionDate = scheduledExecutionDate
    m_scheduledExecutionTime = scheduledExecutionTime
    m_deadlineDateAndTime = deadlineDateAndTime
    m_estimatedExecutionalDuration = estimatedExecutionalDuration
    m_scheduledResttimeDuration = scheduledResttimeDuration
    m_actualBeginningTime = actualBeginningTime
    m_actualFinishTime = actualFinishTime
    m_generationDateAndTime = generationDateAndTime
    m_milestoneId = milestoneId
    m_toDoTypeId = toDoTypeId
    m_priorityId = priorityId
End Sub

