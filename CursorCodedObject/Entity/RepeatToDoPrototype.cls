VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepeatToDoPrototype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' RepeatToDoPrototype Entity

Private m_repPrototypeId As String
Private m_defaultToDoName As String
Private m_defaultScheduledExecutionTime As Date
Private m_defaultSpanUntilDeadline As Long
Private m_defaultEstimatedExecutionalDuration As Date
Private m_defaultScheduledResttimeDuration As Date
Private m_generationDateAndTime As Date

Private m_defaultMilestoneId As String
Private m_defaultToDoTypeId As String
Private m_defaultPriorityId As String

Public Property Get RepPrototypeId() As String
    RepPrototypeId = m_repPrototypeId
End Property

Public Property Get DefaultToDoName() As String
    DefaultToDoName = m_defaultToDoName
End Property

Public Property Get DefaultScheduledExecutionTime() As Date
    DefaultScheduledExecutionTime = m_defaultScheduledExecutionTime
End Property

Public Property Get DefaultSpanUntilDeadline() As Long
    DefaultSpanUntilDeadline = m_defaultSpanUntilDeadline
End Property

Public Property Get DefaultEstimatedExecutionalDuration() As Date
    DefaultEstimatedExecutionalDuration = m_defaultEstimatedExecutionalDuration
End Property

Public Property Get DefaultScheduledResttimeDuration() As Date
    DefaultScheduledResttimeDuration = m_defaultScheduledResttimeDuration
End Property

Public Property Get GenerationDateAndTime() As Date
    GenerationDateAndTime = m_generationDateAndTime
End Property

Public Property Get DefaultMilestoneId() As String
    DefaultMilestoneId = m_defaultMilestoneId
End Property

Public Property Get DefaultToDoTypeId() As String
    DefaultToDoTypeId = m_defaultToDoTypeId
End Property

Public Property Get DefaultPriorityId() As String
    DefaultPriorityId = m_defaultPriorityId
End Property

Public Sub Init( _
    ByVal repPrototypeId As String, _
    ByVal defaultToDoName As String, _
    ByVal defaultScheduledExecutionTime As Date, _
    ByVal defaultSpanUntilDeadline As Long, _
    ByVal defaultEstimatedExecutionalDuration As Date, _
    ByVal defaultScheduledResttimeDuration As Date, _
    ByVal generationDateAndTime As Date, _
    ByVal defaultMilestoneId As String, _
    ByVal defaultToDoTypeId As String, _
    ByVal defaultPriorityId As String)

    m_repPrototypeId = repPrototypeId
    m_defaultToDoName = defaultToDoName
    m_defaultScheduledExecutionTime = defaultScheduledExecutionTime
    m_defaultSpanUntilDeadline = defaultSpanUntilDeadline
    m_defaultEstimatedExecutionalDuration = defaultEstimatedExecutionalDuration
    m_defaultScheduledResttimeDuration = defaultScheduledResttimeDuration
    m_generationDateAndTime = generationDateAndTime
    m_defaultMilestoneId = defaultMilestoneId
    m_defaultToDoTypeId = defaultToDoTypeId
    m_defaultPriorityId = defaultPriorityId
End Sub

