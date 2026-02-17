VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DayScheduleDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' DayScheduleDTO
' Output DTO for schedule calendar (one day).

Private m_dateValue As Date
Private m_timeSlots As Collection  ' of some cell info DTO

Public Property Get DateValue() As Date
    DateValue = m_dateValue
End Property
Public Property Let DateValue(ByVal value As Date)
    m_dateValue = value
End Property

Public Property Get TimeSlots() As Collection
    Set TimeSlots = m_timeSlots
End Property
Public Property Set TimeSlots(ByVal value As Collection)
    Set m_timeSlots = value
End Property

Private Sub Class_Initialize()
    Set m_timeSlots = New Collection
End Sub

