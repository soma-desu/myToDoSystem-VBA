VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExecutedTaskRowDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ExecutedTaskRowDTO
' Output DTO for executed task aggregation per day.

Private m_id As String
Private m_name As String
Private m_dailyMinutes() As Double  ' index 1..31

Public Property Get Id() As String
    Id = m_id
End Property
Public Property Let Id(ByVal value As String)
    m_id = value
End Property

Public Property Get Name() As String
    Name = m_name
End Property
Public Property Let Name(ByVal value As String)
    m_name = value
End Property

Public Property Get DailyMinutes(ByVal dayIndex As Long) As Double
    If dayIndex < LBound(m_dailyMinutes) Or dayIndex > UBound(m_dailyMinutes) Then
        DailyMinutes = 0
    Else
        DailyMinutes = m_dailyMinutes(dayIndex)
    End If
End Property

Public Sub SetDailyMinutes(ByVal dayIndex As Long, ByVal minutes As Double)
    If dayIndex < LBound(m_dailyMinutes) Or dayIndex > UBound(m_dailyMinutes) Then Exit Sub
    m_dailyMinutes(dayIndex) = minutes
End Sub

Private Sub Class_Initialize()
    ReDim m_dailyMinutes(1 To 31)
End Sub

