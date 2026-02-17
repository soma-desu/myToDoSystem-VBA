VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Client Entity
' Immutable value holder for client master.

Private m_clientId As String
Private m_clientName As String
Private m_generationDateAndTime As Date

Public Property Get ClientId() As String
    ClientId = m_clientId
End Property

Public Property Get ClientName() As String
    ClientName = m_clientName
End Property

Public Property Get GenerationDateAndTime() As Date
    GenerationDateAndTime = m_generationDateAndTime
End Property

Public Sub Init(ByVal clientId As String, _
                ByVal clientName As String, _
                ByVal generationDateAndTime As Date)
    m_clientId = clientId
    m_clientName = clientName
    m_generationDateAndTime = generationDateAndTime
End Sub

