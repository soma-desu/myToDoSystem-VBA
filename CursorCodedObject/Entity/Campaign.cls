VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Campaign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Campaign Entity
' Immutable value holder for campaign master.

Private m_campaignId As String
Private m_campaignName As String
Private m_clientId As String
Private m_generationDateAndTime As Date

Public Property Get CampaignId() As String
    CampaignId = m_campaignId
End Property

Public Property Get CampaignName() As String
    CampaignName = m_campaignName
End Property

Public Property Get ClientId() As String
    ClientId = m_clientId
End Property

Public Property Get GenerationDateAndTime() As Date
    GenerationDateAndTime = m_generationDateAndTime
End Property

Public Sub Init(ByVal campaignId As String, _
                ByVal campaignName As String, _
                ByVal clientId As String, _
                ByVal generationDateAndTime As Date)
    m_campaignId = campaignId
    m_campaignName = campaignName
    m_clientId = clientId
    m_generationDateAndTime = generationDateAndTime
End Sub

