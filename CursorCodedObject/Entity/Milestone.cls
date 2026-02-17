VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Milestone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Milestone Entity

Private m_milestoneId As String
Private m_milestoneName As String
Private m_campaignId As String
Private m_generationDateAndTime As Date

Public Property Get MilestoneId() As String
    MilestoneId = m_milestoneId
End Property

Public Property Get MilestoneName() As String
    MilestoneName = m_milestoneName
End Property

Public Property Get CampaignId() As String
    CampaignId = m_campaignId
End Property

Public Property Get GenerationDateAndTime() As Date
    GenerationDateAndTime = m_generationDateAndTime
End Property

Public Sub Init(ByVal milestoneId As String, _
                ByVal milestoneName As String, _
                ByVal campaignId As String, _
                ByVal generationDateAndTime As Date)
    m_milestoneId = milestoneId
    m_milestoneName = milestoneName
    m_campaignId = campaignId
    m_generationDateAndTime = generationDateAndTime
End Sub

