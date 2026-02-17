VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ServiceResult
' Represents outcome of a service use case.

Private m_isSuccess As Boolean
Private m_errors As Collection   ' of String

Public Property Get IsSuccess() As Boolean
    IsSuccess = m_isSuccess
End Property
Public Property Let IsSuccess(ByVal value As Boolean)
    m_isSuccess = value
End Property

Public Property Get Errors() As Collection
    Set Errors = m_errors
End Property

Public Sub AddError(ByVal message As String)
    If m_errors Is Nothing Then Set m_errors = New Collection
    m_errors.Add message
    m_isSuccess = False
End Sub

Private Sub Class_Initialize()
    m_isSuccess = True
    Set m_errors = New Collection
End Sub

