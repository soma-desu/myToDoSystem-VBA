VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ValidationResult
' Immutable-like value object representing validation outcome.

Private m_isValid As Boolean
Private m_errors As Collection   ' of String

Public Property Get IsValid() As Boolean
    IsValid = m_isValid
End Property

Public Property Get Errors() As Collection
    Set Errors = m_errors
End Property

Public Sub AddError(ByVal message As String)
    If m_errors Is Nothing Then Set m_errors = New Collection
    m_errors.Add message
    m_isValid = False
End Sub

Public Function Merge(ByVal other As ValidationResult) As ValidationResult
    Dim result As New ValidationResult
    
    Dim errMsg As Variant
    
    ' Copy current errors
    If Not m_errors Is Nothing Then
        For Each errMsg In m_errors
            result.AddError CStr(errMsg)
        Next errMsg
    End If
    
    ' Copy other errors
    If Not other Is Nothing Then
        If Not other.Errors Is Nothing Then
            For Each errMsg In other.Errors
                result.AddError CStr(errMsg)
            Next errMsg
        End If
    End If
    
    If result.Errors Is Nothing Or result.Errors.Count = 0 Then
        result.m_isValid = True
    Else
        result.m_isValid = False
    End If
    
    Set Merge = result
End Function

Private Sub Class_Initialize()
    m_isValid = True
    Set m_errors = New Collection
End Sub

