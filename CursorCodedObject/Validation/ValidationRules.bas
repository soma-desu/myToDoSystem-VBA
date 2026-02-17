Attribute VB_Name = "ValidationRules"
Option Explicit

' Generic validation helper functions.
' All functions return ValidationResult and are pure (no IO).

Public Function ValidateString(ByVal value As Variant, _
                               ByVal acceptEmpty As Boolean, _
                               ByVal minLen As Long, _
                               ByVal maxLen As Long, _
                               ByRef avoidChars() As String, _
                               ByVal formatPattern As String) As ValidationResult
    Dim result As New ValidationResult
    ' TODO: implement concrete rules
    Set ValidateString = result
End Function

Public Function ValidateIdFormat(ByVal value As Variant, _
                                 ByVal formatPattern As String, _
                                 ByVal expectedLength As Long) As ValidationResult
    Dim result As New ValidationResult
    ' TODO: implement concrete rules
    Set ValidateIdFormat = result
End Function

Public Function ValidateDate(ByVal value As Variant, _
                             ByVal acceptNull As Boolean, _
                             ByVal formatPattern As String, _
                             ByVal minDate As Date, _
                             ByVal maxDate As Date) As ValidationResult
    Dim result As New ValidationResult
    ' TODO: implement concrete rules
    Set ValidateDate = result
End Function

Public Function ValidateTime(ByVal value As Variant, _
                             ByVal acceptNull As Boolean, _
                             ByVal formatPattern As String, _
                             ByVal minTime As Date, _
                             ByVal maxTime As Date) As ValidationResult
    Dim result As New ValidationResult
    ' TODO: implement concrete rules
    Set ValidateTime = result
End Function

Public Function ValidateDateTime(ByVal value As Variant, _
                                 ByVal acceptNull As Boolean, _
                                 ByVal formatPattern As String, _
                                 ByVal minDate As Date, _
                                 ByVal maxDate As Date, _
                                 ByVal minTime As Date, _
                                 ByVal maxTime As Date) As ValidationResult
    Dim result As New ValidationResult
    ' TODO: implement concrete rules
    Set ValidateDateTime = result
End Function

Public Function ValidateNumber(ByVal value As Variant, _
                               ByVal acceptNull As Boolean, _
                               ByVal minValue As Double, _
                               ByVal maxValue As Double) As ValidationResult
    Dim result As New ValidationResult
    ' TODO: implement concrete rules
    Set ValidateNumber = result
End Function

Public Function ValidateBoolean(ByVal value As Variant, _
                                ByVal acceptNull As Boolean) As ValidationResult
    Dim result As New ValidationResult
    ' TODO: implement concrete rules
    Set ValidateBoolean = result
End Function

