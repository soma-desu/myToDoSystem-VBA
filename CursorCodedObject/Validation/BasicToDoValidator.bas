Attribute VB_Name = "BasicToDoValidator"
Option Explicit

' BasicToDoValidator
' Composes field-level rules for BasicToDoInputDTO.

Public Function Validate(ByVal dto As BasicToDoInputDTO) As ValidationResult
    Dim result As New ValidationResult
    ' TODO: apply ValidateString/ValidateDate/ValidateTime/etc. per field and Merge results.
    Set Validate = result
End Function

