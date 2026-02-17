Attribute VB_Name = "RepeatToDoValidator"
Option Explicit

' RepeatToDoValidator
' Composes field-level rules for RepeatToDoInputDTO.

Public Function Validate(ByVal dto As RepeatToDoInputDTO) As ValidationResult
    Dim result As New ValidationResult
    ' TODO: apply ValidateString/ValidateNumber/etc. per field and Merge results.
    Set Validate = result
End Function

