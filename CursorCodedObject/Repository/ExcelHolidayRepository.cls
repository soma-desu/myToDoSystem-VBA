VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelHolidayRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ExcelHolidayRepository
' Implements IHolidayRepository using Excel as storage.

Implements IHolidayRepository

Public Function IHolidayRepository_FindAll() As Holidays
    Dim allH As New Holidays
    ' TODO: read holiday tables and fill Holidays collection
    Set IHolidayRepository_FindAll = allH
End Function

