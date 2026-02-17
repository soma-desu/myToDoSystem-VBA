VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Holidays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Holidays aggregate Entity
' Wraps different kinds of holidays (national, company, personal).

Private m_nationalHolidays As Collection  ' of Date
Private m_companyHolidays As Collection   ' of Date
Private m_personalHolidays As Collection  ' of Date

Public Property Get NationalHolidays() As Collection
    Set NationalHolidays = m_nationalHolidays
End Property

Public Property Get CompanyHolidays() As Collection
    Set CompanyHolidays = m_companyHolidays
End Property

Public Property Get PersonalHolidays() As Collection
    Set PersonalHolidays = m_personalHolidays
End Property

Public Sub Init( _
    ByVal nationalHolidays As Collection, _
    ByVal companyHolidays As Collection, _
    ByVal personalHolidays As Collection)

    Set m_nationalHolidays = nationalHolidays
    Set m_companyHolidays = companyHolidays
    Set m_personalHolidays = personalHolidays
End Sub

