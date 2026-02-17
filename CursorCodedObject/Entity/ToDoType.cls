VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToDoType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ToDoType Entity

Private m_toDoTypeId As String
Private m_toDoTypeName As String
Private m_description As String

Public Property Get ToDoTypeId() As String
    ToDoTypeId = m_toDoTypeId
End Property

Public Property Get ToDoTypeName() As String
    ToDoTypeName = m_toDoTypeName
End Property

Public Property Get Description() As String
    Description = m_description
End Property

Public Sub Init(ByVal toDoTypeId As String, _
                ByVal toDoTypeName As String, _
                ByVal description As String)
    m_toDoTypeId = toDoTypeId
    m_toDoTypeName = toDoTypeName
    m_description = description
End Sub

