VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' EventDispatcher
' Entry point from Worksheet events.

Private m_toDoService As ToDoService
Private m_displayService As DisplayService
Private m_sheetPresenter As SheetPresenter

Public Property Set ToDoService(ByVal svc As ToDoService)
    Set m_toDoService = svc
End Property

Public Property Set DisplayService(ByVal svc As DisplayService)
    Set m_displayService = svc
End Property

Public Property Set SheetPresenter(ByVal presenter As SheetPresenter)
    Set m_sheetPresenter = presenter
End Property

' Main dispatch from RawEvent
Public Sub Dispatch(ByVal eventObj As RawEvent)
    ' TODO:
    ' 1. Inspect eventObj.EventTypeId / SheetName / TargetAddress
    ' 2. Read necessary input from sheet and repositories
    ' 3. Build DTOs and call ToDoService/DisplayService
    ' 4. Use SheetPresenter to write back results
End Sub

