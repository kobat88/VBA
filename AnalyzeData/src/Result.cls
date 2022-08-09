VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Property Let ProcName(myProcName As String)
    Cells(2, 1).Value = myProcName
End Property

Public Property Let ProcResult(myProcResult As Long)
    Cells(2, 2).Value = myProcResult
End Property

Public Property Let ProcDatetime(myProcDatetime As Date)
    Cells(2, 3).Value = myProcDatetime
End Property

Public Property Let ErrDesc(myErrDesc As String)
    Cells(2, 4).Value = myErrDesc
End Property

