Attribute VB_Name = "Module1"
Function LastSaveTime()

    Application.Volatile
    LastSaveTime = ThisWorkbook.BuiltinDocumentProperties("Last save time").Value

End Function

