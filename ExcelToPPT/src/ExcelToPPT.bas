Attribute VB_Name = "Main"
'Option Explicit

Public Sub Main()

    Const SRC_WS_NAME = "Sheet1"
    Const PPT_FILE_PATH = "C:\Users\kobat88\Desktop\VBA\ExcelToPPT\ExcelToPPT.pptx"
    Const LAYOUT_NAME = "質問と回答"

    Dim objPPT As Object
    Dim pptFile As Object
    Dim objOffice As Object
    Dim pptFilePath As Variant, fileFilterStr As String, titleStr As String
    Dim objFSO As Object
    Dim objPres As Object
    Dim presSlide As Object
    Dim cl As Object
    Dim clIdx As Long
    Dim objLayout As Object
    Dim objSlide As Object
    Dim shp As Object
    Dim i As Long

    Debug.Print "Main開始"

    Set objPPT = CreateObject("PowerPoint.Application")
    Set objFSO = CreateObject("scripting.FileSystemObject")

    On Error GoTo ErrHandler1

    '新規プレゼンテーションファイル作成
    'Set objPres = objPPT.Presentations.Add()

    '指定のファイルが存在しなければエラー
    If Dir(PPT_FILE_PATH) = "" Then
        Err.Raise Number:=500, Description:="指定のファイルが存在しません"
    End If

    '指定のファイルが既に開かれていたらエラー
    For Each pptFile In objPPT.Presentations
        If pptFile.FullName = PPT_FILE_PATH Then
            Err.Raise Number:=500, Description:="指定のファイルが開かれています"
        End If
    Next

    'ファイルを開くダイアログ表示
    'カレントフォルダ以外を初期表示したい場合は、ChDir フォルダ する
    fileFilterStr = "PowerPoint プレゼンテーション,*.pptx"
    titleStr = "反映先のパワーポイントファイルを選択してください"
    'pptFilePath = Application.GetOpenFilename(FileFilter = fileFilterStr, Title = titleStr)
    pptFilePath = Application.GetOpenFilename(fileFilterStr, , titleStr)
    If pptFilePath = False Then
        Err.Raise Number:=500, Description:="反映先のファイルが選択されていません"
    End If
    
    '指定のファイルをバックアップ
    Call objFSO.CopyFile(pptFilePath, pptFilePath & Format(Now(), "yyyymmdd-hhmmss") & ".backup")
    
    '指定のファイルを画面非表示で開く
    Set objPres = objPPT.Presentations.Open(pptFilePath, WithWindow:=msoFalse)

    On Error GoTo ErrHandler2

    '既存のスライドを全て削除
    For i = objPres.Slides.Count To 1 Step -1
        objPres.Slides(i).Delete
    Next

    'スライドレイアウトの選択
    clIdx = 0
    For Each cl In objPres.SlideMaster.CustomLayouts
        If cl.Name = LAYOUT_NAME Then
            clIdx = cl.Index
        End If
    Next
    If clIdx = 0 Then
        Err.Raise Number:=500, Description:="指定のレイアウトがありません"
    End If
    
    Set objLayout = objPres.SlideMaster.CustomLayouts(clIdx)

    'シェイプ番号を調べるためのデバッグ情報
    'objPres.Slides.AddSlide 1, objLayout
    'Set objSlide = objPres.Slides(1)
    'i = 1
    'For Each shp In objSlide.Shapes
    '    Debug.Print shp.Name & ",IdxNo=" & i
    '    i = i + 1
    'Next shp
    '↓
    'Title 1, IdxNo = 1
    'Text Placeholder 2,IdxNo=2
    'Text Placeholder 3,IdxNo=3

    On Error GoTo ErrHandler3
    'エクセルのデータをシェイプに書き込む
    With ThisWorkbook.Worksheets(SRC_WS_NAME)
        For i = 1 To .Cells(Rows.Count, 1).End(xlUp).Row - 1
            objPres.Slides.AddSlide(i, objLayout).Shapes(2).TextFrame.TextRange.Text = .Cells(i + 1, 1).Value
            objPres.Slides(i).Shapes(3).TextFrame.TextRange.Text = .Cells(i + 1, 2).Value
        Next i
    End With

    On Error GoTo ErrHandler2
    'ファイルを保存
    objPres.Save

    '保存が完了するまで待つ
    Do Until objPres.Saved
        Debug.Print "プレゼンテーション保存中"
    Loop

Finally:
    On Error Resume Next
    
    'ファイルを閉じる
    objPres.Close
    
    'パワーポイントを終了
    objPPT.Quit
    Set objPPT = Nothing
    Set objFSO = Nothing
       
    Debug.Print "Main終了"
        
    Exit Sub

ErrHandler1:
    Debug.Print Err.Number, Err.Description
    objPPT.Quit
    Set objPPT = Nothing
    Exit Sub

ErrHandler2:
    Debug.Print Err.Number, Err.Description
    Resume Finally
    
ErrHandler3:
    Err.Raise Number:=500, Description:="シェイプへの書込みに失敗しました"
    Debug.Print Err.Number, Err.Description
    Resume Finally
   
End Sub

