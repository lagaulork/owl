Sub inputTestingDate()
    Dim inputDate As Long, objSheet_inputDate As Object
    
     ' -- キャンセル時に発生するエラートラップを無視
    On Error Resume Next
    
     ' -- フォーム設定
    inputDate = Application.InputBox( _
                                                        Title:=" 実施日フォーム", _
                                                        Prompt:=" シートを選択し、 実施日を入力して下さい", _
                                                        Default:=Format(Date, "yyyy/mm/dd"), _
                                                        Left:=50, _
                                                        Top:=50, _
                                                        Type:=1)
                                                    
    ' -- 日付を入力しOKした場合、選択されたシートの実施日セルに入力する
    If Not inputDate = 0 Then
        For Each objSheet_inputDate In ActiveWindow.SelectedSheets
            objSheet_inputDate.Cells(2, 11) = inputDate
        Next objSheet_inputDate
    Else
        Debug.Print ("inputTestingDate ：キャンセル実行")
    End If
    
    ' -- On Error Resume Next  を 解除
    On Error GoTo 0
    
End Sub

 Sub inputTestingUserName()
    Dim inputName As String, objsheet_inputName As Object
    
    ' -- キャンセル時に発生するエラートラップを無視
    On Error Resume Next
    
    inputName = Application.InputBox( _
                                        Title:=" 実施者フォーム", _
                                        Prompt:="値を入力してください！！", _
                                        Left:=50, _
                                        Top:=50, _
                                        Type:=2)
                      
    '-- 実施者を入力し、OKした場合選択されたシートの実施者セルに入力する
    If Not inputName = False Then
        If inputName = "" Then
            MsgBox ("実施者を入力して下さい")
        Else
            For Each objsheet_inputName In ActiveWindow.SelectedSheets
                objsheet_inputName.Cells(3, 11) = inputName
            Next objsheet_inputName
        End If
    Else
        Debug.Print ("inputTestingUserName ：キャンセル実行")
    End If
    
    '-- On Error Resume Next  を 解除
    On Error GoTo 0
    
 End Sub


