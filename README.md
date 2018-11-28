
Sub evidence_check()
    Dim arrFlg As Boolean, rc As Integer, arrCnt As Long, arrSheetName() As String, ws As Worksheet
    arrFlg = False
    
    rc = MsgBox("エビデンスチェックを実行しますか？", vbYesNo Or vbInformation)
    
    If rc = vbYes Then
        ' -- 動的配列を初期化
        arrCnt = 0
        ReDim arrSheetName(arrCnt)
        
        For Each ws In Worksheets
            If ws.Pictures.Count = 0 Then
                ' -- データを残して要素数を変更
                ReDim Preserve arrSheetName(arrCnt)
                ' -- 配列にシート名を格納
                arrSheetName(arrCnt) = ws.Name
                ' -- シート見出しに色を付ける
                ws.Tab.ColorIndex = 44
                
                arrCnt = arrCnt + 1
                
                arrFlg = True
            End If
        Next ws
        
        ' -- ダイアログボックスにメッセージを表示
        If arrFlg Then
            MsgBox "以下のシートで異常が確認されました" & vbCrLf & Join(arrSheetName, "、"), vbExclamation
        Else
            MsgBox "正常にチェックが完了しました", vbInformation
        End If
    Else
        MsgBox "キャンセルしました", vbInformation
    End If
End Sub

