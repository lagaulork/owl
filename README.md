Sub アイテムタイプテイギショツクール()
    Dim newBookName As String
   
    newBookName = createWorkBook()
    Call setFrontAndHistorySheet(newBookName)
    Call setDesignDocSheet(newBookName)
   
End Sub
 
Private Function createWorkBook() As String
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Application.ScreenUpdating = False
   
    With wb
        ' 隠れシート表示
        .Sheets("表紙").Visible = xlSheetVisible
        .Sheets("変更履歴").Visible = xlSheetVisible
        .Sheets("設計書").Visible = xlSheetVisible
 
        ' 新規ブックを作成し、以下シートを張り付ける
        Sheets(Array("表紙", "変更履歴", "設計書")).Copy
   
        ' 再度非表示
        .Sheets("表紙").Visible = xlSheetHidden
        .Sheets("変更履歴").Visible = xlSheetHidden
        .Sheets("設計書").Visible = xlSheetHidden
    End With
   
    Application.ScreenUpdating = True
   
    ' 新規ブック名を取得
    createWorkBook = ActiveWorkbook.Name
   
End Function
 
Private Sub setFrontAndHistorySheet(newBookName As String)
    ' シート(表紙)を処理する
    Workbooks(newBookName).Worksheets(1).Activate
    ' 処理記述
    ' 作成者と日付とかを入力
 
 
    ' シート(変更履歴)を処理する
    Workbooks(newBookName).Worksheets(2).Activate
    ' 処理記述
    ' 作成者と日付とかを入力
 
End Sub
 
Private Sub setDesignDocSheet(newBookName As String)
    ' データベース接続関係変数
    Dim cn As ADODB.Connection, sql As String, rs As Recordset, itmName As String
    Set cn = New ADODB.Connection
   
    ' シート(設計書)を処理する
    Workbooks(newBookName).Worksheets(3).Activate
   
    ' SQLへの接続(クライアントサイドカーソル(3)に変更)
    cn.ConnectionString = "Provider=SQLOLEDB;" & "Data Source=DESKTOP-OR8JOJO\SQLEXPRESS,49172;" & "Initial Catalog =VBA_SampleTest;" & _
    "User ID=user1;" & "Password=Passw0rd;"
    cn.Open
    cn.CursorLocation = 3
   
'    sql = "SELECT * FROM ITEMTYPE WHERE NAME = '" & itmName & "'"
'    Set rs = cn.Execute(sql)
'    If (Not rs.EOF) Then
        ' 各テーブルのフィールド値を設定していく
        Dim tableName As String, recCnt As Long, i As Long
        Call setPropertyRec(cn, sql, rs, tableName, recCnt, i)
        Call setRelationShipTypeRec(cn, sql, rs, tableName, recCnt, i)
'    Else
'        MsgBox ("アイテムタイプの取得に失敗しました")
'        GoTo finallyCul
'    End If
   
finallyCul:
    'データベース切断とSQLで取得したレコード変数（オブジェクト）を開放
    cn.Close
    Set rs = Nothing
   
End Sub
 
Private Sub setPropertyRec(cn As ADODB.Connection, sql As String, rs As Recordset, tableName As String, recCnt As Long, j As Long)
    tableName = "プロパティ"
 
    ' SQLクエリを作成し、実行する
    sql = "SELECT * FROM HELI_IT_Doc"
    Set rs = cn.Execute(sql)
   
    If (Not rs.EOF) Then
        ' 行を挿入して罫線を引く
        recCnt = rs.RecordCount
        Call insRowAndRuledLine(recCnt, tableName)
       
        ' アクティブセルの行番号を取得
        i = ActiveCell.Row
 
        ' SQLで取得したレコードの先頭に移動
       rs.MoveFirst
       
        ' SQLで取得したレコードの最後になるまでLoop処理
        Do Until rs.EOF = True
 
            ' セルにデータを格納
            Cells(i, 2).Value = rs!Name
            Cells(i, 4).Value = rs!Label_ja
 
            ' 次のレコードへ
            rs.MoveNext
            i = i + 1
        Loop
    Else
        Debug.Print ("プロパティレコードの取得件数：0件")
    End If
 
End Sub
Private Sub setRelationShipTypeRec(cn As ADODB.Connection, sql As String, rs As Recordset, tableName As String, recCnt As Long, i As Long)
    tableName = "リレーションシップタイプ"
 
    ' SQLクエリを作成し、実行する
    sql = "SELECT * FROM HELI_RT_Doc"
    Set rs = cn.Execute(sql)
 
    If (Not rs.EOF) Then
        ' 行を挿入して罫線を引く
        recCnt = rs.RecordCount
        Call insRowAndRuledLine(recCnt, tableName)
 
        ' アクティブセルの行番号を取得
        i = ActiveCell.Row
 
        ' SQLで取得したレコードの先頭に移動
        rs.MoveFirst
       
        ' SQLで取得したレコードの最後になるまでLoop処理
        Do Until rs.EOF = True
 
            'セルにデータを格納
            Cells(i, 2).Value = rs!Name
            Cells(i, 4).Value = rs!Label_ja
 
            ' 次のレコードへ
            rs.MoveNext
            i = i + 1
        Loop
    Else
        Debug.Print ("リレーションシップタイプレコードの取得件数：0件")
    End If
End Sub
 
Private Sub insRowAndRuledLine(recCnt As Long, tableName As String)
    Dim tableName_range As Range, tableName_row As Long, startIns As Long, endIns As Long
   
    ' Object型に属するのでSetしてから、テーブル名記載の行を取得
    Set tableName_range = Cells.Find(what:=tableName, LookAt:=xlWhole)
    tableName_row = tableName_range.Row
   
    ' テーブル名が確認ができたら、行追加時の開始行と終了行を定義
    If (Not tableName_range Is Nothing) Then
        Rows(tableName_row + 3).Copy
        startIns = tableName_row + 3
        endIns = startIns + recCnt - 1
    Else
        MsgBox ("テーブル名が見つかりませんでした")
        End
    End If
   
    ' SQLデータ取得件数分、行を挿入
    Rows(startIns & ":" & endIns).Insert xlShiftToRight
   
    ' セルを範囲指定
    Range(Columns(2), Columns(15)).Rows(tableName_row + 2 & ":" & endIns - 1).Select
   
    ' 指定されているセルに罫線を引く
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    ' アクティブセルを指定
    Cells(tableName_row + 2, 2).Select
End Sub
 
 
 
