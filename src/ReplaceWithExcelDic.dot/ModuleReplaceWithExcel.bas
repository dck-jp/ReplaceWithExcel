Attribute VB_Name = "ModuleReplaceWithExcel"
Option Explicit
' <original code>
'   ワード文書の語群をExcelの辞書を使って置換 ver.04
' Author: 渡辺真
' URL: http://makoto-watanabe.main.jp/WordVba_replace.html
'
' modified By D*isuke YAMAKAWA
' last modified: 2012/8/1

Const xlUp = -4162

'Excel 辞書ファイル 関連の変数 _________________
Const colNumBefore As Integer = 1 '置換対象の語句を記載する列
Const colNumAfter As Integer = 2 '置換後の語句を記載する列
Const colNumEnablesWildcards As Integer = 3 'ワイルドカード使用の有無を記載する列
Const IsTrue As String = "有効" 'ワイルドカードの使用を示す文字列

Const rowNumBeginData As Integer = 2 '辞書データの記載が始まる行
' _______________________________________________


Sub Excelの辞書を使って置換()
    Dim filename As String
    Dim dic() As ReplaceData '辞書データを格納する配列
    Dim sw As Stopwatch
    Set sw = New Stopwatch
    
    '辞書ファイルを選択
    filename = SelectDictionary
    If filename = vbNullString Then Exit Sub
    
    sw.Start '処理時間の計測開始
        dic = LoadDictionary(filename) '配列に辞書データを読み込み
        Call SortByWordCount(dic) '辞書データを並べ替え
        Call ReplaceWithDictionary(dic) '辞書データを使って置換
    sw.Stop_ '処理時間の計測終了

    MsgBox "処理を終了しました。" & vbNewLine & "処理時間は、" _
        & Format(sw.Elapsed, "hh時間nn分ss秒") & " でした。"
End Sub

'辞書データを用いて、Wordファイルの置換を行う
'@dictionary() : 辞書データ
Private Sub ReplaceWithDictionary(dictionary() As ReplaceData)
    Application.ScreenUpdating = False
    
    Dim i
    For i = 0 To UBound(dictionary)
        '置換後の語句が、（誤入力で）空白になっていないか確認
        If dictionary(i).afterWord <> "" Then
            With Selection.Find
               .ClearFormatting
               .text = dictionary(i).beforeWord
               .Forward = True
               .Wrap = wdFindContinue
               .Format = False
               .MatchCase = False
               .MatchByte = False
               .MatchSoundsLike = False
               .MatchAllWordForms = False
               .MatchFuzzy = False
               
               If dictionary(i).EnablesWildcard Then
                  .MatchWildcards = True
               Else
                  .MatchWildcards = False
               End If
            
               .Replacement.ClearFormatting
               .Replacement.text = dictionary(i).afterWord
               .Font.Hidden = False
            End With
            
            Selection.Find.Execute Replace:=wdReplaceAll
            
            'Wordの画面が固まるのを防ぐため
            '１つの語句の置換が終わるたびにメッセージポンプを強制的に回す
            DoEvents
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub

'辞書データを語数順（降順）に並び替える
'ソートに用いたアルゴリズム: Gnome Sort
'@jisyo() : 辞書データ
Public Sub SortByWordCount(ByRef jisyo() As ReplaceData)
    Dim temp As ReplaceData
    Dim i As Integer, n As Integer
    
    i = 1
    n = UBound(jisyo) + 1
    
    Do While i < n
        If (jisyo(i - 1).wordCount < jisyo(i).wordCount) Then
            temp = jisyo(i - 1)
            jisyo(i - 1) = jisyo(i)
            jisyo(i) = temp
            i = i - 1
            
            If i = 0 Then
                i = 2
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

'辞書ファイルを配列に格納して、取得する
'@ filename : Excel 辞書ファイル名
Private Function LoadDictionary(filename As String) As ReplaceData()
    'Excelの辞書ファイルを開く
    Dim excelApp, workBook
    Set excelApp = CreateObject("Excel.Application")
    Set workBook = excelApp.Workbooks.Open(filename)
    
        '辞書ファイルにデータが記載されている最終行を取得
        Dim rowNumLastData As Integer
        rowNumLastData = excelApp.cells(workBook.ActiveSheet.Rows.count, 1).End(xlUp).Row
        
        '取得した最終行に基づいて辞書ファイルのデータ数を決定し、
        '辞書データを格納する配列を確保する
        Dim dictionaryData() As ReplaceData
        Dim masterDataCount
        masterDataCount = rowNumLastData - rowNumBeginData + 1
        ReDim dictionaryData(masterDataCount)
        
        '辞書ファイルを一行づつ配列に格納する
        Dim i, currentRowNum
        For i = 0 To masterDataCount
            currentRowNum = i + rowNumBeginData
        
            Dim repData As ReplaceData
            Dim temp
            repData.afterWord = excelApp.cells(currentRowNum, colNumAfter).Value
            repData.beforeWord = excelApp.cells(currentRowNum, colNumBefore).Value
            
            temp = excelApp.cells(currentRowNum, colNumEnablesWildcards).Value
            If temp = IsTrue Then
                repData.EnablesWildcard = True
            Else
                repData.EnablesWildcard = False
            End If
            
            repData.wordCount = Len(repData.beforeWord)
            
            If repData.afterWord <> "" Then
                dictionaryData(i) = repData
            End If
        Next i
    
    '辞書ファイルを閉じる
    workBook.Close SaveChanges:=False
    excelApp.Quit
    Set excelApp = Nothing
    
    LoadDictionary = dictionaryData
End Function

'ファイル選択ダイアログを開いて、選択されたファイル名を取得する
'@return : ファイル名（フルパス）
'          ※ファイルを選択しなかった場合vbNullStringを返す
Private Function SelectDictionary() As String
    MsgBox "Excel 辞書ファイルを選択してください。"
    
    Dim dlg As Dialog
    Dim dlgFind As Dialog
    Set dlg = Dialogs(wdDialogFileOpen)
    Set dlgFind = Dialogs(wdDialogFileFind)

    With dlg
        .Name = "*.xls"
        Select Case .Display
        Case -1 'ファイルが選択されたとき
            dlgFind.Update
            SelectDictionary = dlgFind.SearchPath & "\" & dlg.Name
        Case Else 'キャンセルボタンが押されたとき
            SelectDictionary = vbNullString
        End Select
    
    End With
End Function

