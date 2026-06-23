Option Explicit

Private Const REC_SEP As String = "|||REC|||"
Private Const FIELD_SEP As String = "|||FLD|||"

Sub RunPriceLookupViaPython()

    Dim ws As Worksheet
    Dim folderPath As String
    Dim dbPath As String
    Dim pyScriptPath As String
    Dim inputPath As String
    Dim outputPath As String
    Dim sourceDb As String
    Dim areaText As String
    Dim lastRow As Long
    Dim cmd As String
    Dim exitCode As Long
    Dim shell As Object

    Set ws = ThisWorkbook.Worksheets("入力")

    folderPath = ThisWorkbook.Path
    dbPath = folderPath & "\price_db.db"
    pyScriptPath = folderPath & "\sqlite_price_lookup.py"
    inputPath = folderPath & "\_price_lookup_input.tsv"
    outputPath = folderPath & "\_price_lookup_output.tsv"

    sourceDb = Trim(CStr(ws.Range("B2").Value))

    ' [面積特価対応] E2に面積(㎡)を入力する。空なら面積処理なし（通常価格）。
    ' D2に入力ガイドのラベルを表示しておく。
    ws.Range("D2").Value = "面積(㎡)→"
    areaText = Trim(CStr(ws.Range("E2").Value))

    ' [修正②] ハードコードリスト廃止。空チェックのみ実施し、DB名の妥当性はPython側に任せる
    If sourceDb = "" Then
        MsgBox "B2で検索対象DBを選択してください。", vbExclamation
        Exit Sub
    End If

    If Dir(dbPath) = "" Then
        MsgBox "DBファイルが見つかりません。" & vbCrLf & dbPath, vbCritical
        Exit Sub
    End If

    If Dir(pyScriptPath) = "" Then
        MsgBox "Pythonスクリプトが見つかりません。" & vbCrLf & pyScriptPath, vbCritical
        Exit Sub
    End If

    ' [修正④] 最終行の上限チェック（誤入力・書式残りで最終行が暴走するのを防ぐ）
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    If lastRow < 4 Then
        MsgBox "B4以降に製品名を入力してください。", vbExclamation
        Exit Sub
    End If

    If lastRow > 2000 Then
        MsgBox "対象行が多すぎます（" & lastRow & "行）。" & vbCrLf & _
               "B列に不要なデータや書式が残っていないか確認してください。", vbCritical
        Exit Sub
    End If

    ' [荷姿/単位対応] 列レイアウトをスクショ準拠に変更
    '   C=価格 / D=荷姿 / 単位 / E=結果 / F=品番CD / G=正式材料名
    ws.Range("C3").Value = "価格"
    ws.Range("D3").Value = "荷姿 / 単位"
    ws.Range("E3").Value = "結果"
    ws.Range("F3").Value = "品番CD"
    ws.Range("G3").Value = "正式材料名"

    If lastRow >= 4 Then
        ws.Range("C4:G" & lastRow).ClearContents
    End If

    WriteInputTsv inputPath, ws, lastRow

    cmd = "py " & QuotePath(pyScriptPath) & _
          " --db " & QuotePath(dbPath) & _
          " --source " & QuotePath(sourceDb) & _
          " --input " & QuotePath(inputPath) & _
          " --output " & QuotePath(outputPath)

    ' [面積特価対応] 面積が入力されていれば --area で渡す（空なら渡さない）
    If areaText <> "" Then
        cmd = cmd & " --area " & QuotePath(areaText)
    End If

    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run(cmd, 0, True)

    If exitCode <> 0 Then
        MsgBox "Python処理でエラーが発生しました。" & vbCrLf & _
               "コマンドプロンプトで次のコマンドを実行して内容を確認してください。" & vbCrLf & _
               cmd, vbCritical
        Exit Sub
    End If

    If Dir(outputPath) = "" Then
        MsgBox "出力TSVが作成されませんでした。" & vbCrLf & outputPath, vbCritical
        Exit Sub
    End If

    ReadOutputTsv outputPath, ws

    MsgBox "価格取得が完了しました。", vbInformation

End Sub

' [修正③] BOMなしUTF-8で書き出す
' ADODB.Stream は Charset="utf-8" + WriteText でBOM付きになる既知の挙動があるため、
' テキストストリームで書いた後、バイナリモードでBOM3バイトをスキップして保存する
Private Sub WriteInputTsv(ByVal inputPath As String, ByVal ws As Worksheet, ByVal lastRow As Long)

    Dim txtStream As Object
    Dim binStream As Object
    Dim r As Long
    Dim lineText As String
    Dim productName As String

    Set txtStream = CreateObject("ADODB.Stream")
    txtStream.Type = 2
    txtStream.Charset = "utf-8"
    txtStream.Open

    txtStream.WriteText "row_no" & vbTab & "product_name" & vbCrLf

    For r = 4 To lastRow
        productName = CStr(ws.Cells(r, "B").Value)
        lineText = CStr(r) & vbTab & CleanTsvValue(productName) & vbCrLf
        txtStream.WriteText lineText
    Next r

    ' BOMをスキップしてバイナリとして保存
    txtStream.Position = 0
    txtStream.Type = 1  ' バイナリモードに切り替え
    txtStream.Position = 3  ' UTF-8 BOM は3バイト

    Set binStream = CreateObject("ADODB.Stream")
    binStream.Type = 1
    binStream.Open
    txtStream.CopyTo binStream
    binStream.SaveToFile inputPath, 2
    binStream.Close
    txtStream.Close

End Sub

' [荷姿/単位対応] 出力TSVの列構成が1列増えた
'   0:row_no 1:product_name 2:price 3:memo 4:product_cd 5:official_name 6:unit 7:candidates
Private Sub ReadOutputTsv(ByVal outputPath As String, ByVal ws As Worksheet)

    Dim stream As Object
    Dim text As String
    Dim lines() As String
    Dim cols() As String
    Dim i As Long
    Dim rowNo As Long
    Dim productName As String
    Dim unitVal As String
    Dim candidates As String

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile outputPath

    text = stream.ReadText
    stream.Close

    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)

    lines = Split(text, vbLf)

    For i = 1 To UBound(lines)
        If Trim(lines(i)) <> "" Then
            cols = Split(lines(i), vbTab)

            If UBound(cols) >= 5 Then
                rowNo = CLng(cols(0))
                productName = cols(1)

                unitVal = ""
                candidates = ""
                If UBound(cols) >= 6 Then unitVal = cols(6)
                If UBound(cols) >= 7 Then candidates = cols(7)

                If cols(2) = "候補選択" Then
                    HandleCandidatePopup ws, rowNo, productName, cols(3), candidates
                Else
                    ws.Cells(rowNo, "C").Value = cols(2)    ' 価格
                    ws.Cells(rowNo, "D").Value = unitVal    ' 荷姿 / 単位
                    ws.Cells(rowNo, "E").Value = cols(3)    ' 結果
                    ws.Cells(rowNo, "F").Value = cols(4)    ' 品番CD
                    ws.Cells(rowNo, "G").Value = cols(5)    ' 正式材料名
                End If
            End If
        End If
    Next i

End Sub

' [荷姿/単位対応] 候補レコードのフィールド構成（末尾に unit を追加）
'   0:no 1:product_cd 2:official_name 3:match_name 4:used_source 5:price 6:score 7:note 8:unit
Private Sub HandleCandidatePopup(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal productName As String, ByVal memo As String, ByVal candidates As String)

    Dim records() As String
    Dim fields() As String
    Dim listText As String
    Dim i As Long
    Dim answer As Variant
    Dim selectedValue As String
    Dim found As Boolean
    Dim validList As String
    Dim unitSel As String

    If Trim(candidates) = "" Then
        ws.Cells(rowNo, "C").Value = "候補なし"
        ws.Cells(rowNo, "E").Value = memo
        Exit Sub
    End If

    records = Split(candidates, REC_SEP)

    listText = BuildCandidateMessage(productName, memo, records, validList)

    If validList = "" Then
        ws.Cells(rowNo, "C").Value = "候補解析失敗"
        ws.Cells(rowNo, "E").Value = "候補データを読めません"
        Exit Sub
    End If

RetryInput:
    AppActivate Application.Caption
    answer = Application.InputBox( _
        Prompt:=listText, _
        Title:="候補選択", _
        Type:=2 _
    )

    If VarType(answer) = vbBoolean Then
        If answer = False Then
            ws.Cells(rowNo, "C").Value = "選択キャンセル"
            ws.Cells(rowNo, "E").Value = "候補あり"
            Exit Sub
        End If
    End If

    selectedValue = NormalizeChoice(CStr(answer))

    If selectedValue = "" Then
        MsgBox "候補番号、または品番CDを入力してください。", vbExclamation
        GoTo RetryInput
    End If

    found = False

    For i = LBound(records) To UBound(records)
        If Trim(records(i)) <> "" Then
            fields = Split(records(i), FIELD_SEP)

            If UBound(fields) >= 6 Then
                If IsCandidateMatch(selectedValue, fields(0), fields(1)) Then
                    unitSel = ""
                    If UBound(fields) >= 8 Then unitSel = fields(8)

                    ws.Cells(rowNo, "C").Value = fields(5)  ' 価格
                    ws.Cells(rowNo, "D").Value = unitSel    ' 荷姿 / 単位
                    ws.Cells(rowNo, "E").Value = "選択 / 品番CD:" & fields(1) & " / 参照DB:" & fields(4) & " / 類似度:" & fields(6) & "%"  ' 結果
                    ws.Cells(rowNo, "F").Value = fields(1)  ' 品番CD
                    ws.Cells(rowNo, "G").Value = fields(2)  ' 正式材料名
                    found = True
                    Exit For
                End If
            End If
        End If
    Next i

    If Not found Then
        MsgBox "入力値が候補にありません。" & vbCrLf & _
               "候補番号、または表示されている品番CDを入力してください。" & vbCrLf & _
               "有効候補: " & validList, vbExclamation
        GoTo RetryInput
    End If

End Sub

Private Function BuildCandidateMessage(ByVal productName As String, ByVal memo As String, ByRef records() As String, ByRef validList As String) As String

    Dim fields() As String
    Dim i As Long
    Dim msg As String
    Dim noteStr As String
    Dim unitStr As String

    msg = "[" & productName & "] " & memo & vbCrLf & vbCrLf

    validList = ""

    For i = LBound(records) To UBound(records)
        If Trim(records(i)) <> "" Then
            fields = Split(records(i), FIELD_SEP)

            If UBound(fields) >= 6 Then
                validList = validList & IIf(validList = "", "", "/") & fields(0)

                noteStr = ""
                If UBound(fields) >= 7 Then
                    If Trim(fields(7)) <> "" Then noteStr = " (" & fields(7) & ")"
                End If

                unitStr = ""
                If UBound(fields) >= 8 Then
                    If Trim(fields(8)) <> "" Then unitStr = " [" & fields(8) & "]"
                End If

                msg = msg & fields(0) & ": " & ShortText(fields(2), 16) & noteStr & " " & fields(5) & "円" & unitStr & vbCrLf
            End If
        End If
    Next i

    msg = msg & vbCrLf & "番号入力 (キャンセル=スキップ):"

    BuildCandidateMessage = msg

End Function

Private Function IsCandidateMatch(ByVal selectedValue As String, ByVal candidateNo As String, ByVal productCd As String) As Boolean

    Dim candidateNoNorm As String
    Dim productCdNorm As String

    candidateNoNorm = NormalizeChoice(candidateNo)
    productCdNorm = NormalizeChoice(productCd)

    If selectedValue = candidateNoNorm Then
        IsCandidateMatch = True
        Exit Function
    End If

    If selectedValue = productCdNorm Then
        IsCandidateMatch = True
        Exit Function
    End If

    IsCandidateMatch = False

End Function

' [修正①] "CD" / "ＣＤ" の無条件削除を廃止
' 品番CDに "CD" が含まれる製品（例: CD-001）を入力するとマッチングが壊れるため。
' "品番CD:" / "品番ＣＤ:" のような入力プレフィックスは上流で除去済みのため問題なし。
Private Function NormalizeChoice(ByVal s As String) As String

    s = Trim(CStr(s))

    On Error Resume Next
    s = StrConv(s, vbNarrow)
    On Error GoTo 0

    s = Replace(s, "　", "")
    s = Replace(s, " ", "")
    s = Replace(s, "候補", "")
    s = Replace(s, "番号", "")
    s = Replace(s, "品番CD", "")
    s = Replace(s, "品番ＣＤ", "")
    s = Replace(s, ":", "")
    s = Replace(s, "：", "")

    NormalizeChoice = s

End Function

Private Function ShortText(ByVal s As String, ByVal maxLen As Long) As String
    If Len(s) <= maxLen Then
        ShortText = s
    Else
        ShortText = Left(s, maxLen) & "..."
    End If
End Function

Private Function CleanTsvValue(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    CleanTsvValue = s
End Function

Private Function QuotePath(ByVal s As String) As String
    QuotePath = """" & s & """"
End Function
