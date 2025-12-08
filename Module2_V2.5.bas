Attribute VB_Name = "Module2"
'V_2.5
Option Explicit

' === 設定（ここを変えれば一括調整） ===
Private Const BRACE_WEIGHT_PTS As Single = 5      ' 括弧の線の太さ（pt）
Private Const LABEL_FONT_SIZE As Single = 32      ' ラベルの文字サイズ（pt）
Private Const BRACE_WIDTH As Single = 12          ' 括弧の横幅
Private Const LABEL_X_OFFSET As Single = 12       ' 括弧からラベルまでの横距離

' ラベルの書き込み先設定
' WRITE_WHERE: 1=名前付き範囲→固定セルの順で書込、2=選択範囲の左上セルへ書込
Private Const WRITE_WHERE As Long = 1
Private Const TARGET_NAMED_RANGE As String = "LabelTarget"
Private Const TARGET_CELL_ADDR As String = "Z1"

' 区切り：初回（元の文字列と最初のラベル）と、2回目以降（ラベル同士）
' ※ 初回の区切り（FIRST）が“マーカー”として機能します
Private Const FIRST_APPEND_SEPARATOR As String = "-"    ' 例: 半角スペース
Private Const SUBSEQ_APPEND_SEPARATOR As String = ","  ' 例: 日本語読点

' 図形識別タグ（Shape.AlternativeText に付与）
Private Const BRACE_SHAPE_TAG As String = "BRACE_TAG"

' 一括削除時：シート全体を走査してセルの追記を剥がすか
Private Const SCAN_WHOLE_SHEET As Boolean = True

' 追記ブロックの最大想定長（必要なら調整）
Private Const MAX_LABEL_TAIL As Long = 256   ' ← 64 だと短いケースがあるので広げます

' === ミラー出力の設定 ===
Private Const MIRROR_ENABLED As Boolean = True          ' ← 有効化するか
Private Const MIRROR_SHEETS As String = "製造票 ③ (現場用),検査記録 (現場用)" ' ← 出力先シート名をカンマ区切りで
Private Const MIRROR_WRITE_LABELS As Boolean = False    ' ← ミラー先のセルにもラベル追記するなら True
' === 同一シート内のローカルミラー設定 ===
Private Const LOCAL_MIRROR_ENABLED As Boolean = True   ' 同じシートにも出すなら True
Private Const LOCAL_MIRROR_BASE_ROW As Long = 30       ' ミラー開始の基準行（ここから行数ぶん下へ）
Private Const LOCAL_MIRROR_WRITE_LABELS As Boolean = False ' セル追記もするなら True
' === 同一シート 固定帯ミラー（6～19 → 30～43） ===
Private Const FIXMIR_SRC_ROW_START As Long = 6
Private Const FIXMIR_SRC_ROW_END   As Long = 19
Private Const FIXMIR_DST_ROW_START As Long = 30
Private Const FIXMIR_DST_ROW_END   As Long = 43
Private Const FIXMIR_ROW_OFFSET    As Long = (FIXMIR_DST_ROW_START - FIXMIR_SRC_ROW_START)
Private Const FIXMIR_ENABLED       As Boolean = True          ' ← 有効/無効
Private Const FIXMIR_WRITE_LABELS  As Boolean = False         ' ← ミラー先にセル追記もするなら True


'========================
' 入力マクロ（任意ラベル）
'========================
Public Sub DrawLeftBrace_Label_One()
    Dim rng As Range, ws As Worksheet
    Dim labelTxt As String
    Static lastLabel As String
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    Set rng = Selection
    If rng.Areas.Count > 1 Then
        MsgBox "複数領域は非対応です。1つの連続範囲を選択してください。", vbExclamation
        Exit Sub
    End If
    
    labelTxt = InputBox("括弧に付けるラベルを入力してください（例：A／B／C／検査）", _
                        "Brace Label", IIf(Len(lastLabel) > 0, lastLabel, "A"))
    If StrPtr(labelTxt) = 0 Then Exit Sub
    labelTxt = Trim$(labelTxt)
    If Len(labelTxt) = 0 Then Exit Sub
    lastLabel = labelTxt
    
    Call DrawLeftBraceWithLabelForSelection(labelTxt, rng)
    WriteLabelToCell rng.Worksheet, rng, labelTxt
End Sub


'―― コア：選択範囲に1本の左中括弧＋ラベルを描画 ――'
Private Sub DrawLeftBraceWithLabelForSelection(ByVal labelTxt As String, ByVal rng As Range)
    Dim ws As Worksheet
    Dim leftPos As Single, topPos As Single, h As Single
    Dim shpBrace As Shape, shpLbl As Shape, grp As Shape
    Dim xLabel As Single, yCenter As Single
    
    Set ws = rng.Worksheet
    RemoveOverlappingSameLabelBraces ws, rng, labelTxt
    
    leftPos = Application.Max(2, rng.Left - (BRACE_WIDTH + LABEL_X_OFFSET))
    topPos = rng.Top
    h = rng.Height
    
    Set shpBrace = ws.Shapes.AddShape(msoShapeLeftBrace, leftPos, topPos, BRACE_WIDTH, h)
    With shpBrace
        .Name = NextUniqueShapeName(ws, "Brace_" & SafeToken(labelTxt))
        .Fill.Visible = msoFalse
        .Line.Weight = BRACE_WEIGHT_PTS
        .Line.ForeColor.RGB = RGB(0, 0, 0)           ' 黒
        .Placement = xlMoveAndSize
        .ZOrder msoBringToFront
        .AlternativeText = BRACE_SHAPE_TAG           ' 識別タグ
    End With
    
    xLabel = leftPos - LABEL_X_OFFSET
    yCenter = topPos + h / 2
    Set shpLbl = MakeLabel(ws, labelTxt, xLabel, yCenter)
    With shpLbl
        .AlternativeText = BRACE_SHAPE_TAG
    End With
    
    Set grp = ws.Shapes.Range(Array(shpBrace.Name, shpLbl.Name)).Group
    With grp
        .Name = NextUniqueShapeName(ws, "BraceGroup_" & SafeToken(labelTxt))
        .ZOrder msoBringToFront
        .AlternativeText = BRACE_SHAPE_TAG
    End With
End Sub

'―― ラベルを“特定セル”へ追記（数式保持／初回=FIRST、以降=SUBSEQを確実化）――'
Private Sub WriteLabelToCell(ByVal ws As Worksheet, ByVal rng As Range, ByVal labelTxt As String)
    Dim tgt As Range, cur As String, f As String
    Dim labEsc As String, firstChunk As String, subseqChunk As String
    Dim alreadyAppended As Boolean

    Set tgt = ResolveTargetCell(ws, rng)
    If tgt Is Nothing Then
        MsgBox "書き込み先セルが見つかりませんでした。設定を確認してください。", vbExclamation
        Exit Sub
    End If

    labEsc = Replace$(labelTxt, """", """""")
    firstChunk = Replace$(FIRST_APPEND_SEPARATOR, """", """""") & labEsc
    subseqChunk = Replace$(SUBSEQ_APPEND_SEPARATOR, """", """""") & labEsc

    If tgt.hasFormula Then
        ' --- 数式セル：")&""" が既にあれば「追記済み」扱い ---
        f = tgt.FormulaLocal
        alreadyAppended = (InStr(1, f, ")&""", vbBinaryCompare) > 0)
        If Left$(f, 1) = "=" Then f = Mid$(f, 2)   ' 式本体だけにする

        If alreadyAppended Then
            ' 2回目以降：SUBSEQで連結
            tgt.FormulaLocal = "=(" & f & ")&""" & subseqChunk & """"
        Else
            ' 初回：必ずFIRSTで連結（元の式の見た目が空でもFIRSTを付ける）
            tgt.FormulaLocal = "=(" & f & ")&""" & firstChunk & """"
        End If

    Else
        ' --- 値セル ---
        cur = CStr(tgt.Value)
        ' 「先頭がFIRSTで始まる」または「どこかにSUBSEQがある」＝追記済み
        alreadyAppended = (Left$(cur, Len(FIRST_APPEND_SEPARATOR)) = FIRST_APPEND_SEPARATOR) _
                          Or (InStr(1, cur, SUBSEQ_APPEND_SEPARATOR, vbBinaryCompare) > 0)

        If alreadyAppended Then
            ' 2回目以降：SUBSEQで追加
            tgt.Value = cur & SUBSEQ_APPEND_SEPARATOR & labelTxt
        Else
            ' 初回：必ずFIRSTで入れる（元が空でもFIRSTを付ける）
            tgt.Value = cur & FIRST_APPEND_SEPARATOR & labelTxt
        End If
    End If
End Sub

'' 式に FIRST/SUBSEQ を含むリテラル（"&"で連結された "..."）があるか
'Private Function FormulaHasAnySeparatorLiteral(ByVal f As String) As Boolean
'    Dim s As String, i As Long, p As Long, qOpen As Long, qClose As Long, lit As String
'    If Len(f) = 0 Then Exit Function
'    s = f: If Left$(s, 1) <> "=" Then s = "=" & s
'
'    i = 1
'    Do
'        p = InStr(i, s, ")&""", vbBinaryCompare)
'        If p = 0 Then Exit Do
'        qOpen = p + 3
'        qClose = InStr(qOpen + 1, s, """", vbBinaryCompare)
'        If qClose = 0 Then Exit Do
'
'        lit = Mid$(s, qOpen + 1, qClose - (qOpen + 1))
'        If InStr(1, lit, FIRST_APPEND_SEPARATOR, vbBinaryCompare) > 0 _
'        Or InStr(1, lit, SUBSEQ_APPEND_SEPARATOR, vbBinaryCompare) > 0 Then
'            FormulaHasAnySeparatorLiteral = True
'            Exit Function
'        End If
'        i = qClose + 1
'    Loop
'End Function


'========================
' ★ 一括削除（このマクロが追加した図形＋セル追記のみ）
'========================
Public Sub DeleteAllBracesAndLabels_ActiveSheet()
    Dim ws As Worksheet: Set ws = ActiveSheet
    DeleteAllBraceShapes ws           ' 図形（タグ/名前規則で判定）
    If SCAN_WHOLE_SHEET Then
        StripAppendedLabelsOnSheet ws ' セル（最初の区切りで判定）
    Else
        Dim tgt As Range
        Set tgt = TryGetNamedRange(ws, TARGET_NAMED_RANGE)
        If Not tgt Is Nothing Then StripAppendedLabelFromCell tgt
        On Error Resume Next
        Set tgt = ws.Range(TARGET_CELL_ADDR): On Error GoTo 0
        If Not tgt Is Nothing Then StripAppendedLabelFromCell tgt
        If TypeName(Selection) = "Range" Then StripAppendedLabelFromCell Selection.Cells(1, 1)
    End If
End Sub

'（任意）選択範囲のみ対象の削除
Public Sub DeleteBracesAndLabels_SelectedRange()
    Dim ws As Worksheet, ur As Range, r As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = ActiveSheet

    ' 図形：選択縦範囲と重なる & タグ/名前規則一致のみ削除
    Dim i As Long, sh As Shape, topSel As Single, bottomSel As Single
    topSel = Selection.Top: bottomSel = Selection.Top + Selection.Height
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If IsBraceShape(sh) Then
            If Not (sh.Top + sh.Height < topSel Or sh.Top > bottomSel) Then
                On Error Resume Next: sh.Delete: On Error GoTo 0
            End If
        End If
    Next i

    ' セル：選択範囲内のみ剥がす
    On Error Resume Next: Set ur = Intersect(ws.UsedRange, Selection): On Error GoTo 0
    If ur Is Nothing Then Exit Sub

    On Error Resume Next
    For Each r In ur.SpecialCells(xlCellTypeFormulas): StripAppendedLabelFromCell r: Next r
    On Error GoTo 0
    On Error Resume Next
    For Each r In ur.SpecialCells(xlCellTypeConstants): StripAppendedLabelFromCell r: Next r
    On Error GoTo 0
End Sub


'―― 図形関連 ――'
Public Sub EnablePrintingForBraceShapes_ActiveSheet()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim sh As Shape
    For Each sh In ws.Shapes
        If IsBraceShape(sh) Then
            On Error Resume Next
            sh.PrintObject = True
            On Error GoTo 0
        End If
    Next sh
End Sub

Private Sub DeleteAllBraceShapes(ByVal ws As Worksheet)
    Dim i As Long, sh As Shape
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If IsBraceShape(sh) Then
            On Error Resume Next: sh.Delete: On Error GoTo 0
        End If
    Next i
End Sub

Private Function IsBraceShape(ByVal sh As Shape) As Boolean
    On Error Resume Next
    If sh.AlternativeText = BRACE_SHAPE_TAG Then
        IsBraceShape = True
    Else
        IsBraceShape = (Left$(sh.Name, 12) = "BraceGroup_" Or _
                        Left$(sh.Name, 6) = "Brace_" Or _
                        Left$(sh.Name, 4) = "Lbl_")
    End If
    On Error GoTo 0
End Function

' 1セル分の剥がし（値セルは“末尾ウィンドウ”で最初の区切りを探す）
Private Sub StripAppendedLabelFromCell(ByVal tgt As Range)
    Dim f As String, cur As String
    If tgt Is Nothing Then Exit Sub

    If tgt.hasFormula Then
        ' --- 数式セル（従来どおり） ---
        f = tgt.FormulaLocal
        Dim inner As String
        inner = ExtractOriginalFormulaSmart(f)
        If Len(inner) > 0 Then
            On Error Resume Next
            tgt.FormulaLocal = "=" & inner
            Err.Clear: On Error GoTo 0
        End If
    Else
        ' --- 値セル：末尾ウィンドウ探索で先頭区切りを特定 ---
        cur = CStr(tgt.Value)
        If Len(cur) = 0 Then Exit Sub

        Dim startWin As Long, posFirst As Long, posSub As Long, cutPos As Long
        startWin = Application.Max(1, Len(cur) - MAX_LABEL_TAIL + 1)

        ' 末尾ウィンドウの中だけで「最初に現れる」区切りを探す
        posFirst = InStr(startWin, cur, FIRST_APPEND_SEPARATOR, vbBinaryCompare)
        posSub = InStr(startWin, cur, SUBSEQ_APPEND_SEPARATOR, vbBinaryCompare)

        If posFirst > 0 Then
            cutPos = posFirst
        ElseIf posSub > 0 Then
            cutPos = posSub
        Else
            Exit Sub  ' 末尾側に区切りが見つからない＝追記無しと判断
        End If

        If cutPos <= 1 Then
            tgt.ClearContents
        Else
            tgt.Value = Left$(cur, cutPos - 1)
        End If
    End If
End Sub

' =(<expr>)&"..." が複数段あっても、
' 最初に FIRST/SUBSEQ を含むリテラルが出た箇所で <expr> だけに戻す
Private Function ExtractOriginalFormulaSmart(ByVal f As String) As String
    Dim s As String, i As Long, p As Long, qOpen As Long, qClose As Long
    Dim lit As String

    If Len(f) = 0 Then Exit Function
    s = f: If Left$(s, 1) <> "=" Then s = "=" & s
    If Left$(s, 2) <> "=(" Then Exit Function   ' 想定形以外は返さない

    i = 1
    Do
        p = InStr(i, s, ")&""", vbBinaryCompare)
        If p = 0 Then Exit Do
        qOpen = p + 3
        qClose = InStr(qOpen + 1, s, """", vbBinaryCompare)
        If qClose = 0 Then Exit Do

        lit = Mid$(s, qOpen + 1, qClose - (qOpen + 1)) ' "…" の中身
        If (InStr(1, lit, FIRST_APPEND_SEPARATOR, vbBinaryCompare) > 0) Or _
           (InStr(1, lit, SUBSEQ_APPEND_SEPARATOR, vbBinaryCompare) > 0) Then
            ExtractOriginalFormulaSmart = Mid$(s, 3, (p - 1) - 3) ' “=(”直後～直前の “)”
            Exit Function
        End If
        i = qClose + 1
    Loop

    ' 末尾連結が無く =(<expr>) だけならそのまま返す
    If Right$(s, 1) = ")" Then
        ExtractOriginalFormulaSmart = Mid$(s, 3, Len(s) - 3 - 1)
    End If
End Function

' シート全体を走査して、当マクロ由来（区切りで判定）の追記を剥がす
Private Sub StripAppendedLabelsOnSheet(ByVal ws As Worksheet)
    Dim ur As Range, r As Range
    On Error Resume Next: Set ur = ws.UsedRange: On Error GoTo 0
    If ur Is Nothing Then Exit Sub

    On Error Resume Next
    For Each r In ur.SpecialCells(xlCellTypeFormulas)
        StripAppendedLabelFromCell r
    Next r
    On Error GoTo 0

    On Error Resume Next
    For Each r In ur.SpecialCells(xlCellTypeConstants)
        StripAppendedLabelFromCell r
    Next r
    On Error GoTo 0
End Sub


'========================
' 書き込み先セルの解決・共通ヘルパー
'========================
Private Function ResolveTargetCell(ByVal ws As Worksheet, ByVal rng As Range) As Range
    Dim tgt As Range
    Select Case WRITE_WHERE
        Case 1
            Set tgt = TryGetNamedRange(ws, TARGET_NAMED_RANGE)
            If tgt Is Nothing Then
                On Error Resume Next
                Set tgt = ws.Range(TARGET_CELL_ADDR)
                On Error GoTo 0
            End If
        Case 2
            Set tgt = rng.Cells(1, 1)
        Case Else
            Set tgt = rng.Cells(1, 1)
    End Select
    Set ResolveTargetCell = tgt
End Function

Private Function TryGetNamedRange(ws As Worksheet, ByVal nm As String) As Range
    Dim r As Range
    On Error Resume Next
    Set r = ws.Range(nm)   ' シートスコープ優先
    If r Is Nothing Then Set r = Range(nm) ' ブックスコープ
    On Error GoTo 0
    Set TryGetNamedRange = r
End Function

Private Function MakeLabel(ws As Worksheet, txt As String, x As Single, yCenter As Single) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                   Left:=x, Top:=yCenter - 8, Width:=16, Height:=16)
    With shp
        .TextFrame2.TextRange.text = txt
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = LABEL_FONT_SIZE
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .Placement = xlMoveAndSize
        .Name = NextUniqueShapeName(ws, "Lbl_" & SafeToken(txt))
    End With
    shp.Top = yCenter - shp.Height / 2
    shp.Left = x - shp.Width
    Set MakeLabel = shp
End Function

Private Sub RemoveOverlappingSameLabelBraces(ws As Worksheet, rng As Range, ByVal labelTxt As String)
    Dim i As Long, sh As Shape
    Dim tgtTop As Single, tgtBottom As Single
    Dim nm As String, token As String
    token = SafeToken(labelTxt)
    tgtTop = rng.Top
    tgtBottom = rng.Top + rng.Height
    
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        nm = sh.Name
        If Left$(nm, Len("BraceGroup_")) = "BraceGroup_" Then
            If InStr(1, nm, "_" & token, vbTextCompare) > 0 Then
                If OverlapsVertically(sh, tgtTop, tgtBottom) Then
                    On Error Resume Next: sh.Delete: On Error GoTo 0
                End If
            End If
        End If
    Next i
End Sub

Private Function OverlapsVertically(sh As Shape, tgtTop As Single, tgtBottom As Single) As Boolean
    Dim shTop As Single, shBottom As Single
    shTop = sh.Top: shBottom = sh.Top + sh.Height
    OverlapsVertically = Not (shBottom < tgtTop Or shTop > tgtBottom)
End Function

Private Function ShapeExists(ws As Worksheet, nm As String) As Boolean
    Dim t As Shape
    On Error Resume Next
    Set t = ws.Shapes(nm)
    ShapeExists = Not (t Is Nothing)
    On Error GoTo 0
End Function

Private Function NextUniqueShapeName(ws As Worksheet, base As String) As String
    Dim i As Long, nm As String
    i = 1
    Do
        nm = base & "_" & i
        If Not ShapeExists(ws, nm) Then
            NextUniqueShapeName = nm
            Exit Function
        End If
        i = i + 1
    Loop
End Function

' ラベルを名前に含めても安全な簡易トークン化
Private Function SafeToken(ByVal s As String) As String
    Dim r As String
    r = s
    r = Replace$(r, " ", "")
    r = Replace$(r, "　", "")
    r = Replace$(r, "/", "_")
    r = Replace$(r, "\", "_")
    r = Replace$(r, ":", "_")
    r = Replace$(r, "*", "_")
    r = Replace$(r, "?", "_")
    r = Replace$(r, """", "_")
    r = Replace$(r, "<", "_")
    r = Replace$(r, ">", "_")
    r = Replace$(r, "|", "_")
    SafeToken = r
End Function

'========================
' 公開マクロ：実行すると書き込み先セルに =TEXT(K4,"dd")
'========================
Public Sub Write_TEXT_K4_dd_ToTarget()
    Dim ws As Worksheet, rng As Range
    Set ws = ActiveSheet

    ' WRITE_WHERE=1 の場合は LabelTarget / Z1 を使用
    ' WRITE_WHERE=2 の場合は選択範囲の左上セルを使用
    If WRITE_WHERE = 2 Then
        If TypeName(Selection) <> "Range" Then
            MsgBox "WRITE_WHERE=2 のときはセル範囲を選択してください。", vbExclamation
            Exit Sub
        End If
        Set rng = Selection
    Else
        ' ダミーでシートのA1を渡しても ResolveTargetCell 側で NamedRange/固定セルに振り分けます
        Set rng = ws.Range("A1")
    End If

    SetTargetToTEXTdd ws, rng
End Sub

'========================
' 指定セルへ =TEXT(K4,"dd") を入力
'========================
Private Sub SetTargetToTEXTdd(ByVal ws As Worksheet, ByVal rng As Range)
    Dim tgt As Range
    Set tgt = ResolveTargetCell(ws, rng)
    If tgt Is Nothing Then
        MsgBox "書き込み先セルが見つかりませんでした。WRITE_WHERE / 名前付き範囲 / アドレスを確認してください。", vbExclamation
        Exit Sub
    End If
    tgt.FormulaLocal = "=TEXT(K4,""d"")"
End Sub

'========================
' 図形＋セルラベル一括削除 → 直後に =TEXT(K4,"dd") を入力
'========================
Public Sub DeleteAllBracesAndLabels_AndWriteTEXTdd()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ' 1) 既存の一括削除を実行
    DeleteAllBracesAndLabels_ActiveSheet
    ' 2) 書き込み先セルを解決して TEXTdd を投入
    If WRITE_WHERE = 2 Then
        ' 選択範囲左上セルをターゲットにする運用
        If TypeName(Selection) = "Range" Then
            SetTargetToTEXTdd ws, Selection
        Else
            ' 念のためのフォールバック
            SetTargetToTEXTdd ws, ws.Range("A1")
        End If
    Else
        ' 名前付き範囲／固定セル運用
        SetTargetToTEXTdd ws, ws.Range("A1")  ' ResolveTargetCell 側で NamedRange / 固定セルに振り分け
    End If
End Sub

'========================
' 選択範囲だけ削除 → 直後に =TEXT(K4,"dd") を入力
'========================
Public Sub DeleteBracesAndLabels_SelectedRange_AndWriteTEXTdd()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ' 1) 選択範囲のみの削除
    DeleteBracesAndLabels_SelectedRange
    ' 2) 書き込み先セルに TEXTdd を投入
    If WRITE_WHERE = 2 Then
        If TypeName(Selection) = "Range" Then
            SetTargetToTEXTdd ws, Selection
        Else
            SetTargetToTEXTdd ws, ws.Range("A1")
        End If
    Else
        SetTargetToTEXTdd ws, ws.Range("A1")
    End If
End Sub

' カンマ区切りをもとに、存在するワークシートだけ返す
Private Function MirrorSheetList() As Collection
    Dim col As New Collection
    Dim arr() As String, i As Long, nm As String
    Dim ws As Worksheet
    If Len(Trim$(MIRROR_SHEETS)) = 0 Then
        Set MirrorSheetList = col: Exit Function
    End If
    arr = Split(MIRROR_SHEETS, ",")
    For i = LBound(arr) To UBound(arr)
        nm = Trim$(arr(i))
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(nm)
        On Error GoTo 0
        If Not ws Is Nothing Then col.add ws
        Set ws = Nothing
    Next
    Set MirrorSheetList = col
End Function

' 元の選択範囲と同じ「列」だけを、別シートの指定行範囲で作る
' 元の選択範囲と同じ「行×列」の矩形を別シートに作る
Private Function BuildSameRectRange(ByVal ws As Worksheet, ByVal src As Range) As Range
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    r1 = src.row
    r2 = src.row + src.Rows.Count - 1
    c1 = src.Column
    c2 = src.Column + src.Columns.Count - 1
    Set BuildSameRectRange = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
End Function

' 選択範囲に括弧を描画 → 同一シート30～43行にも（同じ列）→ 他シートへもミラー
Public Sub DrawLeftBrace_Label_One_WithMirror()
    Dim rng As Range, labelTxt As String
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim list As Collection, w As Worksheet, rngT As Range
    Dim rngLocal As Range

    If TypeName(Selection) <> "Range" Or Selection.Areas.Count > 1 Then
        MsgBox "ミラー出力付き：連続したセル範囲を1つ選択して実行してください。", vbExclamation
        Exit Sub
    End If
    Set rng = Selection

    labelTxt = InputBox("括弧に付けるラベルを入力してください（例：A／B／C／検査）", "Brace Label", "A")
    If StrPtr(labelTxt) = 0 Then Exit Sub
    labelTxt = Trim$(labelTxt)
    If Len(labelTxt) = 0 Then Exit Sub

    ' 1) 元シートに作成
    DrawLeftBraceWithLabelForSelection labelTxt, rng
    WriteLabelToCell ws, rng, labelTxt
    
    ' 1.5) 同一シート 固定帯ミラー（6～19 → 30～43、選択の重なる部分だけ）
    If FIXMIR_ENABLED Then
        Dim rngFix As Range
        Set rngFix = MapBand6_19_To30_43(rng)
        If Not rngFix Is Nothing Then
            DrawLeftBraceWithLabelForSelection labelTxt, rngFix
            If FIXMIR_WRITE_LABELS Then
                WriteLabelToCell ws, rngFix, labelTxt
            End If
        End If
    End If

'    ' 2) 同じシート：基準行から“選択と同じ行数”でローカルミラー
'    If LOCAL_MIRROR_ENABLED Then
'        Set rngLocal = BuildSameColumnsRangeByRowCount(ws, rng, LOCAL_MIRROR_BASE_ROW)
'        DrawLeftBraceWithLabelForSelection labelTxt, rngLocal
'        If LOCAL_MIRROR_WRITE_LABELS Then
'            WriteLabelToCell ws, rngLocal, labelTxt
'        End If
'    End If

    ' 3) 他シートへのミラー（既存設定に従う）
    If MIRROR_ENABLED Then
        Set list = MirrorSheetList()
        For Each w In list
            Set rngT = BuildSameRectRange(w, rng) ' ←「選択と同じ行×列」の矩形ミラーを使っている場合
            ' ※ もし「列だけ同じ」にしたいなら BuildSameColumnsRange(w, rng, rng.Row, rng.Row + rng.Rows.Count - 1) に変更
            DrawLeftBraceWithLabelForSelection labelTxt, rngT
            If MIRROR_WRITE_LABELS Then
                WriteLabelToCell w, rngT, labelTxt
            End If
        Next
    End If
End Sub

Public Sub MirrorLastSelectionToOtherSheets()
    Dim rng As Range, ws As Worksheet: Set ws = ActiveSheet
    Dim list As Collection, w As Worksheet, rngT As Range
    Dim labelTxt As String
    Dim rngLocal As Range  ' ★同一シート 30～43 用

    If Not MIRROR_ENABLED Then Exit Sub
    If TypeName(Selection) <> "Range" Or Selection.Areas.Count > 1 Then
        MsgBox "連続したセル範囲を1つ選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    Set rng = Selection

    labelTxt = InputBox("ミラー先に付けるラベルを入力してください", "Brace Label (Mirror Only)", "A")
    If StrPtr(labelTxt) = 0 Then Exit Sub
    labelTxt = Trim$(labelTxt)
    If Len(labelTxt) = 0 Then Exit Sub

    ' ★同一シート 30～43 に、選択範囲と同じ「列幅」で作成
    Set rngLocal = BuildSameColumnsAtRows(ws, rng, 30, 43)
    DrawLeftBraceWithLabelForSelection labelTxt, rngLocal
    ' ミラー先のセル追記も行いたい場合は↓を有効化
    'WriteLabelToCell ws, rngLocal, labelTxt

    ' 他シートへミラー（同じ矩形）
    Set list = MirrorSheetList()
    For Each w In list
        Set rngT = BuildSameRectRange(w, rng)
        DrawLeftBraceWithLabelForSelection labelTxt, rngT
        If MIRROR_WRITE_LABELS Then
            WriteLabelToCell w, rngT, labelTxt
        End If
    Next
End Sub


'========================
' ミラー先も含めて一括削除（シート全体）
'========================
Public Sub DeleteAllBracesAndLabels_WithMirror()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim list As Collection, w As Worksheet

    ' まずアクティブシートを既存ロジックで削除
    DeleteAllBracesAndLabels_ActiveSheet

    ' ミラー先にも同等の削除を適用
    If MIRROR_ENABLED Then
        Set list = MirrorSheetList()
        For Each w In list
            ' 図形は必ず削除
            DeleteAllBraceShapes w
            ' ミラー先のセル追記も消したい場合のみ実施
            If MIRROR_WRITE_LABELS Then
                StripAppendedLabelsOnSheet w
            End If
        Next
    End If
     SetTargetToTEXTdd ws, ws.Range("A1")
End Sub

'========================
' ミラー先も含めて選択範囲のみ削除
'========================
Public Sub DeleteBracesAndLabels_SelectedRange_WithMirror()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim list As Collection, w As Worksheet, rng As Range, rngT As Range

    If TypeName(Selection) <> "Range" Or Selection.Areas.Count > 1 Then
        MsgBox "連続したセル範囲を1つ選択してから実行してください。", vbExclamation
        Exit Sub
    End If
    Set rng = Selection

    ' まずアクティブシートの選択範囲だけ削除
    DeleteBracesAndLabels_SelectedRange

    ' ミラー先：同じ行×列の矩形だけに対して削除
    If MIRROR_ENABLED Then
        Set list = MirrorSheetList()
        For Each w In list
            Set rngT = BuildSameRectRange(w, rng)
            ' 図形：矩形に交差する当マクロの括弧だけを削除
            DeleteBraceShapesInRange w, rngT
            ' セル追記：必要時のみ
            If MIRROR_WRITE_LABELS Then
                StripAppendedLabelsInRange w, rngT
            End If
        Next
    End If
End Sub

'========================
' 図形削除（矩形範囲に交差する当マクロの括弧のみ）
'========================
Private Sub DeleteBraceShapesInRange(ByVal ws As Worksheet, ByVal rng As Range)
    Dim i As Long, sh As Shape
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If IsBraceShape(sh) Then
            If ShapeIntersectsRange(sh, rng) Then
                On Error Resume Next: sh.Delete: On Error GoTo 0
            End If
        End If
    Next i
End Sub

' 図形の外接矩形と範囲の外接矩形が交差するか
Private Function ShapeIntersectsRange(ByVal sh As Shape, ByVal rng As Range) As Boolean
    Dim sL As Double, sT As Double, sR As Double, sB As Double
    Dim rL As Double, rT As Double, rR As Double, rB As Double
    sL = sh.Left: sT = sh.Top: sR = sh.Left + sh.Width: sB = sh.Top + sh.Height
    rL = rng.Left: rT = rng.Top: rR = rng.Left + rng.Width: rB = rng.Top + rng.Height
    ShapeIntersectsRange = Not (sR < rL Or sL > rR Or sB < rT Or sT > rB)
End Function

'========================
' セル追記の剥がし（選択矩形だけ）
'========================
Private Sub StripAppendedLabelsInRange(ByVal ws As Worksheet, ByVal rng As Range)
    Dim ur As Range, r As Range
    On Error Resume Next
    Set ur = Intersect(ws.UsedRange, rng)
    On Error GoTo 0
    If ur Is Nothing Then Exit Sub

    On Error Resume Next
    For Each r In ur.SpecialCells(xlCellTypeFormulas)
        StripAppendedLabelFromCell r
    Next r
    On Error GoTo 0

    On Error Resume Next
    For Each r In ur.SpecialCells(xlCellTypeConstants)
        StripAppendedLabelFromCell r
    Next r
    On Error GoTo 0
End Sub

' 同じ“列”のみ合わせ、行は任意の開始～終了で矩形を作成
Private Function BuildSameColumnsRangeByRows(ByVal ws As Worksheet, ByVal src As Range, _
                                             ByVal rowStart As Long, ByVal rowEnd As Long) As Range
    Dim c1 As Long, c2 As Long
    c1 = src.Column
    c2 = src.Column + src.Columns.Count - 1
    Set BuildSameColumnsRangeByRows = ws.Range(ws.Cells(rowStart, c1), ws.Cells(rowEnd, c2))
End Function

' 元の選択範囲と同じ「列」を、指定行に敷き直した矩形を返す
Private Function BuildSameColumnsAtRows(ByVal ws As Worksheet, ByVal src As Range, _
                                        ByVal rowStart As Long, ByVal rowEnd As Long) As Range
    Dim c1 As Long, c2 As Long
    c1 = src.Columns(1).Column
    c2 = src.Columns(src.Columns.Count).Column
    Set BuildSameColumnsAtRows = ws.Range(ws.Cells(rowStart, c1), ws.Cells(rowEnd, c2))
End Function

' 元の選択範囲と同じ「列」だけを、別の行範囲で作る
Private Function BuildSameColumnsRange(ByVal ws As Worksheet, ByVal src As Range, _
                                       ByVal rowStart As Long, ByVal rowEnd As Long) As Range
    Dim c1 As Long, c2 As Long
    c1 = src.Columns(1).Column
    c2 = src.Columns(src.Columns.Count).Column
    Set BuildSameColumnsRange = ws.Range(ws.Cells(rowStart, c1), ws.Cells(rowEnd, c2))
End Function

' 元の選択範囲と同じ「列」と「行数」で、任意の基準行から作る
Private Function BuildSameColumnsRangeByRowCount( _
    ByVal ws As Worksheet, ByVal src As Range, ByVal baseRow As Long) As Range

    Dim c1 As Long, c2 As Long, r1 As Long, r2 As Long
    c1 = src.Columns(1).Column
    c2 = src.Columns(src.Columns.Count).Column
    r1 = baseRow
    r2 = baseRow + src.Rows.Count - 1
    Set BuildSameColumnsRangeByRowCount = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
End Function
' 6～19 行に重なる「選択の一部」を取り出す
Private Function OverlapWithFixedBand(ByVal src As Range, _
                                      ByVal rowStart As Long, ByVal rowEnd As Long) As Range
    Dim ws As Worksheet: Set ws = src.Worksheet
    Dim band As Range
    Set band = ws.Range(ws.Cells(rowStart, 1), ws.Cells(rowEnd, ws.Columns.Count))
    Set OverlapWithFixedBand = Intersect(src, band)
End Function

' 6～19 に重なる部分を +24 行して 30～43 に収めた矩形を返す（列はそのまま）
Private Function MapBand6_19_To30_43(ByVal src As Range) As Range
    Dim cut As Range, ws As Worksheet
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    
    Set ws = src.Worksheet
    Set cut = OverlapWithFixedBand(src, FIXMIR_SRC_ROW_START, FIXMIR_SRC_ROW_END)
    If cut Is Nothing Then Exit Function  ' 6～19 にかかっていない
    
    ' 元の列幅と重なり行数をそのまま、行だけ +24 → 30～43 にクリップ
    c1 = cut.Columns(1).Column
    c2 = cut.Columns(cut.Columns.Count).Column
    r1 = cut.row + FIXMIR_ROW_OFFSET
    r2 = r1 + cut.Rows.Count - 1
    
    If r1 < FIXMIR_DST_ROW_START Then r1 = FIXMIR_DST_ROW_START
    If r2 > FIXMIR_DST_ROW_END Then r2 = FIXMIR_DST_ROW_END
    If r2 < r1 Then Exit Function         ' クリップで消滅
    
    Set MapBand6_19_To30_43 = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
End Function
