VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'V_2.5
Option Explicit

'=== 基本設定 ===
Private Const ROW_START As Long = 6
Private Const ROW_END   As Long = 19

Private Const LEFT_COL As Long = 4            ' D列（2行分割出力: 上2桁/下3桁）
Private Const RIGHT_COL_START As Long = 7     ' G列（右側結合ブロック先頭）
Private Const RIGHT_COL_END   As Long = 10    ' J列（右側結合ブロック末尾）

'=== 装飾設定（未ヒット時の既定装飾など） ===
Private Const DECOR_CLEAR_BELOW_FIRST As Boolean = False
Private Const DECOR_NOHIT_BEFORE_TEXT As String = ""
Private Const DECOR_NOHIT_AFTER_TEXT  As String = "t"   ' 動作確認用。不要なら空に

'=== 未ヒット時 左列フォールバック ===
Private Const UNMATCHED_LEFT_UPPER As String = "62"
Private Const UNMATCHED_LEFT_LOWER As String = "018"

'=== レイアウト可変対応（名前付き範囲優先 / 固定番地フォールバック） ===
Private Const NR_LEFT_INPUT As String = "AutoInputLeft"
Private Const NR_RIGHT_OUTPUT As String = "AutoInputRight"

Private Function ResolveSheetNamedRange(ByVal nm As String) As Range
    On Error Resume Next
    Set ResolveSheetNamedRange = Me.Range(nm)
    If ResolveSheetNamedRange Is Nothing Then Set ResolveSheetNamedRange = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
End Function

Private Function LeftInputRange() As Range
    Dim r As Range
    Set r = ResolveSheetNamedRange(NR_LEFT_INPUT)
    If r Is Nothing Then Set r = Me.Range(Me.Cells(ROW_START, LEFT_COL), Me.Cells(ROW_END, LEFT_COL))
    Set LeftInputRange = r
End Function

Private Function RightOutputRange() As Range
    Dim r As Range
    Set r = ResolveSheetNamedRange(NR_RIGHT_OUTPUT)
    If r Is Nothing Then Set r = Me.Range(Me.Cells(ROW_START, RIGHT_COL_START), Me.Cells(ROW_END, RIGHT_COL_END))
    Set RightOutputRange = r
End Function

Private Function LayoutRowStart() As Long
    Dim l As Range, r As Range
    Set l = LeftInputRange()
    Set r = RightOutputRange()
    LayoutRowStart = Application.Min(l.row, r.row)
End Function

Private Function LayoutRowEnd() As Long
    Dim l As Range, r As Range
    Set l = LeftInputRange()
    Set r = RightOutputRange()
    LayoutRowEnd = Application.Max(l.row + l.Rows.Count - 1, r.row + r.Rows.Count - 1)
End Function

Private Function LayoutLeftCol() As Long
    LayoutLeftCol = LeftInputRange().Column
End Function

Private Function LayoutRightColStart() As Long
    LayoutRightColStart = RightOutputRange().Column
End Function

Private Function LayoutRightColEnd() As Long
    LayoutRightColEnd = RightOutputRange().Column + RightOutputRange().Columns.Count - 1
End Function

Private Sub ComboBox1_Change()

End Sub

'======================== Worksheet_Change ========================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanUp

    Debug.Print "[CHG] fired: ", Now, " TargetAreas=", Target.Areas.Count, " EnableEvents=", Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    '=== RefCache から参照マップ構築（別ブックは開かない）===
    Dim map34 As Object, map43 As Object
    Set map34 = CreateObject("Scripting.Dictionary")
    Set map43 = CreateObject("Scripting.Dictionary")

    If Not BuildMapsFromCache(map34, map43) Then
        Static warned As Boolean
        If Not warned Then
            warned = True
            MsgBox "参照キャッシュ(RefCache)がありません。" & vbCrLf & _
                   "メニューから『UpdateRefCacheFromRefFile』を実行して更新してください。", vbExclamation
        End If
        GoTo CleanUp
    End If
    RefDebug_Stats "Cache", map34, map43

    '=== 変更セルをルーティング ===
    Dim ar As Range, c As Range
    Dim handled As Boolean: handled = False

    Debug.Print "[Route] Areas=", Target.Areas.Count
    For Each ar In Target.Areas
        For Each c In ar.Cells
            If RouteOneCell(c, map34, map43) Then handled = True
        Next c
    Next ar

    If Not handled Then
        Debug.Print "[INFO] no handled cells in change."
    End If
    
    '=== 5～6列の結合セルで @ → 〃 変換 ===
    Dim tl As Range
    For Each ar In Target.Areas
    For Each c In ar.Cells
        ' 行範囲外はスキップ
        If c.row >= LayoutRowStart() And c.row <= LayoutRowEnd() Then
            ' 結合セルは先頭セルだけ処理
            If c.MergeCells Then
                Set tl = c.MergeArea.Cells(1, 1)
                If c.Address <> tl.Address Then GoTo NextC
            Else
                Set tl = c
            End If

            ' 列5～6だけ処理
            If tl.Column >= 5 And tl.Column <= 6 Then
                Dim beforeVal As String, afterVal As String
                beforeVal = CStr(tl.Value)
                ' 半角/全角の @ を 〃 に置換（全置換）
                afterVal = Replace(beforeVal, "@", "〃")
                afterVal = Replace(afterVal, "＠", "〃")
                If afterVal <> beforeVal Then
                    tl.Value = afterVal
                    ' 他の処理と競合しないが、必要なら handled=True などフラグ更新
                End If
            End If
        End If
NextC:
    Next c
Next ar


CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

'======================== 1セルルーティング ========================
Private Function RouteOneCell(ByVal c As Range, ByVal map34 As Object, ByVal map43 As Object) As Boolean
    Dim tl As Range, tlRow As Long, tlCol As Long

    If c.MergeCells Then
        Set tl = c.MergeArea.Cells(1, 1)
    Else
        Set tl = c
    End If
    tlRow = tl.row: tlCol = tl.Column
    Debug.Print "[Cell] ", tl.Address, " row=", tlRow, " col=", tlCol

    If tlRow < LayoutRowStart() Or tlRow > LayoutRowEnd() Then
        Debug.Print "  [SKIP] row out"
        RouteOneCell = False: Exit Function
    End If

    '==== 左（列4）→右 ====
    If tlCol = LayoutLeftCol() Then
        Dim topRowL As Long: topRowL = PairTopRow(tlRow)
        If tlRow <> topRowL Then
            Debug.Print "  [SKIP L] lower row edited (pairTop=" & topRowL & ")"
            RouteOneCell = False: Exit Function
        End If

        Dim inValL As String, outVal As String, matchedKey As String
        Dim score As Long
        inValL = NormalizeForMatch(CStr(Me.Cells(topRowL, LayoutLeftCol()).Value) & CStr(Me.Cells(topRowL + 1, LayoutLeftCol()).Value))
        Debug.Print "  [L->R] inNorm=[", inValL, "]"

        If Len(inValL) = 0 Then
            WriteRight topRowL, ""
            ApplyDecorations topRowL, ""
            Debug.Print "  [L->R] EMPTY -> clear right & decor"
        Else
            Resolve_OR_WithKey inValL, map34, map43, outVal, matchedKey, score
            Debug.Print "  [L->R] hitKey=[", matchedKey, "] score=", score, " out=[", outVal, "]"

            If score = 0 Then
                WriteRight topRowL, ""
                ApplyDecorations topRowL, ""
            Else
                WriteRight topRowL, outVal
                ApplyDecorations topRowL, matchedKey
            End If
        End If

        RouteOneCell = True: Exit Function
    End If

    '==== 右（列7～10）→左 ====
    If tlCol >= LayoutRightColStart() And tlCol <= LayoutRightColEnd() Then
        Dim topRowR As Long: topRowR = PairTopRow(tlRow)
        If tlRow <> topRowR Then
            Debug.Print "  [SKIP R] lower row edited (pairTop=" & topRowR & ")"
            RouteOneCell = False: Exit Function
        End If

        Dim rightTop As Range
        Set rightTop = GetRowMergeTopLeft(Me, topRowR, LayoutRightColStart(), LayoutRightColEnd())

        Dim inValR As String, outVal2 As String, matchedKey2 As String
        Dim score2 As Long
        inValR = NormalizeForMatch(CStr(rightTop.Value))
        Debug.Print "  [R->L] inNorm=[", inValR, "] (read @", rightTop.Address, ")"

        If Len(inValR) = 0 Then
            WriteLeftSplit topRowR, ""
            ApplyDecorations topRowR, ""
            Debug.Print "  [R->L] EMPTY -> clear left & decor"
        Else
            Resolve_OR_WithKey inValR, map34, map43, outVal2, matchedKey2, score2
            Debug.Print "  [R->L] hitKey=[", matchedKey2, "] score=", score2, " out=[", outVal2, "]"

            If score2 = 0 Then
                WriteLeftFallback topRowR      ' 左は「未」「登録」
                ApplyDecorations topRowR, ""
            Else
                WriteLeftSplit topRowR, outVal2
                ApplyDecorations topRowR, matchedKey2
            End If
        End If

        RouteOneCell = True: Exit Function
    End If

    Debug.Print "  [SKIP] col out"
    RouteOneCell = False
End Function

'======================== BuildMapsFromCache ========================
Private Function GetOrCreateCol(ByVal dict As Object, ByVal k As String) As Collection
    If dict.Exists(k) Then
        Set GetOrCreateCol = dict(k)
    Else
        Dim c As New Collection
        dict.add k, c
        Set GetOrCreateCol = c
    End If
End Function

Private Function BuildMapsFromCache(ByRef map34 As Object, ByRef map43 As Object) As Boolean
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("RefCache")
    On Error GoTo 0
    If ws Is Nothing Then BuildMapsFromCache = False: Exit Function

    Dim last As Long
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If last < 2 Then BuildMapsFromCache = False: Exit Function

    Dim r As Long, kRaw As String, vRaw As String, tRaw As String
    Dim kNorm As String, vNorm As String
    For r = 2 To last
        kRaw = CStr(ws.Cells(r, 1).Value) ' C_col
        vRaw = CStr(ws.Cells(r, 2).Value) ' D_col
        tRaw = CStr(ws.Cells(r, 3).Value) ' Tag
        kNorm = NormalizeForMatch(kRaw)
        vNorm = NormalizeForMatch(vRaw)

        If Len(kNorm) > 0 Then GetOrCreateCol(map34, kNorm).add Array(vRaw, kRaw, tRaw, r)
        If Len(vNorm) > 0 Then GetOrCreateCol(map43, vNorm).add Array(kRaw, vRaw, tRaw, r)
    Next r
    BuildMapsFromCache = True
End Function

'======================== 検索スコア（ゆるめ検索） ========================
' 自動入力（検索）専用：短語OK・記号/空白の揺れ吸収
Private Function ContainsLoose(ByVal textNorm As String, ByVal patternNorm As String) As Boolean
    Dim t As String, p As String
    t = textNorm: p = patternNorm
    If Len(p) = 0 Or Len(t) = 0 Then Exit Function
    If InStr(1, t, p, vbTextCompare) > 0 Then ContainsLoose = True: Exit Function

    Dim t2 As String, p2 As String
    t2 = Replace(Replace(Replace(Replace(t, "-", ""), "_", ""), " ", ""), "/", "")
    p2 = Replace(Replace(Replace(Replace(p, "-", ""), "_", ""), " ", ""), "/", "")
    If InStr(1, t2, p2, vbTextCompare) > 0 Then ContainsLoose = True
End Function

Private Sub Resolve_OR_WithKey(ByVal inNorm As String, ByVal map34 As Object, ByVal map43 As Object, _
                               ByRef outValue As String, ByRef outKey As String, ByRef outScore As Long)
    Dim v34 As String, k34 As String, s34 As Long
    Dim v43 As String, k43 As String, s43 As Long

    Best_OR inNorm, map34, True, v34, k34, s34
    Best_OR inNorm, map43, False, v43, k43, s43

    If s34 = 0 And s43 = 0 Then
        outValue = "": outKey = "": outScore = 0
    ElseIf s34 >= s43 Then
        outValue = v34: outKey = k34: outScore = s34
    Else
        outValue = v43: outKey = k43: outScore = s43
    End If
End Sub

Private Sub Best_OR(ByVal inNorm As String, ByVal dict As Object, ByVal is34 As Boolean, _
                    ByRef outValue As String, ByRef outKey As String, ByRef outScore As Long)
    Dim k As Variant, keyStr As String, tokens As Variant
    Dim i As Long, token As String, matchCount As Long, matchLenSum As Long
    Dim col As Collection, entry As Variant
    Dim bestRow As Long, score As Long

    outValue = "": outKey = "": outScore = 0: bestRow = 0

    For Each k In dict.keys
        keyStr = CStr(k)
        If Len(keyStr) = 0 Then GoTo NextKey

        If InStr(keyStr, " ") > 0 Then
            tokens = Split(keyStr, " ")
        Else
            tokens = Array(keyStr)
        End If

        matchCount = 0: matchLenSum = 0
        If ArrCount(tokens) > 0 Then
            For i = SafeLBound(tokens) To SafeUBound(tokens)
                token = Trim$(CStr(tokens(i)))
                If Len(token) > 0 And ContainsLoose(inNorm, token) Then
                    matchCount = matchCount + 1
                    matchLenSum = matchLenSum + Len(token)
                End If
            Next i
        End If

        If matchCount = 0 Then GoTo NextKey

        Set col = dict(k)
        For Each entry In col
            Dim outRaw As String, keyRaw As String, tagRaw As String
            Dim rowNum As Long
            Dim m2c As Long, m2l As Long, mTc As Long, mTl As Long
            Dim tks2 As Variant, t2 As String, tksT As Variant, tT As String

            outRaw = CStr(entry(0))
            keyRaw = CStr(entry(1))
            tagRaw = CStr(entry(2))
            rowNum = CLng(entry(3))

            m2c = 0: m2l = 0
            tks2 = Split(NormalizeForMatch(outRaw), " ")
            If ArrCount(tks2) > 0 Then
                For i = SafeLBound(tks2) To SafeUBound(tks2)
                    t2 = Trim$(CStr(tks2(i)))
                    If Len(t2) > 0 And ContainsLoose(inNorm, t2) Then
                        m2c = m2c + 1
                        m2l = m2l + Len(t2)
                    End If
                Next i
            End If

            mTc = 0: mTl = 0
            If Len(tagRaw) > 0 Then
                tksT = Split(NormalizeForMatch(tagRaw), " ")
                If ArrCount(tksT) > 0 Then
                    For i = SafeLBound(tksT) To SafeUBound(tksT)
                        tT = Trim$(CStr(tksT(i)))
                        If Len(tT) > 0 And ContainsLoose(inNorm, tT) Then
                            mTc = mTc + 1
                            mTl = mTl + Len(tT)
                        End If
                    Next i
                End If
            End If

            score = (matchCount * 100000) + matchLenSum _
                    + (m2c * 50000) + m2l _
                    + (mTc * 200000) + (mTl * 2)

            If score > outScore Or (score = outScore And outScore > 0 And rowNum < bestRow) Then
                outScore = score: outValue = outRaw: outKey = keyRaw: bestRow = rowNum
            End If
        Next entry
NextKey:
    Next k
End Sub

'======================== 装飾：ユーティリティ ========================
Private Sub AppendUnique(ByRef buf As String, ByVal frag As String, ByVal baseText As String)
    Dim f As String: f = CStr(frag)
    If Len(f) = 0 Then Exit Sub
    If InStr(baseText, f) > 0 Then Exit Sub
    If InStr(buf, f) > 0 Then Exit Sub
    buf = buf & f
End Sub

Private Function EdgeStrip(ByVal s As String, ByVal frag As String, ByVal where As String) As String
    Dim t As String: t = CStr(s)
    Dim f As String: f = CStr(frag)
    If Len(f) = 0 Then EdgeStrip = t: Exit Function

    If LCase$(where) = "before" Or LCase$(where) = "both" Then
        Do While Len(t) >= Len(f) And Left$(t, Len(f)) = f
            t = Mid$(t, Len(f) + 1)
        Loop
    End If
    If LCase$(where) = "after" Or LCase$(where) = "both" Then
        Do While Len(t) >= Len(f) And Right$(t, Len(f)) = f
            t = Left$(t, Len(t) - Len(f))
        Loop
    End If
    EdgeStrip = t
End Function

Private Function RegexTest(ByVal text As String, ByVal pattern As String) As Boolean
    On Error Resume Next
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.pattern = pattern: rx.IgnoreCase = True: rx.Global = False
    RegexTest = rx.Test(text)
    On Error GoTo 0
End Function

Private Function HasWordBoundary(ByVal textNorm As String, ByVal pat As String) As Boolean
    Dim pos As Long: pos = InStr(1, textNorm, pat, vbTextCompare)
    If pos = 0 Then Exit Function
    Dim beforeCh As String, afterCh As String
    If pos > 1 Then beforeCh = Mid$(textNorm, pos - 1, 1) Else beforeCh = " "
    If pos + Len(pat) <= Len(textNorm) Then afterCh = Mid$(textNorm, pos + Len(pat), 1) Else afterCh = " "
    HasWordBoundary = (IsNonWord(beforeCh) And IsNonWord(afterCh))
End Function
Private Function IsNonWord(ByVal ch As String) As Boolean
    IsNonWord = Not (ch Like "[A-Za-z0-9]")
End Function

' 拡張ゆるめ一致（Options対応）
' Options: "nofuzzy;word;min=4" 等
Private Function FuzzyContainsEx(ByVal textNorm As String, ByVal patternNorm As String, ByVal options As String) As Boolean
    Dim opt As String: opt = LCase$(options)
    Dim allowFuzzy As Boolean: allowFuzzy = (InStr(opt, "nofuzzy") = 0)
    Dim needWord As Boolean:   needWord = (InStr(opt, "word") > 0)

    Dim minLen As Long: minLen = 3
    Dim p As Long: p = InStr(opt, "min=")
    If p > 0 Then
        On Error Resume Next
        minLen = CLng(Mid$(opt, p + 4))
        On Error GoTo 0
        If minLen < 0 Then minLen = 0
    End If

    Dim t As String: t = textNorm
    Dim pat As String: pat = patternNorm
    If Len(pat) = 0 Or Len(t) = 0 Then Exit Function

    If Len(pat) < minLen Then Exit Function

    If InStr(1, t, pat, vbTextCompare) > 0 Then
        FuzzyContainsEx = IIf(needWord, HasWordBoundary(t, pat), True)
        Exit Function
    End If

    Dim t2 As String, p2 As String
    t2 = Replace(Replace(Replace(Replace(t, "-", ""), "_", ""), " ", ""), "/", "")
    p2 = Replace(Replace(Replace(Replace(pat, "-", ""), "_", ""), " ", ""), "/", "")
    If Len(t2) >= Len(p2) And InStr(1, t2, p2, vbTextCompare) > 0 Then
        FuzzyContainsEx = IIf(needWord, HasWordBoundary(t, pat), True)
        Exit Function
    End If

    If allowFuzzy And Len(pat) >= 4 And Len(pat) <= 15 And Len(t) >= Len(pat) Then
        Dim i As Long, j As Long, subStr As String, diff As Long
        For i = 1 To Len(t) - Len(pat) + 1
            subStr = Mid$(t, i, Len(pat)): diff = 0
            For j = 1 To Len(pat)
                If Mid$(subStr, j, 1) <> Mid$(pat, j, 1) Then
                    diff = diff + 1: If diff > 1 Then Exit For
                End If
            Next j
            If diff <= 1 Then
                FuzzyContainsEx = IIf(needWord, HasWordBoundary(t, pat), True)
                Exit Function
            End If
        Next i
    End If
End Function

' 後方互換ラッパー
Private Function FuzzyContains(ByVal textNorm As String, ByVal patternNorm As String) As Boolean
    FuzzyContains = FuzzyContainsEx(textNorm, patternNorm, "")
End Function

' --- Patterns 分割：用途別に2種 ---
' A) トークン用途（OR/AND向け）：空白も区切り
Private Function ParsePatternsTokens(ByVal patStr As String) As Variant
    Dim s As String, parts As Variant, out() As String, i As Long, p As String
    s = CStr(patStr)
    On Error Resume Next: s = StrConv(s, vbNarrow): On Error GoTo 0
    s = Replace(s, vbCrLf, ","): s = Replace(s, vbCr, ","): s = Replace(s, vbLf, ",")
    s = Replace(s, "、", ","): s = Replace(s, "，", ",")
    s = Replace(s, ";", ",")
    s = Replace(s, " ", ",")
    Do While InStr(s, ",,") > 0: s = Replace(s, ",,", ","): Loop
    If Left$(s, 1) = "," Then s = Mid$(s, 2)
    If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)

    parts = Split(s, ",")
    ReDim out(LBound(parts) To UBound(parts))
    For i = LBound(parts) To UBound(parts)
        p = Trim$(CStr(parts(i))): out(i) = p
    Next i

    ' すべて空なら空配列を返す
    If UBound(out) >= LBound(out) Then
        Dim allEmpty As Boolean: allEmpty = True
        For i = LBound(out) To UBound(out)
            If Len(out(i)) > 0 Then allEmpty = False: Exit For
        Next i
        If allEmpty Then ParsePatternsTokens = Array(): Exit Function
    End If
    ParsePatternsTokens = out
End Function

' B) フレーズ用途（PHRASE向け）：空白は保持、カンマ/改行のみで分割
Private Function ParsePatternsPhrases(ByVal patStr As String) As Variant
    Dim s As String, parts As Variant, out() As String, i As Long, p As String
    s = CStr(patStr)
    On Error Resume Next: s = StrConv(s, vbNarrow): On Error GoTo 0
    s = Replace(s, vbCrLf, ","): s = Replace(s, vbCr, ","): s = Replace(s, vbLf, ",")
    s = Replace(s, "、", ","): s = Replace(s, "，", ",")
    s = Replace(s, ";", ",")
    Do While InStr(s, ",,") > 0: s = Replace(s, ",,", ","): Loop
    If Left$(s, 1) = "," Then s = Mid$(s, 2)
    If Right$(s, 1) = "," Then s = Left$(s, Len(s) - 1)

    parts = Split(s, ",")
    ReDim out(LBound(parts) To UBound(parts))
    For i = LBound(parts) To UBound(parts)
        p = Trim$(CStr(parts(i))): out(i) = p
    Next i

    ' すべて空なら空配列を返す
    If UBound(out) >= LBound(out) Then
        Dim allEmpty As Boolean: allEmpty = True
        For i = LBound(out) To UBound(out)
            If Len(out(i)) > 0 Then allEmpty = False: Exit For
        Next i
        If allEmpty Then ParsePatternsPhrases = Array(): Exit Function
    End If
    ParsePatternsPhrases = out
End Function

' --- DecorConfig 読み込み ---
Private Function BuildDecorRules() As Collection
    Const CFG_SHEET_NAME As String = "DecorConfig"

    Dim rules As New Collection
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim patStr As String, beforeTxt As String, afterTxt As String
    Dim posStr As String, mt As String, rm As String, opts As String
    Dim pats As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Set BuildDecorRules = rules: Exit Function

    lastRow = Application.Max( _
        ws.Cells(ws.Rows.Count, 1).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 2).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 3).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 4).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 5).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 6).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 7).End(xlUp).row)

    For r = 2 To lastRow
        patStr = CStr(ws.Cells(r, 1).Value)    ' A Patterns
        beforeTxt = CStr(ws.Cells(r, 2).Value) ' B BeforeText
        afterTxt = CStr(ws.Cells(r, 3).Value)  ' C AfterText
        posStr = LCase$(Trim$(CStr(ws.Cells(r, 4).Value)))    ' D Pos
        mt = LCase$(Trim$(CStr(ws.Cells(r, 5).Value)))        ' E MatchType
        rm = LCase$(Trim$(CStr(ws.Cells(r, 6).Value)))        ' F RemoveAt
        opts = CStr(ws.Cells(r, 7).Value)                     ' G Options

        If Len(Trim$(patStr & beforeTxt & afterTxt & posStr & mt & rm & opts)) = 0 Then GoTo nextRow

        If Len(Trim$(patStr)) = 0 Then
            pats = Array()
        ElseIf mt = "phrase" Then
            pats = ParsePatternsPhrases(patStr)     ' 空白保持
        Else
            pats = ParsePatternsTokens(patStr)      ' 空白も区切り
        End If

        Dim rule As Object: Set rule = CreateObject("Scripting.Dictionary")
        rule("patterns") = pats
        If Len(beforeTxt) > 0 Then rule("before_text") = beforeTxt
        If Len(afterTxt) > 0 Then rule("after_text") = afterTxt
        If Len(posStr) > 0 Then rule("pos") = posStr
        If Len(mt) > 0 Then rule("match_type") = mt
        If Len(rm) > 0 Then rule("remove_at") = rm
        If Len(opts) > 0 Then rule("options") = opts
        rules.add rule
nextRow:
    Next r

    Set BuildDecorRules = rules
End Function

' --- ルール命中判定（Options/MatchType対応・短語min自動緩和・空配列ガード） ---
Private Function RuleHit(ByVal srcNorm As String, ByVal rule As Object) As Boolean
    Dim mt As String, opts As String, pats As Variant, i As Long, p As String
    Dim hasShort As Boolean

    If rule.Exists("match_type") Then mt = LCase$(CStr(rule("match_type"))) Else mt = "or"
    If rule.Exists("options") Then opts = LCase$(CStr(rule("options"))) Else opts = ""
    pats = rule("patterns")

    If ArrCount(pats) = 0 Then RuleHit = False: Exit Function

    hasShort = False
    For i = SafeLBound(pats) To SafeUBound(pats)
        p = NormalizeForMatch(CStr(pats(i)))
        If Len(p) > 0 And Len(p) < 3 Then hasShort = True
    Next i
    If hasShort And InStr(opts, "min=") = 0 Then
        opts = IIf(Len(opts) = 0, "min=1", opts & ";min=1")
    End If

    Select Case mt
        Case "or"
        For i = SafeLBound(pats) To SafeUBound(pats)
            p = NormalizeForMatch(CStr(pats(i)))
            If Len(p) > 0 Then
                ' ★短語は境界や揺れをさらに緩和
                If Len(p) <= 2 Then
                    If ContainsLoose(srcNorm, p) Then RuleHit = True: Exit Function
                Else
                    Dim localOpts As String: localOpts = opts
                    ' ↓ ここで通常判定（nofuzzy等はそのまま）
                    If FuzzyContainsEx(srcNorm, p, localOpts) Then RuleHit = True: Exit Function
                End If
            End If
        Next i
        RuleHit = False

        Case "and"
            For i = SafeLBound(pats) To SafeUBound(pats)
                p = NormalizeForMatch(CStr(pats(i)))
                If Len(p) = 0 Then GoTo skipAND
                If Not FuzzyContainsEx(srcNorm, p, opts) Then RuleHit = False: Exit Function
skipAND:
            Next i
            RuleHit = True

        Case "phrase"
        For i = SafeLBound(pats) To SafeUBound(pats)
            p = NormalizeForMatch(CStr(pats(i)))
            If Len(p) > 0 Then
                ' ★短語は緩和、フレーズは nofuzzy で厳密
                If Len(p) <= 2 Then
                    If ContainsLoose(srcNorm, p) Then RuleHit = True: Exit Function
                Else
                    If FuzzyContainsEx(srcNorm, p, "nofuzzy;" & opts) Then RuleHit = True: Exit Function
                End If
            End If
        Next i
        RuleHit = False


        Case "regex"
            RuleHit = RegexTest(srcNorm, CStr(pats(SafeLBound(pats))))

        Case Else ' "or"
            For i = SafeLBound(pats) To SafeUBound(pats)
                p = NormalizeForMatch(CStr(pats(i)))
                If Len(p) > 0 And FuzzyContainsEx(srcNorm, p, opts) Then RuleHit = True: Exit Function
            Next i
            RuleHit = False
    End Select
End Function

' --- 既存装飾だけを剥がしてから、今のルールで付け直す ---
Private Function StripDecorations(ByVal s As String, ByVal rules As Collection) As String
    Dim result As String: result = Trim$(CStr(s))   ' 前後空白を正規化
    Dim rule As Object, frag As String, rm As String, pos As String

    For Each rule In rules
        rm = "edge": If rule.Exists("remove_at") Then rm = LCase$(CStr(rule("remove_at")))
        pos = "both"
        If rule.Exists("pos") Then
            Select Case LCase$(CStr(rule("pos")))
                Case "before": pos = "before"
                Case "after":  pos = "after"
                Case "both", "default_nohit": pos = "both"
            End Select
        End If

        If rule.Exists("before_text") Then
            frag = CStr(rule("before_text"))
            If Len(frag) > 0 Then
                If rm = "edge" Then result = EdgeStrip(result, frag, "before") Else result = Replace(result, frag, "")
            End If
        End If
        If rule.Exists("after_text") Then
            frag = CStr(rule("after_text"))
            If Len(frag) > 0 Then
                If rm = "edge" Then result = EdgeStrip(result, frag, "after") Else result = Replace(result, frag, "")
            End If
        End If
        If rule.Exists("text") Then
            frag = CStr(rule("text"))
            If Len(frag) > 0 Then
                If rm = "edge" Then result = EdgeStrip(result, frag, pos) Else result = Replace(result, frag, "")
            End If
        End If
    Next rule

    If Len(DECOR_NOHIT_BEFORE_TEXT) > 0 Then result = EdgeStrip(result, DECOR_NOHIT_BEFORE_TEXT, "before")
    If Len(DECOR_NOHIT_AFTER_TEXT) > 0 Then result = EdgeStrip(result, DECOR_NOHIT_AFTER_TEXT, "after")

    StripDecorations = Trim$(result)
End Function

' --- 装飾適用（重複防止・既存装飾剥離・PHRASE/Options対応） ---
' 装飾適用：exclude / exclusive / default_nohit を厳密に扱う完全版
' - topRow        : 2行ペアの上段行
' - matchedKeyRaw : 参照ヒット時のキー（未使用なら "" でOK）
Private Sub ApplyDecorations(ByVal topRow As Long, ByVal matchedKeyRaw As String)
    On Error GoTo EH

    Dim belowRow As Long: belowRow = topRow + 1
    If belowRow > LayoutRowEnd() Then Exit Sub

    ' 判定に使う生値（右:上段先頭セル / 左:上段+下段の連結）
    Dim rightTop As String, leftPair As String
    rightTop = CStr(GetRowMergeTopLeft(Me, topRow, LayoutRightColStart(), LayoutRightColEnd()).Value)
    leftPair = CStr(Me.Cells(topRow, LayoutLeftCol()).Value) & CStr(Me.Cells(belowRow, LayoutLeftCol()).Value)

    ' マッチ用の正規化文字列（キー＋右上段＋左2行）
    Dim srcNorm As String
    srcNorm = NormalizeForMatch(matchedKeyRaw & " " & rightTop & " " & leftPair)
    If Len(srcNorm) = 0 Then
        Dim fallbackNorm As String
        fallbackNorm = NormalizeForMatch(rightTop & " " & leftPair)
        If Len(fallbackNorm) = 0 Then Exit Sub
        srcNorm = fallbackNorm
    End If
    Debug.Print "[Decor] row=", topRow, " srcNorm=[", srcNorm, "]"

    ' 出力先：右側結合ブロックの「下段」の結合先頭セル
    Dim cell As Range
    Set cell = GetRowMergeTopLeft(Me, belowRow, LayoutRightColStart(), LayoutRightColEnd())
    If cell Is Nothing Then Exit Sub

    ' ルール読み込み
    Dim rules As Collection: Set rules = BuildDecorRules()

    ' 既存装飾の扱い：本文だけ抽出（または全消し）
    Dim existingRaw As String: existingRaw = CStr(cell.Value)
    Dim baseText As String
    If DECOR_CLEAR_BELOW_FIRST Then
        baseText = ""
    Else
        baseText = StripDecorations(existingRaw, rules)
    End If

    ' ---- 事前スキャン：exclude / exclusive / 通常ヒット有無の確定 ----
    Dim ropts As String
    Dim hasExclusive As Boolean: hasExclusive = False
    Dim anyNormalHit As Boolean: anyNormalHit = False
    Dim isExcluded As Boolean:    isExcluded = False

    Dim rScan As Object
    For Each rScan In rules
        If RuleHit(srcNorm, rScan) Then
            ropts = LCase$(IIf(rScan.Exists("options"), CStr(rScan("options")), ""))
            If InStr(ropts, "exclude") > 0 Then
                isExcluded = True
                Exit For
            End If
            If InStr(ropts, "exclusive") > 0 Then hasExclusive = True
            ' default_nohit 以外が当たれば通常ヒット
            If Not (rScan.Exists("pos") And LCase$(CStr(rScan("pos"))) = "default_nohit") Then
                anyNormalHit = True
            End If
        End If
    Next rScan

    ' ---- 除外ヒット：何も付けずに（必要なら装飾だけ剥がして）終了 ----
    If isExcluded Then
        cell.Value = StripDecorations(CStr(cell.Value), rules)
        Debug.Print "[Decor] excluded (no decoration)"
        Exit Sub
    End If

    ' ---- 適用フェーズ ----
    Dim toPrepend As String, toAppend As String
    toPrepend = "": toAppend = ""

    Dim rule As Object, posVal As String

    If anyNormalHit Then
        ' 通常命中があるとき：default_nohit とグローバル既定は一切付けない
        For Each rule In rules
            If RuleHit(srcNorm, rule) Then
                ' 未ヒット専用はスキップ
                If rule.Exists("pos") Then
                    If LCase$(CStr(rule("pos"))) = "default_nohit" Then GoTo NextRule1
                End If
                ropts = LCase$(IIf(rule.Exists("options"), CStr(rule("options")), ""))
                ' exclusive が一つでもあれば exclusive 以外は無視
                If hasExclusive And InStr(ropts, "exclusive") = 0 Then GoTo NextRule1

                ' 装飾追加（重複は AppendUnique で抑止）
                If rule.Exists("before_text") Then AppendUnique toPrepend, CStr(rule("before_text")), baseText
                If rule.Exists("after_text") Then AppendUnique toAppend, CStr(rule("after_text")), baseText

                ' 旧式 text + pos（互換）
                If rule.Exists("text") Then
                    posVal = LCase$(IIf(rule.Exists("pos"), CStr(rule("pos")), "after"))
                    Select Case posVal
                        Case "before": AppendUnique toPrepend, CStr(rule("text")), baseText
                        Case "after":  AppendUnique toAppend, CStr(rule("text")), baseText
                        Case "both":   AppendUnique toPrepend, CStr(rule("text")), baseText
                                       AppendUnique toAppend, CStr(rule("text")), baseText
                    End Select
                End If
            End If
NextRule1:
        Next rule

    Else
        ' 通常命中ゼロ：default_nohit 行を優先、無ければグローバル既定
        Dim usedDefaultRule As Boolean: usedDefaultRule = False

        For Each rule In rules
            If rule.Exists("pos") Then
                If LCase$(CStr(rule("pos"))) = "default_nohit" Then
                    If RuleHit(srcNorm, rule) Then
                        If rule.Exists("before_text") Then AppendUnique toPrepend, CStr(rule("before_text")), baseText
                        If rule.Exists("after_text") Then AppendUnique toAppend, CStr(rule("after_text")), baseText
                        usedDefaultRule = True
                    End If
                End If
            End If
        Next rule

        If Not usedDefaultRule Then
            If Len(DECOR_NOHIT_BEFORE_TEXT) > 0 Then AppendUnique toPrepend, DECOR_NOHIT_BEFORE_TEXT, baseText
            If Len(DECOR_NOHIT_AFTER_TEXT) > 0 Then AppendUnique toAppend, DECOR_NOHIT_AFTER_TEXT, baseText
        End If
    End If

    ' 書き込み（装飾も本文も空なら何もしない）
    If Len(toPrepend) = 0 And Len(toAppend) = 0 And Len(baseText) = 0 Then Exit Sub

    cell.Value = toPrepend & baseText & toAppend
    Debug.Print "[Decor] write row=", topRow, _
                " hitNormal=", anyNormalHit, " excl=", hasExclusive, _
                " val=[", toPrepend & baseText & toAppend, "]"
    Exit Sub

EH:
    Debug.Print "[Decor][ERR] row=", topRow, " err=", Err.Number, Err.Description
End Sub

'======================== ユーティリティ ========================
Private Function PairTopRow(ByVal anyRow As Long) As Long
    PairTopRow = LayoutRowStart() + ((anyRow - LayoutRowStart()) \ 2) * 2
End Function

Private Sub WriteRight(ByVal topRow As Long, ByVal valueToWrite As String)
    On Error GoTo WR_E
    Dim tgt As Range
    Set tgt = Me.Range(Me.Cells(topRow, LayoutRightColStart()), Me.Cells(topRow, LayoutRightColEnd())).MergeArea.Cells(1, 1)
    tgt.Value = valueToWrite
    Debug.Print "[WRITE-R] row=", topRow, " val=[", valueToWrite, "]"
    Exit Sub
WR_E:
    Debug.Print "[WRITE-R][ERR] row=", topRow, " err=", Err.Number, Err.Description
End Sub

Private Sub WriteLeftSplit(ByVal topRow As Long, ByVal valueToWrite As String)
    On Error GoTo WL_E
    Dim digits As String, upper2 As String, lower3 As String
    digits = ExtractDigitsHalf(valueToWrite)
    If Len(digits) = 0 Then
        upper2 = "": lower3 = ""
    Else
        upper2 = Left$(digits, WorksheetFunction.Min(2, Len(digits)))
        lower3 = Right$(digits, WorksheetFunction.Min(3, Len(digits)))
    End If
    Me.Cells(topRow, LayoutLeftCol()).Value = upper2
    If topRow + 1 <= LayoutRowEnd() Then Me.Cells(topRow + 1, LayoutLeftCol()).Value = lower3
    Debug.Print "[WRITE-L] row=", topRow, " up=[", upper2, "] low=[", lower3, "] src=[", valueToWrite, "]"
    Exit Sub
WL_E:
    Debug.Print "[WRITE-L][ERR] row=", topRow, " err=", Err.Number, Err.Description
End Sub

Private Sub WriteLeftFallback(ByVal topRow As Long)
    Me.Cells(topRow, LayoutLeftCol()).Value = UNMATCHED_LEFT_UPPER
    If topRow + 1 <= LayoutRowEnd() Then Me.Cells(topRow + 1, LayoutLeftCol()).Value = UNMATCHED_LEFT_LOWER
    Debug.Print "[WRITE-L] Fallback @row ", topRow
End Sub

Private Function GetRowMergeTopLeft(ws As Worksheet, ByVal row As Long, ByVal startCol As Long, ByVal endCol As Long) As Range
    Dim c As Long, cell As Range, area As Range
    For c = startCol To endCol
        Set cell = ws.Cells(row, c)
        Set area = cell.MergeArea
        If Not Intersect(area, ws.Rows(row)) Is Nothing Then
            Set GetRowMergeTopLeft = area.Cells(1, 1)
            Exit Function
        End If
    Next c
    Set GetRowMergeTopLeft = ws.Cells(row, startCol)
End Function

Private Function NormalizeForMatch(ByVal s As String) As String
    Dim t As String: t = CStr(s)
    On Error Resume Next: t = StrConv(t, vbNarrow): On Error GoTo 0
    t = Replace(t, ChrW(&H3000), " ")
    t = Replace(Replace(t, vbCr, " "), vbLf, " ")
    t = Replace(t, vbTab, " ")
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    NormalizeForMatch = LCase$(Trim$(t))
End Function

Private Function ExtractDigitsHalf(ByVal s As String) As String
    Dim t As String, i As Long, ch As String, buf As String
    t = s: On Error Resume Next: t = StrConv(t, vbNarrow): On Error GoTo 0
    buf = ""
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "#" Then buf = buf & ch
    Next i
    ExtractDigitsHalf = buf
End Function

Private Sub RefDebug_Stats(ByVal src As String, ByVal map34 As Object, ByVal map43 As Object)
    Dim k As Variant, c34 As Long, c43 As Long
    For Each k In map34.keys: c34 = c34 + map34(k).Count: Next k
    For Each k In map43.keys: c43 = c43 + map43(k).Count: Next k
    Debug.Print "[RefLoad] src=" & src & " 3→4=" & c34 & "件 4→3=" & c43 & "件"
End Sub

'======================== 診断（任意） ========================
Public Sub TraceDecorAtRow(ByVal topRow As Long)
    Dim belowRow As Long: belowRow = topRow + 1
    Dim rightTop As String, leftPair As String, srcNorm As String
    Dim rules As Collection, rule As Object, p As Variant
    Dim ok As Boolean, mt As String, opts As String, pats As Variant, patList As String

    rightTop = CStr(GetRowMergeTopLeft(Me, topRow, LayoutRightColStart(), LayoutRightColEnd()).Value)
    leftPair = CStr(Me.Cells(topRow, LayoutLeftCol()).Value) & CStr(Me.Cells(belowRow, LayoutLeftCol()).Value)
    srcNorm = NormalizeForMatch(rightTop & " " & leftPair)
    Set rules = BuildDecorRules()

    Debug.Print "---- TraceDecorAtRow row=" & topRow & " ----"
    Debug.Print "srcNorm=[", srcNorm, "] rightTop=[", rightTop, "] leftPair=[", leftPair, "]"

    For Each rule In rules
        If rule.Exists("match_type") Then mt = CStr(rule("match_type")) Else mt = "or"
        If rule.Exists("options") Then opts = CStr(rule("options")) Else opts = ""
        pats = rule("patterns")
        patList = "": If ArrCount(pats) > 0 Then For Each p In pats: patList = patList & IIf(Len(patList) > 0, " | ", "") & CStr(p): Next p
        ok = RuleHit(srcNorm, rule)
        Debug.Print " [Rule] mt=" & mt & " opts=" & opts & " pats={" & patList & "} -> " & IIf(ok, "HIT", "MISS")
    Next rule
    Debug.Print "----------------------------------------"
End Sub
