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

    If tlRow < ROW_START Or tlRow > ROW_END Then
        Debug.Print "  [SKIP] row out"
        RouteOneCell = False: Exit Function
    End If

    '==== 左（列4）→右 ====
    If tlCol = LEFT_COL Then
        Dim topRowL As Long: topRowL = PairTopRow(tlRow)
        If tlRow <> topRowL Then
            Debug.Print "  [SKIP L] lower row edited (pairTop=" & topRowL & ")"
            RouteOneCell = False: Exit Function
        End If

        Dim inValL As String, outVal As String, matchedKey As String
        Dim score As Long
        inValL = NormalizeForMatch(CStr(Me.Cells(topRowL, LEFT_COL).Value) & CStr(Me.Cells(topRowL + 1, LEFT_COL).Value))
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
    If tlCol >= RIGHT_COL_START And tlCol <= RIGHT_COL_END Then
        Dim topRowR As Long: topRowR = PairTopRow(tlRow)
        If tlRow <> topRowR Then
            Debug.Print "  [SKIP R] lower row edited (pairTop=" & topRowR & ")"
            RouteOneCell = False: Exit Function
        End If

        Dim rightTop As Range
        Set rightTop = GetRowMergeTopLeft(Me, topRowR, RIGHT_COL_START, RIGHT_COL_END)

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

        If Len(Trim$(patStr & beforeTxt & afterTxt & posStr & mt & rm & opts)) = 0 Then GoTo NextRow

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
NextRow:
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
Private Sub ApplyDecorations(ByVal topRow As Long, ByVal matchedKeyRaw As String)
    On Error GoTo EH

    Dim belowRow As Long: belowRow = topRow + 1
    If belowRow > ROW_END Then Exit Sub

    Dim rightTop As String, leftPair As String
    rightTop = CStr(GetRowMergeTopLeft(Me, topRow, RIGHT_COL_START, RIGHT_COL_END).Value)
    leftPair = CStr(Me.Cells(topRow, LEFT_COL).Value) & CStr(Me.Cells(belowRow, LEFT_COL).Value)

    Dim srcNorm As String
    srcNorm = NormalizeForMatch(matchedKeyRaw & " " & rightTop & " " & leftPair)
    If Len(srcNorm) = 0 Then
        Dim fallbackNorm As String
        fallbackNorm = NormalizeForMatch(rightTop & " " & leftPair)
        If Len(fallbackNorm) = 0 Then Exit Sub
        srcNorm = fallbackNorm
    End If

    Debug.Print "[Decor] row=", topRow, " srcNorm=[", srcNorm, "]"

    Dim cell As Range
    Set cell = GetRowMergeTopLeft(Me, belowRow, RIGHT_COL_START, RIGHT_COL_END)
    If cell Is Nothing Then Exit Sub

    Dim rules As Collection: Set rules = BuildDecorRules()

    Dim existingRaw As String: existingRaw = CStr(cell.Value)
    Dim baseText As String
    If DECOR_CLEAR_BELOW_FIRST Then
        baseText = ""
    Else
        baseText = StripDecorations(existingRaw, rules)
    End If

    Dim toPrepend As String, toAppend As String
    toPrepend = "": toAppend = ""

    Dim rule As Object, hit As Boolean: hit = False
    Dim hasExclusive As Boolean: hasExclusive = False
    Dim ropts As String
    
    ' --- 先に除外命中があるか確認 ---
    Dim isExcluded As Boolean: isExcluded = False
    For Each rule In rules
        If RuleHit(srcNorm, rule) Then
            ropts = LCase$(IIf(rule.Exists("options"), CStr(rule("options")), ""))
            If InStr(ropts, "exclude") > 0 Then
                isExcluded = True
                Exit For
            End If
        End If
    Next rule
    If isExcluded Then
        ' 既存装飾があれば剥がすだけで終了（剥がしたくなければこの2行をコメントアウト）
        Dim baseTextOnly As String
        baseTextOnly = StripDecorations(CStr(cell.Value), rules)
        cell.Value = baseTextOnly
        Debug.Print "[Decor] excluded (no decoration)"
        Exit Sub
    End If
    
    ' --- 先に exclusive 命中の有無だけ確認 ---
    For Each rule In rules
        If RuleHit(srcNorm, rule) Then
            If rule.Exists("options") Then ropts = LCase$(CStr(rule("options"))) Else ropts = ""
            If InStr(ropts, "exclusive") > 0 Then
                hasExclusive = True
                Exit For
            End If
        End If
    Next rule
    
    ' --- 本適用：exclusiveがあればそれだけ、無ければ全部 ---
    For Each rule In rules
        If RuleHit(srcNorm, rule) Then
            If rule.Exists("options") Then ropts = LCase$(CStr(rule("options"))) Else ropts = ""
            If hasExclusive And InStr(ropts, "exclusive") = 0 Then GoTo NextRule
    
            If rule.Exists("before_text") Then AppendUnique toPrepend, CStr(rule("before_text")), baseText
            If rule.Exists("after_text") Then AppendUnique toAppend, CStr(rule("after_text")), baseText
            If rule.Exists("text") Then
                Dim posVal As String
                If rule.Exists("pos") Then posVal = LCase$(CStr(rule("pos"))) Else posVal = "after"
                Select Case posVal
                    Case "before": AppendUnique toPrepend, CStr(rule("text")), baseText
                    Case "after":  AppendUnique toAppend, CStr(rule("text")), baseText
                    Case "both":   AppendUnique toPrepend, CStr(rule("text")), baseText
                                   AppendUnique toAppend, CStr(rule("text")), baseText
                End Select
            End If
        End If
NextRule:
    Next rule


    If Not hit Then
        Dim hasDefault As Boolean: hasDefault = False
        For Each rule In rules
            If rule.Exists("pos") Then
                If LCase$(CStr(rule("pos"))) = "default_nohit" Then
                    hasDefault = True
                    If rule.Exists("before_text") Then AppendUnique toPrepend, CStr(rule("before_text")), baseText
                    If rule.Exists("after_text") Then AppendUnique toAppend, CStr(rule("after_text")), baseText
                End If
            End If
        Next rule

        If Not hasDefault Then
            AppendUnique toPrepend, DECOR_NOHIT_BEFORE_TEXT, baseText
            AppendUnique toAppend, DECOR_NOHIT_AFTER_TEXT, baseText
        End If
    End If

    If Len(toPrepend) = 0 And Len(toAppend) = 0 And Len(baseText) = 0 Then Exit Sub

    cell.Value = toPrepend & baseText & toAppend
    Debug.Print "[Decor] write row=", topRow, " val=[", toPrepend & baseText & toAppend, "]  hit=", hit
    Exit Sub
EH:
    Debug.Print "[Decor][ERR] row=", topRow, " err=", Err.Number, Err.Description
End Sub

'======================== ユーティリティ ========================
Private Function PairTopRow(ByVal anyRow As Long) As Long
    If anyRow Mod 2 = 0 Then PairTopRow = anyRow Else PairTopRow = anyRow - 1
End Function

Private Sub WriteRight(ByVal topRow As Long, ByVal valueToWrite As String)
    On Error GoTo WR_E
    Dim tgt As Range
    Set tgt = Me.Range(Me.Cells(topRow, RIGHT_COL_START), Me.Cells(topRow, RIGHT_COL_END)).MergeArea.Cells(1, 1)
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
    Me.Cells(topRow, LEFT_COL).Value = upper2
    If topRow + 1 <= ROW_END Then Me.Cells(topRow + 1, LEFT_COL).Value = lower3
    Debug.Print "[WRITE-L] row=", topRow, " up=[", upper2, "] low=[", lower3, "] src=[", valueToWrite, "]"
    Exit Sub
WL_E:
    Debug.Print "[WRITE-L][ERR] row=", topRow, " err=", Err.Number, Err.Description
End Sub

Private Sub WriteLeftFallback(ByVal topRow As Long)
    Me.Cells(topRow, LEFT_COL).Value = UNMATCHED_LEFT_UPPER
    If topRow + 1 <= ROW_END Then Me.Cells(topRow + 1, LEFT_COL).Value = UNMATCHED_LEFT_LOWER
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

    rightTop = CStr(GetRowMergeTopLeft(Me, topRow, RIGHT_COL_START, RIGHT_COL_END).Value)
    leftPair = CStr(Me.Cells(topRow, LEFT_COL).Value) & CStr(Me.Cells(belowRow, LEFT_COL).Value)
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


