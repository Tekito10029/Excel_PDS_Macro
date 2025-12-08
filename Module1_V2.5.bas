Attribute VB_Name = "Module1"
'V_2.5
Option Explicit

'==============================================================
' Module1：パス管理 ＋ DecorConfigコピー(A:G) ＋ RefCache更新(C/D[+Tag]→A:C)
'            ＋ 空配列ガード(SafeL/UBound, ArrCount)
'            ＋ 診断ダンプ/ラッパー
'==============================================================

'==== ▼▼▼ 環境に合わせて変更する定数 ▼▼▼ =====================

'--- 装飾ルール(DecorConfig) 側 ---
Private Const DECOR_CFG_PATH_NAME As String = "CFG_DECOR_PATH"   ' 名前定義で保持
Public Const DECOR_CFG_DEFAULT    As String = "\\svr03\鶴見工場\製造\0?.沓\00.製造表 各社\DecorMaster.xlsx"         ' 既定パス(UNC推奨/未使用なら"")
Public Const DECOR_CFG_SHEET      As String = "DecorConfig"      ' マスタ/運用シート名

'--- 参照マスタ(Ref) 側 ---
Private Const REF_CFG_PATH_NAME   As String = "CFG_REF_PATH"     ' 名前定義で保持
Public Const REF_CFG_DEFAULT      As String = "\\svr03\鶴見工場\製造\0?.沓\00.製造表 各社\沓用コード表.xlsx"        ' 既定パス(UNC推奨/未使用なら"")
Public Const REF_SOURCE_SHEET     As String = "沓用コード"           ' ★参照元ブックのシート名に合わせて変更
Public Const REF_COL_KEY1         As Long = 3                    ' ★C列=「3→4」検索キー列
Public Const REF_COL_KEY2         As Long = 4                    ' ★D列=「4→3」検索キー列
Public Const REF_COL_TAG          As Long = 5                    ' ★任意:タグ列(0=無効, 5=E列など)
Public Const REF_CACHE_SHEET      As String = "RefCache"         ' 運用ブック内のキャッシュシート名

'==== ▲▲▲ 必要に応じて変更 ▲▲▲ ===============================


'==============================================================
' 空配列/未配列ガード
'==============================================================
Public Function SafeUBound(v As Variant) As Long
    On Error GoTo EH
    If IsArray(v) Then SafeUBound = UBound(v) Else SafeUBound = -1
    Exit Function
EH:
    SafeUBound = -1
End Function

Public Function SafeLBound(v As Variant) As Long
    On Error GoTo EH
    If IsArray(v) Then SafeLBound = LBound(v) Else SafeLBound = 0
    Exit Function
EH:
    SafeLBound = 0
End Function

Public Function ArrCount(v As Variant) As Long
    If Not IsArray(v) Then ArrCount = 0: Exit Function
    Dim lo As Long, hi As Long
    hi = SafeUBound(v): lo = SafeLBound(v)
    If hi < lo Then ArrCount = 0 Else ArrCount = hi - lo + 1
End Function

'==============================================================
' 共通: パス存在確認（UNC/日本語パス対応）
'==============================================================
Private Function SafeFileExists(ByVal path As String) As Boolean
    On Error Resume Next
    Dim p As String: p = Trim$(path)
    If Len(p) = 0 Then SafeFileExists = False: Exit Function
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    SafeFileExists = fso.FileExists(p)
    If Not SafeFileExists Then SafeFileExists = fso.FolderExists(p)
    If Err.Number <> 0 Then SafeFileExists = False: Err.Clear
    On Error GoTo 0
End Function

'==============================================================
' 名前定義にパス保存（VeryHiddenシートに格納）
'==============================================================
Private Sub SavePathToName(ByVal nm As String, ByVal p As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ConfigHidden")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.add
        ws.Name = "ConfigHidden"
        ws.Visible = xlSheetVeryHidden
    End If
    Select Case nm
        Case DECOR_CFG_PATH_NAME: ws.Range("A1").Value = p: ThisWorkbook.Names.add Name:=DECOR_CFG_PATH_NAME, RefersTo:=ws.Range("A1")
        Case REF_CFG_PATH_NAME:   ws.Range("B1").Value = p: ThisWorkbook.Names.add Name:=REF_CFG_PATH_NAME, RefersTo:=ws.Range("B1")
        Case Else
            ws.Range("C1").Value = p
            ThisWorkbook.Names.add Name:=nm, RefersTo:=ws.Range("C1")
    End Select
End Sub

Private Function LoadPathFromName(ByVal nm As String) As String
    On Error Resume Next
    Dim n As Name: Set n = ThisWorkbook.Names(nm)
    If Not n Is Nothing Then LoadPathFromName = CStr(n.RefersToRange.Value) Else LoadPathFromName = ""
    On Error GoTo 0
End Function

'==============================================================
' ファイル選択共通
'==============================================================
Private Function PickFilePath(ByVal titleText As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = titleText
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.add "Excel Files", "*.xlsx;*.xlsm;*.xlsb"
        If .Show = -1 Then
            PickFilePath = .SelectedItems(1)
        Else
            PickFilePath = ""
        End If
    End With
End Function

'==============================================================
' DecorConfig 用パス取得（保存＞既定＞選択）
'==============================================================
Public Function GetDecorCfgPath() As String
    Dim saved As String, p As String
    saved = LoadPathFromName(DECOR_CFG_PATH_NAME)
    If SafeFileExists(saved) Then
        GetDecorCfgPath = saved: Debug.Print "[DecorCfg] use saved:", saved: Exit Function
    ElseIf Len(saved) > 0 Then
        Debug.Print "[DecorCfg] saved but missing:", saved
    End If

    If SafeFileExists(DECOR_CFG_DEFAULT) Then
        SavePathToName DECOR_CFG_PATH_NAME, DECOR_CFG_DEFAULT
        GetDecorCfgPath = DECOR_CFG_DEFAULT
        Debug.Print "[DecorCfg] use default:", DECOR_CFG_DEFAULT
        Exit Function
    End If

    p = PickFilePath("装飾ルールマスタ(DecorConfig)を選択してください")
    Debug.Print "[DecorCfg] picked:", p, " exists=", SafeFileExists(p)
    If SafeFileExists(p) Then
        SavePathToName DECOR_CFG_PATH_NAME, p
        GetDecorCfgPath = p
    Else
        GetDecorCfgPath = ""
    End If
End Function

'==============================================================
' Ref(参照マスタ) 用パス取得（保存＞既定＞選択）
'==============================================================
Public Function GetRefCfgPath() As String
    Dim saved As String, p As String
    saved = LoadPathFromName(REF_CFG_PATH_NAME)
    If SafeFileExists(saved) Then
        GetRefCfgPath = saved: Debug.Print "[RefCfg] use saved:", saved: Exit Function
    ElseIf Len(saved) > 0 Then
        Debug.Print "[RefCfg] saved but missing:", saved
    End If

    If SafeFileExists(REF_CFG_DEFAULT) Then
        SavePathToName REF_CFG_PATH_NAME, REF_CFG_DEFAULT
        GetRefCfgPath = REF_CFG_DEFAULT
        Debug.Print "[RefCfg] use default:", REF_CFG_DEFAULT
        Exit Function
    End If

    p = PickFilePath("参照マスタブック(C/D列)を選択してください")
    Debug.Print "[RefCfg] picked:", p, " exists=", SafeFileExists(p)
    If SafeFileExists(p) Then
        SavePathToName REF_CFG_PATH_NAME, p
        GetRefCfgPath = p
    Else
        GetRefCfgPath = ""
    End If
End Function

'==============================================================
' DecorConfig（A:G）コピー：マスタ → このブック
'==============================================================
Public Sub UpdateDecorConfigFromFile()
    Dim p As String, wbCfg As Workbook, wbDst As Workbook, wsSrc As Worksheet, wsDst As Worksheet
    Dim needClose As Boolean, lastSrc As Long, wb As Workbook

    Set wbDst = ThisWorkbook

    p = GetDecorCfgPath()
    If p = "" Then
        MsgBox "装飾マスタファイルが選択されなかったため、処理を中止します。", vbExclamation
        Exit Sub
    End If
    If Not SafeFileExists(p) Then
        MsgBox "設定されている装飾マスタファイルが見つかりません。" & vbCrLf & p, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 既に開いている場合は再利用
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, p, vbTextCompare) = 0 Then Set wbCfg = wb: Exit For
    Next wb
    If wbCfg Is Nothing Then
        On Error Resume Next
        Set wbCfg = Workbooks.Open(Filename:=p, ReadOnly:=True)
        needClose = True
        On Error GoTo 0
    End If
    If wbCfg Is Nothing Then
        MsgBox "装飾マスタファイルを開けませんでした。" & vbCrLf & p, vbExclamation
        GoTo DEC_EXIT
    End If
    If wbCfg Is wbDst Then
        MsgBox "このブック自身を装飾マスタとして指定できません。", vbExclamation
        GoTo DEC_EXIT
    End If

    On Error Resume Next
    Set wsSrc = wbCfg.Worksheets(DECOR_CFG_SHEET)
    On Error GoTo 0
    If wsSrc Is Nothing Then
        MsgBox "装飾マスタにシート '" & DECOR_CFG_SHEET & "' がありません。", vbExclamation
        GoTo DEC_EXIT
    End If

    On Error Resume Next
    Set wsDst = wbDst.Worksheets(DECOR_CFG_SHEET)
    On Error GoTo 0
    If wsDst Is Nothing Then
        Set wsDst = wbDst.Worksheets.add(After:=wbDst.Sheets(wbDst.Sheets.Count))
        wsDst.Name = DECOR_CFG_SHEET
        wsDst.Visible = xlSheetHidden
    End If

    ' A:G のどこかに値がある最終行
    lastSrc = Application.Max( _
        wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row, _
        wsSrc.Cells(wsSrc.Rows.Count, 2).End(xlUp).row, _
        wsSrc.Cells(wsSrc.Rows.Count, 3).End(xlUp).row, _
        wsSrc.Cells(wsSrc.Rows.Count, 4).End(xlUp).row, _
        wsSrc.Cells(wsSrc.Rows.Count, 5).End(xlUp).row, _
        wsSrc.Cells(wsSrc.Rows.Count, 6).End(xlUp).row, _
        wsSrc.Cells(wsSrc.Rows.Count, 7).End(xlUp).row)

    wsDst.Cells.ClearContents
    If lastSrc >= 1 Then
        wsDst.Range("A1:G" & lastSrc).Value = wsSrc.Range("A1:G" & lastSrc).Value
    End If

    MsgBox "装飾ルール(DecorConfig)を更新しました。（A?G 列対応）", vbInformation

DEC_EXIT:
    On Error Resume Next
    If needClose And Not wbCfg Is Nothing Then wbCfg.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Sub

'==============================================================
' RefCache 更新：参照マスタ(C/D[+Tag]) → このブック RefCache(A:C)
'==============================================================
Public Sub UpdateRefCacheFromRefFile()
    Dim p As String, wbRef As Workbook, wbDst As Workbook
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim needClose As Boolean, last As Long, r As Long
    Dim k1 As String, k2 As String, tag As String
    Dim wb As Workbook

    Set wbDst = ThisWorkbook

    p = GetRefCfgPath()
    If p = "" Then
        MsgBox "参照マスタファイルが選択されなかったため、処理を中止します。", vbExclamation
        Exit Sub
    End If
    If Not SafeFileExists(p) Then
        MsgBox "設定されている参照マスタファイルが見つかりません。" & vbCrLf & p, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 既に開いている場合は再利用
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, p, vbTextCompare) = 0 Then Set wbRef = wb: Exit For
    Next wb
    If wbRef Is Nothing Then
        On Error Resume Next
        Set wbRef = Workbooks.Open(Filename:=p, ReadOnly:=True)
        needClose = True
        On Error GoTo 0
    End If
    If wbRef Is Nothing Then
        MsgBox "参照マスタファイルを開けませんでした。", vbExclamation
        GoTo REF_EXIT
    End If
    If wbRef Is wbDst Then
        MsgBox "このブック自身を参照マスタとして指定できません。", vbExclamation
        GoTo REF_EXIT
    End If

    On Error Resume Next
    Set wsSrc = wbRef.Worksheets(REF_SOURCE_SHEET)
    On Error GoTo 0
    If wsSrc Is Nothing Then
        MsgBox "参照マスタにシート '" & REF_SOURCE_SHEET & "' がありません。", vbExclamation
        GoTo REF_EXIT
    End If

    ' キャッシュ先シート取得/作成
    On Error Resume Next
    Set wsDst = wbDst.Worksheets(REF_CACHE_SHEET)
    On Error GoTo 0
    If wsDst Is Nothing Then
        Set wsDst = wbDst.Worksheets.add(After:=wbDst.Sheets(wbDst.Sheets.Count))
        wsDst.Name = REF_CACHE_SHEET
        wsDst.Visible = xlSheetHidden
    End If

    ' 見出し
    wsDst.Cells.ClearContents
    wsDst.Range("A1").Value = "C_col"   ' = 参照元3列目 (Key1)
    wsDst.Range("B1").Value = "D_col"   ' = 参照元4列目 (Key2)
    wsDst.Range("C1").Value = "Tag"     ' 任意タグ

    ' 参照元の最終行（Key1/Key2/Tag のいずれかに値がある行まで）
    last = Application.Max( _
        wsSrc.Cells(wsSrc.Rows.Count, REF_COL_KEY1).End(xlUp).row, _
        wsSrc.Cells(wsSrc.Rows.Count, REF_COL_KEY2).End(xlUp).row, _
        IIf(REF_COL_TAG > 0, wsSrc.Cells(wsSrc.Rows.Count, REF_COL_TAG).End(xlUp).row, 1))

    If last < 2 Then
        MsgBox "参照マスタに有効なデータがありません。", vbExclamation
        GoTo REF_EXIT
    End If

    ' 値コピー（A:Cに格納）
    For r = 2 To last
        k1 = CStr(wsSrc.Cells(r, REF_COL_KEY1).Value)
        k2 = CStr(wsSrc.Cells(r, REF_COL_KEY2).Value)
        If REF_COL_TAG > 0 Then
            tag = CStr(wsSrc.Cells(r, REF_COL_TAG).Value)
        Else
            tag = ""
        End If
        wsDst.Cells(r, 1).Value = k1
        wsDst.Cells(r, 2).Value = k2
        wsDst.Cells(r, 3).Value = tag
    Next r

    MsgBox "参照キャッシュ(RefCache)を更新しました。(" & (last - 1) & "件)", vbInformation

REF_EXIT:
    On Error Resume Next
    If needClose And Not wbRef Is Nothing Then wbRef.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Sub

'==============================================================
' 便利：両方まとめて更新
'==============================================================
Public Sub UpdateAllMasterCaches()
    UpdateDecorConfigFromFile
    UpdateRefCacheFromRefFile
End Sub

'==============================================================
' 診断ユーティリティ
'==============================================================
Public Sub Debug_DecorRulesSummary()
    Dim ws As Worksheet, lastRow As Long
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("DecorConfig"): On Error GoTo 0
    If ws Is Nothing Then Debug.Print "[DecorRules] NOT FOUND (DecorConfig)": Exit Sub
    lastRow = Application.Max( _
        ws.Cells(ws.Rows.Count, 1).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 2).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 3).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 4).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 5).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 6).End(xlUp).row, _
        ws.Cells(ws.Rows.Count, 7).End(xlUp).row)
    Debug.Print "[DecorRules] rows(A:G) =", lastRow
    If lastRow >= 2 Then
        Debug.Print "  A2=", CStr(ws.Cells(2, 1).Value), " E2=", CStr(ws.Cells(2, 5).Value), " G2=", CStr(ws.Cells(2, 7).Value)
    End If
End Sub

Public Sub DumpDecorRules()
    Dim ws As Worksheet, r As Long, last As Long
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("DecorConfig"): On Error GoTo 0
    If ws Is Nothing Then Debug.Print "[DecorConfig] not found": Exit Sub
    last = Application.Max(ws.Cells(ws.Rows.Count, 1).End(xlUp).row, ws.Cells(ws.Rows.Count, 7).End(xlUp).row)
    Debug.Print "---- DecorConfig (A:G) rows=", last, " ----"
    For r = 2 To last
        Debug.Print r, _
            "A=" & CStr(ws.Cells(r, 1).Value), _
            "|B=" & CStr(ws.Cells(r, 2).Value), _
            "|C=" & CStr(ws.Cells(r, 3).Value), _
            "|D=" & CStr(ws.Cells(r, 4).Value), _
            "|E=" & CStr(ws.Cells(r, 5).Value), _
            "|F=" & CStr(ws.Cells(r, 6).Value), _
            "|G=" & CStr(ws.Cells(r, 7).Value)
    Next r
    Debug.Print "----------------------------------"
End Sub

' アクティブシート上で TraceDecorAtRow を呼ぶラッパー
Public Sub TraceDecorAtRow_Active(Optional ByVal topRow As Long = 6)
    Dim macroName As String
    macroName = "'" & ThisWorkbook.Name & "'!" & ActiveSheet.CodeName & ".TraceDecorAtRow"
    Application.Run macroName, topRow
End Sub

'===========================================
' EnableEvents 救出ユーティリティ
' - 実行すると、EnableEvents を True に戻し、
'   念のため ScreenUpdating / Calculation も整える
' - 直後に軽くイベントを発火させて動作確認ログを出す
'===========================================
Public Sub Fix_EnableEvents(Optional ByVal alsoResetCalc As Boolean = True, _
                            Optional ByVal verbose As Boolean = True)
    On Error GoTo EH

    Dim oldEE As Boolean, oldSU As Boolean, oldCalc As XlCalculation
    oldEE = Application.EnableEvents
    oldSU = Application.ScreenUpdating
    oldCalc = Application.Calculation

    ' まず描画を止める
    Application.ScreenUpdating = False

    ' 計算モードを手動へ（任意）
    If alsoResetCalc Then Application.Calculation = xlCalculationManual

    ' イベントを確実に有効化
    Application.EnableEvents = True

    ' 軽く DoEvents を入れて反映
    DoEvents

    ' オプション：計算モードを元に戻す
    If alsoResetCalc Then Application.Calculation = oldCalc

    ' 画面更新を元に戻す
    Application.ScreenUpdating = oldSU

    If verbose Then
        Debug.Print "[Fix_EnableEvents] oldEE=", oldEE, " -> ", Application.EnableEvents, _
                    "  Calc=", Application.Calculation, "  ScreenUpdating=", Application.ScreenUpdating
        ' 任意：小さなダミー編集でイベントが動くか確認（アクティブセルに依存しない安全版）
        SafeNudgeActiveCell
    End If
    Exit Sub

EH:
    ' 万一の時でもイベントを有効にして抜ける
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If alsoResetCalc Then Application.Calculation = xlCalculationAutomatic
    If verbose Then Debug.Print "[Fix_EnableEvents][ERR]", Err.Number, Err.Description
End Sub

Public Sub Fix_EnableEvents_Run()
    ' ここで本体を呼ぶ（引数なし）
    Fix_EnableEvents
End Sub

'-------------------------------------------
' アクティブセルを安全に "つついて" Change を誘発（任意）
'-------------------------------------------
Private Sub SafeNudgeActiveCell()
    On Error Resume Next
    Dim c As Range, v As Variant
    Set c = ActiveCell
    If c Is Nothing Then Exit Sub
    v = c.Value
    c.Value = v ' 同値代入（多くのケースで Change は発火しないが、SelectionChange 等の確認用）
    Debug.Print "[Fix_EnableEvents] Nudge @", c.Address(0, 0), " EnableEvents=", Application.EnableEvents
End Sub
