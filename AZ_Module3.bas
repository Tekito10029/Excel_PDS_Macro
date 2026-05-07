Attribute VB_Name = "Module3"
'V_1.5 原紙を一本化 製造番号を自動採番し製造番号ブックに自動入力
Option Explicit

Private Const ORIGINAL_BOOK_NAME As String = "☆沓用AZ製造表_原紙.xlsm"
Private Const COMPANY_MASTER_SHEET As String = "会社設定"
Private Const COMPANY_NAME_CELL As String = "C3"

Private Const SETTINGS_SHEET_NAME As String = "設定"
Private Const LEDGER_PATH_CELL As String = "B1"
Private Const LEDGER_SHEET_CELL As String = "B2"



' === レイアウト可変対応（名前付き範囲優先 / 固定番地フォールバック） ===
Private Const NR_COMPANY_NAME As String = "原紙_会社名"
Private Const NR_K3_DATE As String = "原紙_基準日"
Private Const NR_L1_ORDERNO As String = "原紙_製造番号"
Private Const NR_DETAIL_BLOCK As String = "原紙_明細"
Private Const NR_PROJECT_NAME As String = "原紙_物件名"
Private Const NR_DELIVERY_NOTE As String = "原紙_備考"
Private Const NR_DATE_BLOCK As String = "原紙_日付帯"
Private Const NR_INIT_CHECK_1 As String = "原紙_初期化チェック1"
Private Const NR_INIT_CHECK_2 As String = "原紙_初期化チェック2"
Private Const NR_INIT_CHECK_3 As String = "原紙_初期化チェック3"
Private Const NR_INIT_CHECK_4 As String = "原紙_初期化チェック4"

Private Function ResolveNamedRange(ByVal ws As Worksheet, ByVal nm As String) As Range
    On Error Resume Next
    Set ResolveNamedRange = ws.Range(nm)
    If ResolveNamedRange Is Nothing Then Set ResolveNamedRange = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
End Function

Private Function ResolveLayoutCell(ByVal ws As Worksheet, ByVal nm As String, ByVal fallbackAddr As String) As Range
    Dim r As Range
    Set r = ResolveNamedRange(ws, nm)
    If r Is Nothing Then
        On Error Resume Next
        Set r = ws.Range(fallbackAddr)
        On Error GoTo 0
    End If
    If Not r Is Nothing Then Set ResolveLayoutCell = r.Cells(1, 1)
End Function

Private Function ResolveLayoutRange(ByVal ws As Worksheet, ByVal nm As String, ByVal fallbackAddr As String) As Range
    Dim r As Range
    Set r = ResolveNamedRange(ws, nm)
    If r Is Nothing Then
        On Error Resume Next
        Set r = ws.Range(fallbackAddr)
        On Error GoTo 0
    End If
    Set ResolveLayoutRange = r
End Function

Private Function CompanyNameCell(ByVal ws As Worksheet) As Range
    Set CompanyNameCell = ResolveLayoutCell(ws, NR_COMPANY_NAME, COMPANY_NAME_CELL)
End Function

Private Function K3DateCell(ByVal ws As Worksheet) As Range
    Set K3DateCell = ResolveLayoutCell(ws, NR_K3_DATE, "H5")
End Function

Private Function L1OrderNoCell(ByVal ws As Worksheet) As Range
    Set L1OrderNoCell = ResolveLayoutCell(ws, NR_L1_ORDERNO, "H2")
End Function

Private Function DetailBlockRange(ByVal ws As Worksheet) As Range
    Set DetailBlockRange = ResolveLayoutRange(ws, NR_DETAIL_BLOCK, "D7:H18")
End Function

Private Function ProjectNameCell(ByVal ws As Worksheet) As Range
    Set ProjectNameCell = ResolveLayoutCell(ws, NR_PROJECT_NAME, "C19")
End Function

Private Function DeliveryNoteCell(ByVal ws As Worksheet) As Range
    Set DeliveryNoteCell = ResolveLayoutCell(ws, NR_DELIVERY_NOTE, "E21")
End Function

Private Function DateBlockRange(ByVal ws As Worksheet) As Range
    Set DateBlockRange = ResolveLayoutRange(ws, NR_DATE_BLOCK, "H3:J5")
End Function

Private Function InitCheckCell(ByVal ws As Worksheet, ByVal indexNo As Long) As Range
    Select Case indexNo
        Case 1: Set InitCheckCell = ResolveLayoutCell(ws, NR_INIT_CHECK_1, "Z1")
        Case 2: Set InitCheckCell = ResolveLayoutCell(ws, NR_INIT_CHECK_2, "Z2")
        Case 3: Set InitCheckCell = ResolveLayoutCell(ws, NR_INIT_CHECK_3, "Z3")
        Case 4: Set InitCheckCell = ResolveLayoutCell(ws, NR_INIT_CHECK_4, "Z4")
    End Select
End Function

Private Function InitCheckValue(ByVal ws As Worksheet, ByVal indexNo As Long) As Boolean
    Dim c As Range
    Set c = InitCheckCell(ws, indexNo)
    If c Is Nothing Then Exit Function
    InitCheckValue = IsChecked(c.Value)
End Function

Private Sub SetInitChecks(ByVal ws As Worksheet, ByVal newValue As Boolean)
    Dim i As Long, c As Range
    For i = 1 To 4
        Set c = InitCheckCell(ws, i)
        If Not c Is Nothing Then c.Value = newValue
    Next i
End Sub
Public Sub 原紙から新規ブック作成()

    On Error GoTo EH

    Dim srcWb As Workbook
    Dim saveFolder As String
    Dim baseName As String
    Dim companyName As String
    Dim C3Value As String
    Dim l1Value As String
    Dim newPath As String
    Dim ext As String
    Dim newWb As Workbook
    Dim targetWs As Worksheet
    Dim activeWs As Worksheet
    Dim confirmMsg As String
    Dim k3Date As Date

    Set srcWb = ThisWorkbook

    If Not IsOriginalBook(srcWb) Then
        MsgBox "この機能は原紙ブックでのみ実行できます。", vbExclamation
        Exit Sub
    End If

    If Len(srcWb.path) = 0 Then
        MsgBox "先に原紙ブックを保存してください。", vbExclamation
        Exit Sub
    End If

    If srcWb.ActiveSheet Is Nothing Then
        MsgBox "原紙側のアクティブシートを取得できませんでした。", vbExclamation
        Exit Sub
    End If

    If Not TypeOf srcWb.ActiveSheet Is Worksheet Then
        MsgBox "原紙側のアクティブシートがワークシートではありません。", vbExclamation
        Exit Sub
    End If

    Set activeWs = srcWb.ActiveSheet

    ext = Mid$(srcWb.Name, InStrRev(srcWb.Name, "."))

    Set targetWs = GetVisibleSheetForFileName(srcWb)
    If targetWs Is Nothing Then
        MsgBox "表示シートの C3 が空です。", vbExclamation
        Exit Sub
    End If

    companyName = Trim$(CStr(CompanyNameCell(activeWs).Value))
    If Len(companyName) = 0 Then
        MsgBox "社名セル(" & COMPANY_NAME_CELL & ")が空です。", vbExclamation
        Exit Sub
    End If

    saveFolder = GetSaveFolderByCompany(companyName)

    If Len(saveFolder) = 0 Then
        MsgBox "会社設定シートに保存先が見つかりません。" & vbCrLf & _
               "社名: " & companyName, vbExclamation
        Exit Sub
    End If

    If Not FolderExistsSafe(saveFolder) Then
        MsgBox "保存先フォルダが存在しません。" & vbCrLf & _
               saveFolder, vbExclamation
        Exit Sub
    End If

    C3Value = Trim$(CStr(CompanyNameCell(targetWs).Value))
    C3Value = NormalizeFileName(C3Value)

    If Len(C3Value) = 0 Then
        MsgBox "C4 の値がファイル名に使用できません。", vbExclamation
        Exit Sub
    End If

    If Not IsDate(K3DateCell(targetWs).Value) Then
        MsgBox "H5 に有効な日付が入っていないため採番できません。", vbExclamation
        Exit Sub
    End If
    
    Dim issueBaseDate As Date
    
    k3Date = CDate(targetWs.Range("H5").Value)
    issueBaseDate = GetIssueBaseDateByOption(targetWs, k3Date)
    
    l1Value = GetNextOrderNoFromLedger(issueBaseDate)

    If Len(l1Value) = 0 Then
        MsgBox "次の番号を採番できませんでした。", vbExclamation
        Exit Sub
    End If

    L1OrderNoCell(targetWs).Value = l1Value

    Debug.Print "companyName=[" & companyName & "]"
    Debug.Print "saveFolder=[" & saveFolder & "]"
    Debug.Print "C3Value=[" & C3Value & "]"
    Debug.Print "l1Value=[" & l1Value & "]"
    Debug.Print "ext=[" & ext & "]"

    baseName = C3Value & "_" & l1Value

    Debug.Print "baseName=[" & baseName & "]"

    newPath = GetUniqueFilePath(saveFolder, baseName, ext)

    Debug.Print "newPath=[" & newPath & "]"

    confirmMsg = BuildConfirmMessage(activeWs, baseName, ext, saveFolder, companyName, l1Value)

    If MsgBox( _
        confirmMsg, _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "新規ブック作成確認" _
    ) <> vbYes Then
        Exit Sub
    End If

    srcWb.SaveCopyAs newPath
    Set newWb = Workbooks.Open(newPath)

    AppendToLedger l1Value, companyName

    srcWb.Activate
    activeWs.Activate
    DeleteAllBracesAndLabels_WithMirror_AndWriteTEXTdd

    ResetOriginalForm srcWb, activeWs

    MsgBox "新しいブックを作成しました。" & vbCrLf & newPath, vbInformation
    Exit Sub

EH:
    MsgBox "新規ブック作成中にエラーが発生しました。" & vbCrLf & _
           "No: " & Err.Number & vbCrLf & _
           Err.Description, vbExclamation
End Sub

Public Sub 原紙を手動初期化()

    On Error GoTo EH

    Dim srcWb As Workbook
    Dim ws As Worksheet
    Dim msg As String

    Set srcWb = ThisWorkbook

    If Not IsOriginalBook(srcWb) Then
        MsgBox "この機能は原紙ブックでのみ実行できます。", vbExclamation
        Exit Sub
    End If

    If srcWb.ActiveSheet Is Nothing Then
        MsgBox "アクティブシートを取得できませんでした。", vbExclamation
        Exit Sub
    End If

    If Not TypeOf srcWb.ActiveSheet Is Worksheet Then
        MsgBox "アクティブシートがワークシートではありません。", vbExclamation
        Exit Sub
    End If

    Set ws = srcWb.ActiveSheet

    msg = "現在のシートを手動で初期化します。" & vbCrLf & vbCrLf & _
          "対象シート: " & ws.Name & vbCrLf & vbCrLf & _
          "初期化対象:" & vbCrLf & _
          GetInitTargetSummary(ws) & vbCrLf & vbCrLf & _
          "実行しますか？"

    If MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2, "手動初期化確認") <> vbYes Then
        Exit Sub
    End If

    ResetOriginalForm srcWb, ws

    MsgBox "初期化が完了しました。", vbInformation
    Exit Sub

EH:
    MsgBox "手動初期化中にエラーが発生しました。" & vbCrLf & _
           "No: " & Err.Number & vbCrLf & _
           Err.Description, vbExclamation
End Sub

Private Function IsOriginalBook(ByVal wb As Workbook) As Boolean
    IsOriginalBook = (StrComp(wb.Name, ORIGINAL_BOOK_NAME, vbTextCompare) = 0)
End Function

Private Function GetVisibleSheetForFileName(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim C3Text As String

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            C3Text = Trim$(CStr(CompanyNameCell(ws).Value))
            If Len(C3Text) > 0 Then
                Set GetVisibleSheetForFileName = ws
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function GetSaveFolderByCompany(ByVal companyName As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim keyA As String
    Dim keyC As String
    Dim folderPath As String
    Dim targetName As String

    targetName = NormalizeCompanyName(companyName)
    If Len(targetName) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(COMPANY_MASTER_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Err.Raise vbObjectError + 1200, , "会社設定シートが見つかりません。シート名: " & COMPANY_MASTER_SHEET
    End If

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    For r = 2 To lastRow
        keyA = NormalizeCompanyName(ws.Cells(r, "A").Value)
        keyC = NormalizeCompanyName(ws.Cells(r, "C").Value)
        folderPath = Trim$(CStr(ws.Cells(r, "B").Value))

        If Len(folderPath) > 0 Then
            If keyA = targetName Or keyC = targetName Then
                GetSaveFolderByCompany = folderPath
                Exit Function
            End If
        End If
    Next r
End Function

Private Function NormalizeCompanyName(ByVal s As Variant) As String
    Dim t As String

    t = Trim$(CStr(s))
    If Len(t) = 0 Then
        NormalizeCompanyName = ""
        Exit Function
    End If

    t = Replace$(t, " ", "")
    t = Replace$(t, "　", "")
    t = Replace$(t, "株式会社", "")
    t = Replace$(t, "有限会社", "")
    t = Replace$(t, "㈱", "")
    t = Replace$(t, "(株)", "")
    t = Replace$(t, "（株）", "")

    NormalizeCompanyName = UCase$(t)
End Function

Private Function NormalizeFileName(ByVal s As String) As String
    Dim ngChars As Variant
    Dim i As Long

    s = Trim$(s)
    ngChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(ngChars) To UBound(ngChars)
        s = Replace$(s, CStr(ngChars(i)), "")
    Next i

    s = Replace$(s, vbCr, "")
    s = Replace$(s, vbLf, "")
    s = Replace$(s, vbTab, "")

    s = Trim$(s)
    Do While Len(s) > 0 And (Right$(s, 1) = "." Or Right$(s, 1) = " ")
        s = Left$(s, Len(s) - 1)
    Loop

    NormalizeFileName = s
End Function

Private Function GetUniqueFilePath(ByVal folderPath As String, ByVal baseName As String, ByVal ext As String) As String
    Dim path As String
    Dim i As Long
    Dim fso As Object

    folderPath = Trim$(folderPath)
    baseName = Trim$(baseName)
    ext = Trim$(ext)

    If Len(folderPath) = 0 Then
        Err.Raise vbObjectError + 1100, , "保存先フォルダが空です。"
    End If

    If Len(baseName) = 0 Then
        Err.Raise vbObjectError + 1101, , "ファイル名が空です。"
    End If

    If Len(ext) = 0 Then
        Err.Raise vbObjectError + 1102, , "拡張子が取得できません。"
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    path = folderPath & "\" & baseName & ext
    If Not fso.FileExists(path) Then
        GetUniqueFilePath = path
        Exit Function
    End If

    For i = 2 To 9999
        path = folderPath & "\" & baseName & "_" & CStr(i) & ext
        If Not fso.FileExists(path) Then
            GetUniqueFilePath = path
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 1000, , "保存可能なファイル名を作成できませんでした。"
End Function

Private Function FolderExistsSafe(ByVal folderPath As String) As Boolean
    Dim fso As Object

    If Len(Trim$(folderPath)) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExistsSafe = fso.FolderExists(folderPath)
End Function

Private Function BuildConfirmMessage(ByVal ws As Worksheet, ByVal baseName As String, ByVal ext As String, ByVal saveFolder As String, ByVal companyName As String, ByVal l1Value As String) As String
    Dim s As String

    s = "次の内容で新規ブックを作成します。" & vbCrLf & vbCrLf
    s = s & "社名: " & companyName & vbCrLf
    s = s & "番号: " & l1Value & vbCrLf
    s = s & "ファイル名: " & baseName & ext & vbCrLf
    s = s & "保存先: " & saveFolder & vbCrLf
    s = s & "台帳追記: A列=" & l1Value & " / C列=" & companyName & vbCrLf & vbCrLf
    s = s & "初期化対象:" & vbCrLf
    s = s & GetInitTargetSummary(ws) & vbCrLf & vbCrLf
    s = s & "作成を続行しますか？"

    BuildConfirmMessage = s
End Function

Private Function GetInitTargetSummary(ByVal ws As Worksheet) As String
    Dim lines As String

    lines = ""

    If InitCheckValue(ws, 1) Then
        lines = lines & "・製造票番号（H2）" & vbCrLf
    End If

    If InitCheckValue(ws, 2) Then
        lines = lines & "・品名・サイズ・員数・単価（D7:H18）" & vbCrLf
    End If

    If InitCheckValue(ws, 3) Then
        lines = lines & "・物件名・備考欄（C19:F20）" & vbCrLf
    End If

    If InitCheckValue(ws, 4) Then
        lines = lines & "・入荷日付・納期（H5:J5）" & vbCrLf
    End If

    If Len(lines) = 0 Then
        lines = "・なし" & vbCrLf
    End If

    If Right$(lines, 2) = vbCrLf Then
        lines = Left$(lines, Len(lines) - 2)
    End If

    GetInitTargetSummary = lines
End Function

Private Sub ResetOriginalForm(ByVal wb As Workbook, ByVal ws As Worksheet)
    On Error GoTo EH

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    '製造票番号
    If InitCheckValue(ws, 1) Then
        SafeClearRange L1OrderNoCell(ws)
    End If

    '品名・サイズ
    If InitCheckValue(ws, 2) Then
        SafeClearRange DetailBlockRange(ws)
    End If

    '物件名・備考欄
    If InitCheckValue(ws, 3) Then
        SafeClearRange ProjectNameCell(ws)
        SafeClearRange DeliveryNoteCell(ws)
    End If
    
    '入荷日・納期
    If InitCheckValue(ws, 4) Then
        SafeClearRange DateBlockRange(ws)
    End If

ExitProc:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "原紙の初期化中にエラーが発生しました。" & vbCrLf & _
           "No: " & Err.Number & vbCrLf & _
           Err.Description, vbExclamation
End Sub

Private Function IsChecked(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbBoolean
            IsChecked = CBool(v)
        Case vbString
            IsChecked = (UCase$(Trim$(CStr(v))) = "TRUE")
        Case vbInteger, vbLong, vbSingle, vbDouble, vbByte
            IsChecked = (CDbl(v) <> 0)
        Case Else
            IsChecked = False
    End Select
End Function

Private Sub SafeClearRange(ByVal targetRange As Range)
    Dim c As Range
    Dim doneDict As Object
    Dim key As String
    Dim ma As Range

    Set doneDict = CreateObject("Scripting.Dictionary")

    For Each c In targetRange.Cells
        If c.MergeCells Then
            Set ma = c.MergeArea
            key = ma.Worksheet.Name & "!" & ma.Address(False, False)

            If Not doneDict.Exists(key) Then
                ma.ClearContents
                doneDict.add key, True
            End If
        Else
            c.ClearContents
        End If
    Next c
End Sub

Private Function GetLedgerBookPath() As String
    Dim ws As Worksheet
    Dim p As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    p = Trim$(CStr(ws.Range(LEDGER_PATH_CELL).Value))
    p = Replace$(p, vbCr, "")
    p = Replace$(p, vbLf, "")
    p = Replace$(p, vbTab, "")

    GetLedgerBookPath = p
End Function

Private Function GetLedgerSheetName() As String
    Dim ws As Worksheet
    Dim s As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    s = Trim$(CStr(ws.Range(LEDGER_SHEET_CELL).Value))
    GetLedgerSheetName = s
End Function

Private Sub SaveLedgerBookPath(ByVal bookPath As String)
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    ws.Range(LEDGER_PATH_CELL).Value = bookPath
End Sub

Private Function PickLedgerBookPath() As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "台帳ブックを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.add "Excel Files", "*.xlsx; *.xlsm; *.xls"
        If .Show <> -1 Then Exit Function
        PickLedgerBookPath = .SelectedItems(1)
    End With
End Function

Private Sub AppendToLedger(ByVal orderNo As String, ByVal companyName As String)
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim alreadyOpen As Boolean
    Dim p As String
    Dim sheetName As String
    Dim tmpWb As Workbook
    Dim pickedPath As String

    If Len(Trim$(orderNo)) = 0 Then Exit Sub
    If Len(Trim$(companyName)) = 0 Then Exit Sub

    p = GetLedgerBookPath()
    sheetName = GetLedgerSheetName()

    If Len(sheetName) = 0 Then
        MsgBox "設定シートの台帳シート名が未設定です。" & vbCrLf & _
               SETTINGS_SHEET_NAME & "!" & LEDGER_SHEET_CELL, vbExclamation
        Exit Sub
    End If

    If Len(p) = 0 Then
        pickedPath = PickLedgerBookPath()
        If Len(pickedPath) = 0 Then
            MsgBox "台帳ブックが選択されなかったため、台帳追記を中止しました。", vbExclamation
            Exit Sub
        End If
        SaveLedgerBookPath pickedPath
        p = pickedPath
    End If

    alreadyOpen = False
    Set wb = Nothing

    For Each tmpWb In Application.Workbooks
        If StrComp(NormalizePath(tmpWb.FullName), NormalizePath(p), vbTextCompare) = 0 Then
            Set wb = tmpWb
            alreadyOpen = True
            Exit For
        End If
    Next tmpWb

    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=p, ReadOnly:=False)
        On Error GoTo EH

        If wb Is Nothing Then
            pickedPath = PickLedgerBookPath()
            If Len(pickedPath) = 0 Then
                MsgBox "台帳ブックを開けませんでした。" & vbCrLf & _
                       "保存済みパス: [" & p & "]", vbExclamation
                Exit Sub
            End If

            SaveLedgerBookPath pickedPath
            p = pickedPath

            On Error Resume Next
            Set wb = Workbooks.Open(Filename:=p, ReadOnly:=False)
            On Error GoTo EH

            If wb Is Nothing Then
                MsgBox "選択した台帳ブックも開けませんでした。" & vbCrLf & _
                       "Path=[" & p & "]", vbExclamation
                Exit Sub
            End If
        End If
    End If

    Set ws = wb.Worksheets(sheetName)

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then lastRow = 1
    
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, "A").Value)) = Trim$(orderNo) _
           And Trim$(CStr(ws.Cells(r, "C").Value)) = Trim$(companyName) Then
            wb.Save
            If Not alreadyOpen Then wb.Close SaveChanges:=False
            Exit Sub
        End If
    Next r
    
    nextRow = GetNextLedgerRow(ws, orderNo)
    If nextRow < 2 Then nextRow = 2
    
    ws.Cells(nextRow, "A").Value = orderNo
    ws.Cells(nextRow, "C").Value = companyName

    wb.Save

    If Not alreadyOpen Then
        wb.Close SaveChanges:=False
    End If

    Exit Sub

EH:
    MsgBox "台帳への追記中にエラーが発生しました。" & vbCrLf & _
           "No: " & Err.Number & vbCrLf & _
           "Path=[" & p & "]" & vbCrLf & _
           "Sheet=[" & sheetName & "]" & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation
End Sub

Private Function GetNextLedgerRow(ByVal ws As Worksheet, ByVal newOrderNo As String) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim insertRow As Long
    Dim newPrefix As String
    Dim v As String
    Dim prefix As String
    Dim prevOrderNo As String
    Dim prevPrefix As String

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    If lastRow < 2 Then
        GetNextLedgerRow = 2
        Exit Function
    End If

    newOrderNo = NormalizeOrderNo(newOrderNo)
    If Not IsValidOrderNo(newOrderNo) Then
        GetNextLedgerRow = lastRow + 1
        Exit Function
    End If

    newPrefix = Left$(newOrderNo, 4)

    '既に来月以降の番号が台帳内にある場合は、
    '表の最終行ではなく「同じ年月ブロックの末尾」に差し込む。
    '
    '例:
    '  260400401
    '  （空行）
    '  260500401
    '
    'ここへ 260400402 を追加する場合は、260500401 の前に差し込む。
    For r = 2 To lastRow
        v = NormalizeOrderNo(ws.Cells(r, "A").Value)

        If IsValidOrderNo(v) Then
            prefix = Left$(v, 4)

            If prefix = newPrefix Then
                insertRow = r + 1

            ElseIf CLng(prefix) > CLng(newPrefix) Then
                If insertRow = 0 Then insertRow = r
                Exit For
            End If
        End If
    Next r

    If insertRow > 0 And insertRow <= lastRow Then
        If IsValidOrderNo(NormalizeOrderNo(ws.Cells(insertRow, "A").Value)) Then
            '差し込み先がすでに次月以降の番号行の場合は、
            '新規行＋月区切り用の空行を作ってから書き込む。
            ws.Rows(insertRow).Insert Shift:=xlDown
            ws.Rows(insertRow + 1).Insert Shift:=xlDown
        Else
            '差し込み先が既存の空行の場合は、空行を残すために1行挿入する。
            ws.Rows(insertRow).Insert Shift:=xlDown
        End If

        GetNextLedgerRow = insertRow
        Exit Function
    End If

    'まだ来月以降の番号がない場合は従来通り末尾に追加する。
    prevOrderNo = GetLastOrderNoFromColumnA(ws, lastRow)

    If IsValidOrderNo(prevOrderNo) Then
        prevPrefix = Left$(prevOrderNo, 4)

        If prevPrefix <> newPrefix Then
            GetNextLedgerRow = lastRow + 2
            Exit Function
        End If
    End If

    GetNextLedgerRow = lastRow + 1
End Function

Private Function GetNextOrderNoFromLedger(ByVal baseDate As Date) As String
    On Error GoTo EH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim p As String
    Dim sheetName As String
    Dim tmpWb As Workbook
    Dim pickedPath As String
    Dim alreadyOpen As Boolean
    Dim lastRow As Long
    Dim lastNo As String
    Dim currentPrefix As String
    Dim nextNo As String

    p = GetLedgerBookPath()
    sheetName = GetLedgerSheetName()

    If Len(sheetName) = 0 Then
        MsgBox "設定シートの台帳シート名が未設定です。" & vbCrLf & _
               SETTINGS_SHEET_NAME & "!" & LEDGER_SHEET_CELL, vbExclamation
        Exit Function
    End If

    If Len(p) = 0 Then
        pickedPath = PickLedgerBookPath()
        If Len(pickedPath) = 0 Then Exit Function
        SaveLedgerBookPath pickedPath
        p = pickedPath
    End If

    alreadyOpen = False
    Set wb = Nothing

    For Each tmpWb In Application.Workbooks
        If StrComp(NormalizePath(tmpWb.FullName), NormalizePath(p), vbTextCompare) = 0 Then
            Set wb = tmpWb
            alreadyOpen = True
            Exit For
        End If
    Next tmpWb

    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=p, ReadOnly:=False)
        On Error GoTo EH

        If wb Is Nothing Then
            pickedPath = PickLedgerBookPath()
            If Len(pickedPath) = 0 Then Exit Function

            SaveLedgerBookPath pickedPath
            p = pickedPath

            On Error Resume Next
            Set wb = Workbooks.Open(Filename:=p, ReadOnly:=False)
            On Error GoTo EH

            If wb Is Nothing Then Exit Function
        End If
    End If

    Set ws = wb.Worksheets(sheetName)
    
    Dim maxNo As String
    
    lastRow = GetLastUsedRowInColumnA(ws)
    currentPrefix = Format$(baseDate, "yymm")
    
    maxNo = GetMaxOrderNoByPrefix(ws, currentPrefix, lastRow)
    nextNo = BuildNextOrderNo(maxNo, currentPrefix)
    
    GetNextOrderNoFromLedger = nextNo

    If Not alreadyOpen Then
        wb.Close SaveChanges:=False
    End If

    Exit Function

EH:
    MsgBox "次番号の採番中にエラーが発生しました。" & vbCrLf & _
           "No: " & Err.Number & vbCrLf & _
           Err.Description, vbExclamation
End Function

Private Function GetMaxOrderNoByPrefix(ByVal ws As Worksheet, ByVal targetPrefix As String, ByVal lastRow As Long) As String
    Dim r As Long
    Dim v As String
    Dim maxSeq As Long
    Dim seq As Long

    maxSeq = 0

    For r = 2 To lastRow
        v = NormalizeOrderNo(ws.Cells(r, "A").Value)

        If IsValidOrderNo(v) Then
            If Left$(v, 4) = targetPrefix Then
                seq = CLng(Right$(v, 5))
                If seq > maxSeq Then
                    maxSeq = seq
                    GetMaxOrderNoByPrefix = v
                End If
            End If
        End If
    Next r
End Function

Private Function GetIssueBaseDateByOption(ByVal ws As Worksheet, ByVal srcDate As Date) As Date
    If IsChecked(ws.Range("Z5").Value) Then
        GetIssueBaseDateByOption = DateSerial(Year(srcDate), Month(srcDate) + 1, 1)
    Else
        GetIssueBaseDateByOption = srcDate
    End If
End Function

Private Function GetLastUsedRowInColumnA(ByVal ws As Worksheet) As Long
    GetLastUsedRowInColumnA = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If GetLastUsedRowInColumnA < 1 Then GetLastUsedRowInColumnA = 1
End Function

Private Function GetLastOrderNoFromColumnA(ByVal ws As Worksheet, ByVal startRow As Long) As String
    Dim r As Long
    Dim v As String

    For r = startRow To 2 Step -1
        v = NormalizeOrderNo(ws.Cells(r, "A").Value)
        If IsValidOrderNo(v) Then
            GetLastOrderNoFromColumnA = v
            Exit Function
        End If
    Next r
End Function

Private Function NormalizeOrderNo(ByVal v As Variant) As String
    Dim s As String

    s = Trim$(CStr(v))
    s = Replace$(s, " ", "")
    s = Replace$(s, "　", "")
    s = Replace$(s, "-", "")
    s = Replace$(s, vbCr, "")
    s = Replace$(s, vbLf, "")
    s = Replace$(s, vbTab, "")

    NormalizeOrderNo = s
End Function

Private Function IsValidOrderNo(ByVal s As String) As Boolean
    If Len(s) <> 9 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    IsValidOrderNo = True
End Function

Private Function BuildNextOrderNo(ByVal lastNo As String, ByVal currentPrefix As String) As String
    Dim lastPrefix As String
    Dim lastSeq As Long
    Dim nextSeq As Long

    If Len(currentPrefix) <> 4 Or Not IsNumeric(currentPrefix) Then
        Err.Raise vbObjectError + 2000, , "現在年月の接頭辞が不正です。"
    End If

    If IsValidOrderNo(lastNo) Then
        lastPrefix = Left$(lastNo, 4)

        If lastPrefix = currentPrefix Then
            lastSeq = CLng(Right$(lastNo, 5))
            nextSeq = lastSeq + 1
            BuildNextOrderNo = currentPrefix & Format$(nextSeq, "00000")
            Exit Function
        End If
    End If

    BuildNextOrderNo = currentPrefix & "50001"
End Function

Private Function NormalizePath(ByVal s As String) As String
    Dim t As String

    t = Trim$(s)
    t = Replace$(t, "/", "\")
    t = Replace$(t, vbCr, "")
    t = Replace$(t, vbLf, "")
    t = Replace$(t, vbTab, "")

    Do While Len(t) > 0 And Right$(t, 1) = "\"
        t = Left$(t, Len(t) - 1)
    Loop

    NormalizePath = t
End Function

Public Sub 初期化チェックを一括切替()
    Dim ws As Worksheet
    Dim allOn As Boolean

    Set ws = ActiveSheet

    allOn = InitCheckValue(ws, 1) _
         And InitCheckValue(ws, 2) _
         And InitCheckValue(ws, 3) _
         And InitCheckValue(ws, 4)

    If allOn Then
        SetInitChecks ws, False
    Else
        SetInitChecks ws, True
    End If
End Sub

Public Sub 初期化チェックをすべてON()
    SetInitChecks ActiveSheet, True
End Sub

Public Sub 初期化チェックをすべてOFF()
    SetInitChecks ActiveSheet, False
End Sub

