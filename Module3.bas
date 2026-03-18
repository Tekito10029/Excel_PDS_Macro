Attribute VB_Name = "Module3"
'V_1.1 新規作成時ラベル削除実装
Option Explicit

Private Const ORIGINAL_BOOK_NAME As String = ""

Public Sub 原紙から新規ブック作成()

    On Error GoTo EH

    Dim srcWb As Workbook
    Dim saveFolder As String
    Dim baseName As String
    Dim originalBaseName As String
    Dim l1Value As String
    Dim newPath As String
    Dim ext As String
    Dim newWb As Workbook
    Dim targetWs As Worksheet
    Dim activeWs As Worksheet
    Dim confirmMsg As String

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

    saveFolder = srcWb.path
    ext = Mid$(srcWb.Name, InStrRev(srcWb.Name, "."))

    originalBaseName = Left$(srcWb.Name, InStrRev(srcWb.Name, ".") - 1)

    Set targetWs = GetVisibleSheetForL1(srcWb)
    If targetWs Is Nothing Then
        MsgBox "表示シートの L1 にファイル名として使用できる値がありません。", vbExclamation
        Exit Sub
    End If

    l1Value = Trim$(CStr(targetWs.Range("L1").Value))
    l1Value = NormalizeFileName(l1Value)

    If Len(l1Value) = 0 Then
        MsgBox "L1 の値がファイル名に使用できません。", vbExclamation
        Exit Sub
    End If

    Debug.Print "saveFolder=[" & saveFolder & "]"
    Debug.Print "originalBaseName=[" & originalBaseName & "]"
    Debug.Print "l1Value=[" & l1Value & "]"
    Debug.Print "ext=[" & ext & "]"
    
    Dim p As Long
    
    p = InStrRev(originalBaseName, "_")
    
    If p > 0 Then
        baseName = Left$(originalBaseName, p) & l1Value
    Else
        baseName = originalBaseName & "_" & l1Value
    End If
    
    Debug.Print "baseName=[" & baseName & "]"
    
    newPath = GetUniqueFilePath(saveFolder, baseName, ext)
    
    Debug.Print "newPath=[" & newPath & "]"

    confirmMsg = BuildConfirmMessage(activeWs, baseName, ext, saveFolder)

    If MsgBox( _
        confirmMsg, _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "新規ブック作成確認" _
    ) <> vbYes Then
        Exit Sub
    End If

    srcWb.SaveCopyAs newPath
    Set newWb = Workbooks.Open(newPath)
    
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

Private Function IsOriginalBook(ByVal wb As Workbook) As Boolean
    IsOriginalBook = (StrComp(wb.Name, ORIGINAL_BOOK_NAME, vbTextCompare) = 0)
End Function

Private Function GetVisibleSheetForL1(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim s As String

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            s = Trim$(CStr(ws.Range("L1").Value))
            If Len(s) > 0 Then
                Set GetVisibleSheetForL1 = ws
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function ReplaceLastUnderscoreSuffix(ByVal originalBaseName As String, ByVal newSuffix As String) As String
    Dim p As Long

    p = InStrRev(originalBaseName, "_")

    If p > 0 Then
        ReplaceLastUnderscoreSuffix = Left$(originalBaseName, p) & newSuffix
    Else
        ReplaceLastUnderscoreSuffix = originalBaseName & "_" & newSuffix
    End If
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

Private Function BuildConfirmMessage(ByVal ws As Worksheet, ByVal baseName As String, ByVal ext As String, ByVal saveFolder As String) As String
    Dim s As String

    s = "次の名前で新規ブックを作成します。" & vbCrLf & vbCrLf
    s = s & "ファイル名: " & baseName & ext & vbCrLf
    s = s & "保存先: " & saveFolder & vbCrLf & vbCrLf
    s = s & "初期化対象:" & vbCrLf
    s = s & GetInitTargetSummary(ws) & vbCrLf & vbCrLf
    s = s & "作成を続行しますか？"

    BuildConfirmMessage = s
End Function

Private Function GetInitTargetSummary(ByVal ws As Worksheet) As String
    Dim lines As String

    lines = ""

    If IsChecked(ws.Range("Z1").Value) Then
        lines = lines & "・製造票番号（L1）" & vbCrLf
    End If

    If IsChecked(ws.Range("Z2").Value) Then
        lines = lines & "・品名・サイズ・員数・単重（D6:L19）" & vbCrLf
    End If

    If IsChecked(ws.Range("Z3").Value) Then
        lines = lines & "・物件名（C20）" & vbCrLf
    End If

    If IsChecked(ws.Range("Z4").Value) Then
        lines = lines & "・入荷日・納期（K3:K4）" & vbCrLf
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

    ' Z1=TRUE のとき 製造票番号 を初期化
    If IsChecked(ws.Range("Z1").Value) Then
        SafeClearRange ws.Range("L1")
    End If

    ' Z2=TRUE のとき 品名・サイズ・員数・単重 を初期化
    If IsChecked(ws.Range("Z2").Value) Then
        SafeClearRange ws.Range("D6:L19")
    End If

    ' Z3=TRUE のとき追加範囲を初期化
    If IsChecked(ws.Range("Z3").Value) Then
        SafeClearRange ws.Range("C20")
    End If
    
    ' Z4=TRUE のとき入荷日・納期を初期化
    If IsChecked(ws.Range("Z4").Value) Then
        SafeClearRange ws.Range("K3:K4")
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

Public Sub 初期化チェックを一括切替()
    Dim ws As Worksheet
    Dim allOn As Boolean
    
    Set ws = ActiveSheet
    
    allOn = CBool(ws.Range("Z1").Value) _
         And CBool(ws.Range("Z2").Value) _
         And CBool(ws.Range("Z3").Value) _
         And CBool(ws.Range("Z4").Value)
    
    If allOn Then
        ws.Range("Z1:Z4").Value = False
    Else
        ws.Range("Z1:Z4").Value = True
    End If
End Sub

Public Sub 初期化チェックをすべてON()
    ActiveSheet.Range("Z1:Z4").Value = True
End Sub

Public Sub 初期化チェックをすべてOFF()
    ActiveSheet.Range("Z1:Z4").Value = False
End Sub

