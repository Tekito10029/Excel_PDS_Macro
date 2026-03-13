Attribute VB_Name = "Module3"
Option Explicit

Private Const ORIGINAL_BOOK_NAME As String = "原紙自動入力.xlsm"

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

    ' 保存前に、原紙側でアクティブなシートを保持しておく
    Set activeWs = srcWb.ActiveSheet

    saveFolder = srcWb.path
    ext = Mid$(srcWb.Name, InStrRev(srcWb.Name, "."))

    originalBaseName = Left$(srcWb.Name, InStrRev(srcWb.Name, ".") - 1)

    ' ファイル名用の L1 は表示シートから取得
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

    baseName = ReplaceLastUnderscoreSuffix(originalBaseName, l1Value)
    newPath = GetUniqueFilePath(saveFolder, baseName, ext)

    srcWb.SaveCopyAs newPath
    Set newWb = Workbooks.Open(newPath)

    ' 原紙側で元々アクティブだったシートのみ初期化
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

    path = folderPath & "\" & baseName & ext
    If Dir$(path) = "" Then
        GetUniqueFilePath = path
        Exit Function
    End If

    For i = 2 To 9999
        path = folderPath & "\" & baseName & "_" & CStr(i) & ext
        If Dir$(path) = "" Then
            GetUniqueFilePath = path
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 1000, , "保存可能なファイル名を作成できませんでした。"
End Function

Private Sub ResetOriginalForm(ByVal wb As Workbook, ByVal ws As Worksheet)
    On Error GoTo EH

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    SafeClearRange ws.Range("L1")
    SafeClearRange ws.Range("K3:K4")
    SafeClearRange ws.Range("D6:L19")
    SafeClearRange ws.Range("C20")
    SafeClearRange ws.Range("F21")
    
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

