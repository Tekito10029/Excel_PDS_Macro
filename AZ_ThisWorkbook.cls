VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforePrint(Cancel As Boolean)

    Dim ws As Worksheet
    Dim orderNo As String

    On Error GoTo EH

    ' 原紙ファイルそのものは印刷禁止
    If IsOriginalTemplateBook() Then
        Cancel = True
        MsgBox _
            "このブックは原紙ファイルのため印刷できません。" & vbCrLf & vbCrLf & _
            "先に「原紙から新規ブック作成」を実行し、" & vbCrLf & _
            "製造票番号が採番された新しいブックを印刷してください。", _
            vbExclamation, _
            "原紙の印刷禁止"
        Exit Sub
    End If

    ' 念のため、製造番号が空欄のブックも印刷禁止
    If TypeOf ActiveSheet Is Worksheet Then
        Set ws = ActiveSheet

        orderNo = Trim$(CStr(ws.Range("H2").MergeArea.Cells(1, 1).Value))

        If Len(orderNo) = 0 Then
            Cancel = True
            MsgBox _
                "製造票番号が空欄のため印刷できません。" & vbCrLf & vbCrLf & _
                "原紙を直接印刷している可能性があります。" & vbCrLf & _
                "先に「原紙から新規ブック作成」を実行してください。", _
                vbExclamation, _
                "製造票番号未採番"
            Exit Sub
        End If
    End If

    Exit Sub

EH:
    Cancel = True
    MsgBox _
        "印刷前チェック中にエラーが発生したため、印刷を中止しました。" & vbCrLf & _
        "No: " & Err.Number & vbCrLf & _
        Err.Description, _
        vbExclamation, _
        "印刷中止"

End Sub

Private Function IsOriginalTemplateBook() As Boolean

    Select Case LCase$(ThisWorkbook.Name)
        Case LCase$("☆沓用製造表_原紙自動入力.xlsm"), _
             LCase$("☆沓用AZ製造表_原紙.xlsm")
            IsOriginalTemplateBook = True
    End Select

End Function

