Sub 查找公式单元格()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    Dim resultMsg As String
    Dim countFormula As Long

    ' 提高稳定性
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set ws = ActiveSheet
    Set rng = ws.UsedRange

    resultMsg = ""
    countFormula = 0

    ' 遍历单元格
    For Each cell In rng

        ' 判断是否为公式
        If cell.HasFormula Then

            countFormula = countFormula + 1

            resultMsg = resultMsg & _
                cell.Address(False, False) & _
                " ： " & cell.Formula & vbCrLf

        End If

    Next cell

    ' 输出结果
    If countFormula > 0 Then

        MsgBox _
            "发现公式单元格：" & vbCrLf & vbCrLf & _
            resultMsg & vbCrLf & _
            "公式单元格总数量：" & countFormula, _
            vbInformation, _
            "检查结果"

    Else

        MsgBox _
            "未发现公式单元格。", _
            vbInformation, _
            "检查结果"

    End If

CleanExit:

    ' 清理资源
    Set cell = Nothing
    Set rng = Nothing
    Set ws = Nothing

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.CutCopyMode = False

    Exit Sub

ErrHandler:

    MsgBox _
        "运行出错：" & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical, _
        "宏运行失败"

    Resume CleanExit

End Sub

