Sub 校验表头()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    '========================
    ' 预期数据
    '========================
    Dim checkArr As Variant

    checkArr = Array( _
        Array("A1", "序号"), _
        Array("B1", "姓名"), _
        Array("Z1", "性别") _
    )

    Dim i As Integer

    Dim cellAddr As String
    Dim expectedValue As String
    Dim actualValue As String

    Dim errorMsg As String
    Dim successMsg As String

    Dim errorCount As Integer
    Dim successCount As Integer

    '========================
    ' 开始校验
    '========================
    For i = LBound(checkArr) To UBound(checkArr)

        cellAddr = checkArr(i)(0)
        expectedValue = checkArr(i)(1)

        actualValue = Trim(ws.Range(cellAddr).Value)

        If actualValue = expectedValue Then

            successMsg = successMsg & _
                cellAddr & "，" & actualValue & vbCrLf

            successCount = successCount + 1

        Else

            errorMsg = errorMsg & _
                cellAddr & "，实际值：" & actualValue & _
                "，应为：" & expectedValue & vbCrLf

            errorCount = errorCount + 1

        End If

    Next i

    '========================
    ' 结果弹窗
    '========================
    Dim resultMsg As String

    resultMsg = "错误数据：" & vbCrLf

    If errorMsg = "" Then
        resultMsg = resultMsg & "无" & vbCrLf
    Else
        resultMsg = resultMsg & errorMsg
    End If

    resultMsg = resultMsg & vbCrLf & _
                "正确数据：" & vbCrLf

    If successMsg = "" Then
        resultMsg = resultMsg & "无" & vbCrLf
    Else
        resultMsg = resultMsg & successMsg
    End If

    resultMsg = resultMsg & vbCrLf & _
                "错误数据 " & errorCount & " 个，" & _
                "正确数据 " & successCount & " 个"

    MsgBox resultMsg, vbInformation, "校验结果"

End Sub



