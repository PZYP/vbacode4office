Sub 批量写入成绩_书签版()

    On Error GoTo ErrHandler

    Dim wordApp As Object
    Dim wordDoc As Object

    Dim ws As Worksheet

    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String

    Dim studentName As String

    Dim 语文成绩 As String
    Dim 数学成绩 As String
    Dim 英语成绩 As String

    Dim lastRow As Long
    Dim i As Long
    Dim foundRow As Long

    Dim successCount As Long
    Dim failCount As Long

    ' =========================
    ' Excel工作表
    ' =========================
    Set ws = ThisWorkbook.Sheets(1)

    ' =========================
    ' Word文件夹路径结尾带\
    ' =========================
    folderPath = "C:\Users\Administrator\Desktop\成绩文件夹\"

    ' =========================
    ' 获取最后一行
    ' =========================
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' =========================
    ' 启动Word
    ' =========================
    Set wordApp = CreateObject("Word.Application")

    wordApp.Visible = False
    wordApp.DisplayAlerts = False

    ' =========================
    ' 遍历Word文件
    ' =========================
    fileName = Dir(folderPath & "*.doc*")

    Do While fileName <> ""

        On Error Resume Next

        foundRow = 0

        fullPath = folderPath & fileName

        ' =========================
        ' 从文件名提取姓名
        ' =========================
        studentName = 提取姓名(fileName)

        ' =========================
        ' Excel查找姓名 姓名列(i,数据列)
        ' =========================
        For i = 2 To lastRow

            If Trim(ws.Cells(i, 1).Value) = Trim(studentName) Then

                foundRow = i
                Exit For

            End If

        Next i

        ' =========================
        ' 找到数据 各个成绩所在的列用数字表示
        ' =========================
        If foundRow > 0 Then

            语文成绩 = ws.Cells(foundRow, 3).Value
            数学成绩 = ws.Cells(foundRow, 4).Value
            英语成绩 = ws.Cells(foundRow, 5).Value

            ' =========================
            ' 打开Word
            ' =========================
            Set wordDoc = wordApp.Documents.Open( _
                fileName:=fullPath, _
                ReadOnly:=False, _
                AddToRecentFiles:=False _
            )

            ' =========================
            ' 写入书签
            ' =========================
            Call 写入书签(wordDoc, "语文成绩", 语文成绩)
            Call 写入书签(wordDoc, "数学成绩", 数学成绩)
            Call 写入书签(wordDoc, "英语成绩", 英语成绩)

            ' =========================
            ' 保存
            ' =========================
            wordDoc.Save

            ' =========================
            ' 关闭Word
            ' =========================
            wordDoc.Close SaveChanges:=False

            Set wordDoc = Nothing

            successCount = successCount + 1

        Else

            failCount = failCount + 1

            MsgBox "Excel中未找到：" & studentName, vbExclamation

        End If

        fileName = Dir()

        On Error GoTo ErrHandler

    Loop

    GoTo SafeExit

' =========================
' 错误处理
' =========================
ErrHandler:

    failCount = failCount + 1

    MsgBox _
        "处理出错：" & vbCrLf & _
        "文件：" & fileName & vbCrLf & _
        "错误：" & Err.Description, _
        vbCritical

' =========================
' 安全退出
' =========================
SafeExit:

    On Error Resume Next

    If Not wordDoc Is Nothing Then

        wordDoc.Close SaveChanges:=False
        Set wordDoc = Nothing

    End If

    If Not wordApp Is Nothing Then

        wordApp.Quit
        Set wordApp = Nothing

    End If

    Set ws = Nothing

    Application.CutCopyMode = False

    MsgBox _
        "处理完成！" & vbCrLf & vbCrLf & _
        "成功：" & successCount & " 个" & vbCrLf & _
        "失败：" & failCount & " 个", _
        vbInformation

End Sub


' =========================
' 写入Word书签
' =========================
Sub 写入书签(doc As Object, 书签名 As String, 内容 As String)

    If doc.Bookmarks.Exists(书签名) Then

        Dim rng As Object

        Set rng = doc.Bookmarks(书签名).Range

        rng.Text = 内容

        ' 重新添加书签
        doc.Bookmarks.Add 书签名, rng

    Else

        MsgBox "未找到书签：" & 书签名, vbExclamation

    End If

End Sub


' =========================
' 提取文件名中的姓名
' =========================
Function 提取姓名(fileName As String) As String

    Dim tempName As String
    Dim arr As Variant

    ' =========================
    ' 去掉扩展名
    ' =========================
    tempName = Replace(fileName, ".docx", "")
    tempName = Replace(tempName, ".doc", "")

    ' =========================
    ' 处理：
    ' 1_张三
    ' =========================
    If InStr(tempName, "_") > 0 Then

        arr = Split(tempName, "_")

        If UBound(arr) >= 1 Then

            If IsNumeric(arr(0)) Then

                提取姓名 = Trim(CStr(arr(1)))

            Else

                提取姓名 = Trim(CStr(arr(0)))

            End If

            Exit Function

        End If

    End If

    ' =========================
    ' 处理：
    ' 1-张三
    ' =========================
    If InStr(tempName, "-") > 0 Then

        arr = Split(tempName, "-")

        If UBound(arr) >= 1 Then

            If IsNumeric(arr(0)) Then

                提取姓名 = Trim(CStr(arr(1)))

            Else

                提取姓名 = Trim(CStr(arr(0)))

            End If

            Exit Function

        End If

    End If

    ' =========================
    ' 默认直接返回
    ' =========================
    提取姓名 = Trim(tempName)

End Function

