Sub 批量生成个人档案()

    On Error GoTo 错误处理
    
    Dim excelApp As Object
    Dim excelWB As Object
    Dim excelWS As Object

    Dim wordDoc As Document

    Dim templatePath As String
    Dim excelPath As String
    Dim savePath As String

    Dim lastRow As Long
    Dim i As Long

    Dim 序号 As String
    Dim 姓名 As String
    Dim 性别 As String

    '========================
    ' 文件路径
    '========================

    templatePath = "C:\Users\Administrator\Desktop\新建文件夹\档案模板.doc"
    excelPath = "C:\Users\Administrator\Desktop\新建文件夹\姓名档案信息.xlsm"
    savePath = "C:\Users\Administrator\Desktop\新建文件夹\档案结果\"

    '========================
    ' 打开Excel
    '========================

    Set excelApp = CreateObject("Excel.Application")

    excelApp.Visible = False

    Set excelWB = excelApp.Workbooks.Open(excelPath)

    Set excelWS = excelWB.Sheets(1)

    ' 获取最后一行
    lastRow = excelWS.Cells(excelWS.Rows.Count, 1).End(-4162).Row

    '========================
    ' 开始循环
    '========================

    For i = 2 To lastRow
    
        ' 读取数据
        序号 = excelWS.Cells(i, 1).Value
        姓名 = excelWS.Cells(i, 2).Value
        性别 = excelWS.Cells(i, 6).Value

        '========================
        ' 重新打开模板
        ' 注意：ReadOnly:=True
        '========================

        Set wordDoc = Documents.Add(templatePath)
        
        '========================
        ' 替换占位符
        '========================

        Call 替换文本(wordDoc, "{{姓名}}", 姓名)
        Call 替换文本(wordDoc, "{{性别}}", 性别)
        Call 替换文本(wordDoc, "{{出生年月}}", 出生年月)

        '========================
        ' 保存新文件
        '========================
        
        文件名 = 序号 & "_" & 姓名
        wordDoc.SaveAs2 savePath & 文件名 & ".docx"

        '========================
        ' 关闭文档
        ' False = 不保存模板修改
        '========================

        wordDoc.Close SaveChanges:=False

        Set wordDoc = Nothing

    Next i

    '========================
    ' 关闭 Excel
    '========================
正常结束:

    If Not excelWB Is Nothing Then
        excelWB.Close SaveChanges:=False
    End If
    
    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If

    Set excelWS = Nothing
    Set excelWB = Nothing
    Set excelApp = Nothing
    
    DoEvents

    MsgBox "全部文档生成完成！"
    
    Exit Sub
    
错误处理:

    MsgBox "发生错误：" & Err.Description
    Resume 正常结束

End Sub


'========================
' 查找替换
'========================

Sub 替换文本(doc As Document, 查找内容 As String, 替换内容 As String)

    With doc.Content.Find

        .ClearFormatting
        .Replacement.ClearFormatting

        .Text = 查找内容
        .Replacement.Text = 替换内容

        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False

        .Execute Replace:=wdReplaceAll

    End With

End Sub
