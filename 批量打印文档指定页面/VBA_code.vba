#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub WPS_SumatraPDF_批量指定页打印()

    On Error GoTo ErrHandler

    Dim wordApp As Object
    Dim wordDoc As Object

    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String

    Dim pdfFolder As String
    Dim pdfPath As String

    Dim successCount As Long
    Dim failCount As Long
    Dim failList As String

    Dim sumatraPath As String
    Dim cmd As String

    '========================
    ' SumatraPDF路径
    '========================
    sumatraPath = "C:\Program Files\SumatraPDF\SumatraPDF.exe"

    If Dir(sumatraPath) = "" Then
        MsgBox "未找到 SumatraPDF，请检查安装路径！", vbCritical
        Exit Sub
    End If

    '========================
    ' 选择文件夹
    '========================
    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "请选择Word文档文件夹"

        If .Show <> -1 Then
            Exit Sub
        End If

        folderPath = .SelectedItems(1)

    End With

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    '========================
    ' PDF缓存目录
    '========================
    pdfFolder = folderPath & "PDF_TEMP\"

    If Dir(pdfFolder, vbDirectory) = "" Then
        MkDir pdfFolder
    End If

    '========================
    ' 启动WPS/Word
    '========================
    Set wordApp = CreateObject("Word.Application")

    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    successCount = 0
    failCount = 0
    failList = ""

    '========================
    ' 遍历文档
    '========================
    fileName = Dir(folderPath & "*.doc*")

    Do While fileName <> ""

        On Error Resume Next

        fullPath = folderPath & fileName

        '========================
        ' 打开文档
        '========================
        Set wordDoc = wordApp.Documents.Open( _
            FileName:=fullPath, _
            ReadOnly:=True, _
            AddToRecentFiles:=False)

        If Err.Number <> 0 Then

            failCount = failCount + 1
            failList = failList & vbCrLf & fileName

            Err.Clear

            GoTo NextFile

        End If

        '========================
        ' PDF路径
        '========================
        pdfPath = pdfFolder & fileName

        pdfPath = Replace(pdfPath, ".docx", ".pdf")
        pdfPath = Replace(pdfPath, ".doc", ".pdf")

        '========================
        ' 导出PDF
        '========================
        wordDoc.ExportAsFixedFormat _
            OutputFileName:=pdfPath, _
            ExportFormat:=17

        If Err.Number <> 0 Then

            failCount = failCount + 1
            failList = failList & vbCrLf & fileName

            Err.Clear

            wordDoc.Close False
            Set wordDoc = Nothing

            GoTo NextFile

        End If

        '========================
        ' 关闭文档
        '========================
        wordDoc.Close False
        Set wordDoc = Nothing

        '========================
        ' 等待PDF生成
        '========================
        Sleep 3000

        '========================
        ' SumatraPDF打印指定页
        '========================
        cmd = """" & sumatraPath & """" & _
              " -print-to-default" & _
              " -silent" & _
              " -print-settings ""2,4-5""" & _
              " """ & pdfPath & """"

        Shell cmd, vbHide

        '========================
        ' 等待打印提交
        '========================
        Sleep 5000

        successCount = successCount + 1

NextFile:

        fileName = Dir()

        On Error GoTo ErrHandler

    Loop

CleanExit:

    On Error Resume Next

    If Not wordDoc Is Nothing Then
        wordDoc.Close False
        Set wordDoc = Nothing
    End If

    If Not wordApp Is Nothing Then
        wordApp.Quit
        Set wordApp = Nothing
    End If

    MsgBox _
        "运行完成！" & vbCrLf & vbCrLf & _
        "已完成：" & successCount & "个" & vbCrLf & _
        "失败：" & failCount & "个" & vbCrLf & vbCrLf & _
        "失败名称：" & failList, _
        vbInformation

    Exit Sub

ErrHandler:

    failCount = failCount + 1

    If fileName <> "" Then
        failList = failList & vbCrLf & fileName
    End If

    Resume Next

End Sub