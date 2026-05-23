Sub 批量插入照片到Word文档_书签版()

    On Error GoTo ErrHandler

    Dim wordApp As Object
    Dim wordDoc As Object

    Dim photoFolder As String
    Dim docFolder As String

    Dim fso As Object
    Dim file As Object
    Dim photoFile As Object

    Dim docPath As String
    Dim photoPath As String

    Dim docName As String
    Dim pureDocName As String
    Dim purePhotoName As String

    Dim successCount As Long
    Dim failCount As Long

    Dim failList As String

    Dim matched As Boolean

    '========================
    ' 选择照片文件夹
    '========================
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "请选择照片文件夹"
        If .Show <> -1 Then Exit Sub
        photoFolder = .SelectedItems(1)
    End With

    '========================
    ' 选择Word文档文件夹
    '========================
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "请选择Word文档文件夹"
        If .Show <> -1 Then Exit Sub
        docFolder = .SelectedItems(1)
    End With

    If Right(photoFolder, 1) <> "\" Then photoFolder = photoFolder & "\"
    If Right(docFolder, 1) <> "\" Then docFolder = docFolder & "\"

    Set fso = CreateObject("Scripting.FileSystemObject")

    '========================
    ' 启动Word
    '========================
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    successCount = 0
    failCount = 0
    failList = ""

    '========================
    ' 遍历Word文档
    '========================
    For Each file In fso.GetFolder(docFolder).Files

        If LCase(fso.GetExtensionName(file.Name)) = "doc" _
        Or LCase(fso.GetExtensionName(file.Name)) = "docx" Then

            docPath = file.Path
            docName = fso.GetBaseName(file.Name)

            ' 提取纯姓名
            pureDocName = 提取姓名(docName)

            matched = False
            photoPath = ""

            '========================
            ' 查找对应照片
            '========================
            For Each photoFile In fso.GetFolder(photoFolder).Files

                If LCase(fso.GetExtensionName(photoFile.Name)) = "jpg" _
                Or LCase(fso.GetExtensionName(photoFile.Name)) = "jpeg" _
                Or LCase(fso.GetExtensionName(photoFile.Name)) = "png" Then

                    purePhotoName = 提取姓名(fso.GetBaseName(photoFile.Name))

                    If pureDocName = purePhotoName Then

                        matched = True
                        photoPath = photoFile.Path
                        Exit For

                    End If

                End If

            Next photoFile

            '========================
            ' 插入照片
            '========================
            If matched = True Then

                On Error Resume Next

                Set wordDoc = wordApp.Documents.Open( _
                    FileName:=docPath, _
                    ReadOnly:=False, _
                    AddToRecentFiles:=False)

                If Err.Number <> 0 Then

                    failCount = failCount + 1
                    failList = failList & file.Name & "（文档打开失败）" & vbCrLf
                    Err.Clear

                Else

                    If wordDoc.Bookmarks.Exists("个人照片") Then

                        Dim rng As Object

                        Set rng = wordDoc.Bookmarks("个人照片").Range

                        rng.Text = ""

                        rng.InlineShapes.AddPicture _
                            FileName:=photoPath, _
                            LinkToFile:=False, _
                            SaveWithDocument:=True, _
                            Range:=rng

                        ' 重新添加书签（插入图片后书签会消失）
                        wordDoc.Bookmarks.Add "个人照片", rng

                        wordDoc.Save

                        successCount = successCount + 1

                    Else

                        failCount = failCount + 1
                        failList = failList & file.Name & "（缺少书签：个人照片）" & vbCrLf

                    End If

                    wordDoc.Close False

                End If

                On Error GoTo ErrHandler

            Else

                failCount = failCount + 1
                failList = failList & file.Name & "（未找到对应照片）" & vbCrLf

            End If

        End If

    Next file

    '========================
    ' 关闭Word
    '========================
    wordApp.Quit

    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set fso = Nothing

    '========================
    ' 显示结果
    '========================
    MsgBox _
        "已完成：" & successCount & "个" & vbCrLf & _
        "失败：" & failCount & "个" & vbCrLf & vbCrLf & _
        "失败名称：" & vbCrLf & _
        failList, _
        vbInformation

    Exit Sub

'========================
' 错误处理
'========================
ErrHandler:

    On Error Resume Next

    If Not wordDoc Is Nothing Then
        wordDoc.Close False
    End If

    If Not wordApp Is Nothing Then
        wordApp.Quit
    End If

    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set fso = Nothing

    MsgBox "运行出错：" & Err.Description, vbCritical

End Sub

'=================================================
' 提取姓名
' 支持：
' 1_张三
' 张三_1
' 001_张三_照片
'=================================================
Function 提取姓名(ByVal fileName As String) As String

    Dim arr() As String
    Dim i As Long
    Dim txt As String
    Dim result As String

    fileName = Replace(fileName, "-", "_")
    fileName = Replace(fileName, " ", "_")

    arr = Split(fileName, "_")

    result = ""

    For i = LBound(arr) To UBound(arr)

        txt = Trim(arr(i))

        If txt <> "" Then

            ' 过滤纯数字
            If Not IsNumeric(txt) Then

                ' 拼接中文姓名部分
                If result = "" Then
                    result = txt
                Else
                    result = result & txt
                End If

            End If

        End If

    Next i

    提取姓名 = result

End Function