Sub 批量插入照片_查找文字版()

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

    Dim findText As String

    '========================
    ' 要查找的文字
    '========================
    findText = "1寸红底"

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

            ' 提取姓名
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
            ' 找到照片后处理文档
            '========================
            If matched = True Then

                On Error Resume Next

                Set wordDoc = wordApp.Documents.Open( _
                    fileName:=docPath, _
                    ReadOnly:=False, _
                    AddToRecentFiles:=False)

                If Err.Number <> 0 Then

                    failCount = failCount + 1
                    failList = failList & file.Name & "（文档打开失败）" & vbCrLf
                    Err.Clear

                Else

                    Dim rng As Object

                    Set rng = wordDoc.Content

                    With rng.Find

                        .ClearFormatting
                        .Text = findText
                        .Forward = True
                        .Wrap = 1
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False

                    End With

                    If rng.Find.Execute Then
                    
                        Dim startPos As Long
                        Dim endPos As Long
                        Dim i As Integer
                        
                        '起始位置
                        startPos = rng.Start
                        
                        '从找到的位置开始
                        rng.Collapse 0
                        
                        '插入图片
                        Dim shp As Object
                        
                        '插入浮动图片
                        Set shp = wordDoc.Shapes.AddPicture( _
                            fileName:=photoPath, _
                            LinkToFile:=False, _
                            SaveWithDocument:=True, _
                            Anchor:=rng)
                            
                        With shp
                        
                            '浮于文字上方
                            .WrapFormat.Type = 3
                            
                            '设置尺寸 示例高2.8cm 宽2.7cm 可根据需要调整
                            .LockAspectRatio = False
                            
                            .Width = wordApp.CentimetersToPoints(2.7)
                            .Height = wordApp.CentimetersToPoints(2.8)
                            
                            '=========================
                            '微调位置 示例右移-0.1cm 下移0cm 可根据需要调整
                            '=========================
                            
                            '向右移动
                            .Left = wordApp.CentimetersToPoints(-0.1)
                            
                            '向下移动
                            .Top = wordApp.CentimetersToPoints(0)

                        End With

                        wordDoc.Save

                        successCount = successCount + 1

                    Else

                        failCount = failCount + 1
                        failList = failList & file.Name & "（未找到文字：" & findText & "）" & vbCrLf

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
    ' 退出Word
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
' 提取姓名 直接提取中文字符 过滤非中文字符
' 支持：
' 1_张三
' 张三_1
' 001_张三_照片
' test-张三
' 0(1)张三
'=================================================
Function 提取姓名(ByVal fileName As String) As String

    Dim i As Long
    Dim ch As String
    Dim result As String
    
    result = ""
    
    '去掉扩展名影响

    fileName = Replace(fileName, ".jpg", "")
    fileName = Replace(fileName, ".jpeg", "")
    fileName = Replace(fileName, ".png", "")
    fileName = Replace(fileName, ".doc", "")
    fileName = Replace(fileName, ".docx", "")
    
    '逐字符提取中文
    For i = 1 To Len(fileName)
    
        ch = Mid(fileName, i, 1)
        
        '中文字符范围
        If AscW(ch) >= 19968 And AscW(ch) <= 40869 Then
            result = result & ch
        End If
        
    Next i
    
    提取姓名 = Trim(result)

End Function