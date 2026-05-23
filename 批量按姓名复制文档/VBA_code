Option Explicit

Sub 按姓名复制文档()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Dim nameStr As String
    
    Dim folderA As String
    Dim folderC As String
    
    Dim resultFile As String
    
    Dim successCount As Long
    Dim failCount As Long
    Dim repeatCount As Long
    Dim totalCount As Long
    
    Dim failList As String
    
    Dim targetFile As String
    Dim baseName As String
    Dim extName As String
    Dim newFile As String
    Dim n As Long
    
    Dim resultMsg As String
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    '==================================================
    ' 文件夹路径（末尾必须带 \）
    '==================================================
    
    ' 要搜索的总文件夹（包含所有子文件夹）
    folderA = "D:\文档集文件夹\"
    
    ' 复制到的目标文件夹
    folderC = "D:\拷贝输出文件夹\"
    
    
    '==================================================
    ' 当前工作表
    '==================================================
    Set ws = ActiveSheet
    
    ' B列最后一行
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    
    On Error GoTo ERR_HANDLE
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    successCount = 0
    failCount = 0
    repeatCount = 0
    totalCount = 0
    
    
    '==================================================
    ' 遍历B列姓名
    '==================================================
    For i = 2 To lastRow
    
        nameStr = Trim(ws.Cells(i, "B").Value)
        
        If nameStr <> "" Then
        
            totalCount = totalCount + 1
            
            resultFile = ""
            
            '==================================================
            ' 递归搜索文件夹及所有子文件夹
            '==================================================
            Call RecursiveFindFile(folderA, nameStr, resultFile)
            
            
            '==================================================
            ' 找到则复制
            '==================================================
            If resultFile <> "" Then
            
                targetFile = folderC & fso.GetFileName(resultFile)
                
                baseName = fso.GetBaseName(resultFile)
                extName = fso.GetExtensionName(resultFile)
                
                
                '==================================================
                ' 文件重名自动编号
                '==================================================
                If fso.FileExists(targetFile) Then
                
                    repeatCount = repeatCount + 1
                    
                    n = 1
                    
                    Do
                    
                        newFile = folderC & _
                                  baseName & "_重复" & n & "." & extName
                        
                        n = n + 1
                        
                    Loop While fso.FileExists(newFile)
                    
                    targetFile = newFile
                    
                End If
                
                
                '==================================================
                ' 复制文件
                '==================================================
                FileCopy resultFile, targetFile
                
                successCount = successCount + 1
                
            Else
            
                failCount = failCount + 1
                failList = failList & vbCrLf & nameStr
                
            End If
            
        End If
        
    Next i
    
    
    '==================================================
    ' 恢复Excel状态
    '==================================================
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    
    '==================================================
    ' 保存Excel
    '==================================================
    ThisWorkbook.Save
    
    
    '==================================================
    ' 运行结果提示
    '==================================================
    resultMsg = "处理完成！" & vbCrLf & vbCrLf & _
                "总数量：" & totalCount & vbCrLf & _
                "成功：" & successCount & vbCrLf & _
                "重复文件：" & repeatCount & vbCrLf & _
                "未找到：" & failCount
    
    If failCount > 0 Then
    
        resultMsg = resultMsg & vbCrLf & vbCrLf & _
                    "未找到姓名：" & vbCrLf & _
                    failList
    
    End If
    
    MsgBox resultMsg, vbInformation
    
    
    '==================================================
    ' 关闭Excel进程
    ' 避免下次打开提示“只读”
    '==================================================
    Application.Quit
    
    Exit Sub
    
    
'==================================================
' 错误处理
'==================================================
ERR_HANDLE:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    MsgBox "运行出错：" & vbCrLf & Err.Description, vbCritical

    Application.Quit

End Sub



'========================================================
' 递归搜索文件
'
' 支持：
'   doc
'   docx
'
' 支持：
'   1_张三.docx
'   张三_成绩.doc
'   001_李四_语文.docx
'
' 精准匹配：
'   不会把 张三 匹配成 张三丰
'========================================================
Sub RecursiveFindFile(ByVal folderPath As String, _
                      ByVal personName As String, _
                      ByRef resultFile As String)

    Dim fso As Object
    Dim folderObj As Object
    Dim subFolder As Object
    Dim fileObj As Object
    
    Dim extName As String
    Dim pureName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 已找到则停止递归
    If resultFile <> "" Then Exit Sub
    
    ' 文件夹不存在
    If Not fso.FolderExists(folderPath) Then Exit Sub
    
    Set folderObj = fso.GetFolder(folderPath)
    
    
    '========================================================
    ' 遍历当前文件夹文件
    '========================================================
    For Each fileObj In folderObj.Files
    
        extName = LCase(fso.GetExtensionName(fileObj.Name))
        
        ' 仅处理doc/docx
        If extName = "doc" Or extName = "docx" Then
        
            ' 去掉扩展名
            pureName = fso.GetBaseName(fileObj.Name)
            
            ' 精准姓名匹配
            If IsExactNameMatch(pureName, personName) Then
            
                resultFile = fileObj.Path
                Exit Sub
                
            End If
            
        End If
        
    Next fileObj
    
    
    '========================================================
    ' 递归子文件夹
    '========================================================
    For Each subFolder In folderObj.SubFolders
    
        Call RecursiveFindFile(subFolder.Path, personName, resultFile)
        
        If resultFile <> "" Then Exit Sub
        
    Next subFolder

End Sub



'========================================================
' 精准姓名匹配
'
' 允许：
'   1_张三
'   张三_成绩
'   001_张三_语文
'
' 不允许：
'   张三丰
'   老张三
'========================================================
Function IsExactNameMatch(ByVal fileName As String, _
                          ByVal personName As String) As Boolean

    Dim arr() As String
    Dim item As Variant
    
    ' 统一分隔符
    fileName = Replace(fileName, "-", "_")
    fileName = Replace(fileName, " ", "_")
    
    ' 按下划线拆分
    arr = Split(fileName, "_")
    
    For Each item In arr
    
        If Trim(CStr(item)) = Trim(personName) Then
        
            IsExactNameMatch = True
            Exit Function
            
        End If
        
    Next item
    
    IsExactNameMatch = False

End Function