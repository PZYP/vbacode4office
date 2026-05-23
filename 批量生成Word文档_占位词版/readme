说明：此code可以基于doc/docx文档模板以及Excel表格，批量生成新的文档（以"序号_姓名"作为文件名）；
支持doc/docx文档；
代码执行流程：Excel表格内容与文档模板中指定位置的关系>复制模板-插入内容-生成新的文档；

运行环境：
Windows7/64bit含以上

数据清洗：
要求doc/docx文档，必须预先已经在对应的位置插入关键词，格式{{姓名}}。

运行前微调：
1、必须安装wps；
2、code中修改文档模板路径；
3、code中修改Excel表单文件路径；
4、code中修改输出结果文件夹路径；
5、code中修改Excel成绩列与文档指定成绩填充位置的关键词对应关系；
    例如：
    序列 A列(转换成数1) 对应 文档中的关键词{{序号}}
    姓名 B列(转换成数2) 对应 文档中的关键词{{姓名}}
    性别 F列(转换成数6) 对应 文档中的关键词{{性别}}
    出生年月 G列(转换成数7) 对应 文档中的关键词{{出生年月}}
    代码中示例：
    序号 = excelWS.Cells(i, 1).Value
    姓名 = excelWS.Cells(i, 2).Value
    性别 = excelWS.Cells(i, 6).Value
    出生年月 = excelWS.Cells(i, 7).Value
    ……
    Call 替换文本(wordDoc, "{{姓名}}", 姓名)
    Call 替换文本(wordDoc, "{{性别}}", 性别)
    Call 替换文本(wordDoc, "{{出生年月}}", 出生年月)
6、code应该在doc/docx模板中创建；

备注：
无