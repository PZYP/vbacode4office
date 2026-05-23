说明：此code可以将指定文件夹下文档批量打印指定页面（单页或者多页）；
支持doc/docx文档；
代码执行流程：指定文件夹>wps转换pdf>调用SumatraPDF软件打印指定页面；

运行环境：
Windows7/64bit含以上

数据清洗：
无特别要求

运行前微调：
1、必须安装wps、SumatraPDF；
2、code中修改SumatraPDF的安装路径；
3、code中修改需要打印的页面范围（比如第2页、4-5页）；
4、code在新建空白docx文档中创建；

备注：
WPS只负责：
doc/docx → PDF
真正打印：
交给 SumatraPDF

会在对应文件夹下生成PDF_TEMP的文件夹，相当于打印完成后，你还得到了每一个文档对应的PDF文件。