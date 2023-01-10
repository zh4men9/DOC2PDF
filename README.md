[![Security Status](https://www.murphysec.com/platform3/v3/badge/1611146949546246144.svg?t=1)](https://www.murphysec.com/accept?code=c39119a3ae1ad45230415f6b78ea3b1f&type=1&from=2&t=2)

# DOC2PDF

正如DOC2PDF名字可见，这个工具将DOC文档转换成PDF文件，批量转换，解放双手。详细功能如下：

- 文档转换（doc -> pdf， docx -> pdf）
- 多页合并为单页
- 裁剪页边距
- 合并文件

更详细介绍见博客：[(10条消息) DOC2PDF项目博客_zh4men9的博客-CSDN博客](https://blog.csdn.net/qq_32614873/article/details/126470122?spm=1001.2014.3001.5502)

# 使用

为方便使用工具，已将工具打包，打包后的文件在 `DOC2PDF_APP`中

打包命令: `pyinstaller -D -p DOC2PDF_APP DOC2PDF.py`

使用说明见 `使用说明.txt`

## TO DO

* [ ] 多页合并
* [ ] 选择文件进行操作
* [ ] 支持其他文件转PDF（图片、PPT等）
