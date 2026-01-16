# batchHtmlToWord
批量转换HTML文件成WORD
AI写的，本人不会写代码。网上搜了批量转换的工具未找到，于是找AI帮忙写个工具，自己用了能用，就分享出来了。
网上大部分是在线的，或是JS脚本。找到一个叫“鹰迅办公”的，可以批量转换，但是收费。
将HTML转Word转换器.exe文件与HTML文件放一起，会自动遍历所有子文件夹，并转换成WORD，存放于上一级目录的word文件夹下。
转换后与原目录结构相同。
最开始的目的是将为知笔记导出的html文件导入到其他笔记，为知笔记导出的为html文件。
1.py是源代码，AI写的，担心安全问题的话自己看。我只是打包成.EXE

AI给出的打包命令
# 使用控制台模式打包（推荐）
pyinstaller --onefile --console --name="HTML转Word转换器" --add-data=".;." 1.py

# 或者如果已经打包过，先清理再打包
pyinstaller --clean --onefile --console --name="HTML转Word转换器" 1.py
