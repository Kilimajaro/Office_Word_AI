# 项目说明

项目使用VBA语言开发，延迟极低，深度嵌入Word系统，无需安装任何插件，100%安全。主要面向自定义api和轻量级密集ai交互场景。

本项目包含三个文件：`AI.frm`、`AI.frx` 和 `模块2.bas`。这些文件需要导入到 Microsoft Word 的 Visual Basic 编辑器中以实现AI辅助写作功能。
按照以下提示导入完成后，您只需要在word工具栏的最右侧，右键点击自定义工作区，并将导入的模块2添加到开始中的组（可能需要新建）中，即可正常使用。
![image](https://github.com/user-attachments/assets/a562ec94-4b50-4315-80e8-21b41635685e)



## 获取代码

1. 点击右上角的 **Code** 按钮。
2. 选择 **Download ZIP** 以下载整个项目的压缩包。
3. 解压下载的 ZIP 文件到您的计算机上。

## 导入文件到 Word

### 导入 `AI.frm` 和 `AI.frx` 文件

1. 打开 Microsoft Word。
2. 按 `Alt + F11` 打开 Visual Basic 编辑器。
3. 在菜单栏中选择 **插入** > **模块** 以创建一个新的模块。
4. 在新模块中，按 `Ctrl + G` 打开 **立即窗口**。
5. 在 **立即窗口** 中输入以下代码并按 Enter：

   ```vba
   ThisDocument.VBProject.VBComponents("窗体").Import "C:\path\to\your\AI.frm"
   ThisDocument.VBProject.VBComponents("窗体").Import "C:\path\to\your\AI.frx"
请将 C:\path\to\your\ 替换为您解压文件的实际路径。
### 导入 模块2.bas 文件
1. 在 Visual Basic 编辑器中，选择 插入 > 模块 以创建一个新的模块。
2. 在新模块中，按 Ctrl + G 打开 立即窗口。
3. 在 立即窗口 中输入以下代码并按 Enter：
   ```vba
    ThisDocument.VBProject.VBComponents("模块").Import "C:\path\to\your\模块2.bas"
请将 C:\path\to\your\ 替换为您解压文件的实际路径。
## 使用说明
导入完成后，您可以在 Word 的 Visual Basic 编辑器中查看和编辑这些文件。确保在将您自己的 **API秘钥和地址** 替换到 **AI.frm** 文件中注释标注的地方。
联系我们
如有任何问题或需要进一步的帮助，请通过 GitHub Issues 与我们联系。
