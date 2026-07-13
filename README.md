# Mini DOCX Web Editor

一个面向 Windows 本地使用的轻量 DOCX 编辑器。程序启动本机 HTTP 服务，并在默认浏览器中打开编辑器；也可以打包为带系统托盘的 `MiniDocxTray.exe`。

## 功能

- 打开、另存和导出 `.docx`；保存使用原子写入，并保留最近的本地安全副本
- 字体、字号、加粗、斜体、下划线、文字底色
- 段落样式、对齐、行距、段前段后间距、缩进
- 多级编号、标题大纲、章节折叠和双层导航
- 表格的插入、增删行列与删除表格
- 页面尺寸设置、文档内查找与替换、可配置快捷键、编辑器缩放
- 最近文件、常用目录和浏览器文件系统访问能力支持
- 导入时恢复常见文字格式、表格、编号、页尺寸与内嵌图片

## 快速开始

### 直接运行

建议使用项目开发环境：

```powershell
& 'D:\codes\venvs\docs\.venv\Scripts\python.exe' -m pip install -r requirements.txt
& 'D:\codes\venvs\docs\.venv\Scripts\python.exe' server.py
```

也可双击 [run_windows.bat](run_windows.bat)。首次运行时它会创建 `D:\codes\venvs\docs\.venv`；若该环境缺少依赖，请按上面的命令安装 `requirements.txt`。

### 使用已打包版本

双击 `MiniDocxTray.exe`。程序在系统托盘中运行，可从菜单启动/停止本地服务、重新打开网页或切换端口。

## 开发与构建环境

项目将开发和打包依赖分开：

- `D:\codes\venvs\docs\.venv`：日常运行、调试和测试。
- `D:\codes\venvs\docs\.build-venv`：仅用于 PyInstaller 打包，避免构建依赖影响开发环境。

构建可执行文件：

```powershell
.\build_exe.bat
```

输出文件位于 `dist\MiniDocxTray.exe`。

## DOCX 兼容性说明

这是一个精简编辑器，不等同于 Word 或 WPS 的完整排版引擎。

- 未对已打开文档进行编辑时，10MB 以内的原始 DOCX 会原样导出，以保留本编辑器不建模的部件。
- 编辑后会按当前编辑器模型重建正文；复杂页眉页脚、批注、修订、浮动对象、复杂合并单元格等内容不能保证完全保真。
- DOCX 文件以二进制传输；导入图片在本地服务进程内以临时媒体缓存保存，不经 Base64 JSON 往返。
- 对重要或复杂文档，请先保留原文件，并在 Word/WPS 中复核导出结果。

## 目录说明

- [server.py](server.py)：本地 HTTP 服务、原生文件对话框、保存与临时媒体缓存。
- [docx_io.py](docx_io.py)：DOCX 解析、生成和编号/样式处理。
- [static/index.html](static/index.html)：编辑器页面结构。
- [static/app.js](static/app.js)：编辑器交互、历史记录、文件操作与大纲。
- [static/styles.css](static/styles.css)：响应式界面样式。
- [tests](tests)：DOCX 回归和前端结构测试。

## 测试

使用开发环境运行 DOCX 回归测试：

```powershell
& 'D:\codes\venvs\docs\.venv\Scripts\python.exe' -m unittest discover -s tests -v
```

## 注意事项

- 服务只监听 `127.0.0.1`，不对局域网暴露。
- “清理资源”会请求 Windows 清理多个进程的工作集；该功能可能使其他应用短暂变慢，建议仅在确有需要时使用。
- 资源状态页面在窗口隐藏时会暂停轮询，重新显示后恢复。
