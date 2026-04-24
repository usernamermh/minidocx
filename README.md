# Mini DOCX Web Editor

一个可在 Windows 本地运行的网页版 DOCX 精简编辑器，启动后会开启本地服务并自动在浏览器打开。

## 已实现功能

- 字体、字号、加粗、斜体、下划线
- 段落左对齐、居中、右对齐
- H1 / H2 / H3 标题与左侧导航栏
- 本地图片插入
- 导入 `.docx` 到网页编辑器
- 下载导出 `.docx`

## 运行方式

### 双击启动

直接双击 `run_windows.bat`。

启动成功后会打开本地地址：`http://127.0.0.1:8765`

### 命令行启动

```bat
cd /d D:\path\to\proj_docx
python server.py
```

## 目录说明

- `server.py`：本地 HTTP 服务入口
- `docx_io.py`：原生 DOCX 读写逻辑
- `static/index.html`：编辑器页面
- `static/app.js`：编辑器交互逻辑
- `static/styles.css`：页面样式

## 说明

- 这是精简版编辑器，目标是本地可用、轻依赖、易启动
- DOCX 导入会尽量恢复文字样式、标题、对齐和图片
- 复杂页眉页脚、表格、浮动对象、批注等高级内容不会完整保真
