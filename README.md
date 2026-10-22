# Word_Compare

Windows Word 文档双栏对比工具，提供文字差异、格式与图片变化查看，以及审核记录和报告导出。

## 下载与运行

从 [Releases](https://github.com/moonhuahua/Word_Compare/releases/latest) 下载 `Word_Compare.exe`，或下载包含使用说明和示例的 `Word_Compare-Windows.zip`。无需安装 Microsoft Word。

需要 Windows 和 .NET Framework 4.8。DOC、RTF、ODT 读取及 PDF 导出另需安装 LibreOffice；DOCX、TXT 和其他报告导出不需要 LibreOffice。

## 功能

- 左右按段落和表格行对齐，共用滚动位置，支持差异定位、搜索与仅显示差异。
- 查看内容、常见格式和图片变化，提供字符、词/数字、句子、整段高亮。
- 记录“采纳 / 不采纳”审核意见，保存 JSON 后可重新载入；不合并或改写原文档。
- 支持 DOCX、HTML、CSV、JSON 报告，以及通过 LibreOffice 导出 PDF。
- 保留常见格式、可选暗黑模式及可折叠差异面板。

```text
Word_Compare.exe "原版.docx" "修订版.docx"
Word_Compare.exe --report "原版.docx" "修订版.docx" "对比报告.html"
```

## 构建和验证

使用 Windows 自带的 .NET Framework C# 编译器构建，无需下载第三方运行时 DLL：

```powershell
powershell -NoProfile -File .\Build.ps1
powershell -NoProfile -File .\Build.ps1 -Test
```

`-UI` 可验证 WPF 控件和主题，`-Conversion` 可验证 LibreOffice 转换。测试样例生成脚本需要 Python、python-docx 和 Pillow；正常运行 EXE 不需要 Python。

详见[使用说明](使用说明.md)。视图按内容对齐，不还原 Word 原始分页；复杂浮动对象、公式、域和修订历史等有明确边界。`示例` 提供人工生成的演示文档。
