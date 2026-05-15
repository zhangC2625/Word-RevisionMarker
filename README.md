# Word Revision Marker

将 Word 审阅模式（Track Changes）下的修订标记转换为可视化格式标记，逐条审核，避免误操作。

Convert Word's Track Changes revisions into visible formatting marks, with step-by-step review.

## 效果 / Effect

| 修订类型 | 显示效果 |
|---------|---------|
| 插入的文字 / Insertions | <span style="color:blue"><u>蓝色 + 下划线</u></span> |
| 删除的文字 / Deletions | <span style="color:red"><s>红色 + 删除线</s></span> |

## 安装 / Installation

### 第 1 步：导入 VBA 宏

1. 打开 Word
2. `工具` → `宏` → `宏…`（或 `开发工具` → `宏`）
3. `宏的位置` 选择 **Normal.dotm**
4. 输入任意宏名称 → `创建`
5. 删除自动生成的代码，打开 `RevisionMarker.bas` 全选复制粘贴进去
6. `Cmd+S` 保存，关闭窗口

### 第 2 步（可选）：添加功能区按钮

关闭 Word，终端运行：

```bash
python3 setup_addin.py
```

重启 Word 后，功能区出现「修订工具」选项卡。

## 使用 / Usage

`工具` → `宏` → 选对应宏 → `运行`：

| 宏 / Macro | 功能 |
|-----------|------|
| `DiagnoseRevisions` | 诊断文档中修订的类型和数量 |
| `ReviewAndMarkRevisions_Simple` | 逐条审阅修订标记（推荐） |
| `ClearRevisionFormatting` | 清除蓝/红色格式 |
| `ShowHelp` | 显示帮助 |

逐条审阅时输入数字：**1**=标记  **2**=跳过  **3**=全部标记  **0**=停止

## 原理 / How It Works

使用 `Selection.NextRevision` 导航修订（与 Word 审阅栏的"下一条"按钮完全一致），不依赖 `ActiveDocument.Revisions` 集合遍历，避免 Mac Word 上的兼容性问题。

Uses `Selection.NextRevision` (same as clicking "Next" in Word's Review tab) instead of iterating the `Revisions` collection, avoiding compatibility issues on Mac Word.

## 环境 / Requirements

- macOS / Windows
- Microsoft Word（支持 VBA 宏）
- Python 3（仅安装功能区按钮时需要）

## License

MIT
