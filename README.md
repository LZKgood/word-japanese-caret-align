<p align="center">
  <img src="assets/banner.svg" alt="Word 日语精听 · 光标垂直对齐" width="100%">
</p>

# Word 日语精听 · 光标垂直对齐（CaretDownAligned）

[中文](README.md) | [English](README.en.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Word 2016+](https://img.shields.io/badge/Word-2016%2B-blue.svg)](docs/INSTALL.md)
[![VBA](https://img.shields.io/badge/VBA-macro-green.svg)](src/CaretDownAligned.bas)
[![No telemetry](https://img.shields.io/badge/telemetry-none-brightgreen.svg)](PRIVACY.md)
[![Version](https://img.shields.io/badge/version-1.0.0-informational.svg)](CHANGELOG.md)

在 Microsoft Word 中按 `Alt + ↓`，把光标从日文句子里的任意一个词，垂直落到下一行的同一列。用于日语精听、逐词注释、假名/翻译对齐，不会把上一行后面的文字挤到下一行。

> 本项目会修改 Word 文档的排版（插入全角空格、切换改写模式）。请先备份文档，并阅读 [免责声明](DISCLAIMER.md) 与 [隐私说明](PRIVACY.md)。

---

## 效果

```text
先生と学生が話しています。
        ↑ 光标停在「学生」这里，按 Alt+↓

先生と学生が話しています。
            学生
            ↑ 光标落到下一行同一列
```

同一行可以继续补写其他词：

```text
先生と学生が話しています。
老师  学生  正在说话。
```

## 功能

- **垂直对齐光标**：按 `Alt + ↓` 把光标落到下一行与当前词相同的水平列。
- **不拆行**：只在段落之间新建/定位注释行，不会在光标处插入换行，因此不会把「が話しています。」挤到下一行。
- **同一行多个注释**：第一次生成一条全角空格“网格行”，之后回到日文行换词再按，光标会回到同一行的对应列。
- **自动改写保护**：如果目标列右边已经有注释，宏会自动切到 Word 的“改写”模式，输入时替换空格，不推走后面的注释。
- **没有下一行也能用**：自动在日文段落后面补一条注释行。
- **纯本地**：没有网络请求、没有遥测、没有数据收集。
- **可审计**：完整 VBA 源码，无第三方组件。

## 环境要求

- Windows 10 / 11
- Microsoft Word 2016 / 2019 / 2021 / Microsoft 365 桌面版
- 已启用 VBA 宏（Word 的宏安全设置需允许运行宏）
- 普通**单栏正文**；表格、文本框、分栏、竖排文字不支持
- 注释建议使用**全角中文/日文**；半角英文/数字可能产生半字符偏移
- 未在 WPS、Word for Mac 上正式测试

已在 Windows 上的 **Microsoft Word 16.0** 实测通过：下一行已存在、没有下一行、下一段是普通正文、同一行补写第二个词等场景。

## 安装

### 方案 A：直接导入英文提示版 `.bas`（推荐）

1. 下载 [`src/CaretDownAligned.bas`](src/CaretDownAligned.bas)。
2. 在 Word 中按 `Alt + F11` 打开 VBA 编辑器。
3. 菜单 `文件` → `导入文件…`，选择 `CaretDownAligned.bas`。
4. 菜单 `调试` → `编译 VBAProject`，确认没有报错。
5. 绑定快捷键：
   - `文件` → `选项` → `自定义功能区` → 底部 `键盘快捷方式：自定义…`
   - 类别选「宏」，找到 `CaretDownAlignedFinal`
   - 按 `Alt + ↓`，点「指定」
   - “将更改保存在”选 `Normal.dotm`（所有文档可用）或当前文档
6. 保存模板/文档。

英文 `.bas` 中的提示文字是英文，代码逻辑与中文版完全相同。

### 方案 B：复制中文提示版代码

1. 打开 [`src/CaretDownAligned.zh-CN.vba.txt`](src/CaretDownAligned.zh-CN.vba.txt)。
2. 在 VBA 编辑器中右键工程 → `插入` → `模块`。
3. 如果模块顶部已经有 `Option Explicit`，保留它；把中文文件中的代码从 `' ============================================================` 开始全部粘贴到模块中。
4. 编译，并按上面的方法绑定 `Alt + ↓` 到 `CaretDownAlignedFinal`。

> 为什么提供两个版本？VBA 直接导入 `.bas` 时使用系统 ANSI 编码，中文 `.bas` 在非中文 Windows 上容易乱码。英文 `.bas` 是 ASCII，可直接导入；中文版用 UTF-8 `.txt` 提供，复制粘贴到 VBA 编辑器不会乱码。

更详细的安装、宏安全和卸载说明见 [docs/INSTALL.md](docs/INSTALL.md)。

## 使用

1. 输入日文句子：
   ```text
   先生と学生が話しています。
   ```
2. 把光标放在想注释的词后面（例如「学生」后面），按 `Alt + ↓`。
3. 第一次会在日文段落后面生成一条全角空格网格行，光标落在「学生」的列。直接输入注释：
   ```text
   先生と学生が話しています。
               学生
   ```
4. 回到日文行，把光标放到另一个词后面（例如「先生」），再按 `Alt + ↓`。光标会回到同一注释行的「先生」列；如果右边已经有「学生」，宏会自动打开“改写”模式。输入「老师」：
   ```text
   先生と学生が話しています。
   老师        学生
   ```
5. 完成注释后按 `Insert` 切回插入模式，避免回日文行修改时误覆盖字符。

### 对齐到词的哪一边？

宏严格对齐**当前插入点所在的列**：

- 想对齐到「学生」的**左边缘**：把光标放在「学」前面。
- 想对齐到「学生」的**右边缘**：把光标放在「生」后面。

### 从旧版本升级

1.0 之前的版本用“首行缩进”生成注释行。那种行不能继续用，否则左边没有可放置光标的列。第一次用 1.0 时会弹出提示：

> 检测到旧版生成的缩进注释行。请先删除这一行，再回到日文行重新按 Alt+↓ 生成新的网格注释行。

处理方式：复制旧注释文字 → 整行删除（包括段落标记）→ 回到日文行重新按 `Alt + ↓` → 粘贴/输入注释。

## 工作原理（简版）

Word 的普通文本行没有“虚拟列”：空白行的光标只能停在段落起点（受缩进控制）或已有字符之间。所以这个宏不是单纯移动光标，而是：

1. 记录当前光标相对正文左边界的水平位置。
2. 在日文段落之后生成一条由**全角空格**（U+3000）组成的注释行。一个全角空格正好是一个全角字符的宽度，于是形成一条字符网格。
3. 按目标水平位置把光标放到网格上最接近的字符边界。
4. 如果右侧已有注释，切换到 Word 的“改写”模式，让新输入替换空格而不是插入字符。
5. 只操作段落和格式，不在光标处插入段落标记，所以不会拆分原句。

技术细节、取舍和兼容性说明见 [docs/DESIGN.md](docs/DESIGN.md)。

## 限制与注意事项

| 场景 | 支持 | 说明 |
|---|---|---|
| 普通单栏正文 | ✅ | 主要使用场景 |
| 同一行多个词注释 | ✅ | 建议全角中文/日文 |
| 下一行已存在/不存在 | ✅ | 自动复用或新建注释行 |
| 半角英文/数字注释 | ⚠️ | 宽度不是整列，可能偏移半字符 |
| 表格、文本框、分栏 | ❌ | 直接 Beep，不执行 |
| WPS / Word for Mac | ⚠️ | 未正式测试，对象模型可能不同 |
| 竖排文字 | ❌ | 水平坐标语义不同 |
| 修改注释行字号 | ⚠️ | 网格列宽会变化，可能失准 |
| 很长的注释 | ⚠️ | 可能换行；大量排版建议用表格 |
| 旧版缩进注释行 | ❌ | 删除后重新生成 |

- 注释行里有很多全角空格，这是定位锚点；除非你不再需要对齐，否则不要手动删掉它们。打开“显示编辑标记”时它们会显示为小点/空白。
- 宏会自动切换 Word 的“改写”模式。状态栏会显示提示；写完注释按 `Insert` 切回插入模式。
- 如果注释较多、字体大小会变、或需要正式排版，**两行表格**（原文一行、注释一行，按词/字符分列）比光标宏更稳定，也不需要改写模式。
- 使用前请备份文档。宏会修改段落格式、插入全角空格并切换改写模式。

## 常见问题

- **438「方法和数据成员未找到」**：1.0 之前的版本使用了 Word 不存在的 `Paragraph.Index`；1.0 已改为遍历 `Paragraphs` 集合按 `Range.Start` 查找。请删除所有旧模块，只保留一份新代码。
- **4608「数值超出范围」**：旧版在文档末尾用 `InsertAfter` 插入段落导致；1.0 改用 `InsertParagraphAfter`，已修复。
- **光标没有落到词下面**：确认目标行是新版生成的网格行；旧版缩进行需要删除重建。确认你是从日文行按快捷键，而不是从注释行。
- **438 仍然出现**：检查 VBA 工程里是否还有旧代码、同名过程或重复模块；删除旧模块后重新导入/粘贴。
- **宏列表里有两个同名宏**：只保留一个，删除旧模块，重新绑定 `CaretDownAlignedFinal`。

更多问题见 [docs/FAQ.md](docs/FAQ.md)。

## 项目结构

```text
.
├── src/
│   ├── CaretDownAligned.bas            # 英文提示版，可直接导入
│   └── CaretDownAligned.zh-CN.vba.txt  # 中文提示版，复制粘贴用
├── docs/
│   ├── INSTALL.md                      # 详细安装/卸载
│   ├── FAQ.md                          # 常见问题
│   ├── DESIGN.md                       # 技术设计与取舍
│   └── PUBLISH.md                      # 发布到 GitHub 的步骤
├── assets/banner.svg
├── scripts/validate.ps1                # 静态校验
├── .github/                            # Issue/PR 模板与 CI
├── CHANGELOG.md
├── CONTRIBUTING.md
├── CODE_OF_CONDUCT.md
├── SECURITY.md
├── PRIVACY.md
├── DISCLAIMER.md
├── NOTICE.md
├── LICENSE
└── VERSION
```

## 参与贡献

欢迎提交 Issue 和 PR。请先阅读 [CONTRIBUTING.md](CONTRIBUTING.md) 和 [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)。安全问题请按 [SECURITY.md](SECURITY.md) 私密报告，不要在 Issue 中粘贴敏感文档内容。

## 许可与声明

- 代码采用 [MIT License](LICENSE)。
- 中文法律摘要（许可/免责/隐私/商标）：[docs/LEGAL.zh-CN.md](docs/LEGAL.zh-CN.md)。
- 使用风险、数据备份和商标声明见 [DISCLAIMER.md](DISCLAIMER.md) 与 [NOTICE.md](NOTICE.md)。
- 本项目不收集数据、不联网，详见 [PRIVACY.md](PRIVACY.md)。
- 本项目与 Microsoft 无隶属关系；Microsoft Word 是 Microsoft Corporation 的商标。

版本：`1.0.0` · 变更记录：[CHANGELOG.md](CHANGELOG.md)
