# 发布到 GitHub / Publishing Checklist

本文假设仓库名为 `word-japanese-caret-align`，你可以改成自己喜欢的名字。

## 1. 发布前检查

- [ ] 更新版权持有人（可选）：`LICENSE`、`NOTICE.md` 中的
      `The CaretDownAligned Project Contributors` 可替换为你的名字或组织名。
- [ ] 确认 `VERSION` 和 `CHANGELOG.md` 的版本一致（当前 `1.0.0`）。
- [ ] 在一个**干净的 Word 环境**里测试 `src/CaretDownAligned.bas` 导入：
      - 新建空白文档；
      - `Alt + F11` → 导入 `.bas`；
      - `调试` → `编译 VBAProject`；
      - 按 `README.md` 的示例跑一遍；
      - 再测试中文源码复制粘贴。
- [ ] 确认没有提交个人文档、日志、截图中的个人信息、绝对路径、密钥。
- [ ] 运行静态校验：

  ```powershell
  pwsh -File scripts/validate.ps1
  ```

- [ ] 检查 README 中的相对链接（`src/…`、`docs/…`、`LICENSE` 等）在 GitHub 上可点开。
- [ ] 如需修改项目名，全局搜索 `CaretDownAligned` / `word-japanese-caret-align`，统一 README、NOTICE、PUBLISH 中的名称。
- [ ] 确认 `.bas` 是 ASCII + CRLF，中文 `.txt` 是 UTF-8 无 BOM。
- [ ] 确认没有提交 `~$*.doc*` 等 Word 锁文件（`.gitignore` 已覆盖）。

## 2. 初始化仓库并推送

在项目根目录执行：

```bash
git init
git add .
git commit -m "Initial release v1.0.0"
git branch -M main
```

在 GitHub 上新建空仓库（不要勾选自动生成 README/License），然后：

```bash
git remote add origin https://github.com/<YOUR_NAME>/word-japanese-caret-align.git
git push -u origin main
```

如果使用 SSH：

```bash
git remote add origin git@github.com:<YOUR_NAME>/word-japanese-caret-align.git
git push -u origin main
```

## 3. 仓库设置

建议填写：

**Description**

> Alt+Down vertical caret alignment for Japanese intensive-listening notes in Microsoft Word.

**Topics**

```text
word
vba
microsoft-word
japanese
language-learning
listening
annotation
macro
japanese-study
```

其他：

- 开启 **Issues**。
- 可选开启 **Discussions**。
- 确认 GitHub 自动识别 **MIT License**。
- 把 `assets/banner.svg` 转成 PNG（如有需要）用作 Social preview；GitHub 的 Social preview 上传通常接受 PNG/JPG。
- 如果启用了私有漏洞报告，`SECURITY.md` 中的说明可直接使用。

## 4. 打标签与 Release

```bash
git tag -a v1.0.0 -m "v1.0.0"
git push origin v1.0.0
```

在 GitHub 的 **Releases** 页面：

1. 选择 `v1.0.0` 标签。
2. 标题：`v1.0.0 — Initial release`。
3. 说明可从 [`CHANGELOG.md`](../CHANGELOG.md) 的 1.0.0 段落复制。
4. 可选：把 `src/CaretDownAligned.bas` 和
   `src/CaretDownAligned.zh-CN.vba.txt` 作为 Release assets 附上，方便直接下载。
5. 发布。

## 5. 可选：添加 CI 徽章

仓库创建后，在 `README.md` 顶部的 badges 下加入：

```markdown
[![Validate](https://github.com/<YOUR_NAME>/word-japanese-caret-align/actions/workflows/validate.yml/badge.svg)](https://github.com/<YOUR_NAME>/word-japanese-caret-align/actions/workflows/validate.yml)
```

把 `<YOUR_NAME>` 替换为你的 GitHub 用户名。工作流只做静态检查，不编译 VBA；不要把它宣传成“完整测试”。

## 6. 发布后验证

- 从 GitHub 下载 raw 的 `src/CaretDownAligned.bas`，在一个干净文档里导入并运行。
- 从 GitHub 打开 `src/CaretDownAligned.zh-CN.vba.txt`，复制粘贴到新模块，编译并运行。
- 检查 README 的示例、安装步骤、FAQ 链接。
- 检查 Release 资产和标签。
- 检查 Issue 模板是否可用。

## 7. 版本管理

- 使用语义化版本：修复用 `1.0.1`，兼容性新功能用 `1.1.0`，不兼容变更用 `2.0.0`。
- 每次用户可见变更更新 `CHANGELOG.md`。
- `.bas` 和中文 `.txt` 的逻辑必须同步；发布前对比两边的宏名、辅助函数、错误处理。
- 不要提交 Word 文档、日志、个人学习笔记或带敏感信息的截图。
