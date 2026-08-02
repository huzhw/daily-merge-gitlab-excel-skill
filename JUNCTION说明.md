# JUNCTION 说明 — daily-merge

> 本目录已与全局 skill 目录建立 junction，**实时双向同步，改哪边都一样**。

## 指向关系

| 项 | 路径 |
|----|------|
| 全局路径（junction） | `C:\Users\Administrator\.claude\skills\daily-merge` |
| 实际目录（F 仓库） | `F:\idea-workspase-skills\daily-merge` |
| 创建日期 | 2026-08-02 |

## 说明

- 全局目录是指向 F 仓库的 junction，两侧是**同一个目录**，不是副本。
- 修改 F 仓库的代码/脚本/模板，全局立刻生效；在全局路径下改文件，F 仓库同步变化。
- 日常维护只改 F 仓库（git 提交推送后全局自动一致），**不需要手动复制同步**。

## 检查是否正常

```bash
# 看该目录是否带 <JUNCTION> 标记
cmd /c dir "C:\Users\Administrator\.claude\skills" | findstr daily-merge
```

正常应显示 `<JUNCTION>  ...  daily-merge`。

## 回滚方法（恢复成独立副本）

```bat
:: 1. 删除 junction（只删链接，不删 F 源目录）
rd "C:\Users\Administrator\.claude\skills\daily-merge"

:: 2. 从备份恢复原全局目录
xcopy "F:\idea-workspase-skills\_skills_backup_20260802\daily-merge" "C:\Users\Administrator\.claude\skills\daily-merge" /E /I /Y
```

> 注意：`rd` 不要加 `/s`，否则可能递归进 F 源目录。

## 本 skill 特殊差异

- 全局原本独有的 `templates/generate_template.py`（2026-07-10）已并入 F 仓库 `templates/`。
- 全局 Office 锁文件 `templates/~$日报模板.xlsx` 已删除（Excel 临时文件，不该存在）。
- README.md、`templates/日报模板.xlsx` 以 F 仓库版本为准（F 版 07-31 更新，全局是 07-10 旧版）。
- 备份位置：`F:\idea-workspase-skills\_skills_backup_20260802\daily-merge`（8 个文件）。
