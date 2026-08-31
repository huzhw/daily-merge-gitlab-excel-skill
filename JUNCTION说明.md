# JUNCTION 说明 — daily-merge-gitlab-excel

> 本目录已与全局 skill 目录建立 junction，**实时双向同步，改哪边都一样**。

## 指向关系

| 项 | 路径 |
|----|------|
| 全局路径（junction，Claude Code） | `C:\Users\Administrator\.claude\skills\daily-merge-gitlab-excel` |
| 全局路径（junction，DSH） | `C:\Users\Administrator\.dsh\skills\daily-merge-gitlab-excel` |
| 全局路径（junction，Codex） | `C:\Users\Administrator\.codex\skills\daily-merge-gitlab-excel` |
| 全局路径（junction，Zcode） | `C:\Users\Administrator\.zcode\skills\daily-merge-gitlab-excel` |
| 实际目录（F 仓库） | `F:\idea-workspase-skills\daily-merge-gitlab-excel` |
| 更名记录 | 原 `daily-merge`，2026-08-28 更名为 `daily-merge-gitlab-excel` |

## 说明

- 全局两个目录都是指向 F 仓库的 junction，两侧是**同一个目录**，不是副本。
- 修改 F 仓库的代码/脚本/模板，全局立刻生效；在全局路径下改文件，F 仓库同步变化。
- 日常维护只改 F 仓库（git 提交推送后全局自动一致），**不需要手动复制同步**。

## 检查是否正常

```bash
cmd /c dir "C:\Users\Administrator\.claude\skills" | findstr daily-merge-gitlab-excel
cmd /c dir "C:\Users\Administrator\.dsh\skills"    | findstr daily-merge-gitlab-excel
cmd /c dir "C:\Users\Administrator\.codex\skills" | findstr daily-merge-gitlab-excel
cmd /c dir "C:\Users\Administrator\.zcode\skills" | findstr daily-merge-gitlab-excel
```

正常应显示 `<JUNCTION>  ...  daily-merge-gitlab-excel`。

## 回滚方法（恢复成独立副本）

```bat
rd "C:\Users\Administrator\.claude\skills\daily-merge-gitlab-excel"
rd "C:\Users\Administrator\.dsh\skills\daily-merge-gitlab-excel"
rd "C:\Users\Administrator\.codex\skills\daily-merge-gitlab-excel"
rd "C:\Users\Administrator\.zcode\skills\daily-merge-gitlab-excel"
```

> 注意：`rd` 不要加 `/s`，否则可能递归进 F 源目录。删除 junction 只删链接，不删 F 源目录。

## 本 skill 特殊差异

- 含 `templates/`（generate_template.py、日报模板.xlsx），随目录整体迁移。
- 更名前 README、`templates/日报模板.xlsx` 以 F 仓库版本为准；备份位于 `F:\idea-workspase-skills\_skills_backup_20260802\daily-merge`。