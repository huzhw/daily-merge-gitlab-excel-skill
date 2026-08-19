# daily-merge — 日报 Excel 合并

日报流程第二步：读取 `daily-record` 生成的 md 需求记录，合并到当月 Excel 日报表。

## 相关技能
- [daily-record](https://github.com/huzhw/daily-record-skill)：日报记录（第一步）
- [git-commit](https://github.com/huzhw/git-commit-skill)：Git 提交规范
- [coding-rules](https://github.com/huzhw/coding-rules)：AI 编码协作规范
- [reread-claude-md](https://github.com/huzhw/reread-claude-md-skill)：重新加载 CLAUDE.md 规则
- [token-3000](https://github.com/huzhw/token-3000-skill)：API 一键切换（公司免费 ↔ 自己花钱）
- [service-manager](https://github.com/huzhw/service-manager)：桌面服务管理工具
- [code-check](https://github.com/huzhw/code-check-skill)：增量代码隐患检查

---

## 解决了什么问题

**md 写了需求，Excel 还要手动抄一遍。** 每天把 md 表格里的内容复制粘贴到 Excel，调格式、改序号、对齐日期——纯体力活。这个脚本一键完成：读 md → 追加到当月 Excel → 冻结表头 → 列宽自适应 → 日期格式匹配。

## 功能清单

- 自动找到当天 md 文件和当月 Excel 基准文件
- 序号全局递增（同一任务跨天复制时序号保持一致）
- G 列按 8h/工作日叠加，跳过双休和法定节假日（含调休）
- A 列同日期自动合并 + 居中
- D/N 列自动换行，行高自适应
- E 列写百分比（新任务 `0%`，完成手动改 `100%`）
- 仓库名→项目名映射（如 `workingpaper-v5.5` → `中信底稿v5`）
- 数据行和备注之间保留间距，备注原地不动
- 冻结表头行
- 去重：重复跑不会翻倍

## 使用

**每月第一天：**
```bash
python scripts/new_month.py
```

**其他日期：**
```bash
python scripts/daily.py
```

或在 AI 编码助手里说「合并日报」自动判断执行。

## 依赖

```bash
pip install openpyxl chinesecalendar
```

> `chinesecalendar` 用于 G 列跳过法定节假日（含调休）。未安装时自动降级为仅跳过双休并打印警告。

## 文件结构

```
daily-merge/
├── SKILL.md              ← 技能指令
├── README.md             ← 本文档
├── JUNCTION说明.md        ← 与全局 skill 目录的 junction 同步说明
├── templates/
│   ├── 日报模板.xlsx       ← 空白模板（月初无上月基准时兜底）
│   └── generate_template.py ← 重新生成模板
└── scripts/
    ├── new_month.py      ← 月初：读上月 Excel，复制未完成任务 + 写入当天数据
    ├── daily.py          ← 每日：昨日追加
    └── shared.py         ← 公用工具（workday 判断、描述格式化、说明区合并修复等）
```

## 列映射

| md 列 | Excel 列 | 说明 |
|--------|----------|------|
| 仓库 | B (项目名称) | 支持映射表转换 |
| 需求概述+涉及模块+备注 | D (任务描述) | 同仓库多任务合并：`N、概述\n模块\n备注`，任务间破折号分隔 |
| 状态 | E (完成百分比) | 新任务固定 0%，完成手动改 100% |
| 人工工时 | H (预计/实际工时) | 同仓库多任务求和 |
| AI辅助工时 | N (备注/说明) | 写入「预估AI辅助工时(h)：x」 |
| — | A (日期) | 当天，同日期合并单元格 |
| — | F (任务创建时间) | 当天 |
| — | G (预计完成时间) | 按 H/8 个工作日叠加，跳过双休和法定节假日 |

## 安装

```bash
git clone https://github.com/huzhw/daily-merge-skill.git ~/.claude/skills/daily-merge
```

安装后在 AI 编码助手里说「合并日报」触发。

## 许可

MIT
