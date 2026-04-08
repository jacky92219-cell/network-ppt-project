# network-ppt-project

## 專案概覽

以 python-pptx 程式化生成 CSI（Channel State Information）網路卡技術簡報。
主線版本：`output/network-card-csi-v5.1.pptx`（Metal Gray 主題）。

## 工作慣例

- 修改後只執行 `python src/create_ppt.py`（主線），不重新生成色彩變體
- 所有顏色常數集中在 `src/theme.py`
- 版面配置函式在 `src/builders.py`

## Harness：PPT 圖片生成（ppt-image-gen）

**目標：** 分析 slide 內容 → 設計 fal.ai prompt → 生成 AI 插圖 → 嵌入 PPTX

**Agent 團隊：**
| Agent | 角色 |
|-------|------|
| `slide-analyzer` | 分析所有 slide，找出適合插圖的頁面（最多 6 張） |
| `prompt-designer` | 為每張 slide 設計 fal.ai 圖片生成 prompt（深色科技風） |
| `ppt-integrator` | 下載圖片，修改 builders.py/create_ppt.py 嵌入圖片 |

**Skill：**
| Skill | 用途 |
|-------|------|
| `ppt-image-gen` | 完整圖片生成流程 orchestrator |

**執行規則：**
- 收到「生成圖片」、「插圖」、「fal.ai」、「幫 slide 加圖」等請求 → 使用 `ppt-image-gen` skill
- Phase 3（圖片生成）由 orchestrator（Claude）直接呼叫 `mcp__fal__generate_image`，不委派給 agent
- 所有 agent 使用 `model: "opus"`
- 中間產出：`_workspace/` 目錄，最終圖片：`output/images/`

**目錄結構：**
```
.claude/
├── agents/
│   ├── slide-analyzer.md
│   ├── prompt-designer.md
│   └── ppt-integrator.md
└── skills/
    └── ppt-image-gen/
        └── SKILL.md
```

**變更歷史：**
| 日期 | 變更內容 | 對象 | 原因 |
|------|---------|------|------|
| 2026-04-08 | 初始建立 ppt-image-gen harness | 全體 | 整合 fal.ai MCP 圖片生成至 PPTX 工作流程 |
