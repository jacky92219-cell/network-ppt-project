---
name: ppt-image-gen
description: >
  為 network-ppt-project 的 PPTX 簡報生成 AI 插圖並整合。分析各 slide 內容、
  設計 fal.ai image generation prompt、呼叫 fal.ai 生成圖片、將圖片嵌入 PPTX。
  觸發時機：「生成圖片」、「插圖」、「幫 slide 加圖」、「AI 生成圖」、
  「fal.ai 圖片」、「更新插圖」、「重新生成圖片」、「修改圖片」等請求。
---

# PPT Image Generation Orchestrator

## 目標

協調 3 個 agent + 直接呼叫 fal.ai MCP 工具，為 `network-card-csi-v5.1.pptx` 生成技術風格插圖。

---

## Phase 0：情境確認

執行前先確認現有狀態：

1. 檢查 `_workspace/` 是否存在：
   - **不存在** → 初次執行，從 Phase 1 開始
   - **存在且有 `01_slide_analysis.md`** → 詢問使用者是否重新分析或跳過
   - **存在且有 `03_generated_images.md`** → 可直接跳至 Phase 4（圖片整合）

2. 確認 fal.ai 工具可用（`mcp__fal__generate_image` 應在工具清單中）

3. 建立工作目錄：
   ```
   mkdir -p _workspace
   mkdir -p output/images
   ```

---

## Phase 1：Slide 分析

**啟動 slide-analyzer agent：**

```
讀取 src/content.py 和 src/create_ppt.py，分析所有 slide 的內容類型，
找出最適合插入 AI 生成圖片的 6 張以內的 slide，
輸出分析結果至 _workspace/01_slide_analysis.md。
使用 agents/slide-analyzer.md 的角色定義執行此任務。
```

等待 agent 完成後讀取 `_workspace/01_slide_analysis.md`。

---

## Phase 2：Prompt 設計

**啟動 prompt-designer agent：**

```
讀取 _workspace/01_slide_analysis.md，
為每張推薦插圖的 slide 設計 fal.ai 圖片生成 prompt，
輸出至 _workspace/02_image_prompts.md。
使用 agents/prompt-designer.md 的角色定義執行此任務。
```

等待 agent 完成後讀取 `_workspace/02_image_prompts.md`。

---

## Phase 3：圖片生成（Orchestrator 直接執行）

**重要：此 Phase 由 orchestrator（Claude）直接呼叫 fal.ai MCP 工具，不委派給 agent。**
（原因：MCP 工具在主 context 中才有保證可用）

對每個 prompt：

1. 呼叫 `mcp__fal__generate_image`：
   - `model`: `fal-ai/flux/dev`
   - `prompt`: 從 02_image_prompts.md 讀取的完整 prompt
   - `image_size`: 按分析報告選擇（`landscape_16_9` / `square`）
   - `num_inference_steps`: 28
   - `guidance_scale`: 3.5
   - `num_images`: 1

2. 記錄返回的圖片 URL

3. 全部生成完畢後，將結果寫入 `_workspace/03_generated_images.md`：

```markdown
# Generated Images

| Slide | 標題 | 圖片 URL | 位置建議 | 狀態 |
|-------|------|---------|---------|------|
| 3 | CSI 量測原理 | https://... | 右側 1/3 | 成功 |
```

**錯誤處理：**
- API 失敗（rate limit / timeout）→ 等待 5 秒後重試 1 次
- 重試失敗 → 記錄失敗並繼續下一張
- 所有圖片生成完畢（或失敗記錄完畢）後才進入 Phase 4

---

## Phase 4：PPT 整合

**啟動 ppt-integrator agent：**

```
讀取 _workspace/03_generated_images.md 中的圖片 URL 清單，
下載圖片至 output/images/，修改 src/builders.py 新增圖片 helper，
修改 src/create_ppt.py 將圖片嵌入對應 slide，
執行 python src/create_ppt.py 確認無錯誤，
記錄整合結果至 _workspace/04_integration_log.md。
使用 agents/ppt-integrator.md 的角色定義執行此任務。
```

---

## Phase 5：結果彙報

整合完成後向使用者彙報：

1. 顯示 `_workspace/04_integration_log.md` 中的整合結果表格
2. 說明哪些 slide 成功加入圖片
3. 若有失敗，說明原因
4. 提醒執行 `python src/create_ppt.py` 重新生成 PPTX（如 ppt-integrator 未執行）
5. 詢問是否需要重新生成某張圖片或調整位置

---

## 資料流

```
src/content.py ──→ [slide-analyzer] ──→ _workspace/01_slide_analysis.md
                                                    │
                                                    ↓
                                        [prompt-designer]
                                                    │
                                                    ↓
                                    _workspace/02_image_prompts.md
                                                    │
                                                    ↓
                                    [Orchestrator: mcp__fal__generate_image]
                                                    │
                                                    ↓
                                    _workspace/03_generated_images.md
                                                    │
                                                    ↓
                                        [ppt-integrator]
                                                    │
                                    ┌───────────────┴──────────────┐
                                    ↓                              ↓
                            output/images/              src/builders.py (modified)
                            slide_N_*.png               src/create_ppt.py (modified)
```

---

## 測試情境

### 正常流程
- 輸入：「幫簡報生成 AI 插圖」
- 期望：自動分析 → 設計 prompt → 生成圖片 → 整合至 PPTX

### 部分重新執行
- 輸入：「第 5 張 slide 的圖片重新生成」
- 期望：只重新設計第 5 張 prompt，呼叫 fal.ai，更新 PPTX

### 錯誤流程
- fal.ai rate limit → 重試 1 次，失敗則跳過記錄
- python-pptx 嵌入失敗 → 記錄錯誤，不中止整體流程
