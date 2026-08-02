---
name: gwrite-wepost-ppt-skill
description: >
  PPT/PDF 一键转公众号文章全流程：输入 1-N 份 PPT/PDF + 格式框架（md）→ 自动提取幻灯片图与文字 → 贝叶斯过滤视频页（黑色 >80%）→ 按框架嵌入成文（中文单语）→ 反AI校对 → 拷打 → gwrite 预览 → 交付。仅 1 次打断，不部署不提交。
---

# gwrite-wepost-ppt-skill

用户说"把这份 PPT/PDF 变成公众号文章""按这个模板写一下我的 deck" → 执行以下流程。**只在 ⬜ 处弹窗**（用 `question` 工具），其余全部自动完成。

## TL;DR（老手速查）

```
需求对话（文件+框架） → 自动提取幻灯片[打断1: 视频页过滤确认]
→ 贝叶斯筛选页 → 嵌入框架成文（zh） → 校对（反AI≥45）
→ 拷打(自动50/50) → 自动gwrite预览 → 交付总结（不部署不提交）
```

## 前置准备

- 工具：`python3`（含 `PIL`、`python-pptx`；`pymupdf` 由脚本自动安装）、LibreOffice（`soffice`，PPT 输入必需）、`git`（仅复盘升级用）
- 项目已 clone：`/Users/gaothinkin/Coding/gaothink.in`
- 提取脚本：本 skill `scripts/extract_slides.py`
- 术语表与禁用词表复用 `gwrite-weekly-skill`（§A1/A2）
- 预览复用本 repo `scripts/write-server.mjs`（gwrite 预览，本地图自动 base64 内联）

---

## §1 方法论：贝叶斯 + JTBD + 奥卡姆剃刀

三原理不是口号，每一处流程都有对应落地：

| 原理 | 落在哪一步 | 怎么用 |
|------|-----------|--------|
| **贝叶斯** | Step 2 视频页过滤 | 先验 = 黑像素占比 > 80% 的视频页信号；证据 = 该页文本为空或仅时间码；双重证据才判死，避免误删深色设计页 |
| **贝叶斯** | Step 3 页筛选 | 每页先验（页类型：封面/议程/正文/过渡）+ 证据（文字量、图数量）→ 后验分级：**进正文 / 折叠成一句衔接 / 跳过** |
| **贝叶斯** | Step 4 校对修正 | 2+ 处同类型问题才批量改；单处偶发按"最省事解释"处理（奥卡姆），不动全局 |
| **JTBD** | 全流程 | 用户雇佣这个 skill 是"把讲过的 deck 变成 1 篇可直接用的公众号文章，不丢内容、不搬视频页、不替我提交"；正文每页回答读者 3 问：**这一页想说什么 / 凭什么 / 读者该带走什么** |
| **JTBD** | 交互设计 | 只打断 1 次（过滤确认）；不部署不提交，产出即交付 |
| **奥卡姆** | 写作 | 每页 = 1 图 + 1 标题 + 1-3 句解说 + 1 句"带走什么"；能少写不堆字 |
| **奥卡姆** | 流程设计 | 复用 gwrite 家族规范（禁用词/五维/预览），不重复发明；去掉部署链路，因为用户不需要它 |

### 1.1 视频页过滤（黑 > 80%，双重证据）

**规则**：渲染后的页面图像转灰度，亮度 < 30 的像素占比 **> 80%** → 判为视频播放页（黑屏）→ 过滤，不进正文。

**贝叶斯双证据**（防误杀深色设计页，如全黑背景 + 白色大字）：
1. 黑像素占比 > 80%（先验强信号）
2. 该页提取文本为空 或 仅含时间码/播放标记（如 `00:00`、`▶`）（证据）

两者同时满足才 **确定过滤**；只有黑占比超但文本丰富（如黑底白字标题页）→ **保留但标注**，交给 Step 3 的页筛选决定。

**输出**：过滤结果写入 manifest 并在 ⬜ 打断1 展示（含每页黑占比），用户可复核。

---

## §2 输入与输出格式规范

### 2.1 输入

| 输入 | 说明 |
|------|------|
| PPT/PDF 文件 | 1-N 份。`.pdf` 直接处理；`.ppt/.pptx` 经 LibreOffice 转 PDF |
| 格式框架 | md 格式约束文件（用户提供，含章节结构/占位符/字数要求）。**无框架** → 用 §2.2 内置模板 |

**输出**：`content/blog/zh/{date}-{slug}.mdx`（中文单语）+ 同目录 `.assets/` 幻灯片图。**不 git 提交、不部署、不上传 CDN**。

### 2.2 内置模板（无框架时的默认骨架）

```mdx
---
type: long
slug: "{date}-{slug}"
title: "PPT 拆解｜{一句话主题}？"   # 用追问形式
date: YYYY-MM-DD
tags: [PPT拆解]
locale: zh
---

## 卷首语      # 从第一页/演讲主旨提炼，不重复 title

## 一、主题章节一

> *副标题*

### 01 {小节标题}

![slide-01 一句话说明](../assets/{date}-{slug}.assets/slide-01.png)

解说 1-3 句（该页想说什么 + 凭什么 + 读者带走什么），每段 [N] 标注来源（对应页号）[1]。

---

## 附录：原始页对照

| 原始页 | 幻灯片 | 状态 |
|--------|--------|------|
| 1-5 | slide-01~05 | 进正文 |
| 6, 9 | （无） | 视频页，已过滤 |
```

### 2.3 用户格式框架（优先）

- 框架给出章节结构/占位符 → **以框架为准**，幻灯片按逻辑分组嵌入对应槽位
- 框架无图片槽位说明 → 默认每张有效幻灯片一张图：`![slide-{NN} {一句话}](../assets/{date}-{slug}.assets/slide-{NN}.png)`
- 框架约束冲突 → 询问一次（归入打断1 的选项）

### 2.4 幻灯片图规范

- 命名：`slide-{NN}.png`（NN 为过滤后连续序号），存 `content/blog/assets/{date}-{slug}.assets/`（扁平铁律）
- 渲染 dpi：≥150，导出后检查短边 ≥800px，文件 <20KB 视为渲染失败重转
- manifest 中 `original_page` 字段保留原始页号，正文按需标注
- 图片质量检查：`python3 -c "from PIL import Image; print(Image.open('file').size)"`

### 2.5 写作与校对规范

- **文风 = 随笔感悟**：以「我」的视角谈看法、联想、取舍，活人感优先；**只写感悟，不写经历**（禁止编造"当时现场""我讲了"等经历）
- **每段 [N] 标注**：正文引用置于句号前 `文本[1]。`，N = 原始页号或章节号
- **数字用粗体**：关键数字（百分比、人数、价格、时长）加粗
- **术语处理**：核心术语首次出现中英对照，后续用中文简称
- **反 AI 文学腔**：禁用「划时代」「里程碑」「颠覆性」等吹捧词；评价可证伪
- **破折号** ≤10 个/篇
- **五维评分**（复用 weekly §A2，每项 1-10，总分 ≥45 通过）：Directness / Rhythm / Trust / Authenticity / Density
- **事实核查**（复用 weekly §A3）：讲者姓名/头衔/机构、数字、引文逐一与幻灯片原文核对（禁止凭记忆补数字），命中修正记录 `| 位置 | 原文 | 修正 | 依据 |`
- 幻灯片里**没有**的信息禁止补充编造；确需补充 → 标注「讲者未提及，仅供参考」

---

## §3 工作流

**交互原则**（JTBD + 奥卡姆）：只打断 1 次：**过滤确认**。异常才弹窗（反AI <45、拷打 <50、预览失败）。每个 `question` 都标 `(Recommended)`。

### Step 1: 需求对话

收集以下信息（用户没说全才追问，**最多追 1 问**）：

| 维度 | 说明 |
|------|------|
| 文件 | PPT/PDF 路径或目录。给了文件但没给框架 → 用内置模板，不追问 |
| 格式框架 | md 框架文件路径（可选） |

信息足够（文件给了 + 无框架）→ 不弹框直接进入 Step 2。

### Step 2: 提取 + 视频页过滤 ⬜（第 1 次打断）

#### 2.1 运行提取脚本

```bash
python3 ~/.agents/skills/gwrite-wepost-ppt-skill/scripts/extract_slides.py \
  -i /path/to/deck.pdf -o /tmp/ppt-extract/ 2>&1
```

支持多文件：`-i a.pdf -i b.pptx`（多份 deck 时各自独立 manifest，输出按文件分目录）。

脚本自动：PPTX→PDF（soffice）→ 逐页渲染 PNG → 黑占比检测 → 文本提取 → 写 `manifest.json`。

#### 2.2 读取 manifest

```
manifest.json:
- pages_total       总页数
- slides            [{ index(过滤后), original_page, file, dark_ratio, is_video, text }]
- video_pages_excluded  被过滤的原始页号列表（含原因：黑占比 + 无文本）
```

#### 2.3 展示过滤结果 ⬜ 弹窗

```
question tool:
  header: "视频页过滤确认"
  question: "共 N 页，检测到 M 个视频页（黑屏，如原始页 3, 9, 12），已过滤。确认继续？"
  options:
    - "确认，开始成文" (Recommended) → 进入 Step 3
    - "恢复部分视频页" → 用户指定页号，改为保留并在正文标注
    - "调整阈值/重跑" → 改脚本参数（--dark-ratio）重跑后回到此确认
```

### Step 3: 嵌入框架成文（zh）

1. **页筛选（贝叶斯）**：对每张有效幻灯片按 §1.1 思路过一遍：
   - 封面/议程页 → 信息并入卷首语，不单独成图（除非框架要求）
   - 正文页 → 进正文，每页一图
   - 过渡/致谢页 → 折叠成一句衔接或跳过
2. **框架匹配**：有框架 → 按框架槽位填入；无框架 → 按 §2.2 内置模板分组成文（按章节标题/内容聚类分组）
3. **写作**：按 §2.5 规范写 zh 版 → `content/blog/zh/{date}-{slug}.mdx`
4. **反 AI 润色**：五维评分自评 ≥45 自动通过，<45 弹框干预

### Step 4: 校对 + 拷打 + 预览 + 交付（全自动）

#### 4.1 提交前拷打（自动 50/50，每项 ≥9）

| 维度 | 拷打要点 |
|------|---------|
| 完整性 | 每张有效幻灯片都进了正文或被明确记录（附录对照表）？ |
| 事实准确 | 数字/人名能否从幻灯片原文找到出处？没编造补充？ |
| 视频页过滤 | 过滤判定符合双重证据？误删的深色页恢复了？ |
| 反 AI 味 | 吹捧词/破折号/编造经历清干净？五维 ≥45？ |
| 结构完整 | 框架约束全部满足？图路径/alt/附录表齐全？ |

每轮必须产生实际改动（diff）。**50/50 → 自动通过**；<50 → 弹框干预。

#### 4.2 gwrite 预览（自动执行，不弹框）

```bash
node scripts/write-server.mjs &
sleep 2
BODY=$(python3 -c "import json;print(json.dumps(open('content/blog/zh/{file}.mdx',encoding='utf-8').read()))")
RESP=$(curl -sS -X POST "http://127.0.0.1:45678/api/preview-store?target=blog&locale=zh&file={file}.mdx" \
  -H "Content-Type: application/json" -d "{\"body\":$BODY,\"title\":\"t\"}")
ID=$(echo "$RESP" | python3 -c "import sys,json;print(json.load(sys.stdin).get('id',''))")
curl -sS "http://127.0.0.1:45678/api/preview-html?id=$ID" -o /tmp/preview-zh.html
```

自动判断：✅ 通过（预览 HTML 含全部 N 张图）→ 进入交付总结；❌ 失败 → 弹框（重试 / 跳过预览直接交付）。

#### 4.3 交付总结（不部署不提交）

产出就绪后向用户汇报（不弹框）：

```
✓ 文章就绪（未提交、未部署）
  文件：content/blog/zh/{file}.mdx
  图片：N 张（content/blog/assets/{date}-{slug}.assets/slide-*.png）
  视频页已过滤：M 个（原始页 3, 9, 12）
  预览：http://127.0.0.1:45678/api/preview-html?id={ID}（服务关闭即失效）
```

**明确不做**：`git add/commit/push`、`pnpm assets:upload`、`Deploy` tag。

**后续发布指引**（告知用户，不执行）：确认满意后手动 `pnpm assets:upload` + 提交 + 推送，即可按正常博客流程上线；或直接复制正文到公众号编辑器粘贴图片。

---

## §5 复盘与自升级（每次调用完成后必做）

部署验证通过后执行。**收集 → 三原理过滤 → 弹窗确认 → 升级推送**，四步走（同 gwrite-weekly-skill §5）：

### 5.1 收集观察

| 来源 | 收集什么 |
|------|---------|
| 弹窗 | 用户选了非推荐项？手动输入了什么？ |
| 流程 | 中断/重试几次？视频页误判几页？脚本报错类型？ |
| 反馈 | 用户口头纠正、抱怨、额外要求 |
| 环境 | soffice/pymupdf 行为变化；repo 结构变化；新踩的坑 |

### 5.2 三原理过滤（候选 → 升级项，全部通过才算）

1. **贝叶斯**——真实模式还是单次噪声？仅本次出现 → 不升级；2+ 次复现或有机制性根因 → 升级
2. **JTBD**——是否让用户额外弹窗/等待？修复后是否更接近"一句话启动、全程自动"？
3. **奥卡姆**——最小改动？收益 > 维护成本（含误触发）才加

### 5.3 弹窗确认

- **无升级项** → 一句话告知，不弹窗
- **有升级项** → 弹窗列出候选（证据/贝叶斯/JTBD/改动四项说明），选项：全部升级（推荐）/ 部分升级 / 暂不升级

### 5.4 升级执行

```bash
git -C ~/.agents/skills/gwrite-wepost-ppt-skill add -A
git -C ~/.agents/skills/gwrite-wepost-ppt-skill commit -m "upgrade: <一句话依据>"
git -C ~/.agents/skills/gwrite-wepost-ppt-skill push origin main
```

升级只改 SKILL.md/脚本自身，不新建文档；踩坑补进 §6 A4；每次升级独立 commit，回滚 = `git revert`。运行中的问题先修本次输出，改不改 skill 规则由复盘决定。

---

## §6 附录（按需查阅）

### A1 环境准备

```bash
# 脚本依赖（pymupdf 自动安装；其余可手动提前装）
python3 -m pip install --user pymupdf Pillow python-pptx

# PPTX → PDF 转换器（.ppt/.pptx 输入必需；只收 PDF 可跳过）
brew install --cask libreoffice
```

### A2 提取脚本用法

```bash
python3 ~/.agents/skills/gwrite-wepost-ppt-skill/scripts/extract_slides.py \
  -i deck.pdf [-i deck2.pptx ...] \
  -o /tmp/ppt-extract/ \
  [--dark-ratio 0.80] [--brightness-threshold 30] [--dpi 150]
```

- `--dark-ratio`：黑像素占比阈值（默认 0.80）
- `--brightness-threshold`：亮度阈值（默认 30）
- `--dpi`：渲染分辨率（默认 150）
- 输出：`{outdir}/{name}/slides/slide-{NN}.png` + `manifest.json`

### A3 黑页检测原理

每页渲染为灰度图 → 缩放到宽 160px（加速）→ 统计亮度 < 阈值像素占比 → 与文本提取结果联合判定：

```
dark_ratio > 0.80 且 text 为空/仅时间码 → video（过滤）
dark_ratio > 0.80 且 text 丰富          → 深色设计页（保留，标注）
```

### A4 踩坑

- **pymupdf 安装失败**（无网络/权限）：需用户手动装或换机器；`brew install poppler` 后可用 pdftoppm 替代（脚本暂未集成）
- **soffice 首次转换慢**：LibreOffice 首次启动约 10-30s，属正常；转换失败检查文件是否加密/损坏
- **黑底白字设计页误判**：严格按双重证据（黑占比 + 文本），只在两者都满足才过滤；打断1 会列出全部黑占比 >80% 的页供用户复核
- **多份 deck**：每份独立 manifest，正文按文件分章节，附录表分文件列出
- **gwrite 预览 500**：`file` 参数不含 locale 前缀（`?target=blog&locale=zh&file={slug}.mdx`，不能写 `zh/{slug}.mdx`）；预览服务关闭后 ID 失效，重跑预览即可

### A5 元规则

- **本 skill 不产出 git 操作**：交付即止，不 add/commit/push、不上传 CDN、不打 Deploy tag
- **内容用 MDX 后缀但纯 Markdown**：正文不能写 JSX 组件，只能用 Markdown + GFM
- **同源仓库同步**：本 skill 主仓库 `Gao-thinking/gwrite-wepost-ppt-skill`，更新后同步一份到主 repo `scripts/gwrite-wepost-ppt-skill.md`（git add 精确文件 + 无 Deploy tag）
