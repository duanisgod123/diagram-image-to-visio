# Diagram Image to Visio

> 将流程图、架构图等图片一键转为可编辑的 Microsoft Visio `.vsdx` 文件

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://learn.microsoft.com/en-us/powershell/)
[![Visio](https://img.shields.io/badge/Visio-2016+-purple.svg)](https://www.microsoft.com/en-us/microsoft-365/visio/flowchart-software)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

---

## 这是什么？

你手上有一张别人发来的流程图截图、PPT 导出的架构图、或者论文里的技术路线图——想改几个字、换个箭头方向、补一个节点，但没有源文件，只能从头画。

**Diagram Image to Visio** 解决了这个问题：上传一张图片，它会用 AI 视觉识别图中的每一个形状、文字、连线和容器关系，生成结构化中间描述（JSON），再通过 PowerShell 调用 Visio COM 自动化引擎，逐个绘制出完全可编辑的 `.vsdx` 文件。

核心理念：**优先保证可编辑性、正确的连接拓扑和核心样式，而非像素级视觉还原。**

---

## 效果

| 输入 | 输出 |
|------|------|
| PNG / JPG / BMP 格式的流程图截图 | 同名 `.vsdx` 文件，Visio 可直接打开编辑 |
| PPT 导出的框架图 | 每个框都是独立可拖拽的形状 |
| 论文技术路线图 | 连线保持动态路由，文字完全可选可改 |

---

## 前置要求

- **Windows 10 / 11**
- **PowerShell 5.1 或更高版本**
- **Microsoft Visio 2016 或更高版本**（桌面版，已激活）
- Visio COM 组件可用（运行 `Test-VisioEnvironment.ps1` 可检测）

---

## 快速开始

### 1. 检测环境

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Test-VisioEnvironment.ps1
```

正常输出：

```json
{"installed":true,"name":"Microsoft Visio","version":"16.0","message":"Visio COM available."}
```

### 2. 准备图片

将你的流程图、架构图等截图放在任意可写目录下，支持 PNG / JPG / JPEG / BMP 格式。

### 3. 让 AI 提取规格

将图片上传给支持视觉识别的 AI（如 Claude Codex），AI 会按照 `references/diagram-spec.md` 中定义的 JSON Schema 输出结构化的图表规格文件。

### 4. 生成 Visio 文件

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Convert-DiagramImageToVisio.ps1 `
    -ImagePath "D:\path\to\diagram.png" `
    -SpecPath "D:\path\to\diagram.json"
```

生成的文件默认与源图片同名，扩展名为 `.vsdx`：

```
diagram.png  →  diagram.vsdx
```

### 可选参数

| 参数 | 说明 | 可选值 |
|------|------|--------|
| `-Orientation` | 页面方向 | `Auto` / `Landscape` / `Portrait` |
| `-PageSize` | 页面尺寸 | `Auto` / `A4` / `A3` / `Letter` / `Legal` |
| `-PreserveColor` | 保留颜色 | `$true`（默认） / `$false` |
| `-Visible` | 显示 Visio 窗口（调试用） | 开关 |
| `-CleanupIntermediate` | 生成后清理中间 JSON 文件 | 开关 |

---

## 工作流

```
┌──────────┐     ┌──────────────┐     ┌───────────────┐     ┌──────────┐
│  上传图片  │ ──► │ AI 视觉识别   │ ──► │ 生成 JSON 规格 │ ──► │  Visio   │
│ (PNG/JPG) │     │ + 结构化提取  │     │ (diagram-spec)│     │  .vsdx   │
└──────────┘     └──────────────┘     └───────────────┘     └──────────┘
```

1. **视觉识别**：AI 读取图片，推断是否在适用范围内（流程图、架构图、技术路线图等）
2. **规格提取**：按照 `diagram-spec.md` 的 Schema 生成 JSON，包含所有节点、容器、连线、标注和布局约束
3. **环境检测**：运行 `Test-VisioEnvironment.ps1` 确认 Visio COM 可用
4. **生成 Visio**：调用 `Convert-DiagramSpecToVisio.ps1` 通过 COM 自动化逐形状绘制

---

## 适用与不适用

### 适合的图片类型

- 流程图（Flowchart）
- 业务流程地图（Business Process Map）
- 研究框架图（Research Framework）
- 技术架构图（Architecture Diagram）
- PPT / Word 导出的路径图
- 分层框图（Layered Box Diagram）

### 不太适合的图片类型

- 手绘草图
- 密集的科技图表（标签极小、箭头重叠）
- 重渐变 / 阴影的营销图
- 自由曲线为主的复杂图示

---

## 项目结构

```
diagram-image-to-visio/
├── README.md                              # 本文件
├── SKILL.md                               # AI Skill 完整指令（给 AI Agent 阅读）
├── agents/
│   └── openai.yaml                        # OpenAI Codex 集成配置
├── references/
│   ├── diagram-spec.md                    # 中间 JSON Schema 定义（284 行）
│   └── visio-generation-notes.md          # Visio 生成策略：形状映射、连线策略等
└── scripts/
    ├── Test-VisioEnvironment.ps1          # COM 环境检测（54 行）
    ├── Convert-DiagramImageToVisio.ps1    # 图片→Visio 包装脚本（79 行）
    └── Convert-DiagramSpecToVisio.ps1     # Visio COM 主生成器（1561 行）
```

---

## 中间规格 Schema 速览

以下是 `diagram-spec.md` 中定义的核心结构：

```json
{
  "title": "示例图表",
  "canvas": { "width": 1880, "height": 1158 },
  "page": { "orientation": "Landscape", "page_size": "Auto" },
  "nodes": [
    {
      "id": "node-order",
      "shape": "rounded-rectangle",
      "text": "订单处理",
      "x": 972, "y": 366,
      "width": 164, "height": 68,
      "fill_color": "#D9F2A7",
      "line_color": "#333333"
    }
  ],
  "containers": [ /* 虚线大框容器 */ ],
  "connectors": [ /* 箭头连线，支持 solid/dashed/dotted */ ],
  "annotations": [ /* 标题、图例等标注 */ ],
  "images": [ /* 小图标裁剪嵌入 */ ],
  "layout": [ /* 对齐、等距、偏移等布局约束 */ ]
}
```

支持的形状类型：`rectangle`、`rounded-rectangle`、`diamond`、`ellipse`、`circle`、`terminator`、`block-arrow-right`、`block-arrow-down`。

---

## 设计原则

| 原则 | 说明 |
|------|------|
| 优先可编辑 | 文字、形状、连线全部保留为 Visio 原生对象，而非截图粘贴 |
| 不确定性透明 | 箭头归属不明时用 `#FF00FF` 品红色标注，并标记 `uncertain: true` |
| 保守提取 | 宁可标注不确定，也不静默猜测 |
| 布局约束优先 | 能用 `align`/`distribute`/`same_size` 表达的对齐关系，不写死坐标 |
| 渐进式还原 | 拥挤图表按 容器→主节点→主连线→次要标注 的顺序逐层恢复 |

---

## 常见问题

**Q: 为什么需要 Visio 桌面版？**
A: 本工具通过 Visio COM 自动化接口编程绘制形状，需要完整的桌面版 Visio（非网页版/Plan 2 自带桌面版也可以）。

**Q: 生成的效果不理想怎么办？**
A: 可以查看 `SKILL.md` 末尾的「Regression Lessons」章节，包含了大量的调优经验。核心思路是：导出 `.vsdx` 为 PNG 对比检查 → 定位问题区域 → 调整中间 JSON 规格 → 重新生成。

**Q: 支持 Visio 以外的输出格式吗？**
A: 当前仅支持 `.vsdx`。生成的 `.vsdx` 可在 Visio 中另存为 PDF、SVG、PNG 等格式。

---

## License

MIT