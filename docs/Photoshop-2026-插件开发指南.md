# Adobe Photoshop 2026 插件开发指南

> 编写日期：2026-11-11。目标环境：macOS / Apple Silicon，安装目录 `/Applications/Adobe Photoshop 2026`，本机 Photoshop 版本 **27.4.0**。
>
> 本文以 **UXP 面板插件**为主线，涵盖技术选型、环境配置、完整代码、Photoshop API、像素处理、工程化、调试、发布，以及 C++ 原生插件、Hybrid、CEP 和脚本。示例代码为本文编写；本次已核对本机安装信息和 Adobe 官方文档，但未在 Photoshop / UDT 中执行示例，不能把静态检查当成宿主运行验证。

## 目录

1. [本机环境与版本基线](#environment)
2. [选择插件技术路线](#choices)
3. [理解 Photoshop 安装目录](#installation)
4. [搭建 UXP 开发环境](#setup)
5. [理解 UXP 的执行模型](#runtime)
6. [完整示例：图层工具面板](#example)
7. [Manifest 配置详解](#manifest)
8. [Photoshop DOM 与对象定位](#dom)
9. [模态执行、撤销与取消](#modal)
10. [使用 batchPlay 补足 DOM](#batchplay)
11. [文件、目录与持久化](#storage)
12. [保存图像与批量导出](#export)
13. [Imaging API 与像素处理](#imaging)
14. [UI、生命周期与事件](#ui)
15. [网络服务与本地程序集成](#network)
16. [工程结构、构建与版本兼容](#engineering)
17. [调试与故障定位](#debugging)
18. [测试与验收](#testing)
19. [打包、安装与发布](#distribution)
20. [C++ 原生插件开发路线](#native)
21. [UXP Hybrid 开发路线](#hybrid)
22. [脚本、CEP 迁移与 Generator](#legacy)
23. [推荐实施顺序](#roadmap)
24. [官方资料与核验说明](#references)

<a id="environment"></a>

## 1. 本机环境与版本基线

### 1.1 已检查的信息

以下信息来自本机文件读取，不代表所有 Photoshop 2026 安装都相同。

| 项目 | 本机结果 | 开发意义 |
| --- | --- | --- |
| 安装目录 | `/Applications/Adobe Photoshop 2026` | 是产品目录，内部还包含 `.app` |
| 应用包 | `Adobe Photoshop 2026.app` | 原生可执行文件和内置资源位于包内 |
| Bundle ID | `com.adobe.Photoshop` | macOS 应用标识 |
| `CFBundleShortVersionString` | `27.4.0` | 设置兼容范围、定位 API 行为时使用 |
| `CFBundleVersion` | `27.4.0` | 本机 Info.plist 中记录的构建版本字段 |
| 当前命令环境架构 | `arm64` | 本文优先考虑 Apple Silicon 开发 |
| 主程序架构 | `arm64` + `x86_64` Universal Binary | 文件包含两种架构；不等于已经检查正在运行的进程架构 |
| `LSMinimumSystemVersion` | `13.0` | 仅是包内声明值；正式系统要求仍查 Adobe 对应发行说明 |
| Xcode | `/Applications/Xcode.app` 存在 | 原生开发工具存在，工具链版本未做构建验证 |
| UXP Developer Tool | `/Applications` 顶层未发现名称包含 UXP 的应用 | 不能据此断言其他位置没有安装；开发前在 Creative Cloud 中确认 |
| Photoshop C++ SDK | 应用包中未发现 `PIFilter.h`、`PIFormat.h` 等匹配文件 | Photoshop 安装包不能代替开发 SDK |

读取版本的命令如下；这条命令只读文件，不会启动或修改 Photoshop：

```bash
/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' \
  '/Applications/Adobe Photoshop 2026/Adobe Photoshop 2026.app/Contents/Info.plist'
```

### 1.2 四种版本号必须分开

| 名称 | 本文使用值 | 含义 |
| --- | --- | --- |
| Photoshop 产品版本 | `27.4.0` | 宿主应用版本，“2026”是产品名称的一部分 |
| UXP 运行时版本 | 官方对应表列出 Photoshop 27.4 → UXP 9.2.0 | 决定 HTML、CSS、JS 和平台 API 能力 |
| `manifestVersion` | 主示例为 `5` | 插件清单格式版本 |
| `host.data.apiVersion` | `2` | Photoshop API 执行模型版本 |

UXP 版本对应关系来自 [Adobe UXP 版本表](https://developer.adobe.com/uxp/uxp-api/versions)。最终以 UDT 连接目标 Photoshop 后显示的版本为准。不要用 `dvauxpui.framework` 或 `dvauxphost.framework` 的 Bundle 版本推算 UXP 运行时版本。

本文主示例将 `host.minVersion` 设为 `27.4.0`，目的是明确本机开发基线。这是本文的支持范围选择，不表示示例涉及的每个 API 都要到 27.4 才存在。要支持旧版，先检查 API 的 `Min Version` 并在最低支持版本上验收，再降低这个值。

<a id="choices"></a>

## 2. 选择插件技术路线

### 2.1 能力与交付形式

| 技术 | 适合的需求 | 技术栈 | 常见交付形式 | 主要成本 |
| --- | --- | --- | --- | --- |
| UXP 插件 | 常驻面板、图层操作、资源管理、导出工具、服务集成 | JavaScript、HTML、CSS | `.ccx` | 学习宿主 API、权限和异步模型 |
| UXP 脚本 | 一次性处理、内部自动化、验证 API | 现代 JavaScript | `.psjs` | 无常驻面板生命周期，执行结束即退出 |
| UXP Hybrid | UXP 界面加高性能本地算法、复用 C++ 库 | UXP + C++ addon | 包含 `.uxpaddon` 的 `.ccx` | SDK、架构、线程、签名和安装验证 |
| Photoshop C++ SDK | 注册原生滤镜、新文件格式、底层宿主扩展 | C/C++，原生工具链 | macOS `.plugin`；Windows 按类型如 `.8bf`、`.8bi` | 宿主 ABI、PiPL、像素协议、跨平台发布 |
| ExtendScript | 维护已有批处理脚本、兼容旧自动化 | 旧式 JavaScript / JSX | `.jsx`、`.jsxbin` | 旧语言和旧 DOM，与 UXP 不直接兼容 |
| CEP 扩展 | 维护已有 HTML 面板和 JSX 桥接项目 | HTML/JS、CEP、ExtendScript | `.zxp` | 迁移成本、架构兼容、旧平台生命周期 |
| Generator | 维护已有资源生成流水线 | Node.js / Generator | 专用模块和安装方式 | 历史接口，单独的服务与生命周期 |
| Photoshop 云 API | 服务器批处理、云端工作流 | REST API / 后端服务 | 服务部署 | 认证、任务调度、数据传输；不在桌面 PS 中加载 |

以上分类依据 [Photoshop 官方开发入口](https://developer.adobe.com/photoshop/) 和 [Photoshop Hybrid 与 C++ 插件的区别](https://developer.adobe.com/photoshop/uxp/2022/guides/hybrid-plugins/)。Windows 扩展名表示插件类型，不是 UXP 的包格式。

### 2.2 实际选型顺序

1. **要做面板、批量命名、导出、素材库：从 UXP 开始。**
2. 只需执行一次、无需常驻面板：先做 `.psjs`。
3. 需要处理像素：先验证 UXP Imaging API 是否满足吞吐量和格式要求。
4. 算法已有成熟 C++ 实现，或 JS 性能不足：评估 Hybrid。
5. 必须让 Photoshop 识别新的文件格式，或正式注册滤镜菜单项：评估 Photoshop C++ SDK。
6. 已有 CEP 项目：做功能盘点后分阶段迁移，避免把 `CSInterface`、Node 模块和 JSX 原样塞进 UXP。

例如，“面板里调用一个图像算法”与“让算法成为 Photoshop 原生滤镜”有不同的宿主集成要求。Hybrid 提供 JS 与原生代码组合能力，并不会自动完成传统滤镜或文件格式的注册。

### 2.3 2026 年的 CEP 规划

截至本文核验日期，Adobe 公布的 Photoshop CEP 迁移路线包含：2027 年第二季度停止接收新的 CEP 面板上架，2027 年 12 月计划默认禁用 CEP 面板；跨产品的移除目标为 2029 年 12 月。这些是官方路线图目标，不能当作每个已有版本立即停止工作的结论，也不能扩展理解为同日删除全部 JSX 或原生 C++ 插件支持。新面板项目应优先采用 UXP。[Adobe CEP 迁移中心](https://developer.adobe.com/uxp/migration-center/)

<a id="installation"></a>

## 3. 理解 Photoshop 安装目录

### 3.1 本机目录对应关系

```text
/Applications/Adobe Photoshop 2026/
├── Adobe Photoshop 2026.app/
│   └── Contents/
│       ├── Info.plist
│       ├── MacOS/
│       │   └── Adobe Photoshop 2026
│       ├── Required/
│       │   ├── UXP/                  # Adobe 内置 UXP 组件
│       │   ├── Plug-ins/Generator/   # 内置 Generator 相关组件
│       │   └── ...
│       ├── PlugIns/Required/
│       │   ├── Filters/              # 内置原生滤镜
│       │   ├── File Formats/         # 内置文件格式模块
│       │   ├── Automate/
│       │   └── ...
│       └── Resources/CEP/extensions/
├── Plug-ins/
│   └── Generator/
└── Presets/
    └── Scripts/                      # 本机存在的 JSX 脚本目录
```

本机可见 `Average.plugin`、`JPEG2000.plugin` 等原生模块，也存在 `Export Layers To Files.jsx`、`Image Processor.jsx` 等脚本。它们说明不同扩展机制可以共存，不能据此推导为同一种加载方式。

### 3.2 开发代码应放在哪里

将自己的插件工程放在普通可写开发目录，通过 UDT 加载 `manifest.json`。发布后使用 `.ccx` 安装。**不要把 UXP 项目复制到应用内的 `Required/UXP`，也不要把 `.ccx` 解压到原生 `Plug-ins` 目录。** UDT 加载和正式安装是两条工作流。[UDT 工作流](https://developer.adobe.com/uxp/guides/how-to/udt-deep-dive/plugin-workflows)、[UXP 安装方式](https://developer.adobe.com/uxp/guides/how-to/distribution/install/)

对内置组件的合理使用方式是观察版本、目录和公开清单结构。其 JS、内部模块、特殊权限和启动字段可能仅对 Adobe 内置插件有效，不应复制到第三方产品。

本机 `com.adobe.photoshop.exportAs` 的清单使用 Manifest 5 和 `host.data.apiVersion: 2`。这一观察只能验证本机内置组件的配置，不能替代第三方插件的公开接口文档。

<a id="setup"></a>

## 4. 搭建 UXP 开发环境

### 4.1 所需工具

| 工具 | 用途 | 是否为主示例必需 |
| --- | --- | --- |
| Photoshop 27.4.0 | 执行插件 | 是 |
| Adobe Creative Cloud 桌面应用 | 安装 UDT、安装发布包 | 是 |
| UXP Developer Tool，简称 UDT | 加载、重载、调试、打包 | 是 |
| 文本编辑器，如 VS Code | 编辑工程 | 是 |
| Node.js / npm | 构建、类型检查、纯逻辑测试 | 主示例不需要构建；后续工程化可使用 |
| Git | 源码和版本管理 | 建议 |
| Xcode / Visual Studio | 编译原生插件或 addon | 仅原生 / Hybrid 路线需要 |

采用 Adobe 当前提供且支持目标宿主的 UDT 版本。当前通用操作指南以 UDT 2.2 或更高版本展开，不要把多年前教程中的最低版本当成推荐安装版本。[UDT 指南](https://developer.adobe.com/uxp/guides/how-to/udt-deep-dive/)

### 4.2 安装与连接

1. 在 Creative Cloud 桌面应用中搜索 **UXP Developer Tools** 并安装。
2. 启动 UDT，按界面启用 Developer Mode；这一步可能出现系统管理员授权提示。
3. 启动目标 Photoshop，而不是另一套旧版或 Beta 版。
4. 在 Photoshop 的“设置 / 首选项 → 增效工具（Plugins）”中启用 Developer Mode；具体中文名称以当前界面为准，按提示重启。
5. 在 UDT 的 Connected Apps / Connected Applications 中确认 Photoshop 及对应版本。
6. 使用 Add Plugin 选择工程根目录的 `manifest.json`。
7. 执行 Load 或 Load & Watch，然后在 Photoshop 的 Plugins 菜单中打开面板。

开发者模式需要在 UDT 和宿主侧分别确认；UDT 中“已添加工程”不等于已经加载到 Photoshop。[开发工具安装](https://developer.adobe.com/uxp/guides/how-to/developer-tools/)、[开发者模式 FAQ](https://developer.adobe.com/uxp/faq/)

### 4.3 修改后如何重载

| 修改内容 | 建议操作 |
| --- | --- |
| JS、HTML、CSS | 使用 Watch 或 Reload，并查看 Console |
| Manifest 的权限、入口、宿主范围 | Unload，再 Load |
| 构建配置 | 重新构建，再加载构建输出 |
| 原生 addon / 插件二进制 | 按 SDK 流程重新加载；必要时退出并重启宿主 |

当前 UDT 文档明确说明 Watch 不处理 Manifest 变更；只修改文件后等待热更新可能让旧权限继续生效。[UDT 工作流](https://developer.adobe.com/uxp/guides/how-to/udt-deep-dive/plugin-workflows)

<a id="runtime"></a>

## 5. 理解 UXP 的执行模型

### 5.1 三个层次

```text
HTML / CSS / 面板事件
        ↓
插件业务逻辑：校验、生成方案、队列、错误处理
        ↓
┌──────────────────────┬────────────────────────────┐
│ require("photoshop") │ require("uxp")             │
│ 文档、图层、动作、像素 │ 文件、入口、存储等平台能力  │
└──────────────────────┴────────────────────────────┘
        ↓
Photoshop 宿主 + UXP 运行时
```

基础导入示例：

```javascript
const { app, core, action, imaging, constants } = require("photoshop");
const { entrypoints, storage } = require("uxp");
```

`photoshop` 和 `uxp` 是宿主提供的模块。它们不是要通过 npm 安装才能运行的两个普通库，也不能在终端执行 `node main.js` 时直接获得 Photoshop 文档对象。[UXP 平台与宿主 API](https://developer.adobe.com/uxp/uxp-api/)

### 5.2 UXP 与浏览器、Node.js 的边界

- UXP 提供受支持的 Web 技术子集，Chrome 中能运行的界面不一定能直接运行。
- `require()` 存在，不代表具备完整 Node.js 运行时或能加载任意 `.node` 原生模块。
- UXP 自身有部分路径式文件 API，但其权限和覆盖面不能按 Node 的 `fs` 推断。
- npm 可用于开发时构建；依赖要根据实际运行时能力审核。
- `document` 通常指面板 HTML 文档；本文将 Photoshop 文档命名为 `psDoc`，避免混淆。
- 面板业务会长期存在，但用户可随时切换、关闭或修改文档。不能无限期缓存“当前图层”代理并假定仍有效。

界面能力参见 [Spectrum / HTML 的区别](https://developer.adobe.com/photoshop/uxp/2022/uxp/reference-spectrum/)，文件 API 的选择参见 [UXP 文件操作](https://developer.adobe.com/uxp/guides/how-to/recipes/filesystem-operations/)。

### 5.3 读写宿主状态的基本规则

读取普通元数据通常可以直接调用 DOM；创建图层、修改名称、保存文档等改变宿主状态的操作，要放入 `core.executeAsModal()`。所有异步宿主调用应正确 `await`，并处理失败。模态执行提供互斥访问，撤销合并则是单独的历史管理功能。[模态执行文档](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/executeasmodal)

<a id="example"></a>

## 6. 完整示例：图层工具面板

### 6.1 功能与约定

这个示例包含四个文件，无 npm 依赖。按下列文件名复制到同一个新目录后，可通过 UDT 加载。

功能：读取当前文档和选中图层；按前缀批量重命名；将当前选择的元数据导出为 JSON。

行为约定：

- 重命名按 **图层 ID 升序**编号，例如 `asset_001`、`asset_002`，不承诺与图层面板从上到下的视觉顺序相同。
- 仅处理直接选中的图层；选中组时只修改组名，不递归修改子图层。
- 对锁定图层采用保守策略：任意选中图层报告 `locked` 时拒绝整批操作。
- 一次成功重命名形成一个撤销步骤；中途失败或取消时回滚本次悬挂的文档历史。
- 在模态排队前记录目标文档和选择，获得模态后再次校验，避免改错文档。
- JSON 导出记录点击导出时的快照；在用户选择文件位置期间，不持续更新快照。

```text
layer-toolkit/
├── manifest.json
├── index.html
├── styles.css
└── main.js
```

### 6.2 `manifest.json`

```json
{
  "manifestVersion": 5,
  "id": "com.example.layer-toolkit-dev",
  "name": "图层工具开发示例",
  "version": "0.1.0",
  "main": "index.html",
  "host": {
    "app": "PS",
    "minVersion": "27.4.0",
    "data": {
      "apiVersion": 2
    }
  },
  "entrypoints": [
    {
      "type": "panel",
      "id": "layerToolkitPanel",
      "label": {
        "default": "图层工具"
      },
      "minimumSize": {
        "width": 280,
        "height": 340
      },
      "preferredDockedSize": {
        "width": 320,
        "height": 480
      },
      "preferredFloatingSize": {
        "width": 360,
        "height": 500
      }
    }
  ],
  "requiredPermissions": {
    "localFileSystem": "request"
  }
}
```

这里的 ID 用于本地示例。正式分发前换成属于自己的插件 ID；Marketplace 分发必须与 Developer Distribution 中的条目对应。示例没有图标声明，因此不需要额外创建图片文件。

### 6.3 `index.html`

```html
<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8" />
    <link rel="stylesheet" href="styles.css" />
  </head>
  <body>
    <div id="app">
      <sp-heading size="S">图层工具</sp-heading>
      <sp-body>读取选择、批量命名，并导出图层清单。</sp-body>

      <sp-button id="refreshButton">刷新文档信息</sp-button>
      <sp-body id="documentInfo">尚未读取文档。</sp-body>

      <sp-textfield id="prefixInput" value="asset">
        <sp-label slot="label">名称前缀</sp-label>
      </sp-textfield>
      <sp-body size="S">按图层 ID 升序生成 asset_001 等名称。</sp-body>

      <sp-button id="renameButton" variant="primary">
        重命名选中图层
      </sp-button>
      <sp-button id="exportButton">导出选择清单 JSON</sp-button>

      <sp-body id="status">就绪。</sp-body>
    </div>
    <script src="main.js"></script>
  </body>
</html>
```

此处使用 UXP 内置 Spectrum widgets，没有导入独立 Spectrum Web Components 包。两者标签相似，但初始化、功能和版本要求不完全相同。[Spectrum UXP 参考](https://developer.adobe.com/photoshop/uxp/2022/uxp/reference-spectrum/)

### 6.4 `styles.css`

```css
body {
  padding: 16px;
}

#app {
  display: flex;
  flex-direction: column;
}

sp-heading,
sp-body,
sp-textfield,
sp-button {
  margin-bottom: 12px;
}

sp-textfield,
sp-button {
  width: 100%;
}

#documentInfo,
#status {
  white-space: pre-wrap;
}
```

### 6.5 `main.js`

```javascript
const { app, core } = require("photoshop");
const { entrypoints, storage } = require("uxp");

const fileSystem = storage.localFileSystem;
let busy = false;

function setStatus(message) {
  document.getElementById("status").textContent = message;
}

function setBusy(value) {
  busy = value;
  for (const id of [
    "refreshButton",
    "renameButton",
    "exportButton",
    "prefixInput"
  ]) {
    document.getElementById(id).disabled = value;
  }
}

function requireDocument() {
  if (app.documents.length === 0) {
    throw new Error("请先打开一个 Photoshop 文档。");
  }
  return app.activeDocument;
}

function selectedLayersByID(psDoc) {
  return Array.from(psDoc.activeLayers).sort((a, b) => a.id - b.id);
}

function getSelectionSnapshot() {
  const psDoc = requireDocument();
  return {
    schemaVersion: 1,
    capturedAt: new Date().toISOString(),
    photoshopVersion: app.version,
    document: {
      id: psDoc.id,
      title: psDoc.title,
      width: psDoc.width,
      height: psDoc.height
    },
    layers: selectedLayersByID(psDoc).map((layer) => ({
      id: layer.id,
      name: layer.name,
      kind: String(layer.kind),
      visible: layer.visible,
      locked: layer.locked
    }))
  };
}

function refreshInfo() {
  const target = document.getElementById("documentInfo");
  try {
    const snapshot = getSelectionSnapshot();
    target.textContent = [
      `文档：${snapshot.document.title}`,
      `尺寸：${snapshot.document.width} × ${snapshot.document.height} px`,
      `已选图层：${snapshot.layers.length}`
    ].join("\n");
  } catch (error) {
    target.textContent = error.message || String(error);
  }
}

async function runTask(task) {
  if (busy) return;
  setBusy(true);
  setStatus("正在处理……");
  try {
    await task();
  } catch (error) {
    console.error("[layer-toolkit]", error);
    setStatus(`未完成：${error.message || String(error)}`);
  } finally {
    setBusy(false);
    refreshInfo();
  }
}

async function renameSelection() {
  const prefix = String(
    document.getElementById("prefixInput").value || ""
  ).trim();
  if (!prefix || prefix.length > 40 || /[\r\n]/.test(prefix)) {
    throw new Error("前缀应为 1～40 个字符，且不能包含换行。");
  }

  // 在等待模态锁之前固定用户此次点击的目标。
  const snapshot = getSelectionSnapshot();
  if (snapshot.layers.length === 0) {
    throw new Error("请至少选择一个图层。");
  }
  const expectedIDs = snapshot.layers.map((layer) => layer.id);

  await core.executeAsModal(async (context) => {
    const psDoc = requireDocument();
    if (psDoc.id !== snapshot.document.id) {
      throw new Error("当前文档已经改变，请重新执行。");
    }

    const layers = selectedLayersByID(psDoc);
    const sameSelection = layers.length === expectedIDs.length &&
      layers.every((layer, index) => layer.id === expectedIDs[index]);
    if (!sameSelection) {
      throw new Error("图层选择已经改变，请重新执行。");
    }
    if (layers.some((layer) => layer.locked)) {
      throw new Error("选择中存在锁定图层，请调整选择或先解锁。");
    }
    if (context.isCancelled) {
      throw new Error("用户已取消操作。");
    }

    const suspension = await context.hostControl.suspendHistory({
      documentID: psDoc.id,
      name: "批量重命名图层"
    });

    try {
      for (let index = 0; index < layers.length; index += 1) {
        if (context.isCancelled) {
          throw new Error("用户已取消操作。");
        }
        const number = String(index + 1).padStart(3, "0");
        layers[index].name = `${prefix}_${number}`;
        context.reportProgress({ value: (index + 1) / layers.length });
        // 让长列表的循环有机会处理取消信号。
        if ((index + 1) % 20 === 0) {
          await new Promise((resolve) => setTimeout(resolve, 0));
        }
      }
      if (context.isCancelled) {
        throw new Error("用户已取消操作。");
      }
    } catch (error) {
      try {
        await context.hostControl.resumeHistory(suspension, false);
      } catch (rollbackError) {
        console.error("[layer-toolkit:rollback]", rollbackError);
      }
      // 必须将失败继续抛到模态范围之外。
      throw error;
    }

    await context.hostControl.resumeHistory(suspension, true);
  }, {
    commandName: "批量重命名图层",
    timeOut: 2
  });

  setStatus(`已重命名 ${expectedIDs.length} 个图层，可一次撤销。`);
}

async function exportSelection() {
  // 在弹出保存对话框之前生成普通 JSON 数据，避免持有图层代理。
  const snapshot = getSelectionSnapshot();
  const file = await fileSystem.getFileForSaving("layer-selection.json", {
    types: ["json"]
  });
  if (!file) {
    setStatus("已取消导出。");
    return;
  }

  await file.write(JSON.stringify(snapshot, null, 2), {
    format: storage.formats.utf8
  });
  setStatus(`已导出：${file.name}`);
}

entrypoints.setup({
  panels: {
    layerToolkitPanel: {
      show() {
        refreshInfo();
      }
    }
  }
});

document.getElementById("refreshButton").addEventListener("click", () => {
  refreshInfo();
  setStatus("信息已刷新。");
});
document.getElementById("renameButton").addEventListener("click", () => {
  void runTask(renameSelection);
});
document.getElementById("exportButton").addEventListener("click", () => {
  void runTask(exportSelection);
});
refreshInfo();
```

示例中的 `layer.locked` 是读取属性；`layer.name` 是可写属性。锁定检查是本示例的产品策略，不能据此断言所有锁类型都会阻止所有改名操作。[Layer API](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/classes/layer)

### 6.6 运行与验收

1. 通过 UDT Add Plugin 选择 `layer-toolkit/manifest.json`，然后 Load。
2. 在 Photoshop Plugins 菜单打开“图层工具”。
3. 新建测试文档，创建三个未锁定的普通图层，选中其中两个。
4. 点击刷新，确认显示当前文档和两个选中图层。
5. 使用前缀 `test` 执行重命名，确认所选图层变为 `test_001`、`test_002`。
6. 按一次撤销，确认两个名称同时恢复。
7. 导出 JSON，检查文档 ID、图层 ID、名称与点击导出时一致。
8. 再次导出并取消选择，确认显示“已取消导出”，无异常。
9. 选择锁定图层执行改名，确认提示原因且没有部分改名。
10. 关闭全部文档后点击按钮，确认显示可理解的提示。

本例没有监听所有 Photoshop 事件，因此手动切换文档或选择后请点击刷新；真正执行操作时仍会重新读取并校验目标。文件选择和 JSON 写入使用 UXP 文件 API，不需要为了写一个普通报告而获取 Photoshop 模态锁。[文件 API](https://developer.adobe.com/photoshop/uxp/2022/uxp-api/reference-js/modules/uxp/persistent-file-storage/storage)

<a id="manifest"></a>

## 7. Manifest 配置详解

### 7.1 常用字段

| 字段 | 作用 | 常见错误 |
| --- | --- | --- |
| `manifestVersion` | 清单格式；主示例使用数字 `5` | 写成 Photoshop 版本，或只改版本号却不迁移字段 |
| `id` | 插件身份 | 开发版和发布版随意换 ID，导致重复安装或数据隔离 |
| `name` | 插件显示名称 | 与用户看到的市场条目不一致 |
| `version` | 插件自身版本，如 `1.2.3` | 与宿主版本混淆 |
| `main` | 相对插件根目录的入口文件 | 指向源码，但打包目录只有构建后的文件 |
| `host.app` | Photoshop 使用 `PS` | 写成 `Photoshop`、`PHSP` 或照抄其他宿主的值 |
| `host.minVersion` | 最低支持的 Photoshop 版本 | 宣称支持未测试的旧版 |
| `host.data.apiVersion` | Photoshop 专用 API 模型 | 漏配后照用 API v2 示例 |
| `entrypoints` | 面板 / 命令入口 | Manifest ID 与 `entrypoints.setup()` 中的 key 不一致 |
| `requiredPermissions` | 文件、网络等平台权限 | 配置后未重新加载，或申请无关权限 |
| `icons` | 插件和入口图标 | 引用未打包图片、大小写不匹配 |

清单字段参见 [UXP Manifest](https://developer.adobe.com/uxp/guides/explanation/concepts/manifest/)。`host.data` 中 Photoshop 专用字段另见 [Photoshop Manifest 属性](https://developer.adobe.com/photoshop/uxp/2022/guides/uxp-guide/uxp-misc/manifest-v4/photoshop-manifest/)；页面路径含 `manifest-v4` 不表示新项目必须退回 v4。

### 7.2 权限按功能声明

| 功能 | 常用声明 | 设计建议 |
| --- | --- | --- |
| 插件自己的数据 / 临时目录 | 默认沙箱访问 | 通过 UXP 获取目录，不硬编码安装位置 |
| 用户选择文件或目录 | `localFileSystem: "request"` | 文件选择器和 Entry 是主线 |
| 业务要求广泛路径访问 | `localFileSystem: "fullAccess"` | 只在需求确实需要时使用；仍受 OS 权限限制 |
| 访问服务器 | `network.domains` | 指定实际需要的协议、域名和端口 |
| 剪贴板 | `clipboard: "read"` 或 `"readAndWrite"` | 与功能相匹配 |
| 打开外部 URL / 文件 | `launchProcess` | 不等价于完整 shell 或进程 API |
| 内嵌网页 | `webview` | 核对目标运行时的容器、加载和通信限制 |
| 原生 addon | `enableAddon` 及 addon 配置 | 进入 Hybrid 的 SDK 与发布流程 |

网络权限片段示例，需合并到完整 Manifest 的同一个 `requiredPermissions` 对象中：

```json
{
  "requiredPermissions": {
    "localFileSystem": "request",
    "network": {
      "domains": ["https://api.example.com"]
    }
  }
}
```

上面是配置片段，不是可独立加载的清单。`api.example.com` 是占位域名，部署时替换为自己的服务。权限模型详见 [Manifest v5](https://developer.adobe.com/photoshop/uxp/2022/guides/uxp-guide/uxp-misc/manifest-v5/)。

<a id="dom"></a>

## 8. Photoshop DOM 与对象定位

### 8.1 常用对象

```text
app
├── documents
├── activeDocument
│   ├── id / title / width / height
│   ├── layers          # 文档顶层图层集合
│   │   └── group.layers
│   ├── activeLayers    # 当前选中的图层
│   ├── selection       # 像素选区，与 activeLayers 不同
│   ├── historyStates
│   └── saveAs
└── 其他应用级能力
```

Photoshop DOM 与 ExtendScript DOM 并不完全一致。使用 [Document](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/classes/document) 和 [Layer](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/classes/layer) 的当前参考确认属性名、读写权限、同步 / 异步和最低版本。

### 8.2 遍历组内图层

下面是独立工具函数，返回普通对象。大型文档应按需求减少字段和读取次数。

```javascript
function collectLayerTree(psDoc) {
  const { constants } = require("photoshop");
  const result = [];

  function visit(layers, parentID, depth) {
    for (const layer of layers) {
      result.push({
        id: layer.id,
        parentID,
        depth,
        name: layer.name,
        kind: String(layer.kind)
      });
      if (layer.kind === constants.LayerKind.GROUP) {
        visit(layer.layers, layer.id, depth + 1);
      }
    }
  }

  visit(psDoc.layers, null, 0);
  return result;
}
```

### 8.3 定位对象的原则

- 名称可以重复、可以修改；不要把名称作为唯一标识。
- 索引会随图层移动、插入和删除改变；不要保存“第 3 个图层”作为长期目标。
- 在当前文档会话中优先使用文档 ID 和图层 ID，并在每次操作前确认仍存在。
- 文档关闭、重开以及某些转换后，旧 ID 或代理对象可能不再适用；跨会话数据需要独立业务标识与重新匹配策略。
- 区分“用户点击时选中的文档”和“异步工作完成时的活动文档”。后台网络请求返回后，后者可能已经变化。
- 图层 DOM 属性读取可能跨越宿主边界；不要在每次 UI 重绘时读取整棵树的全部属性。

<a id="modal"></a>

## 9. 模态执行、撤销与取消

### 9.1 三件不同的事

| 机制 | 解决的问题 | 不负责的事情 |
| --- | --- | --- |
| `executeAsModal` | 修改宿主状态时获得互斥执行范围 | 不自动将任意业务变成单个事务 |
| `suspendHistory` / `resumeHistory` | 合并或回滚指定文档的历史 | 不回滚网络请求或已写入的磁盘文件 |
| `busy` / 任务队列 | 防止本插件重复执行或交叉操作 | 不替代宿主模态锁 |

模态函数完成或抛出后才退出模态范围。创建没有 `await` 的 Promise 然后提前返回，会让后续写操作跑到范围外。[executeAsModal](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/executeasmodal)

### 9.2 模态范围应尽量短

推荐流程：

```text
读取输入 → 用户选择文件 / 获取服务数据 → 校验结果
       → 申请模态 → 重新校验目标 → 修改文档 → 完成历史 → 释放模态
       → 更新 UI / 写业务日志
```

不要在获得模态锁后等待用户登录、长时间 HTTP 请求或无期限的轮询。耗时计算若能在宿主外完成，应先计算，再短时间写回。

`executeAsModal` 的 `timeOut` 从 Photoshop 25.10 开始提供，单位是**秒**，控制遇到其他模态状态时的等待尝试时间；它不是操作总时长上限。[ExecuteAsModalOptions](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/objects/returnobjects/executeasmodaloptions)

### 9.3 取消与回滚的边界

长任务检查 `context.isCancelled`，适时报告进度，并允许事件循环处理取消。捕获异常后继续抛出，不要先吞掉异常又把操作标记为成功。

历史挂起只作用于指定文档。多文档批处理应为每个文档设计独立的历史或副本策略；已经导出的 PNG、已经上传的数据、已经关闭的其他文件，不会因为文档历史回滚而自动恢复。

正式插件还要决定：取消后删除哪些临时文件、保留哪些成功输出，以及向用户报告“全部取消”还是“部分完成”。这些是业务语义，应写进任务结果，而不是由一个笼统的 `catch` 隐藏。

<a id="batchplay"></a>

## 10. 使用 batchPlay 补足 DOM

### 10.1 使用顺序

优先查 DOM；DOM 没有覆盖时，使用 `require("photoshop").action.batchPlay()` 调用 Action Manager 描述符。它采用 actionJSON，不能靠猜属性名替代录制与验证。[batchPlay 官方说明](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/batchplay)

获取描述符的推荐方法：在 Photoshop 动作面板录制最小操作，选中动作步骤，使用开发者模式下的 **Copy As JavaScript**；也可使用 Plugins → Development 中的动作命令录制功能。录制完清理文档名、固定 ID、绝对路径和与目标功能无关的参数。

### 10.2 必须同时处理抛错和结果内错误

某些调用失败会拒绝 Promise，某些 Photoshop 命令失败会在结果中返回 `_obj: "error"`。封装时同时检查两种情况：

```javascript
async function batchPlayChecked(descriptors) {
  const { action } = require("photoshop");
  const results = await action.batchPlay(descriptors, {});
  const failure = results.find((item) => item && item._obj === "error");
  if (failure) {
    const error = new Error(failure.message || "Photoshop 命令执行失败。");
    error.number = failure.result;
    throw error;
  }
  return results;
}

async function readDocumentTitle(documentID) {
  const [result] = await batchPlayChecked([{
    _obj: "get",
    _target: [
      { _property: "title" },
      { _ref: "document", _id: documentID }
    ],
    _options: { dialogOptions: "dontDisplay" }
  }]);
  return result.title;
}
```

这里读取标题只是演示结构，实际业务直接使用 `psDoc.title` 更清楚。修改状态的描述符仍应在模态范围内执行。

### 10.3 常见字段与维护要求

| 字段 | 常见含义 |
| --- | --- |
| `_obj` | 命令或描述符对象类型，如 `get`、`set`、`make` |
| `_target` | 目标引用链 |
| `_ref` | 引用对象类型，如 `document`、`layer` |
| `_id` | 对象 ID |
| `_enum` / `_value` | 枚举类型及值 |
| `_unit` / `_value` | 单位及数值 |
| `_options.dialogOptions` | 命令对话框行为 |

将描述符集中放在宿主适配层，并标明录制的 Photoshop 版本、输入条件和验证样本。不要为方便而默认打开 `synchronousExecution`，也不要把多个写命令用 `Promise.all()` 并发执行。需要单步撤销时使用历史挂起机制；旧 `historyStateInfo` 不是新项目的推荐方案。[batchPlay 选项](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/batchplay)

<a id="storage"></a>

## 11. 文件、目录与持久化

### 11.1 先区分三类路径

| 位置 | 取得方式 | 适合保存 |
| --- | --- | --- |
| 插件安装目录 | `getPluginFolder()` | 随包分发的只读资源 |
| 插件数据目录 | `getDataFolder()` | 设置、缓存索引、迁移后的用户配置 |
| 插件临时目录 | `getTemporaryFolder()` | 中间文件，不依赖长期保留 |
| 用户选择的文件 | `getFileForOpening()` / `getFileForSaving()` | 导入文件、导出结果 |
| 用户选择的目录 | `getFolder()` | 批量导出目录 |

主示例采用 Entry API：文件和文件夹以 UXP `File` / `Folder` 表示，权限来自沙箱或选择器。取得 `nativePath` 只说明有一个路径字符串，不等于这个字符串可以传给所有 Photoshop 或 Node 风格接口。[UXP 文件存储参考](https://developer.adobe.com/photoshop/uxp/2022/uxp-api/reference-js/modules/uxp/persistent-file-storage/storage)

### 11.2 持久化自己的配置

下面的函数可以独立使用。只写插件数据目录，不涉及用户文档：

```javascript
async function savePreferences(preferences) {
  const { storage } = require("uxp");
  const folder = await storage.localFileSystem.getDataFolder();
  const file = await folder.createFile("preferences.json", {
    overwrite: true
  });
  await file.write(JSON.stringify({
    schemaVersion: 1,
    preferences
  }, null, 2), {
    format: storage.formats.utf8
  });
}
```

生产版本应处理首次运行、文件损坏和配置版本升级。需要可靠持久化时，采用临时文件与替换策略，并在写入失败时保留上一份有效配置。不要将密码、长期服务密钥或未经处理的文档内容写入普通配置文件。

### 11.3 Session token 与 Persistent token

| 类型 | 用途 | 生命周期 / 限制 |
| --- | --- | --- |
| `createSessionToken(entry)` | 某些 `batchPlay` 文件描述符使用的会话引用 | 当前会话使用，不能作为下次启动的授权凭证 |
| `createPersistentToken(entry)` | 保存将来重新取得 Entry 的引用 | 可能因移动、删除或权限变化失效 |
| `getEntryForPersistentToken(token)` | 将持久令牌恢复为 Entry | 恢复成功后，真正读写仍可能失败 |

DOM 如 `app.open()`、`saveAs.png()` 通常直接接受 UXP File Entry；动作描述符若要求 `_path` token，则按该描述符的规则生成 session token。不要把两种调用约定混用。[文件系统 token 参考](https://developer.adobe.com/photoshop/uxp/2022/uxp/reference-js/Modules/uxp/Persistent%20File%20Storage/FileSystemProvider/)

### 11.4 记住用户选择的导出目录

```javascript
async function getExportFolder() {
  const fs = require("uxp").storage.localFileSystem;
  const key = "layer-toolkit.export-folder";
  const savedToken = localStorage.getItem(key);

  if (savedToken) {
    try {
      const entry = await fs.getEntryForPersistentToken(savedToken);
      if (entry.isFolder) return entry;
    } catch (error) {
      console.warn("已保存目录需要重新授权。", error.message);
    }
    localStorage.removeItem(key);
  }

  const folder = await fs.getFolder();
  if (!folder) return null;
  const token = await fs.createPersistentToken(folder);
  localStorage.setItem(key, token);
  return folder;
}
```

调用者在 `createFile()` 或写入失败时，也要提供“重新选择目录”的入口。不要无限重试旧 token。批量生成文件名时应处理路径分隔符、Windows 保留名、末尾空格 / 点、大小写冲突和重复名称；同名文件的覆盖策略应让用户事先理解。

<a id="export"></a>

## 12. 保存图像与批量导出

### 12.1 导出当前文档的 PNG 副本

下面是一个独立的 UXP 插件函数，不会自动添加按钮。若接到第 6 节面板，应在 `runTask` 的任务中调用它，并根据返回结果更新状态。该函数的示范范围限定为 **RGB / 8 bit**，不自动改变源文档的色彩模式或位深。

```javascript
async function exportCurrentPNG() {
  const { app, core, constants } = require("photoshop");
  const fs = require("uxp").storage.localFileSystem;

  if (app.documents.length === 0) {
    throw new Error("请先打开文档。");
  }
  const sourceID = app.activeDocument.id;
  const entry = await fs.getFileForSaving("document-copy.png", {
    types: ["png"]
  });
  if (!entry) return { cancelled: true };

  await core.executeAsModal(async () => {
    if (app.documents.length === 0 ||
        app.activeDocument.id !== sourceID) {
      throw new Error("源文档已关闭或当前文档已经改变。");
    }
    const psDoc = app.activeDocument;
    if (psDoc.mode !== constants.DocumentMode.RGB ||
        psDoc.bitsPerChannel !== constants.BitsPerChannelType.EIGHT) {
      throw new Error("此示例仅导出 RGB / 8 bit 文档。");
    }

    await psDoc.saveAs.png(entry, {
      method: constants.PNGMethod.QUICK,
      compression: 6,
      interlaced: false
    }, true);
  }, { commandName: "导出 PNG 副本" });

  return { cancelled: false, name: entry.name };
}
```

第三个参数 `true` 表示另存副本；PNG 的压缩参数是压缩率 / 文件体积相关选项，不是 JPEG 式的有损画质百分比。[Document.saveAs](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/classes/document)、[PNGSaveOptions](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/objects/saveoptions/pngsaveoptions)、[常量参考](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/modules/constants)

注意：将此函数接入第 6 节时，按钮任务需要根据返回值调用 `setStatus()`，例如输出成功或取消信息。文件选择器可能先创建目标文件；后续保存失败时可能留下空文件或不完整结果，正式导出器需要定义临时文件和失败清理策略。

### 12.2 导出前必须定义的图像语义

| 问题 | 需要明确的选择 |
| --- | --- |
| 范围 | 全画布、选区、图层边界、画板，还是图层效果边界 |
| 可见性 | 当前合成结果，还是只导出指定图层 |
| 透明背景 | 保留 alpha，或用指定颜色铺底 |
| 色彩 | 保留配置文件，还是转换至指定 RGB 配置文件 |
| 尺寸 | 原始像素、缩放比例，还是固定宽高 |
| 智能对象 / 文字 | 保留可编辑数据，还是只交付栅格结果 |
| 文件覆盖 | 跳过、追加序号、报错，还是用户明确选择覆盖 |
| 源文档 | 在副本处理，还是对源文档做可撤销修改 |

“保存整个文档的 PNG”不等价于“导出当前选中图层”。后者可能需要复制到临时文档、管理可见性、裁切与坐标，并处理组、剪贴蒙版、调整图层和混合模式对合成的影响。

### 12.3 批量处理流程

建议将批量导出拆成以下阶段：

1. 收集目标并生成导出计划，预先展示数量、尺寸和命名冲突。
2. 让用户一次选择输出目录。
3. 按序处理每个文档或图层，限制同时驻留的临时文档数量。
4. 对每个对象记录成功、跳过、失败、取消及原因。
5. 输出完成报告，支持重试失败项。

若为导出建立临时文档，可研究 `context.hostControl.registerAutoCloseDocument()` 来管理模态退出时的关闭，具体提交和清理行为在目标版本验证。关闭动作只针对本插件创建的临时文档，不能用“关闭当前文档”替代明确目标。[自动关闭文档](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/executeasmodal)

<a id="imaging"></a>

## 13. Imaging API 与像素处理

### 13.1 什么时候使用

需要预览缩略图、读取像素、分析颜色、处理选区 / 蒙版或写入像素层时，使用 `require("photoshop").imaging`。这比将整份 PSD 序列化给前端再解析更直接。只有实际需要的像素范围才传到 JS。

常见入口包括 `getPixels()`、`putPixels()`、`getLayerMask()`、`getSelection()`、`createImageDataFromBuffer()`。写入像素属于宿主修改，需要相应的模态和历史管理。[Imaging API](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/imaging)

### 13.2 读取合成图缩略像素

以下函数固定目标文档 ID，将最长边限制为约 256 像素，返回尺寸等信息。它只示范读取和释放，不创建 UI 图片，也不将结果写回：

```javascript
async function inspectThumbnail() {
  const { app, imaging } = require("photoshop");
  if (app.documents.length === 0) {
    throw new Error("请先打开文档。");
  }
  const psDoc = app.activeDocument;
  const targetSize = psDoc.width >= psDoc.height
    ? { width: Math.min(256, psDoc.width) }
    : { height: Math.min(256, psDoc.height) };

  const result = await imaging.getPixels({
    documentID: psDoc.id,
    targetSize,
    colorSpace: "RGB",
    colorProfile: "sRGB IEC61966-2.1",
    componentSize: 8,
    applyAlpha: false
  });

  try {
    const imageData = result.imageData;
    const bytes = await imageData.getData({ chunky: true });
    return {
      width: imageData.width,
      height: imageData.height,
      components: imageData.components,
      hasAlpha: imageData.hasAlpha,
      byteLength: bytes.byteLength,
      sourceBounds: result.sourceBounds,
      level: result.level
    };
  } finally {
    result.imageData.dispose();
  }
}
```

图层、色彩与内存规则以 [Imaging API 参数和数据类型](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/imaging) 为准：不要默认总有四通道；16 bit 的默认值域与常见全 16 bit 范围有区别；32 bit HDR 可能超过 0～1；读取范围可能被裁剪，缩略结果还可能来自缓存层级。`applyAlpha: true` 会合成到白色并移除 alpha，不能把它理解成“保留透明度”。

### 13.3 算法接入的工程约定

进入算法前建立明确的数据契约：宽高、通道布局、分量位数、alpha 语义、色彩空间、配置文件、坐标原点。返回时校验数据长度与契约一致，避免把 RGB 当 RGBA 或把浮点缓冲区当 8 bit。

大图内存可先用下式估计，不含宿主副本、临时数组、历史状态和编码开销：

```text
像素缓冲字节数 = 宽度 × 高度 × 通道数 × 每通道字节数

8000 × 8000 × 4 × 1 = 256,000,000 bytes，约 244 MiB
8000 × 8000 × 4 × 4 = 1,024,000,000 bytes，约 977 MiB
```

限制并发、避免反复转 Base64、处理后及时释放图像对象。若算法只分析预览，不要读取原分辨率；若算法必须处理全图，可评估分块及边界重叠策略。性能优化以端到端测量为准，不能只比较 JS 与 C++ 单次循环耗时。

<a id="ui"></a>

## 14. UI、生命周期与事件

### 14.1 界面技术选择

最小项目使用内置 Spectrum widgets，减少依赖。复杂项目可以评估 Spectrum Web Components 或已在目标 UXP 版本验证过的框架。不要因为某个 React / Vue 组件在网页里正常，就假定其 DOM、CSS、Canvas、Worker 或运行时依赖都可用。

先检查目标 UXP 版本的能力，再决定布局和框架。使用 Flexbox、简单间距和原生主题适配作为起点；CSS 新特性、WebView 和 SWC 的 feature flag 应与对应文档一起管理，不凭旧文章认定某功能永远不支持。[Spectrum UXP](https://developer.adobe.com/photoshop/uxp/2022/uxp/reference-spectrum/)、[UXP 版本说明](https://developer.adobe.com/uxp/uxp-api/versions)

### 14.2 生命周期职责

| 阶段 | 推荐工作 |
| --- | --- |
| 加载 / create | 建立 UI、读取轻量配置、创建状态管理对象 |
| show | 快速刷新可见数据，按需启动任务 |
| hide | 暂停无意义刷新、隐藏面板相关工作 |
| destroy | 取消定时器、移除监听、结束可中止任务 |
| 用户执行命令 | 校验输入，提交一个可追踪业务任务 |

不要在生命周期回调中等待一次长网络请求或全量扫描。Manifest v5 的相关回调有较短超时约束，且宿主生命周期行为可能有局限；对重要状态的正确性，应在命令执行时重新确认，而非仅依赖 show / hide。[Manifest 生命周期](https://developer.adobe.com/photoshop/uxp/2022/guides/uxp-guide/uxp-misc/manifest-v5/)

### 14.3 监听宿主事件

Photoshop 的 `action.addNotificationListener(events, callback)` 可订阅动作事件；卸载时使用同一回调引用移除监听。生产插件只订阅所需事件。开发者模式下的 `all` 监听更适合探索描述符。[photoshopAction](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/photoshopaction)

建议采用以下方式组织状态同步：

```text
收到 open / close / select 等已验证事件
    → 将相关状态标记为失效
    → 合并短时间内重复事件
    → 重新读取少量必要数据
    → 更新面板
```

不要把事件回调里的临时描述符当作长期完整文档模型；也不要每收到一次选择事件就立即修改文档，否则容易形成“插件写入 → 事件 → 再写入”的循环。需要自动执行时，应加入来源识别、去重和显式业务状态。

### 14.4 面板可用性

- 为无文档、无选择、处理中、取消、失败、部分完成提供明确状态。
- 执行期间禁用冲突按钮，但保留有效的取消途径。
- 长列表分批渲染，不在每次输入时遍历所有图层。
- 用 `textContent` 展示图层名和服务返回值；不要把不可信字符串拼入 `innerHTML`。
- 测试深浅主题、停靠 / 浮动、窄面板、HiDPI、中英文和长名称。
- 文本输入应考虑输入法组合事件；不要在用户尚未完成中文输入时触发破坏性操作。

<a id="network"></a>

## 15. 网络服务与本地程序集成

### 15.1 典型服务流程

```text
用户选择目标 → 明确上传内容 → 生成输入数据 → 请求业务服务
    → 验证响应 → 再次确认目标文档 → executeAsModal 写回结果
```

网络访问在 Manifest 中声明实际域名。下面仅示范调用结构，`api.example.com`、接口路径和返回协议都需要由业务系统实现：

```javascript
async function fetchPresetList(accessToken) {
  const response = await fetch("https://api.example.com/v1/presets", {
    method: "GET",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Accept": "application/json"
    }
  });
  if (!response.ok) {
    throw new Error(`读取预设失败，HTTP ${response.status}`);
  }
  const data = await response.json();
  if (!data || !Array.isArray(data.items)) {
    throw new Error("服务返回的数据结构不符合约定。");
  }
  return data.items;
}
```

客户端只能携带业务允许下发的令牌。需要保密的服务端密钥应留在后端，不能写进 JS、Manifest 或构建产物。生产实现还需超时、中止、限流、响应体大小上限和 schema 校验；是否支持某种中止 API 要核对目标 UXP 运行时。

访问失败应分别检查 Manifest 域名、代理、证书、服务端鉴权、重定向目的地和服务端策略。网络权限不能替代服务器认证。[Manifest 网络权限](https://developer.adobe.com/uxp/guides/explanation/concepts/manifest/)

### 15.2 本地服务、Python、命令行程序

如果已有 Python 图像服务或本地 C++ 进程，常见选择有：

| 方式 | 适用情况 | 需要处理 |
| --- | --- | --- |
| 独立程序提供 localhost HTTP / WebSocket | 已有独立服务或跨应用复用 | 服务发现、端口、会话认证、断连和版本协议 |
| UXP Hybrid | 需要在插件内直接使用原生代码 | addon ABI、线程、打包和签名 |
| Photoshop C++ 插件与 UXP 消息桥 | 已有宿主原生插件 | 组件 ID、消息协议与双方生命周期 |

`shell.openExternal()` / `openPath()` 主要用于打开 URI 或文件，不应当成 `child_process.exec()` 的替代。通用迁移文档也明确指出 UXP 没有完整 CEP 进程 API 的对应物。[CEP 到 UXP 的能力映射](https://developer.adobe.com/uxp/migration-center/)

本地服务建议只绑定回环地址，使用每次会话的随机凭证，并限制可执行操作和路径范围。即使服务在 localhost，端口也可能被其他本地进程访问；不能仅以“在本机”作为认证条件。

<a id="engineering"></a>

## 16. 工程结构、构建与版本兼容

### 16.1 建议目录

```text
plugin-project/
├── manifest.json
├── package.json
├── package-lock.json
├── src/
│   ├── main.js
│   ├── ui/                  # 面板、控件和状态展示
│   ├── domain/              # 命名、计划、校验等纯逻辑
│   ├── photoshop/           # DOM、batchPlay、模态、像素访问
│   ├── services/            # 网络、文件、配置适配
│   └── diagnostics/         # 日志和错误分类
├── static/                  # HTML、CSS、图标、内置资源
├── dist/                    # 实际加载 / 打包的完整输出
└── tests/                   # 纯逻辑和集成验证资源
```

把 Photoshop 依赖集中在适配层，纯逻辑返回操作计划。这样可以在 Node 中测试命名、文件冲突、坐标和数据验证，再在 Photoshop 中测试真正的读写行为。

### 16.2 构建器需要做什么

无构建示例最容易入门。引入 TypeScript、框架或大量依赖后，再选择构建器，并确认：

1. 保留 `require("photoshop")`、`require("uxp")` 等宿主导入，按构建器语法配置 external。
2. 不将一个假 `photoshop` npm 包打入产物，不自动填充大量不兼容 Node 模块。
3. 将代码输出为目标 UXP 支持的格式，避免未验证的 ESM、动态加载或基于 `eval` 的开发模式。
4. 复制 Manifest、HTML、CSS、图标和本地资源，保证路径相对于最终包根目录正确。
5. source map 在开发时可用于定位；发布时评估是否包含源码或敏感路径。
6. 构建应从确定的输入生成产物，版本与锁文件一起提交。

使用 UDT 当前提供的 Advanced / distribution folder 设置，把源 Manifest 作为稳定工程入口、让 UDT 加载构建输出，可避免构建时删除 `dist` 导致工程引用失效。也可加载固定且不会被反复移除的产物目录，关键是始终确认实际运行哪份代码。[UDT 与构建器](https://developer.adobe.com/uxp/guides/how-to/udt-deep-dive/plugin-workflows)

### 16.3 类型检查

Adobe 提供平台类型包 `@adobe/cc-ext-uxp-types`。它覆盖 UXP 平台，不等于覆盖全部 Photoshop 宿主 API。宿主类型定义按 Adobe 当前指南另行配置。不要同时粗暴引入完整浏览器 DOM 类型与 UXP DOM 类型，避免 `Document`、`HTMLElement` 等全局定义冲突。[TypeScript 与 IntelliSense](https://developer.adobe.com/uxp/guides/how-to/typescript/)

类型能减少拼写错误，但不能证明当前色彩模式允许某操作，也不能保证一个 `batchPlay` 描述符在所有版本都有效。把最低版本、能力探测和宿主回归作为独立检查。

### 16.4 兼容策略

| 层次 | 应固定和记录的内容 |
| --- | --- |
| 产品 | Photoshop 最低 / 最高已测试版本、系统和架构 |
| 运行时 | UXP 版本及所需功能 |
| API | 新增方法的最低版本、异常行为、动作描述符差异 |
| 数据 | 配置 schema、缓存格式、服务器协议 |
| 发布 | 插件 ID、插件版本、包类型、分发渠道 |
| 原生 | SDK 版本、编译器、ABI、依赖库及架构 |

如果一个可选功能不可用，应隐藏或禁用该功能并说明原因。必需 API 缺失则拒绝运行，不能继续执行“差不多”的替代操作。比较版本要解析数字段；字符串比较可能把 `27.10` 排在 `27.4` 前面。

<a id="debugging"></a>

## 17. 调试与故障定位

### 17.1 固定诊断信息

可以在插件 Console 执行以下代码记录环境：

```javascript
function logRuntimeInfo() {
  const { app, core } = require("photoshop");
  const { versions } = require("uxp");
  console.log({
    photoshop: app.version,
    photoshopApiVersion: core.apiVersion,
    uxp: versions.uxp,
    plugin: versions.plugin
  });
}
```

版本属性参见 [宿主环境信息](https://developer.adobe.com/uxp/guides/how-to/recipes/host-info/) 和 [Photoshop core](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/photoshopcore)。再记录操作名称、文档 ID、图层数量和错误编号，避免默认记录整幅像素、完整文档路径或访问令牌。

### 17.2 按症状定位

| 症状 | 优先检查 | 处理方向 |
| --- | --- | --- |
| UDT 看不到宿主 | Photoshop 是否运行、版本、双方开发者模式 | 确认连接目标，查看 UDT 日志 |
| Manifest validation failed | JSON、字段类型、入口路径、权限、版本 | 使用 Details 中的具体错误定位 |
| 面板空白 | `main`、JS 语法错误、资源路径、入口 ID | 打开 Console 和 Elements |
| `require("photoshop")` 失败 | 是否在浏览器 / Node 中运行，构建是否改写导入 | 回到真实 UXP 环境 |
| 修改操作要求模态状态 | 是否在 `executeAsModal` 中，异步调用是否被等待 | 修复调用边界 |
| Host is in a modal state | Photoshop 对话框、其他插件、重复点击、漏 `await` | 串行任务，保留可理解的重试入口 |
| `batchPlay` 没抛错但无结果 | 返回数组中是否有 `_obj: "error"` | 将结果内错误转为失败 |
| 操作“不可用” | 活动文档、图层类型、锁、模式、位深、选区 | 缩小到最小样本并手工验证命令 |
| 文件权限错误 | Manifest、Entry 来源、token、OS 访问权限 | 重新选择文件 / 目录并重新加载权限 |
| 重载后仍是旧代码 | UDT 实际加载目录、Watch、构建输出 | 打印版本标识，确认源码到产物链 |
| 任务偶尔改错文档 | 异步前后依赖 `activeDocument` | 固定 ID，获得模态后再校验 |
| 内存持续增长 | 图像对象、缓冲副本、监听、缓存 | `dispose`、清理引用、限制并发 |
| 开发正常但 `.ccx` 失败 | 包缺文件、ID 冲突、绝对路径、开发机依赖 | 在没有开发加载的环境安装测试 |
| 原生模块无法加载 | 架构、签名、依赖库、SDK / ABI、安装位置 | 查看宿主日志及原生加载错误 |

UDT 调试器提供 Console、Sources、Elements 等工具；启动阶段错误可使用 Break on start。更改 Manifest 后重新加载，不要只依赖 Watch。[UDT 调试操作](https://developer.adobe.com/uxp/guides/how-to/udt-deep-dive/plugin-workflows)

### 17.3 缩小问题的方法

先用一个 RGB / 8 bit、小尺寸、单个未锁定普通图层的测试文档复现；再逐步加入组、蒙版、智能对象、位深和色彩模式。将网络请求、界面和算法暂时替换为固定输入，确认失败发生在纯逻辑、UXP 平台还是 Photoshop API 边界。

上报问题时提供完整版本、最小插件、最小文档、步骤、预期、实际结果及脱敏日志。截图只能描述现象，无法代替代码和触发条件。

<a id="testing"></a>

## 18. 测试与验收

### 18.1 分层验证

| 层次 | 验证对象 | 执行位置 |
| --- | --- | --- |
| 静态检查 | JSON、JS 语法、路径和入口一致性 | 开发机 / CI |
| 纯逻辑测试 | 命名、排序、坐标、格式验证、冲突策略 | Node / 常规测试框架 |
| 宿主集成 | DOM、batchPlay、历史、选择、像素、保存 | 实际 Photoshop |
| 安装验证 | 包内容、权限、升级、卸载、原生依赖 | 干净用户环境 / 测试机 |
| 人工体验 | 主题、输入法、取消、进度、错误提示 | 实际面板 |

不要用伪造的 `photoshop` mock 来证明 Photoshop API 正确；mock 只适合检查自己的分支与数据流。文档代码通过 `node --check` 只代表 JS 语法可解析，不能证明 UXP UI 或宿主调用成功。

### 18.2 建议测试矩阵

| 维度 | 至少覆盖 |
| --- | --- |
| 文档状态 | 无文档、单文档、多文档、未保存文档、处理中被关闭 |
| 图层 | 普通层、背景层、组、文字、智能对象、隐藏、锁定、多选 |
| 图像 | RGB / CMYK / 灰度；8 / 16 / 32 bit；透明与不透明 |
| 大小 | 小图、数千图层、大尺寸图，按业务目标决定上限 |
| 文件 | 中文、空格、emoji、同名、只读、移动后 token、磁盘写入失败 |
| 时序 | 连续点击、模态冲突、用户取消、切换文档、重载插件 |
| 网络 | 离线、超时、401、429、5xx、无效 JSON、异常大响应 |
| 环境 | 最低支持宿主、本机 27.4、其他宣称支持的系统 / 架构 |
| 发布 | 初装、升级、同 ID 开发版冲突、正式包资源完整性 |

并非每个功能都必须支持所有模式；对于不支持的情况，验收目标是提前提示或明确拒绝，且不会留下部分破坏。

### 18.3 发布前完成条件

- 关键操作在明确目标上执行，失败不误报成功。
- 撤销、取消和临时文档清理符合产品约定。
- 对用户文件的覆盖行为明确，失败输出可识别。
- 最低支持版本经过实际宿主测试。
- `.ccx` 安装后的代码与开发加载版本行为一致。
- 原生二进制在所声明的架构和干净机器上验证。
- 对依赖网络服务的功能，有离线和认证失效处理。

<a id="distribution"></a>

## 19. 打包、安装与发布

### 19.1 UXP 发布流程

1. 确定插件 ID、名称、版本、最低宿主版本和分发渠道。
2. 使用生产构建生成完整目录，检查 Manifest 入口、图标、静态资源和依赖。
3. 在 UDT 中选择目标插件，执行 Package，生成 `.ccx`。
4. 在 UDT 中 Unload 开发版本，再在独立测试环境安装该 `.ccx`，避免同时运行同 ID 工程。
5. 通过 Creative Cloud 桌面应用验证安装、启动、升级和卸载。
6. 按需提交 Marketplace，或进行独立 / 企业分发。

`.ccx` 是 ZIP 结构的安装包，优先用 UDT 创建。普通 UXP 包不沿用 CEP 的 ZXP 签名流程；Hybrid 内的原生二进制另有签名要求。[UXP 打包指南](https://developer.adobe.com/uxp/guides/how-to/distribution/package/)

### 19.2 Marketplace 与独立分发

| 渠道 | 需要准备 |
| --- | --- |
| Marketplace | Developer Distribution 条目、匹配的插件 ID、包、截图、说明、支持渠道和必要的隐私信息 |
| 独立 `.ccx` | 用户可获得的安装包、更新策略、兼容说明和支持入口 |
| 企业部署 | 与管理员确认 Creative Cloud / UPIA / Admin Console 对应部署方式 |

发布时以门户当前提交规则为准，不承诺固定审核时间。安装路径、管理方式和许可关联也不应通过手动复制内部目录替代。[安装渠道](https://developer.adobe.com/uxp/guides/how-to/distribution/install/)、[提交与审核](https://developer.adobe.com/developer-distribution/creative-cloud/docs/guides/submission/overview)

同一渠道的更新保持稳定 ID。若同时进行 Marketplace 付费销售与外部分发，应检查渠道 ID / 授权关系，不能假设市场 entitlement 对外部购买者也有效；Adobe 当前打包指南专门说明了这类冲突。[插件 ID 与多渠道分发](https://developer.adobe.com/uxp/guides/how-to/distribution/package/)

### 19.3 升级与数据迁移

配置文件使用 `schemaVersion`，在升级时迁移旧字段并保留可恢复数据。缓存可重建，用户设置应单独对待。不要依赖安装目录可写或保存用户数据；应用升级、插件更新都可能替换安装内容。

建议每次发布记录：插件版本、构建提交、目标 Photoshop / UXP、构建工具和依赖版本、验收结果、已知限制。发生回归时，才能准确复现用户运行的版本。

<a id="native"></a>

## 20. C++ 原生插件开发路线

### 20.1 先确认需要哪一种宿主扩展

| 需求 | 原生方向 | 需要重点研究 |
| --- | --- | --- |
| 图像算法作为正式滤镜 | Filter 插件 | 滤镜协议、像素访问、模式支持、参数记录 |
| Photoshop 直接打开 / 保存自定义格式 | File Format 插件 | 格式探测、读写、图像元数据、错误恢复 |
| 导入 / 导出外部数据 | 对应 Import / Export 接口 | 数据流、文件协议、用户交互 |
| 原生自动化或宿主 suite 能力 | Automation 等 SDK 接口 | suite 获取、事件与生命周期 |
| 仅在面板内调用自己的算法库 | 优先评估 Hybrid | 不必为此自动选择传统滤镜插件 |

Photoshop C++ SDK 与云端 Photoshop API SDK、UXP Hybrid SDK 是不同产品。原生路线的具体类型、能力和支持范围以下载包为准。[Photoshop C++ 官方入口](https://developer.adobe.com/photoshop/)、[Hybrid 选型说明](https://developer.adobe.com/photoshop/uxp/2022/guides/hybrid-plugins/)

### 20.2 获取 SDK

通过 Adobe Developer 的 Photoshop 入口进入 SDK 下载页面，登录后取得对应平台 SDK。Adobe Tech Blog 提供了定位 Photoshop C++ SDK 的官方说明；该说明发表较早，控制台按钮名称可能变化，应寻找 Photoshop SDK / Downloads，不要误选云 API。[Adobe：定位和下载 C++ SDK](https://medium.com/adobetech/locate-and-download-the-photoshop-c-sdk-4f0e55f091ae)

下载后优先阅读：

1. README、发行说明、支持的宿主和编译器版本。
2. 与目标插件类型对应的示例工程。
3. 入口函数、参数结构、错误码、内存归属及线程约束。
4. PiPL 资源和平台资源生成规则。
5. 构建、安装、调试和签名说明。

常见头文件如 `PIFilter.h`、`PIFormat.h`、`PIGeneral.h` 属于 SDK，不应从 Photoshop 内部二进制猜测或从不明版本的散落头文件拼装。本文未下载需要登录获取的 C++ SDK，因此以下为工程实施指南，不提供声称已编译通过的 SDK 项目。

### 20.3 从官方样例构建，而不是空白工程

推荐第一次执行的目标是：**原样编译一个官方样例，并让本机 Photoshop 27.4 成功识别它。**

| 步骤 | 要确认的结果 |
| --- | --- |
| 选择接近需求的 SDK 样例 | 插件类型正确，符合目标宿主 |
| 使用 SDK 说明要求的工具链 | 不凭 Photoshop 年份推断 Xcode / Visual Studio 版本 |
| 配置 include、资源和链接设置 | 头文件、PiPL、平台资源都参与构建 |
| 构建 Apple Silicon 产物 | 不是只有 x86_64 的旧二进制 |
| 安装到单一开发测试位置 | 没有另一份同名旧插件干扰 |
| 启动宿主并验证入口 | 插件出现、能执行、能退出 |
| 再修改算法与 UI | 出错时可以与原样样例比较 |

一个 macOS 原生插件通常是 `.plugin` bundle，其中包含可执行文件、Info.plist 和资源。仅把普通 `.dylib` 改名为 `.plugin`，不等于实现 Photoshop 插件协议。

### 20.4 PiPL 与入口协议

PiPL 用于向宿主描述插件种类及能力，和 UXP 的 `manifest.json` 分属不同机制。传统原生开发一般需要同时正确完成：

- 宿主可识别的插件类型、名称和菜单分类等资源。
- 对应平台 / 架构的入口与导出配置。
- 插件声明支持的颜色模式、位深与操作条件。
- 与插件类型匹配的入口函数和生命周期响应。

Photoshop 的公开桥接文档展示了 Filter 使用 `PluginMain`、Automation 使用 `AutoPluginMain` 的区别，并通过 SDK 中的 `SPBasicSuite` 获取宿主 suites。函数签名、调用约定和消息结构以目标 SDK 头文件为准，不能跨插件类型复用同一个入口。[C++ 插件与 UXP 通信](https://developer.adobe.com/photoshop/uxp/ps_reference/media/cpp-pluginsdk/)

### 20.5 滤镜开发的处理阶段

可以按下面的阶段理解传统滤镜；这是职责示意，实际 selector 名称、调用顺序和可选能力以 SDK 样例为准：

```text
关于 / 参数处理
    → 校验输入与资源需求
    → 开始处理
    → 申请输入区域 → 处理像素块 → 提交输出区域
    → 根据协议继续下一块或结束
    → 释放本次调用的资源
```

像素处理最容易出错的地方：

| 问题 | 实现要求 |
| --- | --- |
| 行跨度 | 使用 SDK 返回的 row bytes / stride，不假定一行等于宽度乘四 |
| 通道 | 区分颜色分量、透明度、蒙版和当前处理 plane |
| 像素深度 | 分别处理 8 / 16 / 32 bit；值域和转换遵循 SDK |
| 选区 | 明确只处理选区还是由宿主合成，正确处理蒙版 |
| 大文档 | 使用 SDK 提供的大坐标 / 分块机制，避免窄整数溢出 |
| 邻域算法 | 为模糊、卷积等请求边缘扩展区域，避免分块接缝 |
| 取消 | 定期检查宿主的中止机制，及时释放资源 |
| 内存 | 明确宿主内存与插件内存的所有权，不跨调用持有临时指针 |
| 异常 | 不让 C++ 异常跨越 C ABI 边界，映射为宿主错误结果 |
| 重复调用 | 参数恢复、动作回放和重复执行都能得到一致结果 |

不要把“滤镜能处理当前普通图层”当成“已经支持智能滤镜、动作记录和所有图像模式”。这些都需要对应 SDK 能力、参数序列化与专项验收。

### 20.6 macOS 安装与架构检查

Adobe 列出的传统插件目录包括：

```text
版本目录：/Applications/Adobe Photoshop 2026/Plug-ins/
共享目录：/Library/Application Support/Adobe/Plug-Ins/CC/
```

开发阶段选择一种位置，避免同一个插件被重复扫描。需要系统目录写权限时通过正常授权或安装器处理；不要把测试模块塞进应用包的 `Contents/PlugIns/Required`。[Adobe 插件目录说明](https://helpx.adobe.com/photoshop/desktop/app-integrations/supported-plugins-and-extensions.html)

下面是检查命令模板，路径中的插件和可执行文件名需要替换为自己的产物；它们不执行安装：

```bash
file '/path/to/ImageFilter.plugin/Contents/MacOS/ImageFilter'
lipo -info '/path/to/ImageFilter.plugin/Contents/MacOS/ImageFilter'
otool -L '/path/to/ImageFilter.plugin/Contents/MacOS/ImageFilter'
codesign --verify --strict --verbose=2 '/path/to/ImageFilter.plugin'
```

Apple Silicon 原生 Photoshop 不能加载只有 Intel 架构的插件。支持两种 Mac 架构时，需要匹配的插件及其全部原生依赖；主程序本身是 Universal Binary 不能弥补插件缺少 arm64。[Adobe Apple Silicon 兼容说明](https://helpx.adobe.com/photoshop/desktop/troubleshoot/tool-feature-issues/photoshop-features-not-working-on-apple-silicon.html)

### 20.7 调试与发布

使用 Xcode / LLDB 调试自己的二进制，保留符号并确认加载的是当前构建。附加到 Photoshop 可能受到 macOS 调试权限和宿主签名约束，应遵循 SDK 的调试说明。不要把修改 Photoshop 主程序签名、关闭系统安全机制当成常规开发步骤。

发布时检查所有依赖库的安装路径、最低系统版本和架构；在没有 Xcode、没有自定义库搜索路径的机器上验证。面向 macOS 的正式分发应按 Apple Developer ID、签名与 notarization 的当前流程处理实际交付物，并保存对应构建符号。[Apple 公证说明](https://developer.apple.com/documentation/security/notarizing-macos-software-before-distribution)

<a id="hybrid"></a>

## 21. UXP Hybrid 开发路线

### 21.1 分工模型

```text
UXP 面板与业务调度
    ├── Photoshop DOM / batchPlay / Imaging
    └── require("image-core.uxpaddon")
            └── C++ 算法、编解码、已有原生库
```

Hybrid 的历史最低要求是 Photoshop 24.2.0，当前项目仍须按具体 SDK 与宿主组合选择。它使用专门的 UXP addon SDK；常规 Node 原生 `.node` 文件不能直接当作 `.uxpaddon` 加载。[Photoshop Hybrid 概述](https://developer.adobe.com/photoshop/uxp/2022/guides/hybrid-plugins/)、[UXP Hybrid SDK](https://developer.adobe.com/uxp/guides/how-to/hybrid-plugins/)

### 21.2 建立最小工程

1. 先完成一个普通 UXP 插件的加载、调试和打包。
2. 下载目标 Photoshop / UXP 支持的 Hybrid SDK，阅读其 README。
3. 加载 SDK 自带的 `template-plugin`，验证宿主、工具和目录布局。
4. 编译 `template-dev`，先导出一个小型、确定性的原生函数。
5. 在 JS 调用该函数，确认输入输出、错误和内存生命周期。
6. 再接入真实图像缓冲和算法，最后处理耗时任务与取消。

SDK 中的 `UxpAddon.h`、`UxpAddonShared.h` 等负责注册与桥接；不要手写猜测的 ABI 或只重命名一个任意动态库。JavaScript 的 `require("image-core.uxpaddon")` 必须对应实际导出的模块和 SDK 要求的布局。[构建 Hybrid 插件](https://developer.adobe.com/uxp/guides/how-to/hybrid-plugins/build)

### 21.3 Manifest 版本的文档差异

本次核验发现：

- Photoshop 的普通插件文档和本指南主示例使用 Manifest 5。
- 新的通用 Hybrid 构建页写明 Manifest 6 或以上，示例宿主为 Premiere。
- 本机 Adobe 内置 `com.adobe.cai.uxp.panel` 带有 addon 字段，但使用 Manifest 5；内置组件可能享有特殊加载条件。

因此，**不要用普通插件的 v5 模板直接推导所有第三方 Hybrid 项目都应使用 v5，也不要仅凭通用页面就宣称本机已验证 v6 Hybrid。** 应采用与 Photoshop 27.4 / UXP 9.2 匹配的 SDK 模板，确认 `host.app` 为 `PS`，并在 UDT 通过清单校验和原生加载。本文未执行该 SDK 验证。[普通 Manifest](https://developer.adobe.com/uxp/guides/explanation/concepts/manifest/)、[通用 Hybrid 清单要求](https://developer.adobe.com/uxp/guides/how-to/hybrid-plugins/build)

需要关注的 addon 配置片段如下；它不指定整个清单版本，不能单独作为插件加载：

```json
{
  "addon": {
    "name": "image-core.uxpaddon"
  },
  "requiredPermissions": {
    "enableAddon": true
  }
}
```

当前通用文档给出的多平台布局如下，实际仍与所用 SDK README 对照：

```text
plugin-root/
├── manifest.json
├── index.html
├── main.js
├── mac/
│   ├── arm64/image-core.uxpaddon
│   └── x64/image-core.uxpaddon
└── win/
    └── x64/image-core.uxpaddon
```

### 21.4 JS 与 C++ 的接口设计

建议原生接口只接收经过校验的普通数据或像素缓冲，返回明确结果。接口至少应约定：

| 项目 | 建议约定 |
| --- | --- |
| 协议 | 版本号、功能名、参数范围、结果结构 |
| 像素 | 宽高、通道数、位深、布局、色彩空间、alpha |
| 内存 | 是复制还是共享、谁持有、何时释放 |
| 异步 | Promise / 回调由桥接层如何完成，任务 ID 如何关联 |
| 取消 | 请求取消后何时停止、已分配内存如何回收 |
| 错误 | 错误码、用户可读消息、可重试状态 |
| 版本 | JS 包与二进制不匹配时明确拒绝 |

把算法主体做成不依赖 Photoshop 的独立 C++ 库，可以在常规单元测试和命令行程序中验证数值正确性，再通过薄 addon 层接入宿主。这是架构建议，不改变 SDK 对线程与内存的要求。

### 21.5 线程与 Photoshop SDK 调用

JS 调用原生函数不等于该函数就在 Photoshop 主线程，也不意味着任何线程都可以调用宿主 suites。通用 Hybrid 文档区分初始化 / 终止与 JS 调用所在的线程；使用 SDK 提供的异步与线程调度机制，避免阻塞脚本线程或持有无效宿主指针。[Hybrid 线程模型](https://developer.adobe.com/uxp/guides/how-to/hybrid-plugins/build)

Photoshop 的 Hybrid 指南还提供了接入 Photoshop C++ SDK 的专门入口路径。只有在实现相应接口并满足宿主协议后，addon 才能使用那些原生能力；它不会自动把算法注册为所有用户可见的 Filter 菜单项。[Hybrid 与 Photoshop C++ SDK](https://developer.adobe.com/photoshop/uxp/2022/guides/hybrid-plugins/getting-started/)

### 21.6 发布要求

按当前 Adobe Hybrid 发布说明，Marketplace 提交需要覆盖 macOS arm64、macOS x64 和 Windows x64；macOS 原生执行文件需使用有效 Developer ID 签名并公证。只支持部分架构的包应限定在相应独立 / 企业分发范围，并明确其支持条件。[Photoshop Hybrid 分发](https://developer.adobe.com/photoshop/uxp/2022/guides/hybrid-plugins/distribute/)、[Hybrid FAQ](https://developer.adobe.com/uxp/guides/how-to/hybrid-plugins/faq)

原生依赖必须随适当的产物交付。尤其在 Windows 上，Debug 构建依赖的调试运行库可能让插件只在开发机成功；正式包应按 SDK 要求构建 Release 并在干净机器测试。Hybrid 安装和更新可能请求系统管理员凭据，应在安装说明中讲清楚。[Hybrid 加载故障](https://developer.adobe.com/uxp/guides/how-to/hybrid-plugins/faq)

<a id="legacy"></a>

## 22. 脚本、CEP 迁移与 Generator

### 22.1 UXP 脚本 `.psjs`

如果只想快速读取文档，无需建立面板，可将下面的完整代码保存为 `inspect-document.psjs`，通过 Photoshop 的 File → Scripts → Browse 执行：

```javascript
const { app, core } = require("photoshop");

if (app.documents.length === 0) {
  await core.showAlert("当前没有打开的文档。");
} else {
  const psDoc = app.activeDocument;
  await core.showAlert([
    `文档：${psDoc.title}`,
    `尺寸：${psDoc.width} × ${psDoc.height} px`,
    `选中图层：${psDoc.activeLayers.length}`
  ].join("\n"));
}
```

`.psjs` 从 Photoshop 23.5 起提供，脚本运行上下文与面板插件有区别，支持在脚本执行中使用异步 API。脚本通常不需要插件 Manifest，也不承载常驻面板；涉及 UI 生命周期、权限和文档修改时，应阅读脚本专用说明，而非机械复制面板的全部框架。[脚本快速开始](https://developer.adobe.com/photoshop/uxp/2022/scripting/getting-started/)、[Photoshop UXP Scripting](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/uxpscripting)

### 22.2 ExtendScript / JSX

本机 `Presets/Scripts` 下仍有 JSX。维护旧脚本时要记住：其全局 `app`、`File`、`Folder`、`ActionDescriptor`、旧式 DOM 和语言版本，与 UXP 模块体系不同。`$.writeln()`、`#target photoshop`、`executeAction()` 等不能直接作为 UXP API 使用。

迁移时先识别业务意图，再改写：

| 旧实现 | UXP 方向 |
| --- | --- |
| JSX 全局 `app` | `require("photoshop").app`，重新核对 DOM |
| `executeAction()` / ActionDescriptor | DOM 优先，必要时 actionJSON / `batchPlay` |
| `File` / `Folder` | UXP File / Folder Entry 或有权限的路径 API |
| 同步顺序脚本 | 异步调用、正确 `await`、模态与取消 |
| ScriptUI 窗口 | UXP 面板或对话框 |
| JSX 字符串执行 | 模块和函数调用，避免动态拼接代码 |

Adobe 的扩展能力概览将 ExtendScript 标为旧式脚本路线；其代码不能只通过把文件后缀改为 `.psjs` 完成迁移。[Photoshop 扩展技术概览](https://developer.adobe.com/photoshop/)

### 22.3 CEP 项目的结构与迁移

传统 CEP 项目一般包括如下组成；这里用于识别旧项目，不是新 UXP 项目的模板：

```text
cep-extension/
├── CSXS/manifest.xml
├── index.html
├── js/
│   ├── CSInterface.js
│   └── panel.js
└── jsx/
    └── host.jsx
```

面板 JS 常通过 `CSInterface.evalScript()` 调用宿主 JSX；UI 与宿主脚本分属不同执行环境。官方资源、CEP 版本对应表和打包工具位于 [Adobe CEP Resources](https://github.com/Adobe-CEP/CEP-Resources)。不要根据某份老教程随意写固定的 CSXS 版本调试开关，也不要仅因应用包里有 CEP 资源就宣称第三方面板在 Apple Silicon 原生模式可用。

迁移任务可以按下面的表拆解：

| 盘点项 | 迁移动作 |
| --- | --- |
| `manifest.xml` | 重建 UXP Manifest 和入口 / 权限 |
| `CSInterface.evalScript` | 将业务改为直接 UXP DOM / batchPlay 调用 |
| `window.cep.fs` | 替换为 UXP 文件 API 与授权流程 |
| Node / `child_process` | 审查依赖；必要时改外部服务或 Hybrid |
| CEF / 浏览器界面 | 按 UXP 支持范围改写并测试 |
| CEP / CSXS 事件 | 重新设计生命周期与宿主通知订阅 |
| `.zxp` 安装链 | 改为 `.ccx` 和 UXP 分发流程 |
| 自动化参数与用户配置 | 设计迁移路径，保留业务含义 |

对于无法立即迁移的模块，先保留独立适配边界，不必一次重写算法和全部 UI。正式兼容声明应列清 Photoshop 版本及运行架构。[Adobe 技术迁移指南](https://developer.adobe.com/uxp/migration-center/uxp-for-cep-devs/technical-migration-guide/)

### 22.4 已有原生插件如何加 UXP UI

若已有 Photoshop C++ 插件，可以评估官方消息桥：C++ 侧使用 `PIUXPSuite`，UXP 侧使用 `photoshop.messaging`。UXP 的插件 ID 与 C++ PiPL 中的 component 标识属于不同身份，需要准确匹配。[原生插件与 UXP 消息桥](https://developer.adobe.com/photoshop/uxp/ps_reference/media/cpp-pluginsdk/)

协议应包含版本、请求 ID、操作名、结果与错误，不要依赖消息到达顺序隐式关联任务。大像素缓冲不宜默认转成巨大 JSON；根据 SDK 能力设计专门的数据传递方式。

### 22.5 Generator

本机同时存在外部 `Plug-ins/Generator` 和内置 Generator 资源，但这不等于系统安装的任意 Node 版本都能作为 Photoshop Generator 运行时。

只有维护现有 Generator 项目时，才进一步核对对应模块协议、启动配置和宿主兼容性。新资源管理或像素处理工具先评估 UXP / Imaging；Adobe 的开发概览将 Generator 列为历史路线，不能把它当成新的通用 Node 插件系统。[Photoshop 技术路线表](https://developer.adobe.com/photoshop/)

<a id="roadmap"></a>

## 23. 推荐实施顺序

### 阶段一：先跑通一个业务闭环

完成第 6 节示例的加载、文档读取、改名、撤销和 JSON 导出。在 UDT 中能看到日志并能定位错误。此阶段的完成标准是实际在 Photoshop 执行通过，而不只是面板显示出来。

### 阶段二：替换成自己的核心需求

选择一个清晰的动作，例如“导出指定画板 PNG”或“对选中图层生成清单”。定义输入、输出、排序、命名、冲突策略和不支持的图层类型，先做少量样本。

### 阶段三：补齐真实场景

加入多文档、取消、进度、部分失败、目录授权失效、色彩和位深处理。将宿主读写与纯逻辑分层，建立最小回归样本集。

### 阶段四：验证性能与原生需求

记录耗时分布：宿主 API、像素传输、算法、编码、文件写入分别占多少。只有在确认瓶颈和能力缺口后引入 Hybrid 或原生 C++ 插件，避免先承担多平台二进制成本却没有实际收益。

### 阶段五：交付安装包

稳定插件 ID 和配置格式，生成 `.ccx`，在没有开发加载的环境验证安装与更新，补充用户说明、支持范围和已知限制，再进入所选分发渠道。

<a id="references"></a>

## 24. 官方资料与核验说明

### 24.1 常用参考入口

下表供继续开发时检索，正文已经在相关技术说明旁附上对应链接。

| 主题 | 官方资料 |
| --- | --- |
| Photoshop 扩展总览 | [Photoshop Developer](https://developer.adobe.com/photoshop/) |
| UXP / 宿主版本对应 | [UXP Versions](https://developer.adobe.com/uxp/uxp-api/versions) |
| 开发工具 | [Set Up Developer Tools](https://developer.adobe.com/uxp/guides/how-to/developer-tools/) |
| 加载、调试、Watch | [UDT Workflows](https://developer.adobe.com/uxp/guides/how-to/udt-deep-dive/plugin-workflows) |
| Manifest | [Manifest Reference](https://developer.adobe.com/uxp/guides/explanation/concepts/manifest/) |
| Photoshop 专用字段 | [Photoshop Manifest](https://developer.adobe.com/photoshop/uxp/2022/guides/uxp-guide/uxp-misc/manifest-v4/photoshop-manifest/) |
| 文档 / 图层 | [Document](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/classes/document)、[Layer](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/classes/layer) |
| 模态与历史 | [executeAsModal](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/executeasmodal) |
| 动作描述符 | [batchPlay](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/batchplay) |
| 文件系统 | [UXP Storage](https://developer.adobe.com/photoshop/uxp/2022/uxp-api/reference-js/modules/uxp/persistent-file-storage/storage) |
| 像素 | [Imaging API](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/imaging) |
| UI | [Spectrum UXP](https://developer.adobe.com/photoshop/uxp/2022/uxp/reference-spectrum/) |
| 类型 | [TypeScript and IntelliSense](https://developer.adobe.com/uxp/guides/how-to/typescript/) |
| 发布包 | [Package a UXP Plugin](https://developer.adobe.com/uxp/guides/how-to/distribution/package/) |
| Hybrid | [Photoshop Hybrid](https://developer.adobe.com/photoshop/uxp/2022/guides/hybrid-plugins/)、[Build a Hybrid Plugin](https://developer.adobe.com/uxp/guides/how-to/hybrid-plugins/build) |
| C++ / UXP 桥接 | [C++ Plugin SDK Messaging](https://developer.adobe.com/photoshop/uxp/ps_reference/media/cpp-pluginsdk/) |
| 脚本 | [UXP Scripting Quickstart](https://developer.adobe.com/photoshop/uxp/2022/scripting/getting-started/) |
| 迁移 | [CEP Migration Center](https://developer.adobe.com/uxp/migration-center/) |

### 24.2 本文验证范围

| 项目 | 状态 |
| --- | --- |
| 本机应用版本、Bundle ID、主程序架构 | 已读取本机文件验证 |
| UXP、CEP、原生插件、JSX 目录 | 已检查本机目录 |
| 公开 API 与发布流程 | 已查阅正文所链接的 Adobe 官方文档 |
| 普通 JS / JSON 示例 | 已进行静态语法检查；PSJS 按允许顶层异步的脚本上下文检查 |
| 目录锚点及主示例文件 / 入口引用 | 已检查一致性 |
| Photoshop 内加载、运行、撤销和文件导出 | 未执行；需按第 6、18 节在宿主验收 |
| C++ SDK 编译、Hybrid addon 加载 | 未执行；需取得匹配 SDK 并在目标架构验证 |
| `.ccx` 打包、安装、Marketplace 提交 | 未执行；本文提供操作流程 |

文档路径中出现 `2022` 是 Adobe 文档站点的路径组织方式，不能据此认定页面只适用于 Photoshop 2022。兼容判断应结合页面的最低版本、变更记录、目标 SDK 和实际宿主运行结果。
