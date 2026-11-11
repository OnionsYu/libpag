# Photoshop PAGX 导出插件

真正运行在 Photoshop 中的 UXP Hybrid 面板，C++ addon 直接链接本仓库的 PAGX 实现。导出时不启动外部 CLI，也不依赖云端 PSD 解析。

## 当前能力

- **导出当前文档**：复制当前内存文档，包含未保存的修改。原文档的内容、历史和图层选择不变。
- **选择 PSD 并导出**：由 Photoshop 打开磁盘 PSD，导出后自动关闭。若 Photoshop 返回已打开的文档，拒绝继续，提示改用当前文档入口，避免把未保存内容误当成磁盘内容。
- 选择输出目录，生成自包含的 `.pagx`，PNG 图像嵌入文件。同名文件自动编号。
- v0.2 对经过验证的 RGB/sRGB/8 位形状导出 **原生贝塞尔路径、复合路径、纯色填充和实线描边**，保留小数坐标、旋转和越界几何。像素层保留为位图。单色文字使用 Photoshop 生成的矢量轮廓；混色、变形、缺字体文字及其他未映射的局部外观由 Photoshop 渲染；涉及背景的合成暂回退到完整文档。隐藏图层不参与当前画面。
- 支持普通画布、方形像素的 8/16 位源文档；画板、非方形像素、32 位 HDR 和 1 位位图明确拒绝。8 位画布上限 64M 像素，16 位上限 32M，编码图像总计上限 128 MiB，XML 上限 256 MiB。转换为 8 位意味着有精度损失。
- 已实测 RGB 8/16 位；其他颜色模式和不同 ICC 配置仍需补充宿主样例，不能视为已全面验证。

**DOM 的 normal 模式及普通 layer get 都不足以证明高级混合为默认值**。适配器同时读取 action 的完整 `json` 属性和图层/路径描述符。已在 Photoshop 27.4、JSON 1.6.1 中验证 Blend If、挖空、内部混合、剪贴混合及透明形状标记可被完整 JSON 检出；缺失能力、其他宿主或 JSON 版本保持 unknown，安全回退。实线描边的非默认透明度、虚线、像素对齐及未支持的填充有明确回退。

## 分层与协议

```text
UXP 面板 main.js
  → ExportService（事务、取消、进度）
  → PhotoshopHost（DOM / Imaging，临时文档）
  → photoshop::Document（版本化 PS 图层规范层）
  → PhotoshopImporter（分析计划、转换器）
  → PAGXDocument / PAGXExporter / PAGXImporter（PAGX 规范、序列化、回读校验）
  → OutputFiles（完整写入临时文件，改名发布）
```

三个规范/转换层均不依赖 UXP。详细语义见 [PS → PAGX 映射规范](../../spec/photoshop_pagx_mapping.zh_CN.md) 和 [PSD 图层规范](../../spec/psd_layer_spec.zh_CN.md)。C++ 输入类型位于 `include/photoshop/Document.h`，JSON 解析器拒绝未知字段、未来版本、缺少字段和非法数据。

桥接协议版本为 1、源快照版本为 2（原生解析器仍接受 v1）：`info()` 握手 → `begin(snapshotJSON, forceComposite)` 返回采集计划和会话 → 按计划顺序 `addRaster(session, metadataJSON, ArrayBuffer)` → `finish(session)` 返回已回读校验的 XML；失败调用幂等的 `discard(session)`。像素必须为实际返回区域的 chunky RGB/RGBA、sRGB、straight alpha。8 位输入用 `Uint8Array`；16 位输入用原生字节序 `Uint16Array`，`fullRange=true`（0..65535），由 C++ 按最近值量化到 8 位。不能将请求区域当作实际区域，也不能重复乘图层透明度。像素只在回调内访问，PNG 编码后释放宿主图像资源。

原生调用是有大小上限的同步调用，取消在步骤之间生效；编码期间不能保证即时响应。宿主操作在 `executeAsModal` 中，临时文档立即注册 `registerAutoCloseDocument`，包括 Esc 取消都由 Photoshop 清理。没有写入任何源 PSD 的路径。

## 构建和加载

目标基线为 Photoshop 27.4.0、Manifest 6、Photoshop API 2。需本仓库可正常构建、Node.js 20+、CMake 和平台 C++ 工具链。

1. 从 [Adobe Developer Console](https://developer.adobe.com/console/) 下载并解压 UXP Hybrid Plugin SDK。SDK 不进仓库。
2. 在此目录构建：

   ```sh
   UXP_SDK_DIR=/path/to/uxp-hybrid-plugin-sdk npm run build:native
   npm test
   npm run package
   ```

   `PAGX_BUILD_DIR` 可指定 CMake 目录；`PAGX_ARCH` 可为 `arm64` 或 `x64`，需要相应架构的仓库依赖。Windows 构建需配置环境变量，并使用 Visual Studio x64 工具链。

3. 使用官方 **UXP Developer Tool**，按 Adobe 提示启用开发者模式，添加本目录的 `manifest.json`，Load 后在 Photoshop 的“增效工具”菜单中打开 **PAGX 导出**。
4. 构建脚本将 addon 放在 `mac/arm64`、`mac/x64` 或 `win/x64`；`dist/plugin` 是仅包含运行文件的加载目录；`dist/*-dev.ccx` 是本地开发包。Windows 用 UDT 对 `dist/plugin` 打包。

macOS 构建脚本仅做开发用 ad-hoc 签名。正式分发需 Developer ID 签名、公证以及目标架构实测；本地开发包不代表已通过 Marketplace 审核。当前验证记录仅覆盖本机 macOS arm64，不能据此宣称 Windows / Intel 已验证。

## 分阶段验证

- 第一阶段：C++ 模型和转换器，`PAGXPhotoshopTest.*` 最初 7 项通过，提交 `e6a1c7d8b`。
- 第二阶段：正式 Adobe SDK 编译 addon；快照 JSON 校验；Node 业务测试覆盖资源释放、取消、并发、像素视图、同名文件和源文档生命周期，提交 `53d149a03`。模拟宿主测试不等同于 Photoshop 实测。
- 第三阶段：在 Photoshop 27.4.0 中加载真实插件，修复异步 addon 加载、UXP 宿主版本读取、空图层组边界和 16 位像素采集。验证结果如下。

### 本机验证记录（2026-11-11）

环境：macOS 15.7.9、Apple Silicon、Photoshop 27.4.0、Adobe UXP Hybrid SDK 6.5.0。开发者模式已开启，UDT 加载状态为 Loaded。用户确认浮动面板可见；两个按钮通过 UDT 控制台触发实际 DOM click，文件/目录选择使用 Photoshop 原生对话框。

| 验证 | 结果 |
| --- | --- |
| C++ `PAGXPhotoshopTest.*` | 8 项通过 |
| Node 业务及加载回归 | 12 项通过 |
| 当前已保存文档导出 | 64×48、透明边距、半透明红色、蓝色均正确 |
| 当前未保存修改 | 50% 图层透明度导出后 alpha 为 64/128；原文档 ID、历史数、图层选择、dirty 状态不变 |
| 选择另一份 PSD | 导出后自动关闭临时文档并恢复原活动文档 |
| 选择已打开的 PSD | 显示明确错误，未关闭或保存原文档 |
| 8/16 位、全透明、复杂图层 | 宿主脚本 4 个文件场景通过，16 位 PSD 由 Photoshop 生成 |
| 取消、原生异常、失败后重试 | 3 个场景通过，无临时文档或输出残留 |
| PAGX 回读和渲染 | 结构校验通过；8/16 位与未保存修改样本的可见像素及 alpha 与独立基准一致；复杂样本的渲染与采集图像一致 |

复杂样本使用本机 Adobe Generator 测试资源 `all-layer-types.psd` 的副本，包含文字、形状、调整层、智能对象、蒙版、效果、空组和越界图层；该第三方资源不提交进仓库。未执行 Windows / Intel、HDR、画板、超大文件压力测试及其他颜色模式完整矩阵。取消测试在采集边界设置取消信号，未将人工 Esc 中断测为通过。

宿主发现：本机 `getPixels(componentSize: 8)` 读取 16 位合成图会失败；读取 16 位后原生量化可正常导出。`executeAsModal` 重新包装异常后可能不保留 JS 原型或 `.message`，错误展示使用字符串兜底。

### 复现宿主验证

在仓库根目录创建独立测试目录（不要用生产 PSD）：

```sh
python3 tools/photoshop-exporter/test/make-fixtures.py .tmp/photoshop-validation
```

用 UDT 加载源码目录的插件，在其 Debug 控制台运行下式，并选择上述目录。脚本将生成 16 位 PSD 副本、PAGX 和 `host-report-*.json`，检查原有文档状态和失败清理；可选将复杂测试 PSD 复制到该目录并命名为 `all-layer-types.psd`。

```js
require("./test/host-smoke.js").run().catch(console.error)
```

安装 Pillow，并确保项目 `pagx` CLI 在 PATH 中，在有图形设备访问权限的环境运行：

```sh
python3 tools/photoshop-exporter/test/verify-exports.py .tmp/photoshop-validation --render
```

当前未保存修改另用 `alpha-fixture.psd` 手工测试：将图层不透明度改为 50%，点击“导出当前文档”，不保存原 PSD，再给像素验证脚本传 `--dirty-export <输出.pagx>`。完全透明像素的 RGB 不定义，比较时仅忽略 alpha 为 0 的 RGB，所有 alpha 值仍逐像素比较。测试脚本、SDK 和样本不进入运行包。

官方依据：[Hybrid 构建](https://developer.adobe.com/uxp/guides/how-to/hybrid-plugins/build)、[Imaging](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/imaging)、[模态执行和自动关闭](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/media/executeasmodal)、[打包](https://developer.adobe.com/uxp/guides/how-to/distribution/package/)。

### 矢量阶段验证（2026-11-11）

新增 `test/host-vectors.js`，从 UDT 调用 `probe(folder)`、`geometryFixtures(folder)`，由
Photoshop 创建矩形、椭圆、孔洞、内描边和旋转/平移样例。在任何导出变换前独立采集
RGBA 基准；`exportFixture(folder, name)` 写出源快照、PAGX 及导出诊断。

```sh
python3 tools/photoshop-exporter/test/verify-vectors.py /path/to/fixture-folder
```

验证原生路径且 **Image 数为 0**，执行 PAGX 回读和 1×/4× 渲染，分别比较 alpha、黑底、
白底。不同渲染器的边缘覆盖有差异：允许平均误差 0.5/255、边缘峰值 64/255，同时
要求平坦区域误差不超过 2/255；不是逐像素完全相同的承诺。孔洞中心另断言透明。
CLI 的矩形/椭圆路径优化建议单独记录，其他校验诊断均失败。PS 的 Fill 仅影响填充，
不影响原生形状描边；图层 Opacity 在两者合成后应用。测试发现的非默认描边透明度
不能直接等价映射，宿主适配器因此保留回退。

### 文字轮廓阶段验证（2026-11-11）

`textFixtures(folder)` 创建 72/300 PPI 中英文多行文字，分别保存**原始文字**和 Photoshop
**转换后轮廓**的独立基准。导出器只对额外复制的文字层调用 `convertToShape()`，读取后
删除该复制层；即使随后回退，采集的仍是原始文字。复制层清理失败会终止事务。

19 项 Node 测试通过。两个宿主文字文件均为 1 个 Path、0 个 Image、0 个 Text；与宿主
轮廓比较的 alpha MAE 均为 0.085/255、峰值 14/255，72/300 PPI 输出像素大小一致，
没有重复进行 PPI 换算。验证命令增加 `--names text-72 text-300 --reference-suffix=-outline`。

轮廓并不保持 Photoshop 字体微调和文字抗锯齿：本例相对于原始文字的 alpha MAE 为
1.908/255、峰值 196/255，差异位于字形边缘。面板明确提供“矢量优先（文字转轮廓）”
与“外观快照（整幅位图）”，不能将轮廓描述为可编辑文字或像素完全一致的文字。多色、
变形、未知字体和未支持的文字绘制方式显式回退。当前未实现 PAGX Text 的原生排版映射。
