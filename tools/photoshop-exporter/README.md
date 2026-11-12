# Photoshop PAGX 导出插件

真正运行在 Photoshop 中的 UXP Hybrid 面板，C++ addon 直接链接本仓库的 PAGX 实现。导出时不启动外部 CLI，也不依赖云端 PSD 解析。

## 当前能力

- **导出当前文档**：复制当前内存文档，包含未保存的修改。原文档的内容、历史和图层选择不变。
- **选择 PSD 并导出**：由 Photoshop 打开磁盘 PSD，导出后自动关闭。若 Photoshop 返回已打开的文档，拒绝继续，提示改用当前文档入口，避免把未保存内容误当成磁盘内容。
- 选择输出目录，生成自包含的 `.pagx`，PNG 图像嵌入文件，并生成同名 `.pagx.report.json`。
  报告列出源颜色/位深、输出方式、矢量与位图计数、合成范围、图层名称和降级原因。
  面板显示这些统计及前 12 条诊断；完整诊断均保留在报告中。同名文件自动编号。
- v0.2 对经过验证的 RGB/sRGB/8 位形状导出 **原生贝塞尔路径、复合路径、纯色填充和实线描边**，保留小数坐标、旋转和越界几何。像素层保留为位图。单色文字使用 Photoshop 生成的矢量轮廓；混色、变形、缺字体文字及其他未映射的局部外观由 Photoshop 渲染。调整、剪贴、特殊混合按隔离边界计算连续依赖范围，范围外继续保留矢量；未知合成和挖空仍回退整图。隐藏图层不参与当前画面。
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

桥接协议版本为 2、源快照版本为 2（原生解析器仍接受 v1）：`info()` 握手 → `begin(snapshotJSON, forceComposite)` 返回采集计划、统计、诊断和会话 → 按计划顺序 `addRaster(session, metadataJSON, ArrayBuffer)` → `finish(session)` 返回已回读校验的 XML；失败调用幂等的 `discard(session)`。采集请求包含连续兄弟范围 `layerIds` 及其父组 `parentId`；协议不匹配会拒绝加载。像素必须为实际返回区域的 chunky RGB/RGBA、sRGB、straight alpha。8 位输入用 `Uint8Array`；16 位输入用原生字节序 `Uint16Array`，`fullRange=true`（0..65535），由 C++ 按最近值量化到 8 位。不能将请求区域当作实际区域，也不能重复乘图层透明度。像素只在回调内访问，PNG 编码后释放宿主图像资源。

原生调用是有大小上限的同步调用，取消在步骤之间生效；编码期间不能保证即时响应。宿主操作在 `executeAsModal` 中，临时文档立即注册 `registerAutoCloseDocument`，包括 Esc 取消都由 Photoshop 清理。没有写入任何源 PSD 的路径。

输出和报告先完整写入同目录的临时文件，再发布报告，最后发布 PAGX；失败时回滚本次
创建的文件，已存在文件不覆盖。两个文件不构成文件系统级原子事务，进程崩溃可能留下
孤立报告；再次导出会避让该报告的名称。报告不含源 PSD 路径、原文内容或像素字节。

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

### 局部合成阶段验证（2026-11-11）

14 项 C++ 测试和 21 项 Node 测试通过，覆盖普通隔离组、穿透组向外的背景依赖、
重叠剪贴范围合并、组蒙版、隐藏后代及组透明度只应用一次。协议升为 2，旧面板与
新原生模块不会混用。已新增 `test/host-compositing.js`，生成独立 Photoshop 基准、
源快照及导出计划；使用 `run(folder)`，再运行：

```sh
python3 tools/photoshop-exporter/test/verify-vectors.py /path/to/fixture-folder \
  --names isolated-group pass-through clipping group-mask adjustment --images 1
```

本阶段提交时宿主窗口访问返回 `cgWindowNotFound`，因此这些新增局部合成样例的
真实 Photoshop 像素验证尚未完成；上述自动化测试不能替代该宿主验证。
调整层分类采用 [Adobe LayerKind](https://developer.adobe.com/photoshop/uxp/2022/ps-reference/modules/constants)
的已知枚举；不存在的枚举、未知类型及读取异常不推断为普通像素层。

### 导出报告阶段验证（2026-11-11）

24 项 Node 测试通过，新增范围/父组名称及统计报告、PAGX 与报告共用编号、报告创建/
写入/重命名失败、PAGX 发布失败和发布中取消的回滚测试。导出成功不再隐藏降级信息。
报告计数按源图层和合成范围统计，不等于去重后的 PNG 资源数或 XML 节点总数。

### 当前核心重放和打包验证（2026-11-11）

最终 C++ 测试为 15 项、Node 测试为 24 项，全部通过。额外检查剪贴链必须引用最近
的未剪贴基底、路径运算枚举有效，以及原生像素元数据的坐标不能是布尔值。
`host-smoke.js` 新增布尔坐标异常注入；这一新增宿主场景仍待窗口访问恢复后执行。

为了检查后续核心改动，新增仅用于验证的 `photoshop-vector-replay`：读取先前由真实
Photoshop 采集的源快照，调用**当前**转换核心重新生成 PAGX 并回读校验。它对需要
位图采集的快照报错，不以旧 PAGX 中的图像替代缺失的宿主采集。

```sh
cmake --build cmake-build-photoshop --target photoshop-vector-replay -j 8
python3 tools/photoshop-exporter/test/verify-vectors.py /path/to/fixture-folder \
  --replay cmake-build-photoshop/tools/photoshop-exporter/native/photoshop-vector-replay \
  --report geometry-replay-verification.json
python3 tools/photoshop-exporter/test/verify-vectors.py /path/to/fixture-folder \
  --names text-72 text-300 --reference-suffix=-outline \
  --replay cmake-build-photoshop/tools/photoshop-exporter/native/photoshop-vector-replay \
  --report text-replay-verification.json
```

4 个形状样例和 2 个文字样例重放通过，均为 0 个 Image；当时完成 1× 独立基准比较及
4× 渲染，**没有 4× 独立基准比较**。形状 alpha MAE 最大 0.314/255，轮廓文字为
0.086/255 以下。原始文字与
轮廓之间的差异仍按文字阶段说明，不因重放通过而消失。开发验证产物归档于
`.tmp/photoshop-vector-validation/`（不进入 Git 或插件包）。

macOS arm64 原生模块已重建并通过 ad-hoc 签名检查，开发 CCX 已生成，包内仅含 15 个
运行文件/目录项，不包含 SDK、测试脚本或样本。**最终协议 2 插件的重新加载、报告发布
及新增局部合成场景尚未完成宿主验收**，不得将本段构建和重放结果解释为已通过完整
Photoshop 回归。渐变、独立蒙版语义、可编辑文字仍未原生映射；使用已说明的局部/整图
回退或明确拒绝，不能宣称任意 PSD 都输出纯矢量。

### 正确性修复与真实宿主回归（2026-11-12）

第一阶段修复复合路径 EvenOdd 规则在内外描边中丢失的问题，增加 RGB 混合 gamma
门禁。17 项 C++ 测试、25 项 Node 测试通过；C++ 像素测试覆盖直接 EvenOdd 孔洞与
减法孔洞、内/外描边、XML 往返及 1×/4×。

已在 Photoshop 27.4 重新加载重建的 arm64 模块，完成 5 个几何样例的 1×/4× 独立
像素比较（含孔洞内外描边），均为 0 个 Image；最大 alpha MAE 为 0.362/255，
非边缘误差不超过 1/255。4× 参考来自分层文档副本重新缩放渲染，明确修正 Photoshop
保留未锁定描边宽度的行为，并校验最终宽度；不是对 1× 合成图插值。
5 个合成样例（隔离组、穿透组、剪贴、组蒙版、调整层）完成协议 2 的真实导出和
1× 像素比较，均只采集一个范围、保留无关矢量，非边缘误差不超过 1/255。
报告随 PAGX 一起发布。生成物及比较报告在 `.tmp/photoshop-vector-stage1/`。

`verify-vectors.py` 默认要求 1× 和 4× 独立参考，缺少任意参考即失败；旧版仅有 1× 的
档案和含回退位图的样例必须明确传 `--scales 1`。用 `--cli` 指定本轮构建的可执行文件：

```sh
python3 tools/photoshop-exporter/test/verify-vectors.py .tmp/photoshop-vector-stage1 \
  --cli ./cmake-build-debug/pagx \
  --names hole stroke transformed hole-inner-stroke hole-outer-stroke
python3 tools/photoshop-exporter/test/verify-vectors.py .tmp/photoshop-vector-stage1 \
  --cli ./cmake-build-debug/pagx --scales 1 --images 1 \
  --names clipping group-mask adjustment
```

### Paint 源模型与描边单位验证（2026-11-12）

插件版本为 0.3.0，源快照升级到 v3，原生模块继续读取 v1/v2；桥接协议仍为 2。
JS 与原生模块在加载时校验 schema 版本，必须一起更新。Paint 分离源色标、中点、
透明度标记、插值方式和坐标矩阵；能采集属性不表示已经支持相应原生输出。
当前新增无蒙版纯色填充层和单组件居中虚线，保留 Fill 图层类型与实际点单位。

18 项 C++ 和 27 项 Node 测试通过。真实 Photoshop 27.4 的纯色填充、平头/圆头/
方头虚线及 300 PPI 虚线共 5 个样例完成 1×/4× 比较，均为 0 个 Image，最大 alpha
MAE 小于 0.085/255，非边缘误差小于 1/255，沿用 MAE≤0.5、边缘最大差≤64 的门槛。
密集端帽重叠在 4× 出现超阈值覆盖率差，未放宽阈值；该类情况、布尔或非居中虚线、
非默认描边透明度继续局部回退。透明度描述符约 35% 而实际宿主描边仍不透明的样例
已记录；不能仅看字段名直接赋给 PAGX alpha。

重叠端帽和非默认描边透明度的回退完成 1× 比较，alpha 误差为 0；Classic/Linear/
Perceptual 三种渐变先保留源模型再回退，1× 回退像素误差均为 0。报告为
`.tmp/photoshop-vector-stage1/stage2-*.json`。生成与验收入口：

```js
await require("./test/host-appearance.js").run(folder);
await require("./test/host-gradients.js").run(folder);
```

```sh
python3 tools/photoshop-exporter/test/verify-vectors.py .tmp/photoshop-vector-stage1 \
  --cli ./cmake-build-debug/pagx --names solid-fill dash dash-round dash-square dash-300
python3 tools/photoshop-exporter/test/verify-vectors.py .tmp/photoshop-vector-stage1 \
  --cli ./cmake-build-debug/pagx --scales 1 --images 1 \
  --names dash-overlap stroke-opacity gradient-classic gradient-linear gradient-perceptual
```

### 原生混合与组隔离（2026-11-12）

19 项 C++ 和 27 项 Node 测试通过。Multiply、Screen、Darken、Lighten、Overlay、
Hard Light、Exclusion 以及隔离组、穿透组共 9 个真实 Photoshop 样例完成 1×/4× 比较，
均为 0 个 Image；三通道最大 MAE 为 0.427/255 以下，非边缘最大差小于 1.60/255。
参考同时覆盖背景 alpha、图层 Opacity、Fill 和组透明度。
Soft Light 的原生结果超出 MAE≤0.5 门槛，继续采集其背景范围；回退后 1× 验证通过。
剪贴、像素组蒙版和调整层的局部回退回归也通过，仍保留范围外矢量。

```js
await require("./test/host-blends.js").run(folder);
await require("./test/host-compositing.js").run(folder);
```

```sh
python3 tools/photoshop-exporter/test/verify-vectors.py .tmp/photoshop-vector-stage1 \
  --cli ./cmake-build-debug/pagx --report stage3-blends.json \
  --names blend-multiply blend-screen blend-darken blend-lighten blend-overlay \
          blend-hardlight blend-exclusion isolated-group pass-through
python3 tools/photoshop-exporter/test/verify-vectors.py .tmp/photoshop-vector-stage1 \
  --cli ./cmake-build-debug/pagx --images 1 --scales 1 \
  --names blend-softlight clipping group-mask adjustment
```
