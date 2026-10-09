# ISIR 图纸识别 / Drawing Recognition Skill

## 中文描述

供 Codex 使用的工业图纸识别与 ISIR 检验报告工作流。支持 PDF、图片、扫描图及 CAD 导出图，提取主体视图中的尺寸、螺纹、角度、倒角、半径和形位公差，并识别材质、外观及适用的未注公差规则。生成连续气泡编号、可追溯参数映射和固定模板的 ISIR Excel，保留原图证据、人工复核标记和空白实测栏；配套脚本检查气泡布局与工作簿一致性，检验报告 PDF 导出须验证物理 A4 尺寸。

## English description

A Codex skill for industrial drawing recognition and ISIR (Initial Sample Inspection Report) preparation. It supports PDFs, images, scans and CAD exports, extracting principal-view dimensions, threads, angles, chamfers, radii and geometric dimensioning and tolerancing (GD&T), together with material, appearance and applicable general-tolerance rules. It produces consecutive balloon numbers, traceable parameter mappings and ISIR Excel workbooks using a fixed template. Source evidence and manual-review flags are preserved, while measurement fields remain blank when no actual measurements are provided. Included scripts check balloon placement and workbook consistency; exported inspection-report PDFs require physical ISO A4 page-size verification.

Invoke it in Codex with `$isir-drawing-recognition` after supplying a drawing. This package provides an AI workflow, validation scripts and a blank template; final inspection items, balloons and reports require human review.

## 能力与范围

- 支持 PDF、图片、扫描图及 CAD 导出图的主体视图参数识别。
- 将材质与外观作为表格末尾两项；未注公差按本图规则补充到相应尺寸。
- 保留参数页码、坐标、原文和人工复核标记；形位公差可使用原图完整截图。
- 输出原始图纸 PDF、ISIR Excel 和编号图 PDF；检验报告 PDF 版式验证要求为物理 A4。

这是 AI 工作流与校验工具包。识别、裁剪、编号与 Excel 生成由支持 skill 的 AI 环境完成；仓库中的两份脚本用于校验，不是独立的一键 OCR 程序。最终参数、气泡和报告需要人工复核。

## 环境

- 支持 SKILL.md 的 Codex 环境，以及可用的 PDF 渲染、图像查看、OCR/视觉识别工具。
- Python 3.9 或以上：气泡映射校验脚本仅使用标准库。
- Windows PowerShell 5.1 或以上，以及已安装的桌面 Microsoft Excel：工作簿校验通过 Excel COM 执行。
- Excel 生成优先直接修改模板副本，保留表格、复选框和图片结构；禁止覆盖模板。

## 安装

下载或克隆本仓库，将整个目录作为 `isir-drawing-recognition` skill 安装到你的 Codex 技能发现目录；若使用能力库，可将目录放入能力库 `skills` 下，再按本机 Codex 配置建立目录关联。重新打开 Codex 会话后确认该 skill 已出现在列表中。

```text
isir-drawing-recognition/
  SKILL.md
  agents/openai.yaml
  references/excel-template-contract.md
  scripts/validate_isir_mapping.py
  scripts/verify_isir_workbook.ps1
  assets/ISIR-golden-reference.xlsx
```

## 使用示例

在 Codex 中提供图纸后输入：

> 使用 $isir-drawing-recognition 识别这份图纸，先输出可复核参数与来源映射，再生成编号图及 ISIR。没有实测数据时保持测量值和判定为空，对不明确公差标记待复核。

运行映射校验：

```powershell
python scripts/validate_isir_mapping.py work/mapping.json --radius 32 --max-gap 24
```

脚本要求 `source_document`、`source_image` 指向存在的文件；路径相对于运行时工作目录。完整坐标和气泡规则见 SKILL.md；不同图像分辨率须调整半径和最大间距。只有完成实际目视检查后才可标记 `visual_confirmed`。

运行 Excel 校验（参数需替换为当前图纸实际信息）：

```powershell
./scripts/verify_isir_workbook.ps1 -WorkbookPath output/ISIR.xlsx -MappingPath work/mapping.json -ExpectedCharacteristics 20 -ExpectedPartName '当前件名' -ExpectedDrawingNo '当前图号' -ExpectedRevision '当前版本' -ExpectedMaterial '当前材质' -ExpectedSupplier ''
```

现有固定模板为三张工作表，每张数据表容量为 21 项，总容量 42 项。超过容量时需要先扩展模板和校验逻辑，不得直接套用。

## 模板说明

`assets/ISIR-golden-reference.xlsx` 是发布用空白模板，已移除原母版中的公司、零件、检验数据和示例图，并设为 A4。SHA-256：

`304FDCB75B2094B35EE1E279B6CC0D9FE46FCF12A2C1BCD00DB00D860B5B77D5`

空白模板本身不等于最终合格报告；生成后必须插入当前编号图，按映射填写当前元数据与参数，并执行校验和目视复核。Excel 校验脚本不验证 PDF 实际纸张尺寸，A4 导出另按模板合同检查。

## 文件与使用权

仓库只包含技能规则、校验脚本和空白模板，不包含客户图纸、实际检验报告、API 密钥或本机个人配置。尚未指定开源许可证；公开可见不代表授予再分发或商业使用许可。
