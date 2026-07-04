# TemplateToDocV2

## 中文说明

TemplateToDocV2 是一个面向 Word 模板复用的文档生成工具，适用于在既有 `.docx` 模板中填写项目字段、替换章节正文和更新图片。工具会以原始模板为基础生成 `final.docx`，尽量延续模板已有的页眉页脚、页码、样式、表格、图片、浮动对象、页边距、分节设置和域设置。

典型使用流程包括导入 Word 模板、生成模板识别配置、填写项目内容、构建 Word 文件，以及按需执行渲染 QA 检查。Markdown 内容或大模型生成结果可以作为章节正文来源，并写回到模板中对应的位置。

仓库中附带了 `空管自动化软件设计说明_扩写版.docx` 示例模板，README 中的命令和路径均以该模板为示例。

### 适用场景

- 已有固定 Word 模板，希望只替换局部字段、章节正文和图片。
- 模板中包含复杂表格、页眉页脚、图文混排或已有图片，需要保持模板已有排版。
- 需要基于 DeepSeek 自动生成某些章节正文，同时保留原模板版式。
- 需要对输出 Word 做渲染检查，确认生成后的页面与原模板结构接近。

### 功能概览

- 导入 `.docx` 模板，并自动生成 `profile.yaml`。
- 支持导入 `.doc` 模板，前提是本机安装了 LibreOffice。
- 自动识别模板中的表格字段、占位符、章节段落和图片。
- 生成可编辑的 `project.yaml`，用于填写项目字段、章节提示词、正文和图片替换设置。
- 支持保留原图片，也支持上传新图片替换模板内图片。
- 支持 DeepSeek 生成缺失章节或在 Web 页面中逐节生成。
- 支持使用 LibreOffice/Poppler 渲染 Word 并进行 QA 对比。
- 提供命令行和本地 Web 两种使用方式。

### 环境准备

建议使用 Python 3.10 及以上版本。

```powershell
git clone git@github.com:zy199336/TemplateToDoc.git
cd TemplateToDoc
python -m pip install -e .
```

运行测试：

```powershell
python -m pytest tests
```

可选依赖：

- DeepSeek 自动生成：设置 `DEEPSEEK_API_KEY`，或在 Web 页面中填写 API Key。
- `.doc` 模板导入：需要安装 LibreOffice，并确保 `soffice` 在 PATH 中。
- QA 渲染对比：需要 LibreOffice `soffice` 和 Poppler `pdftoppm`。

### 命令行使用

#### 1. 导入模板

```powershell
template-to-doc-v2 import-template .\空管自动化软件设计说明_扩写版.docx --profile atc_sdd --project-id atc_sdd_demo
```

执行后会生成：

```text
profiles/atc_sdd/reference.docx
profiles/atc_sdd/profile.yaml
projects/atc_sdd_demo/project.yaml
```

如果没有安装为命令行脚本，也可以直接运行源码入口：

```powershell
python src\template_to_doc_v2\cli.py import-template .\空管自动化软件设计说明_扩写版.docx --profile atc_sdd --project-id atc_sdd_demo
```

#### 2. 初始化项目

如果导入模板时没有传 `--project-id`，可以单独创建项目：

```powershell
template-to-doc-v2 init-project atc_sdd atc_sdd_demo
```

项目文件会写入：

```text
projects/atc_sdd_demo/project.yaml
```

#### 3. 编辑 project.yaml

`project.yaml` 是主要输入文件，常见结构如下：

```yaml
project_id: atc_sdd_demo
profile_id: atc_sdd
global:
  topic: 空管自动化软件设计说明
  background: 本文档用于描述空管自动化软件 CSCI 的设计说明内容。
  materials: 软件功能说明、接口设计资料、运行态势数据、需求追踪表。
  requirements: 保持模板的技术文档语气、标题层级、表格结构和图文版式。
fields:
  paragraph_field_01: 空管自动化软件设计说明
  paragraph_field_02: 文档控制
  paragraph_field_03: 目 录
sections:
  1_4:
    prompt: 编写“范围”章节，说明空管自动化软件设计说明的适用边界。
    content: 本文件规定空管自动化软件 CSCI 的设计说明内容，用于指导软件开发、集成、验证、部署和运行维护活动。
images:
  image1:
    mode: keep
    file: null
```

说明：

- `global`：全局背景，会作为章节生成时的共同上下文。
- `fields`：模板表格或占位符字段。
- `sections.<id>.prompt`：给模型的本节写作提示。
- `sections.<id>.content`：最终写回 Word 的本节正文。
- `images.<id>.mode: keep`：保留模板原图。
- `images.<id>.mode: replace`：使用新图片替换模板图片，Web 页面上传图片时会自动填充 `file`。

#### 4. 构建 Word

仅使用已填写内容构建：

```powershell
template-to-doc-v2 build atc_sdd atc_sdd_demo
```

输出文件：

```text
projects/atc_sdd_demo/final.docx
projects/atc_sdd_demo/replacement_report.yaml
```

#### 5. 使用 DeepSeek 自动生成缺失章节

```powershell
$env:DEEPSEEK_API_KEY="sk-..."
template-to-doc-v2 build atc_sdd atc_sdd_demo --generate --model deepseek-v4-flash
```

如果某个章节的 `content` 为空，系统会结合 `global`、本节 `prompt` 和模板样例段落生成正文，然后写回模板。

#### 6. 构建并执行 QA 检查

```powershell
template-to-doc-v2 build atc_sdd atc_sdd_demo --qa
```

或对已有结果单独比较：

```powershell
template-to-doc-v2 compare atc_sdd atc_sdd_demo
```

QA 输出通常包括：

```text
projects/atc_sdd_demo/qa_compare.yaml
projects/atc_sdd_demo/render/candidate/
projects/atc_sdd_demo/render/reference/
projects/atc_sdd_demo/render/diff/
```

### Web 页面使用

启动本地 Web 服务：

```powershell
template-to-doc-v2 serve --port 8770
```

浏览器打开：

```text
http://127.0.0.1:8770/
```

Web 流程：

1. 在“导入模板”区域选择本地 `.docx/.doc` 文件，或填写模板路径。
2. 输入 `Profile ID` 和 `Project ID`，点击“导入模板”。
3. 填写全局信息、项目字段和各章节提示词。
4. 可以逐节点击“生成本节”，也可以根据全局信息整体生成。
5. 图片区域默认保留原图；如需替换，选择“上传新图片替换”并上传图片。
6. 点击“生成 Word”下载 `final.docx`。
7. 如果安装了 LibreOffice 和 Poppler，可以点击“生成并审核”执行渲染检查。

### 使用样例

下面的示例以 [空管自动化软件设计说明_扩写版.docx](空管自动化软件设计说明_扩写版.docx) 作为模板。运行命令后会在本地创建 `projects/atc_sdd_demo`，生成 `final.docx`，并将部分页面渲染成 PNG，便于快速查看效果。

#### 样例命令

```powershell
cd TemplateToDoc
template-to-doc-v2 import-template .\空管自动化软件设计说明_扩写版.docx --profile atc_sdd --project-id atc_sdd_demo
template-to-doc-v2 build atc_sdd atc_sdd_demo --qa
```

#### 样例输出路径

```text
projects/atc_sdd_demo/final.docx
projects/atc_sdd_demo/replacement_report.yaml
projects/atc_sdd_demo/qa_compare.yaml
projects/atc_sdd_demo/render/candidate/final.pdf
projects/atc_sdd_demo/render/candidate/final-01.png
projects/atc_sdd_demo/render/candidate/final-03.png
projects/atc_sdd_demo/render/candidate/final-05.png
```

#### 使用效果查看

运行 `--qa` 后，可在以下文件中查看渲染效果：

- `projects/atc_sdd_demo/render/candidate/final-01.png`：封面标题、封面图片、文档编号表格和页脚页码保留效果。
- `projects/atc_sdd_demo/render/candidate/final-03.png`：目录层级、点引导线和页码保留效果。
- `projects/atc_sdd_demo/render/candidate/final-05.png`：正文标题层级、图文混排和表格版式保留效果。

这些渲染结果属于本地生成产物，仓库只保留模板和代码，不提交生成结果。

### 使用效果说明

生成后的 `final.docx` 具有以下特点：

- 原模板的页眉、页脚、页码和章节样式会被保留。
- 原模板中的表格结构不会被 Markdown 重新生成，表格边框、合并单元格和宽度更稳定。
- 识别出的表格字段会按 `project.yaml` 中的 `fields` 写回。
- 识别出的章节正文会按 `sections.<id>.content` 在原位置替换。
- 原模板图片默认保留；配置为 `replace` 后可替换指定图片。
- `replacement_report.yaml` 会记录字段、章节、图片的替换数量以及未解析项。
- 可选 QA 会把候选文档和参考模板分别渲染，便于检查页面结构差异。

### 目录结构

```text
TemplateToDoc/
  README.md
  pyproject.toml
  空管自动化软件设计说明_扩写版.docx
  profiles/
    <profile_id>/
      reference.docx
      profile.yaml
  projects/
    <project_id>/
      project.yaml
      final.docx
      replacement_report.yaml
      qa_compare.yaml
      render/
  src/template_to_doc_v2/
  tests/
```

### 注意事项

- 不建议把真实 API Key 写入仓库；优先使用环境变量或 Web 页面临时输入。
- `.doc` 模板会先通过 LibreOffice 转成 `.docx`，转换质量取决于本机 LibreOffice。
- 如果模板结构极其复杂，首次导入后建议检查 `profile.yaml` 中识别出的字段、章节和图片。
- 如果同级存在旧版 `TemplateToDoc/src/template_to_doc`，V2 会复用其中稳定的 QA 渲染比较能力；不存在时不影响模板导入和 Word 生成。

---

## English

TemplateToDocV2 is a document generation tool for reusing existing Word templates. It fills project fields, replaces section content, and updates images inside an existing `.docx` template, then exports `final.docx`. The generated file is based on the original template package, so headers, footers, page numbers, styles, tables, images, floating objects, margins, sections, and field settings are preserved as much as possible.

The typical workflow is to import a Word template, generate the template profile, edit project content, build the Word document, and optionally run a rendering QA check. Markdown text or LLM-generated content can be used as section draft content and written back to the corresponding locations in the template.

This repository includes `空管自动化软件设计说明_扩写版.docx` as the sample template used by the README commands and examples.

### Use Cases

- You already have a fixed Word template and only need to replace selected fields, sections, or images.
- The template contains complex tables, headers, footers, images, or mixed text-image layout that should keep its existing formatting.
- You want DeepSeek to generate section text while keeping the original Word formatting.
- You need a rendering check to compare the generated document against the reference template.

### Features

- Import `.docx` templates and create a `profile.yaml`.
- Import `.doc` templates when LibreOffice is available.
- Detect table fields, placeholders, section bodies, and images.
- Create an editable `project.yaml` for project fields, section prompts, section content, and image replacement settings.
- Keep original images by default or replace selected images.
- Generate missing sections with DeepSeek.
- Render and compare Word outputs through LibreOffice and Poppler.
- Provide both CLI and local Web workflows.

### Setup

Python 3.10 or later is recommended.

```powershell
git clone git@github.com:zy199336/TemplateToDoc.git
cd TemplateToDoc
python -m pip install -e .
```

Run tests:

```powershell
python -m pytest tests
```

Optional dependencies:

- DeepSeek generation: set `DEEPSEEK_API_KEY`, or enter the API key in the Web UI.
- `.doc` import: install LibreOffice and make sure `soffice` is available in PATH.
- QA rendering: install LibreOffice `soffice` and Poppler `pdftoppm`.

### CLI Usage

#### 1. Import a Template

```powershell
template-to-doc-v2 import-template .\空管自动化软件设计说明_扩写版.docx --profile atc_sdd --project-id atc_sdd_demo
```

This creates:

```text
profiles/atc_sdd/reference.docx
profiles/atc_sdd/profile.yaml
projects/atc_sdd_demo/project.yaml
```

If the package has not been installed as a console script, run the source entry directly:

```powershell
python src\template_to_doc_v2\cli.py import-template .\空管自动化软件设计说明_扩写版.docx --profile atc_sdd --project-id atc_sdd_demo
```

#### 2. Initialize a Project

If you did not pass `--project-id` during import:

```powershell
template-to-doc-v2 init-project atc_sdd atc_sdd_demo
```

The project file is written to:

```text
projects/atc_sdd_demo/project.yaml
```

#### 3. Edit project.yaml

`project.yaml` is the main input file:

```yaml
project_id: atc_sdd_demo
profile_id: atc_sdd
global:
  topic: ATC automation software design description
  background: This document describes the CSCI-level design of ATC automation software.
  materials: Software functions, interface design notes, operational data, and traceability tables.
  requirements: Preserve the technical-document tone, heading hierarchy, table structure, and mixed text-image layout.
fields:
  paragraph_field_01: ATC Automation Software Design Description
  paragraph_field_02: Document Control
  paragraph_field_03: Table of Contents
sections:
  1_4:
    prompt: Write the Scope section and explain the boundary of the ATC automation software design description.
    content: This document specifies the CSCI-level design description for ATC automation software...
images:
  image1:
    mode: keep
    file: null
```

Key fields:

- `global`: shared context used by section generation.
- `fields`: table or placeholder fields detected from the template.
- `sections.<id>.prompt`: local prompt for a section.
- `sections.<id>.content`: final text written back into the Word template.
- `images.<id>.mode: keep`: keep the original template image.
- `images.<id>.mode: replace`: replace the template image with an uploaded image.

#### 4. Build Word

Build with the content already filled in:

```powershell
template-to-doc-v2 build atc_sdd atc_sdd_demo
```

Outputs:

```text
projects/atc_sdd_demo/final.docx
projects/atc_sdd_demo/replacement_report.yaml
```

#### 5. Generate Missing Sections with DeepSeek

```powershell
$env:DEEPSEEK_API_KEY="sk-..."
template-to-doc-v2 build atc_sdd atc_sdd_demo --generate --model deepseek-v4-flash
```

When a section has empty `content`, TemplateToDocV2 uses `global`, the section `prompt`, and the template sample paragraph to generate text and write it back into the template.

#### 6. Build and Run QA

```powershell
template-to-doc-v2 build atc_sdd atc_sdd_demo --qa
```

Or compare an existing output:

```powershell
template-to-doc-v2 compare atc_sdd atc_sdd_demo
```

QA artifacts usually include:

```text
projects/atc_sdd_demo/qa_compare.yaml
projects/atc_sdd_demo/render/candidate/
projects/atc_sdd_demo/render/reference/
projects/atc_sdd_demo/render/diff/
```

### Web Usage

Start the local Web server:

```powershell
template-to-doc-v2 serve --port 8770
```

Open:

```text
http://127.0.0.1:8770/
```

Web workflow:

1. Select a local `.docx/.doc` file or enter a template path.
2. Enter `Profile ID` and `Project ID`, then import the template.
3. Fill global context, project fields, and section prompts.
4. Generate one section at a time, or generate all sections from the global context.
5. Keep original images by default, or upload a replacement image for selected targets.
6. Click “Generate Word” to download `final.docx`.
7. If LibreOffice and Poppler are installed, use the QA build option to render and compare pages.

### Example

The following example uses [空管自动化软件设计说明_扩写版.docx](空管自动化软件设计说明_扩写版.docx) as the template. Running the commands creates `projects/atc_sdd_demo` locally, builds `final.docx`, and renders selected pages to PNG for quick inspection.

#### Example Commands

```powershell
cd TemplateToDoc
template-to-doc-v2 import-template .\空管自动化软件设计说明_扩写版.docx --profile atc_sdd --project-id atc_sdd_demo
template-to-doc-v2 build atc_sdd atc_sdd_demo --qa
```

#### Example Outputs

```text
projects/atc_sdd_demo/final.docx
projects/atc_sdd_demo/replacement_report.yaml
projects/atc_sdd_demo/qa_compare.yaml
projects/atc_sdd_demo/render/candidate/final.pdf
projects/atc_sdd_demo/render/candidate/final-01.png
projects/atc_sdd_demo/render/candidate/final-03.png
projects/atc_sdd_demo/render/candidate/final-05.png
```

#### Rendered Results

After running `--qa`, inspect these generated files:

- `projects/atc_sdd_demo/render/candidate/final-01.png`: preserved cover title, cover image, document metadata table, and footer pagination.
- `projects/atc_sdd_demo/render/candidate/final-03.png`: preserved table-of-contents hierarchy, dot leaders, and page numbers.
- `projects/atc_sdd_demo/render/candidate/final-05.png`: preserved heading hierarchy, mixed text-image layout, and table formatting.

These rendered files are local build artifacts. The repository keeps only the template and source code, not generated outputs.

### Output Behavior

The generated `final.docx` has the following behavior:

- Original headers, footers, page numbers, and section styles are preserved.
- Tables are not rebuilt from Markdown, so borders, merged cells, and widths remain more stable.
- Detected table fields are written from `project.yaml` `fields`.
- Detected section bodies are replaced from `sections.<id>.content`.
- Template images are kept by default and can be replaced when `mode` is set to `replace`.
- `replacement_report.yaml` records replaced fields, sections, images, and unresolved items.
- Optional QA renders the candidate and reference documents for visual comparison.

### Directory Layout

```text
TemplateToDoc/
  README.md
  pyproject.toml
  空管自动化软件设计说明_扩写版.docx
  profiles/
    <profile_id>/
      reference.docx
      profile.yaml
  projects/
    <project_id>/
      project.yaml
      final.docx
      replacement_report.yaml
      qa_compare.yaml
      render/
  src/template_to_doc_v2/
  tests/
```

### Notes

- Do not commit real API keys. Prefer environment variables or temporary Web input.
- `.doc` templates are converted through LibreOffice before import, so conversion quality depends on LibreOffice.
- For highly complex templates, inspect `profile.yaml` after the first import and adjust detected fields, sections, or images if needed.
- If a legacy `TemplateToDoc/src/template_to_doc` package exists beside this repository, V2 can reuse its stable QA rendering utilities; otherwise template import and Word generation still work.
