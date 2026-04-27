# OpenAI Document Extension for Swamp

A swamp extension that reads, modifies, and generates Microsoft Office documents (`.pptx` and `.docx`) using the OpenAI API. Text content can be analyzed, rewritten, or generated from scratch while keeping the original formatting and design intact, and a separate `design` method composes pixel-perfect PDFs from natural-language instructions by having GPT produce a complete HTML+CSS document and rendering it through headless Chromium.

This is built for the case where you have a working template — a proposal, a deck, a report — and you want AI to fill it in, polish it, or restructure its content without producing the unstyled, default-formatted output you get from typical AI document tools. The extension parses OOXML directly, so it works with any compliant `.pptx` or `.docx` file regardless of the editor that produced it.

## Prerequisites

- An OpenAI API account ([platform.openai.com](https://platform.openai.com)) with access to GPT-4o or GPT-4.1 models
- Swamp installed and a repository initialized
- For the `design` method, `chromium-browser` available on `PATH` (used to render HTML to PDF). On Debian/Ubuntu/WSL: `sudo apt install chromium-browser`. On other platforms, install Chromium and ensure the binary is named `chromium-browser`.

The other three methods (`analyze`, `enhance`, `generate`) operate purely on OOXML and don't require Chromium.

## Installation

```bash
swamp extension pull @dougschaefer/openai-document
```

## Setup

Create a vault and store your API key:

```bash
swamp vault create local_encryption openai
swamp vault put openai api-key
```

Create a model instance wired to the vault:

```bash
swamp model create @dougschaefer/openai-document docs \
  --global-arg 'apiKey=${{ vault.get("openai", "api-key") }}'
```

## Methods

| Method | Description |
|--------|-------------|
| `analyze` | Parse a `.pptx` or `.docx` and return its structure (slides, sections, paragraphs, speaker notes) as a JSON-stringified content tree |
| `enhance` | Rewrite the text of an existing `.pptx` or `.docx` per natural-language instructions, preserving every aspect of the formatting, layout, and design |
| `generate` | Populate a `.pptx` or `.docx` template with new content. Treats the source file as a structural template and replaces its text with GPT-generated content per your instructions |
| `design` | Design a new PDF from scratch. GPT writes a complete HTML+CSS document based on your design and content instructions, and headless Chromium renders it to a pixel-perfect PDF. Optionally pulls source content from an existing `.docx` |

## Usage

### Analyze a Document

```bash
swamp model method run docs analyze --input '{
  "filePath": "/absolute/path/to/proposal.docx"
}'
```

Returns the document's heading structure, paragraph count, and full text content. The output artifact has a `content` field (JSON-stringified) that downstream models or workflows can read with CEL expressions to drive further automation.

### Enhance an Existing Document

```bash
swamp model method run docs enhance --input '{
  "filePath": "/absolute/path/to/draft.docx",
  "instructions": "Tighten the executive summary, remove passive voice throughout, and replace any reference to dates with the actual date format MMMM D, YYYY."
}'
```

The original file's formatting, fonts, colors, images, and layout are preserved. Only the text content inside the OOXML text elements is replaced with the enhanced versions.

### Generate from a Template

```bash
swamp model method run docs generate --input '{
  "templatePath": "/absolute/path/to/template.pptx",
  "instructions": "Sales deck for Acme Corp covering our managed services offering for AV-IT convergence",
  "context": "Acme has 47 conference rooms across 6 sites, mixed Cisco and MTR endpoints."
}'
```

The slide layouts, master slides, theme colors, and image placements stay; the text in each slide is replaced with content GPT generates from your instructions and context.

### Design a PDF From Instructions

```bash
swamp model method run docs design --input '{
  "instructions": "Modern proposal for Snap Inc, Kentucky blue and white color scheme, clean two-column layout with geometric accents. Cover page with full-bleed color block and logo placement, then sections for executive summary, scope, timeline, and pricing.",
  "context": "Snap Inc is a 2,000-employee SaaS company evaluating our managed AV services for their 3 offices.",
  "model": "gpt-4.1"
}'
```

Use `contentSource` to seed the design with content extracted from an existing `.docx`:

```bash
swamp model method run docs design --input '{
  "contentSource": "/absolute/path/to/old-proposal.docx",
  "instructions": "Restructure this proposal into a modern two-column design with our brand colors (#1E40AF, #FFFFFF, #F59E0B). Keep all the content but reflow it for visual clarity.",
  "model": "gpt-4.1"
}'
```

## Model Selection

All methods that call OpenAI accept a `model` parameter. Defaults vary per method:

- `enhance` and `generate` default to `gpt-4o` — sufficient for text rewriting tasks
- `design` defaults to `gpt-4.1` — produces noticeably better HTML/CSS output for PDF rendering

Available models: `gpt-4o`, `gpt-4o-mini`, `gpt-4.1`, `gpt-4.1-mini`. The `-mini` variants are cheaper but produce lower-quality output for complex layouts and long-form text.

## Outputs

Generated files are written to disk as standard `.pptx`, `.docx`, or `.pdf` files alongside the source. Resource artifacts are tracked with version history and the file paths are recorded as data attributes that downstream models can reference via CEL.

## Cost

OpenAI charges per token. A typical `enhance` run on a 10-page document costs $0.05–$0.20 depending on the model. The `design` method uses more tokens (it produces complete HTML) and costs $0.30–$1.00 per document. Current pricing is at [openai.com/api/pricing](https://openai.com/api/pricing).

## Quality and Testing

This extension has been tested against the OpenAI API in the American Sound integration lab. American Sound is solely responsible for this integration. OpenAI does not provide direct support for third-party swamp extensions.

## License

MIT. See [LICENSE](LICENSE) for details.
