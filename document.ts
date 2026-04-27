import { z } from "npm:zod@4.3.6";
import JSZip from "npm:jszip@3.10.1";
// PDF rendering uses Deno.Command to invoke headless Chromium

// ============================================================================
// Text extraction helpers (shared by analyze, enhance, generate)
// ============================================================================

function extractOoxmlText(xml: string, tag: string): string[] {
  const results: string[] = [];
  const re = new RegExp(`<${tag}(?:\\s[^>]*)?>([^<]*)</${tag}>`, "g");
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    if (m[1].trim()) results.push(m[1]);
  }
  return results;
}

function replaceTextInXml(
  xml: string,
  tag: string,
  reps: Map<string, string>,
): string {
  let result = xml;
  for (const [orig, repl] of reps) {
    const escaped = orig.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    result = result.replace(
      new RegExp(`(<${tag}(?:\\s[^>]*)?>)(${escaped})(</${tag}>)`, "g"),
      `$1${repl}$3`,
    );
  }
  return result;
}

// ============================================================================
// Document structure analysis
// ============================================================================

interface DocSection {
  heading: string;
  headingStyle: string;
  paragraphs: string[];
}

interface SlideContent {
  index: number;
  title: string;
  textBlocks: string[];
  speakerNotes: string;
  xmlPath: string;
}

function analyzeDocxStructure(docXml: string): {
  sections: DocSection[];
  allText: { index: number; text: string }[];
} {
  const sections: DocSection[] = [];
  const allText: { index: number; text: string }[] = [];
  const paraRe = /<w:p[ >].*?<\/w:p>/gs;
  let m: RegExpExecArray | null;
  let paraIdx = 0;
  let current: DocSection = {
    heading: "(start)",
    headingStyle: "none",
    paragraphs: [],
  };

  while ((m = paraRe.exec(docXml)) !== null) {
    const para = m[0];
    const texts = extractOoxmlText(para, "w:t");
    const combined = texts.join("").trim();
    const styleMatch = para.match(/<w:pStyle\s+w:val="([^"]*)"/);
    const style = styleMatch?.[1] ?? "";
    const isHeading = /^Heading\d/.test(style) || style === "Title";

    if (combined) allText.push({ index: paraIdx, text: combined });

    if (isHeading && combined) {
      if (current.heading !== "(start)" || current.paragraphs.length > 0) {
        sections.push(current);
      }
      current = { heading: combined, headingStyle: style, paragraphs: [] };
    } else if (combined) {
      current.paragraphs.push(combined);
    }
    paraIdx++;
  }
  sections.push(current);
  return { sections, allText };
}

async function parsePptx(zip: JSZip) {
  const slides: SlideContent[] = [];
  const slideFiles = Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) =>
      parseInt(a.match(/slide(\d+)/)?.[1] ?? "0") -
      parseInt(b.match(/slide(\d+)/)?.[1] ?? "0")
    );

  for (let i = 0; i < slideFiles.length; i++) {
    const xmlPath = slideFiles[i];
    const xml = await zip.file(xmlPath)?.async("string");
    if (!xml) continue;
    const texts = extractOoxmlText(xml, "a:t");
    let speakerNotes = "";
    const noteNum = xmlPath.match(/slide(\d+)/)?.[1];
    const notesXml = await zip.file(`ppt/notesSlides/notesSlide${noteNum}.xml`)
      ?.async("string");
    if (notesXml) {
      speakerNotes = extractOoxmlText(notesXml, "a:t").join(" ").trim();
    }
    slides.push({
      index: i + 1,
      title: texts[0] ?? "",
      textBlocks: texts.slice(1),
      speakerNotes,
      xmlPath,
    });
  }
  return {
    slides,
    metadata: { format: "pptx", slideCount: String(slideFiles.length) },
  };
}

// ============================================================================
// OpenAI chat completions
// ============================================================================

async function chatCompletion(
  apiKey: string,
  model: string,
  systemPrompt: string,
  userMessage: string,
  maxTokens?: number,
): Promise<string> {
  const body: Record<string, unknown> = {
    model,
    messages: [{ role: "system", content: systemPrompt }, {
      role: "user",
      content: userMessage,
    }],
    temperature: 0.7,
  };
  if (maxTokens) body.max_tokens = maxTokens;

  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  if (!response.ok) {
    const err = await response.text();
    throw new Error(`OpenAI API error (${response.status}): ${err}`);
  }
  const result = await response.json();
  return result.choices[0].message.content;
}

// ============================================================================
// Schema
// ============================================================================

const GlobalArgsSchema = z.object({
  apiKey: z
    .string()
    .meta({ sensitive: true })
    .describe("OpenAI API key. Use: ${{ vault.get('openai', 'api-key') }}"),
});

// ============================================================================
// Model definition
// ============================================================================

export const model = {
  type: "@dougschaefer/openai-document",
  version: "2026.04.17.2",
  globalArguments: GlobalArgsSchema,
  resources: {
    analysis: {
      description: "Parsed document structure and content",
      schema: z.object({
        filePath: z.string(),
        fileType: z.enum(["pptx", "docx"]),
        slideCount: z.number().optional(),
        sectionCount: z.number().optional(),
        paragraphCount: z.number().optional(),
        content: z.string().describe("JSON stringified content structure"),
      }),
      lifetime: "infinite" as const,
      garbageCollection: 10,
    },
    document: {
      description: "Designed, generated, or enhanced document file",
      schema: z.object({
        sourceFile: z.string().optional(),
        outputFile: z.string(),
        fileType: z.enum(["pptx", "docx", "pdf", "html"]),
        instructions: z.string(),
        model: z.string(),
      }),
      lifetime: "infinite" as const,
      garbageCollection: 5,
    },
  },
  methods: {
    // ================================================================
    // analyze
    // ================================================================
    analyze: {
      description:
        "Parse a .pptx or .docx file and return its structure, sections, and text content.",
      arguments: z.object({
        filePath: z.string().describe(
          "Absolute path to the .pptx or .docx file",
        ),
        outputName: z.string().default("analysis"),
      }),
      execute: async (args, context) => {
        const lower = args.filePath.toLowerCase();
        const fileType = lower.endsWith(".pptx")
          ? "pptx"
          : lower.endsWith(".docx")
          ? "docx"
          : null;
        if (!fileType) {
          throw new Error("Unsupported file type. Must be .pptx or .docx");
        }

        context.logger.info("Analyzing {type}: {path}", {
          type: fileType,
          path: args.filePath,
        });
        const zip = await JSZip.loadAsync(await Deno.readFile(args.filePath));

        let content: unknown;
        let slideCount: number | undefined;
        let sectionCount: number | undefined;
        let paragraphCount: number | undefined;

        if (fileType === "pptx") {
          const { slides, metadata } = await parsePptx(zip);
          slideCount = slides.length;
          content = { fileType, slides, metadata };
        } else {
          const docXml = await zip.file("word/document.xml")?.async("string");
          if (!docXml) throw new Error("Invalid .docx");
          const { sections, allText } = analyzeDocxStructure(docXml);
          sectionCount = sections.length;
          paragraphCount = allText.length;
          content = {
            fileType,
            sections,
            metadata: { sectionCount, paragraphCount },
          };
        }

        const handle = await context.writeResource(
          "analysis",
          args.outputName,
          {
            filePath: args.filePath,
            fileType,
            slideCount,
            sectionCount,
            paragraphCount,
            content: JSON.stringify(content, null, 2),
          },
        );
        return { dataHandles: [handle] };
      },
    },

    // ================================================================
    // design — HTML/CSS → PDF via GPT + headless Chromium
    // ================================================================
    design: {
      description:
        "Design a professional document as PDF. GPT generates complete HTML with inline CSS based on your design instructions and content, then headless Chromium renders it to a pixel-perfect PDF. Supports designing from scratch, from a content description, or from an existing .docx file's extracted content.",
      arguments: z.object({
        instructions: z.string().describe(
          "Design and content direction — describe the visual style, layout, colors, and what content to include. E.g., 'Modern proposal for Snap Inc, Kentucky blue and white color scheme, clean layout with geometric accents, two-column sections where appropriate. Cover page with large color block and company name.'",
        ),
        contentSource: z.string().optional().describe(
          "Optional: absolute path to a .docx file to extract content from. The design method will pull the text and restructure it into the new design. If omitted, GPT generates content based on instructions alone.",
        ),
        context: z.string().optional().describe(
          "Additional context — company info, client details, scope data, brand guidelines, etc.",
        ),
        model: z
          .enum(["gpt-4o", "gpt-4o-mini", "gpt-4.1", "gpt-4.1-mini"])
          .default("gpt-4.1")
          .describe("OpenAI model for HTML generation"),
        outputName: z.string().default("designed"),
      }),
      execute: async (args, context) => {
        const g = context.globalArgs;

        // Extract content from source document if provided
        let sourceContent = "";
        if (args.contentSource) {
          context.logger.info("Extracting content from: {path}", {
            path: args.contentSource,
          });
          const zip = await JSZip.loadAsync(
            await Deno.readFile(args.contentSource),
          );
          const docXml = await zip.file("word/document.xml")?.async("string");
          if (!docXml) throw new Error("Invalid .docx source file");
          const { sections } = analyzeDocxStructure(docXml);
          sourceContent = sections
            .map((s) => {
              const body = s.paragraphs.join("\n\n");
              return s.headingStyle !== "none"
                ? `## ${s.heading}\n\n${body}`
                : body;
            })
            .join("\n\n");
          context.logger.info("Extracted {chars} characters of content", {
            chars: sourceContent.length,
          });
        }

        const systemPrompt =
          `You are an expert document designer who creates stunning, professional business documents using HTML and CSS. You produce complete, self-contained HTML files that render beautifully when printed to PDF.

OUTPUT: Return ONLY a complete HTML document. No explanation, no markdown fences, no commentary — just the HTML starting with <!DOCTYPE html> and ending with </html>.

DESIGN REQUIREMENTS:
- The document must be print-optimized for US Letter (8.5" × 11") pages
- Use @page CSS rules to set margins and page size
- Use page-break-before/page-break-after to control page breaks between major sections
- All CSS must be inline in a <style> tag — no external resources except Google Fonts
- You may use Google Fonts via @import for modern typography (e.g., Inter, Poppins, Montserrat, Open Sans)
- Use CSS for ALL visual design: backgrounds, geometric shapes (via pseudo-elements, borders, gradients), column layouts (CSS grid/flexbox), whitespace, color blocks
- No JavaScript

DESIGN QUALITY — this is critical:
- Study professional document templates from design agencies. The output should look like it was designed in InDesign, not generated by a computer.
- Use large geometric color blocks to define visual zones on each page (like a colored left third, or a top banner spanning the page)
- Create visual hierarchy through size, weight, color, and spacing — not just bold/italic
- Generous whitespace. Don't pack content. Let the design breathe.
- Section headings should be design elements themselves, not just bigger text — use background colors, overlapping shapes, accent lines
- Use two-column layouts where appropriate for body content
- Cover page should be visually striking — large color block, company name prominent, minimal text
- Tables should be clean and modern — no heavy borders, use alternating row shading and header styling
- Footer with page numbers on each page
- Consistent color theme throughout — primary color for major elements, accent for highlights, neutrals for backgrounds

PAGE STRUCTURE for a typical business proposal:
1. Cover page (visually bold, large color block, project title, company names)
2. Letter/introduction page
3. Content sections (scope, phases, deliverables, etc.)
4. Pricing/table pages
5. Signature/closing page

CSS TECHNIQUES to use:
- background-color on divs for color blocks
- CSS Grid for multi-column layouts
- position: relative/absolute for overlapping geometric elements
- border-left: 4px solid <color> for accent bars
- clip-path or border-radius for geometric shapes
- linear-gradient for subtle backgrounds
- box-shadow for depth on cards/callouts
- ::before/::after pseudo-elements for decorative shapes`;

        let userMsg = `DESIGN INSTRUCTIONS: ${args.instructions}`;

        if (args.context) {
          userMsg += `\n\nADDITIONAL CONTEXT:\n${args.context}`;
        }

        if (sourceContent) {
          userMsg +=
            `\n\nDOCUMENT CONTENT TO INCORPORATE (restructure and design this content — preserve all factual details, pricing, and scope items but improve the presentation):\n\n${sourceContent}`;
        }

        context.logger.info("Generating HTML design with {model}", {
          model: args.model,
        });
        const htmlContent = await chatCompletion(
          g.apiKey,
          args.model,
          systemPrompt,
          userMsg,
          16000,
        );

        // Strip any markdown fences if GPT wraps it
        const cleanHtml = htmlContent
          .replace(/^```html\n?/i, "")
          .replace(/\n?```$/i, "")
          .trim();

        // Write HTML file
        const outputDir = `${context.repoDir}/.swamp/generated-documents`;
        await Deno.mkdir(outputDir, { recursive: true });
        const timestamp = Date.now();
        const htmlPath = `${outputDir}/${args.outputName}-${timestamp}.html`;
        const pdfPath = `${outputDir}/${args.outputName}-${timestamp}.pdf`;

        await Deno.writeTextFile(htmlPath, cleanHtml);
        context.logger.info("HTML saved to {path} ({size} chars)", {
          path: htmlPath,
          size: cleanHtml.length,
        });

        // Render to PDF via headless Chromium
        context.logger.info("Rendering PDF via headless Chromium");
        const chromiumCmd = new Deno.Command("chromium-browser", {
          args: [
            "--headless",
            "--disable-gpu",
            "--no-sandbox",
            "--disable-software-rasterizer",
            `--print-to-pdf=${pdfPath}`,
            "--print-to-pdf-no-header",
            htmlPath,
          ],
          stdout: "piped",
          stderr: "piped",
        });
        const chromiumResult = await chromiumCmd.output();

        if (!chromiumResult.success) {
          // Chromium often returns non-zero in WSL due to DBus errors but still produces the PDF
          try {
            await Deno.stat(pdfPath);
            context.logger.info(
              "Chromium returned non-zero but PDF was created (WSL DBus noise)",
            );
          } catch {
            const stderr = new TextDecoder().decode(chromiumResult.stderr);
            throw new Error(`Chromium PDF rendering failed: ${stderr}`);
          }
        }

        const pdfStat = await Deno.stat(pdfPath);
        context.logger.info("PDF rendered: {path} ({size} bytes)", {
          path: pdfPath,
          size: pdfStat.size,
        });

        // Copy to Windows Downloads for easy access
        const winDownloadsPath = "/mnt/c/Users/DougSchaefer/Downloads";
        try {
          await Deno.stat(winDownloadsPath);
          const winPdfName = `${args.outputName}-${timestamp}.pdf`;
          const winPdfPath = `${winDownloadsPath}/${winPdfName}`;
          await Deno.copyFile(pdfPath, winPdfPath);
          context.logger.info("Copied to Windows Downloads: {path}", {
            path: winPdfPath,
          });
        } catch {
          // Windows path not accessible, skip
        }

        const handle = await context.writeResource(
          "document",
          args.outputName,
          {
            sourceFile: args.contentSource,
            outputFile: pdfPath,
            fileType: "pdf",
            instructions: args.instructions,
            model: args.model,
          },
        );

        return { dataHandles: [handle] };
      },
    },

    // ================================================================
    // enhance — text content modification on .docx/.pptx
    // ================================================================
    enhance: {
      description:
        "Enhance or modify text content of a .pptx or .docx using GPT, preserving formatting and design.",
      arguments: z.object({
        filePath: z.string().describe(
          "Absolute path to the .pptx or .docx file",
        ),
        instructions: z.string().describe("What to change in the text content"),
        model: z.enum(["gpt-4o", "gpt-4o-mini", "gpt-4.1", "gpt-4.1-mini"])
          .default("gpt-4o"),
        outputName: z.string().default("enhanced"),
      }),
      execute: async (args, context) => {
        const g = context.globalArgs;
        const lower = args.filePath.toLowerCase();
        const fileType = lower.endsWith(".pptx") ? "pptx" : "docx";
        context.logger.info("Enhancing {type} with {model}", {
          type: fileType,
          model: args.model,
        });
        const zip = await JSZip.loadAsync(await Deno.readFile(args.filePath));

        if (fileType === "pptx") {
          const { slides } = await parsePptx(zip);
          const content = slides.map((s) =>
            `Slide ${s.index}: Title: "${s.title}"\nContent: ${
              s.textBlocks.join(" | ")
            }`
          ).join("\n\n");
          const resp = await chatCompletion(
            g.apiKey,
            args.model,
            `You are a document editor. Return ONLY a JSON array: [{"slide": <num>, "original": "<text>", "replacement": "<text>"}]. No markdown fences.`,
            `CONTENT:\n${content}\n\nINSTRUCTIONS: ${args.instructions}`,
          );
          for (
            const rep of JSON.parse(
              resp.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim(),
            )
          ) {
            const slide = slides.find((s) => s.index === rep.slide);
            if (!slide) continue;
            const xml = await zip.file(slide.xmlPath)?.async("string");
            if (!xml) continue;
            zip.file(
              slide.xmlPath,
              replaceTextInXml(
                xml,
                "a:t",
                new Map([[rep.original, rep.replacement]]),
              ),
            );
          }
        } else {
          let docXml = await zip.file("word/document.xml")?.async("string");
          if (!docXml) throw new Error("Invalid .docx");
          const { allText } = analyzeDocxStructure(docXml);
          const content = allText.map((t) => `[${t.index}] ${t.text}`).join(
            "\n",
          );
          const resp = await chatCompletion(
            g.apiKey,
            args.model,
            `You are a document editor. Return ONLY a JSON array: [{"original": "<text>", "replacement": "<text>"}]. No markdown fences.`,
            `CONTENT:\n${content}\n\nINSTRUCTIONS: ${args.instructions}`,
          );
          const reps = JSON.parse(
            resp.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim(),
          );
          docXml = replaceTextInXml(
            docXml,
            "w:t",
            new Map(
              reps.map((
                r: { original: string; replacement: string },
              ) => [r.original, r.replacement]),
            ),
          );
          zip.file("word/document.xml", docXml);
        }

        const outputDir = `${context.repoDir}/.swamp/generated-documents`;
        await Deno.mkdir(outputDir, { recursive: true });
        const ext = fileType === "pptx" ? ".pptx" : ".docx";
        const fileName = `${args.outputName}-${Date.now()}${ext}`;
        const outputPath = `${outputDir}/${fileName}`;
        await Deno.writeFile(
          outputPath,
          await zip.generateAsync({ type: "uint8array" }),
        );
        context.logger.info("Enhanced document saved to {path}", {
          path: outputPath,
        });
        const handle = await context.writeResource(
          "document",
          args.outputName,
          {
            sourceFile: args.filePath,
            outputFile: outputPath,
            fileType,
            instructions: args.instructions,
            model: args.model,
          },
        );
        return { dataHandles: [handle] };
      },
    },

    // ================================================================
    // generate — populate template with GPT content
    // ================================================================
    generate: {
      description:
        "Generate a new document from a .docx/.pptx template by populating placeholders with GPT-generated content.",
      arguments: z.object({
        templatePath: z.string().describe(
          "Absolute path to the template .pptx or .docx file",
        ),
        instructions: z.string().describe("What content to generate"),
        context: z.string().optional().describe("Additional context"),
        model: z.enum(["gpt-4o", "gpt-4o-mini", "gpt-4.1", "gpt-4.1-mini"])
          .default("gpt-4o"),
        outputName: z.string().default("generated"),
      }),
      execute: async (args, context) => {
        const g = context.globalArgs;
        const lower = args.templatePath.toLowerCase();
        const fileType = lower.endsWith(".pptx") ? "pptx" : "docx";
        context.logger.info("Generating {type} from template", {
          type: fileType,
        });
        const zip = await JSZip.loadAsync(
          await Deno.readFile(args.templatePath),
        );

        if (fileType === "pptx") {
          const { slides } = await parsePptx(zip);
          const structure = slides.map((s) =>
            `Slide ${s.index}: Title: "${s.title}", Content: [${
              s.textBlocks.map((t) => `"${t}"`).join(", ")
            }]`
          ).join("\n");
          const resp = await chatCompletion(
            g.apiKey,
            args.model,
            `You are a content generator. Return ONLY a JSON array: [{"slide": <num>, "original": "<placeholder>", "replacement": "<content>"}]. No markdown fences.`,
            `TEMPLATE:\n${structure}\n\nINSTRUCTIONS: ${args.instructions}${
              args.context ? `\n\nCONTEXT:\n${args.context}` : ""
            }`,
          );
          for (
            const rep of JSON.parse(
              resp.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim(),
            )
          ) {
            const slide = slides.find((s) => s.index === rep.slide);
            if (!slide) continue;
            const xml = await zip.file(slide.xmlPath)?.async("string");
            if (!xml) continue;
            zip.file(
              slide.xmlPath,
              replaceTextInXml(
                xml,
                "a:t",
                new Map([[rep.original, rep.replacement]]),
              ),
            );
          }
        } else {
          let docXml = await zip.file("word/document.xml")?.async("string");
          if (!docXml) throw new Error("Invalid .docx");
          const { allText } = analyzeDocxStructure(docXml);
          const structure = allText.map((t) => `[${t.index}] "${t.text}"`).join(
            "\n",
          );
          const resp = await chatCompletion(
            g.apiKey,
            args.model,
            `You are a content generator. Return ONLY a JSON array: [{"original": "<placeholder>", "replacement": "<content>"}]. No markdown fences.`,
            `TEMPLATE:\n${structure}\n\nINSTRUCTIONS: ${args.instructions}${
              args.context ? `\n\nCONTEXT:\n${args.context}` : ""
            }`,
          );
          const reps = JSON.parse(
            resp.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim(),
          );
          docXml = replaceTextInXml(
            docXml,
            "w:t",
            new Map(
              reps.map((
                r: { original: string; replacement: string },
              ) => [r.original, r.replacement]),
            ),
          );
          zip.file("word/document.xml", docXml);
        }

        const outputDir = `${context.repoDir}/.swamp/generated-documents`;
        await Deno.mkdir(outputDir, { recursive: true });
        const ext = fileType === "pptx" ? ".pptx" : ".docx";
        const fileName = `${args.outputName}-${Date.now()}${ext}`;
        const outputPath = `${outputDir}/${fileName}`;
        await Deno.writeFile(
          outputPath,
          await zip.generateAsync({ type: "uint8array" }),
        );
        context.logger.info("Generated document saved to {path}", {
          path: outputPath,
        });
        const handle = await context.writeResource(
          "document",
          args.outputName,
          {
            sourceFile: args.templatePath,
            outputFile: outputPath,
            fileType,
            instructions: args.instructions,
            model: args.model,
          },
        );
        return { dataHandles: [handle] };
      },
    },
  },
};
