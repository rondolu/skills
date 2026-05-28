#!/usr/bin/env node
/**
 * generate_from_structure.js
 *
 * Reads a presentation structure JSON (produced by extract_presentation_structure.py)
 * and generates an editable PPTX file using PptxGenJS, applying the Cathay Design System.
 *
 * Usage:
 *   node generate_from_structure.js <input.json> <output.pptx> [--no-cathay]
 *     [--cover-background <image-path>] [--content-background <image-path>]
 *
 * Options:
 *   --no-cathay   Skip Cathay Design System overrides; use original PDF styles as-is
 *   --cover-background <path>    Background image for cover (first) slide
 *   --content-background <path>  Background image for content slides
 *
 * Dependencies:
 *   npm install -g pptxgenjs
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const RESOURCE_FIELDS = ["image_file", "icon_file", "svg_file", "asset_file", "resource_file"];

// ---------------------------------------------------------------------------
// Cathay Design System
// ---------------------------------------------------------------------------

const CATHAY = {
  colors: {
    darkBlue:    "0843AD",
    techBlue:    "186AFF",
    dataTeal:    "36CBDA",
    textBlack:   "262626",
    subtitleGray:"A5A5A5",
    canvasWhite: "FFFFFF",
    sectionGray: "F2F2F2",
    energyYellow:"FFD600",
    actionOrange:"F75801",
    lightBlue:   "B2D9FC",
  },
  fonts: {
    english: "Arial",
    chinese: "Microsoft JhengHei",
  },
  roles: {
    title:          { fontSize: 30, bold: true,  color: "0843AD" },
    subtitle:       { fontSize: 22, bold: false, color: "186AFF" },
    section_header: { fontSize: 28, bold: true,  color: "262626" },
    body:           { fontSize: 14, bold: false, color: "262626" },
    bullet:         { fontSize: 14, bold: false, color: "262626" },
    annotation:     { fontSize: 12, bold: false, color: "A5A5A5" },
  },
  table: {
    headerFill:   "0843AD",
    headerColor:  "FFFFFF",
    rowFill:      "F2F2F2",
    altRowFill:   "FFFFFF",
    borderColor:  "A5A5A5",
    borderPt:     0.5,
  },
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Detect if text contains CJK characters → use Chinese font.
 */
function pickFont(text) {
  const cjkRegex = /[\u4E00-\u9FFF\u3400-\u4DBF\uF900-\uFAFF\u2E80-\u2FDF\u3000-\u303F]/;
  return cjkRegex.test(text) ? CATHAY.fonts.chinese : CATHAY.fonts.english;
}

/**
 * Scale coordinates from source page dimensions to target slide dimensions.
 */
function makeScaler(srcW, srcH, tgtW, tgtH) {
  const sx = tgtW / srcW;
  const sy = tgtH / srcH;
  // Use uniform scale to avoid distortion, fit within target
  const scale = Math.min(sx, sy);
  const offsetX = (tgtW - srcW * scale) / 2;
  const offsetY = (tgtH - srcH * scale) / 2;
  return {
    x: (v) => v * scale + offsetX,
    y: (v) => v * scale + offsetY,
    w: (v) => v * scale,
    h: (v) => v * scale,
    fontSize: (v) => Math.round(v * scale * 10) / 10,  // keep 1 decimal
  };
}

/**
 * Build a fresh shadow object (never reuse across calls — PptxGenJS mutates).
 */
function makeShadow() {
  return { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.10 };
}

/**
 * Convert image file to base64 data URI if it exists.
 * Tries the path as-is first, then resolves relative to jsonDir.
 */
function imageToBase64(filePath, jsonDir) {
  if (!filePath) return null;
  let resolved = filePath;
  if (!fs.existsSync(resolved) && jsonDir) {
    resolved = path.resolve(jsonDir, filePath);
  }
  if (!fs.existsSync(resolved)) return null;
  const ext = path.extname(resolved).toLowerCase().replace(".", "");
  const mime = ext === "jpg" ? "jpeg" : ext;
  const data = fs.readFileSync(resolved);
  return `image/${mime};base64,${data.toString("base64")}`;
}

/**
 * Check if an image element covers most of the source page (full-page image).
 */
function isFullPageImage(elem, srcW, srcH) {
  return elem.w > srcW * 0.8 && elem.h > srcH * 0.8;
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function buildBundlePaths(outputPptx) {
  const resolvedOutput = path.resolve(outputPptx);
  const parentDir = path.dirname(resolvedOutput);
  const ext = path.extname(resolvedOutput);
  const baseName = path.basename(resolvedOutput, ext || ".pptx");
  const bundleDir = path.join(parentDir, baseName);

  return {
    bundleDir,
    deckPath: path.join(bundleDir, `${baseName}.pptx`),
    assetsDir: path.join(bundleDir, "assets"),
    resourceManifestPath: path.join(bundleDir, "assets", "resource-manifest.json"),
  };
}

function collectResourceReferences(pages, extraRefs = []) {
  const refs = [];

  for (const page of pages) {
    for (const elem of page.elements || []) {
      for (const field of RESOURCE_FIELDS) {
        if (typeof elem[field] === "string" && elem[field].trim()) {
          refs.push({ field, originalPath: elem[field].trim() });
        }
      }
    }
  }

  for (const ref of extraRefs) {
    if (ref && typeof ref.originalPath === "string" && ref.originalPath.trim()) {
      refs.push({ field: ref.field || "resource_file", originalPath: ref.originalPath.trim() });
    }
  }

  return refs;
}

function makeSafeFileName(fileName) {
  return fileName.replace(/[<>:"/\\|?*]/g, "_");
}

function resolveResourcePath(resourcePath, jsonDir) {
  if (path.isAbsolute(resourcePath)) {
    return path.normalize(resourcePath);
  }
  return path.resolve(jsonDir, resourcePath);
}

function resolveBackgroundPath(backgroundPath, jsonDir) {
  if (!backgroundPath) return null;
  const resolved = path.isAbsolute(backgroundPath)
    ? path.normalize(backgroundPath)
    : path.resolve(jsonDir, backgroundPath);

  if (!fs.existsSync(resolved)) {
    console.warn(`  Warning: Background image not found: ${backgroundPath}`);
    return null;
  }

  return resolved;
}

function makeRelativeAssetPath(originalPath, resolvedPath, jsonDir) {
  if (!path.isAbsolute(originalPath)) {
    const normalized = path.normalize(originalPath);
    if (!normalized.startsWith("..") && !path.isAbsolute(normalized)) {
      return normalized;
    }
  }

  const relativeToJson = path.relative(jsonDir, resolvedPath);
  if (!relativeToJson.startsWith("..") && !path.isAbsolute(relativeToJson)) {
    return relativeToJson;
  }

  return path.basename(resolvedPath);
}

function uniquifyRelativePath(relativePath, usedPaths) {
  const parsed = path.parse(relativePath);
  let candidate = relativePath;
  let counter = 2;

  while (usedPaths.has(candidate.toLowerCase())) {
    candidate = path.join(parsed.dir, `${parsed.name}-${counter}${parsed.ext}`);
    counter += 1;
  }

  usedPaths.add(candidate.toLowerCase());
  return candidate;
}

function copyPresentationResources(resourceRefs, jsonDir, assetsDir, manifestPath) {
  ensureDir(assetsDir);

  const usedPaths = new Set();
  const manifest = [];
  const seenSourcePaths = new Set();

  for (const ref of resourceRefs) {
    const resolvedPath = resolveResourcePath(ref.originalPath, jsonDir);
    const sourceKey = resolvedPath.toLowerCase();

    if (!fs.existsSync(resolvedPath) || seenSourcePaths.has(sourceKey)) {
      continue;
    }

    seenSourcePaths.add(sourceKey);

    const rawRelativePath = makeRelativeAssetPath(ref.originalPath, resolvedPath, jsonDir)
      .split(path.sep)
      .map(makeSafeFileName)
      .join(path.sep);
    const relativeAssetPath = uniquifyRelativePath(rawRelativePath, usedPaths);
    const destinationPath = path.join(assetsDir, relativeAssetPath);

    ensureDir(path.dirname(destinationPath));
    fs.copyFileSync(resolvedPath, destinationPath);

    manifest.push({
      field: ref.field,
      source: ref.originalPath,
      copiedTo: path.join("assets", relativeAssetPath).replace(/\\/g, "/"),
    });
  }

  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2), "utf-8");
  return manifest;
}

// ---------------------------------------------------------------------------
// Element renderers
// ---------------------------------------------------------------------------

function renderShape(slide, elem, scaler, applyCathay) {
  const pres = slide._slideLayout ? slide._slideLayout._presLayout : null;

  const x = scaler.x(elem.x);
  const y = scaler.y(elem.y);
  const w = scaler.w(elem.w);
  const h = scaler.h(elem.h);

  if (elem.shape_type === "line") {
    const opts = {
      x, y, w, h: 0,
      line: {
        color: elem.stroke_color || CATHAY.colors.subtitleGray,
        width: Math.max(0.5, elem.line_width * 72),  // convert back to pt
      },
    };
    slide.addShape("line", opts);
    return;
  }

  // Rectangle
  const opts = { x, y, w, h };

  if (elem.fill_color) {
    opts.fill = { color: applyCathay ? mapColorToCathay(elem.fill_color) : elem.fill_color };
  }
  if (elem.stroke_color) {
    opts.line = {
      color: elem.stroke_color,
      width: Math.max(0.5, elem.line_width * 72),
    };
  }

  // Large rectangles covering most of the slide → treat as background decoration
  // Apply subtle shadow for card-like smaller shapes
  if (w < 8 && h < 4 && w > 0.5 && h > 0.3) {
    opts.shadow = makeShadow();
  }

  slide.addShape("rect", opts);
}


function renderImage(slide, elem, scaler, jsonDir) {
  const imgData = imageToBase64(elem.image_file, jsonDir);
  if (!imgData) {
    // No image file available — add a placeholder rectangle
    slide.addShape("rect", {
      x: scaler.x(elem.x),
      y: scaler.y(elem.y),
      w: scaler.w(elem.w),
      h: scaler.h(elem.h),
      fill: { color: CATHAY.colors.sectionGray },
      line: { color: CATHAY.colors.subtitleGray, width: 0.5 },
    });
    slide.addText("[Image]", {
      x: scaler.x(elem.x),
      y: scaler.y(elem.y),
      w: scaler.w(elem.w),
      h: scaler.h(elem.h),
      fontSize: 10,
      color: CATHAY.colors.subtitleGray,
      align: "center",
      valign: "middle",
    });
    return;
  }

  slide.addImage({
    data: imgData,
    x: scaler.x(elem.x),
    y: scaler.y(elem.y),
    w: scaler.w(elem.w),
    h: scaler.h(elem.h),
  });
}


function renderTextBlock(slide, elem, scaler, applyCathay) {
  const role = elem.role || "body";
  const cathayStyle = CATHAY.roles[role] || CATHAY.roles.body;

  // Build rich text array from runs
  const textParts = [];
  const runs = elem.runs || [];

  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    const text = run.text || "";
    if (!text) continue;

    const opts = {};

    if (applyCathay) {
      opts.fontSize = scaler.fontSize(cathayStyle.fontSize);
      opts.bold = cathayStyle.bold || run.bold;
      opts.italic = run.italic || false;
      opts.color = cathayStyle.color;
      opts.fontFace = pickFont(text);
    } else {
      opts.fontSize = scaler.fontSize(run.size || 14);
      opts.bold = run.bold || false;
      opts.italic = run.italic || false;
      opts.color = run.color || "262626";
      opts.fontFace = pickFont(text);
    }

    // Add breakLine between runs only if they represent separate visual lines
    if (i < runs.length - 1) {
      const nextRun = runs[i + 1];
      // Heuristic: if this run ends with newline-like content or next run
      // starts a new visual element, add breakLine
      if (text.endsWith("\n") || text.endsWith("\r")) {
        opts.breakLine = true;
      }
    }

    textParts.push({ text, options: opts });
  }

  if (textParts.length === 0) return;

  const blockOpts = {
    x: scaler.x(elem.x),
    y: scaler.y(elem.y),
    w: scaler.w(elem.w),
    h: scaler.h(elem.h),
    align: elem.align || "left",
    valign: "top",
    margin: 0,
    wrap: true,
  };

  // Bullets
  if (role === "bullet") {
    textParts.forEach((tp) => {
      tp.options.bullet = true;
    });
  }

  slide.addText(textParts, blockOpts);
}


function renderTable(slide, elem, scaler, applyCathay) {
  const rows = elem.rows || [];
  if (rows.length === 0) return;

  const tableData = rows.map((row, rowIdx) => {
    return row.map((cell) => {
      const cellOpts = {
        fontSize: scaler.fontSize(12),
        fontFace: pickFont(cell || ""),
        color: CATHAY.colors.textBlack,
        valign: "middle",
        margin: [2, 4, 2, 4],
      };

      if (applyCathay && rowIdx === 0) {
        // Header row
        cellOpts.fill = { color: CATHAY.table.headerFill };
        cellOpts.color = CATHAY.table.headerColor;
        cellOpts.bold = true;
        cellOpts.fontSize = scaler.fontSize(13);
      } else if (applyCathay && rowIdx % 2 === 0) {
        cellOpts.fill = { color: CATHAY.table.rowFill };
      }

      return { text: cell || "", options: cellOpts };
    });
  });

  const numCols = Math.max(...rows.map((r) => r.length));
  const totalW = scaler.w(elem.w);
  const colW = Array(numCols).fill(totalW / numCols);

  slide.addTable(tableData, {
    x: scaler.x(elem.x),
    y: scaler.y(elem.y),
    w: totalW,
    colW,
    border: { pt: CATHAY.table.borderPt, color: CATHAY.table.borderColor },
    autoPage: false,
  });
}


// ---------------------------------------------------------------------------
// Color mapping to Cathay palette
// ---------------------------------------------------------------------------

/**
 * Map an arbitrary hex color to the nearest Cathay Design System color.
 * Only used for shape fills, not text (text colors are set by role).
 */
function mapColorToCathay(hex) {
  if (!hex) return CATHAY.colors.sectionGray;
  const h = hex.toUpperCase();

  // Already a Cathay color
  const cathayValues = Object.values(CATHAY.colors).map((c) => c.toUpperCase());
  if (cathayValues.includes(h)) return hex;

  // Parse to RGB
  const r = parseInt(h.substring(0, 2), 16);
  const g = parseInt(h.substring(2, 4), 16);
  const b = parseInt(h.substring(4, 6), 16);
  if (isNaN(r) || isNaN(g) || isNaN(b)) return hex;

  // Luminance
  const lum = 0.299 * r + 0.587 * g + 0.114 * b;

  // Very dark → Dark Blue
  if (lum < 50) return CATHAY.colors.darkBlue;
  // Very light → Canvas White or Section Gray
  if (lum > 240) return CATHAY.colors.canvasWhite;
  if (lum > 220) return CATHAY.colors.sectionGray;

  // Blue-ish → Tech Blue or Data Teal
  if (b > r && b > g) {
    return (g > r) ? CATHAY.colors.dataTeal : CATHAY.colors.techBlue;
  }
  // Red/Orange-ish → Action Orange
  if (r > 200 && g < 150 && b < 100) return CATHAY.colors.actionOrange;
  // Yellow-ish → Energy Yellow
  if (r > 200 && g > 180 && b < 100) return CATHAY.colors.energyYellow;
  // Gray
  if (Math.abs(r - g) < 30 && Math.abs(g - b) < 30) {
    return lum > 150 ? CATHAY.colors.sectionGray : CATHAY.colors.subtitleGray;
  }

  // Default: keep original
  return hex;
}


// ---------------------------------------------------------------------------
// Main generation
// ---------------------------------------------------------------------------

function generate(inputJson, outputPptx, applyCathay = true, options = {}) {
  const data = JSON.parse(fs.readFileSync(inputJson, "utf-8"));
  const pages = data.pages || [];
  const isOcrSource = data.ocr_source === true;
  const jsonDir = path.dirname(path.resolve(inputJson));
  const bundlePaths = buildBundlePaths(outputPptx);
  const coverBackgroundPath = resolveBackgroundPath(options.coverBackground, jsonDir);
  const contentBackgroundPath = resolveBackgroundPath(options.contentBackground, jsonDir);

  if (pages.length === 0) {
    console.error("ERROR: No pages found in input JSON.");
    process.exit(1);
  }

  if (isOcrSource) {
    console.log("  OCR-sourced JSON detected.");
    console.log("  Segmented visual objects + editable text overlays.\n");
  }

  ensureDir(bundlePaths.bundleDir);

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "PDF-to-PPT Converter";
  pres.title = data.source_pdf || "Converted Presentation";

  // Target slide dimensions (LAYOUT_16x9)
  const TGT_W = 10;
  const TGT_H = 5.625;

  for (const pageData of pages) {
    console.log(`  Generating slide ${pageData.page_number} ...`);

    const srcW = pageData.width_inches || 10;
    const srcH = pageData.height_inches || 7.5;
    const scaler = makeScaler(srcW, srcH, TGT_W, TGT_H);

    const slide = pres.addSlide();
    const isCoverSlide = pageData.page_number === 1;
    const selectedBackgroundPath = isCoverSlide
      ? (coverBackgroundPath || contentBackgroundPath)
      : contentBackgroundPath;

    // Slide background
    if (selectedBackgroundPath) {
      slide.background = { path: selectedBackgroundPath };
    } else if (isOcrSource && pageData.bg_color) {
      // Use the detected background color from OCR segmentation
      slide.background = { color: pageData.bg_color };
    } else if (applyCathay) {
      slide.background = { color: CATHAY.colors.canvasWhite };
    }

    // Render elements in order (shapes → images → tables → text)
    const elements = pageData.elements || [];

    // Sort: shapes first, then images, then tables, then text (z-order)
    const order = { shape: 0, image: 1, table: 2, text_block: 3 };
    const sorted = [...elements].sort(
      (a, b) => (order[a.type] ?? 9) - (order[b.type] ?? 9)
    );

    for (const elem of sorted) {
      try {
        switch (elem.type) {
          case "shape":
            renderShape(slide, elem, scaler, applyCathay);
            break;
          case "image":
            renderImage(slide, elem, scaler, jsonDir);
            break;
          case "text_block":
            renderTextBlock(slide, elem, scaler, applyCathay);
            break;
          case "table":
            renderTable(slide, elem, scaler, applyCathay);
            break;
          default:
            console.warn(`  Unknown element type: ${elem.type}`);
        }
      } catch (err) {
        console.warn(`  Warning: Failed to render ${elem.type} on page ${pageData.page_number}: ${err.message}`);
      }
    }
  }

  // Write output
  pres.writeFile({ fileName: bundlePaths.deckPath })
    .then(() => {
      const manifest = copyPresentationResources(
        collectResourceReferences(pages, [
          options.coverBackground ? { field: "cover_background", originalPath: options.coverBackground } : null,
          options.contentBackground ? { field: "content_background", originalPath: options.contentBackground } : null,
        ]),
        jsonDir,
        bundlePaths.assetsDir,
        bundlePaths.resourceManifestPath
      );

      console.log(`\n  Presentation saved: ${bundlePaths.deckPath}`);
      console.log(`  Bundle folder: ${bundlePaths.bundleDir}`);
      console.log(`  Resource files copied: ${manifest.length}`);
      console.log(`  Resource manifest: ${bundlePaths.resourceManifestPath}`);
      console.log(`  Total slides: ${pages.length}`);
    })
    .catch((err) => {
      console.error(`ERROR writing PPTX: ${err.message}`);
      process.exit(1);
    });
}


// ---------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------

function main() {
  const args = process.argv.slice(2);

  const positional = [];
  let applyCathay = true;
  let coverBackground = null;
  let contentBackground = null;

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];

    if (arg === "--no-cathay") {
      applyCathay = false;
      continue;
    }

    if (arg === "--cover-background") {
      coverBackground = args[i + 1] || null;
      i += 1;
      continue;
    }

    if (arg === "--content-background") {
      contentBackground = args[i + 1] || null;
      i += 1;
      continue;
    }

    positional.push(arg);
  }

  if (positional.length < 2) {
    console.log("Usage: node generate_from_structure.js <input.json> <output.pptx> [--no-cathay] [--cover-background <image-path>] [--content-background <image-path>]");
    console.log("Writes the deck to <output-name>/<output-name>.pptx and copies referenced assets into <output-name>/assets/.");
    process.exit(1);
  }

  const inputJson = positional[0];
  const outputPptx = positional[1];

  if (!fs.existsSync(inputJson)) {
    console.error(`ERROR: File not found: ${inputJson}`);
    process.exit(1);
  }

  console.log(`Generating PPTX from: ${inputJson}`);
  console.log(`Output bundle name: ${path.basename(outputPptx, path.extname(outputPptx) || ".pptx")}`);
  console.log(`Cathay Design System: ${applyCathay ? "ON" : "OFF"}`);
  if (coverBackground) {
    console.log(`Cover background: ${coverBackground}`);
  }
  if (contentBackground) {
    console.log(`Content background: ${contentBackground}`);
  }

  generate(inputJson, outputPptx, applyCathay, {
    coverBackground,
    contentBackground,
  });
}

main();
