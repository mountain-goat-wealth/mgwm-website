import sharp from 'sharp';
import { readdir, stat } from 'fs/promises';
import { join, basename, extname } from 'path';

// ─────────────────────────────────────────────────────────────────────────────
// SIZE_RULES — edit this to change output targets.
//
// Each rule is matched against the source filename (first match wins, unless
// a file has multiple rules — see logo-badge-source below for an example).
//
// Special fields:
//   extract       — crop the source before resizing: { left, top, width, height }
//   transform     — 'luminance-to-alpha': grayscale value → alpha, pixel → white
//                   (used to convert SVG-embedded logo masks to white-on-transparent)
//   outputFormat  — 'webp' (default) | 'png'
//   outputName    — override the output filename (default: <stem>.<outputFormat>)
//   multi         — set to true if this file has multiple rules (all will run)
// ─────────────────────────────────────────────────────────────────────────────
const SIZE_RULES = [
  { match: /^headshot-/i,          width: 640  },
  // Hero / full-width backgrounds
  { match: /^climber-hero/i,       width: 1920 },
  { match: /^happy-hikers/i,       width: 1920 },
  { match: /^mountain\.jpg/i,      width: 1920 },
  { match: /^summit-is-close/i,    width: 1920 },
  // Content images — 2× largest rendered width, capped at 1600
  { match: /^millcreek-overlook/i, width: 1440 }, // 720px × 2
  { match: /^mountain-goat/i,      width: 1600 }, // 880px × 2 → cap 1600
  { match: /^assessment-preview/i, width: 900  }, // 450px × 2
  { match: /^9-step/i,             width: 240  }, // 120px × 2
];

const DEFAULT_WIDTH = 1600;
const QUALITY      = 78;
const SRC_DIR      = new URL('../_image-originals/', import.meta.url).pathname;
const OUT_DIR      = new URL('../assets/images/', import.meta.url).pathname;

function getRules(filename) {
  const matches = SIZE_RULES.filter(r => r.match.test(filename));
  if (matches.length === 0) return [{ width: DEFAULT_WIDTH }];
  // If the first match has multi:true, return all matching rules; otherwise first only.
  return matches[0].multi ? matches : [matches[0]];
}

// Convert a grayscale PNG (white=opaque, black=transparent) to a white RGBA image.
async function luminanceToAlpha(srcPath, extract) {
  let pipeline = sharp(srcPath).grayscale();
  if (extract) pipeline = pipeline.extract(extract);
  const { data, info } = await pipeline.raw().toBuffer({ resolveWithObject: true });

  const pixels = info.width * info.height;
  const rgba   = Buffer.alloc(pixels * 4);
  for (let i = 0; i < pixels; i++) {
    rgba[i * 4]     = 255; // R
    rgba[i * 4 + 1] = 255; // G
    rgba[i * 4 + 2] = 255; // B
    rgba[i * 4 + 3] = data[i * info.channels]; // A = luminance
  }
  return sharp(rgba, { raw: { width: info.width, height: info.height, channels: 4 } });
}

async function processRule(srcPath, file, rule) {
  const stem      = basename(file, extname(file));
  const format    = rule.outputFormat ?? 'webp';
  const outName   = rule.outputName ?? `${stem}.${format}`;
  const outPath   = join(OUT_DIR, outName);
  const width     = rule.width ?? DEFAULT_WIDTH;

  let pipeline;
  if (rule.transform === 'luminance-to-alpha') {
    pipeline = await luminanceToAlpha(srcPath, rule.extract ?? null);
  } else {
    // .rotate() with no args auto-orients from EXIF before anything else.
    pipeline = sharp(srcPath).rotate();
    if (rule.extract) pipeline = pipeline.extract(rule.extract);
  }

  // Use post-rotation metadata for correct source dimensions (90°/270° swaps w/h).
  const rotatedMeta  = await sharp(srcPath).rotate().metadata();
  const meta         = rotatedMeta;
  const sourceWidth  = rule.extract ? rule.extract.width : (meta.width ?? width);
  const targetWidth  = Math.min(width, sourceWidth);

  pipeline = pipeline.resize({ width: targetWidth, withoutEnlargement: true });

  if (format === 'png') {
    pipeline = pipeline.png({ compressionLevel: 9 });
  } else {
    pipeline = pipeline.webp({ quality: QUALITY });
  }

  await pipeline.toFile(outPath);

  const { size: inSize }  = await stat(srcPath);
  const { size: outSize } = await stat(outPath);
  const pct = Math.round((1 - outSize / inSize) * 100);
  console.log(`${file.padEnd(52)} → ${outName.padEnd(26)} ${kb(inSize).padStart(8)} → ${kb(outSize).padStart(8)}  (${pct}% smaller)`);
}

async function main() {
  const files  = await readdir(SRC_DIR);
  const images = files.filter(f => /\.(png|jpe?g|webp|avif|tiff?)$/i.test(f));

  for (const file of images) {
    const srcPath = join(SRC_DIR, file);
    const rules   = getRules(file);
    for (const rule of rules) {
      await processRule(srcPath, file, rule);
    }
  }
}

function kb(bytes) { return `${Math.round(bytes / 1024)} KB`; }

main().catch(err => { console.error(err); process.exit(1); });
