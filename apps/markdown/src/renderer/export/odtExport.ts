/**
 * md document (PM JSON) -> brand-new .odt bytes, fully local. The block/list/
 * table walk mirrors docxExport.ts (shared helpers imported from there) with
 * two odt-specific simplifications:
 *  - Lists need no numId allocation: odt-engine's <text:list> wrapping
 *    restarts numbering per contiguous run of same-kind items on its own, so
 *    every ordered list can share one constant numId (docx's numbering.xml
 *    has no such auto-restart, which is why docxExport.ts allocates one).
 *  - Block math has no ODF OMML equivalent, so it always keeps the LaTeX
 *    source visible as text (docx only falls back when the conversion fails).
 */
import type { JSONContent } from '@tiptap/core'
import { ODT_BULLET_NUM_ID, ODT_ORDERED_NUM_ID, saveOdt } from '@genoffice/odt-engine'
import type { OdtSaveBlock } from '@genoffice/odt-engine'
import type { GeneratedBlock, ParaFormat } from '@genoffice/docx-engine'
import {
  CODE_FILL,
  CODE_FONT,
  INDENT_STEP,
  MAX_LIST_LEVEL,
  mapTable,
  mergeFormat,
  plainText,
  runsFromInline,
} from './docxExport'
import type { ImageLoader } from './docxExport'

interface WalkContext {
  blocks: OdtSaveBlock[]
  loadImage: ImageLoader
  pendingImages: Array<{ index: number; src: string; alt: string }>
}

function pushText(ctx: WalkContext, block: GeneratedBlock): void {
  ctx.blocks.push({ kind: 'text', block })
}

function walkList(
  ctx: WalkContext,
  node: JSONContent,
  kind: 'bullet' | 'ordered',
  ilvl: number,
  base?: ParaFormat,
): void {
  const numId = kind === 'ordered' ? ODT_ORDERED_NUM_ID : ODT_BULLET_NUM_ID
  for (const item of node.content ?? []) {
    if (item.type !== 'listItem' && item.type !== 'taskItem') continue
    let firstPara = true
    for (const child of item.content ?? []) {
      if (child.type === 'paragraph') {
        const runs = runsFromInline(child.content)
        if (item.type === 'taskItem') {
          runs.unshift({ text: item.attrs?.checked ? '☑ ' : '☐ ' })
        }
        if (firstPara && item.type !== 'taskItem') {
          pushText(ctx, {
            type: 'listItem',
            list: { kind, numId, ilvl: Math.min(ilvl, MAX_LIST_LEVEL) },
            runs,
            format: base,
          })
        } else {
          pushText(ctx, {
            type: 'paragraph',
            runs,
            format: mergeFormat(base, { indentLeft: INDENT_STEP * (ilvl + 1) }),
          })
        }
        firstPara = false
      } else if (child.type === 'bulletList' || child.type === 'taskList') {
        walkList(ctx, child, 'bullet', ilvl + 1, base)
      } else if (child.type === 'orderedList') {
        walkList(ctx, child, 'ordered', ilvl + 1, base)
      } else {
        walkBlock(ctx, child, mergeFormat(base, { indentLeft: INDENT_STEP * (ilvl + 1) }))
      }
    }
  }
}

function walkBlock(ctx: WalkContext, node: JSONContent, base?: ParaFormat): void {
  switch (node.type) {
    case 'paragraph':
      pushText(ctx, { type: 'paragraph', runs: runsFromInline(node.content), format: base })
      break
    case 'heading': {
      const level = Math.min(Math.max(Number(node.attrs?.level) || 1, 1), 6) as number
      pushText(ctx, { type: 'heading', level, runs: runsFromInline(node.content), format: base })
      break
    }
    case 'bulletList':
      walkList(ctx, node, 'bullet', 0, base)
      break
    case 'orderedList':
      walkList(ctx, node, 'ordered', 0, base)
      break
    case 'taskList':
      walkList(ctx, node, 'bullet', 0, base)
      break
    case 'blockquote':
      for (const child of node.content ?? []) {
        walkBlock(ctx, child, mergeFormat(base, { indentLeft: INDENT_STEP, borders: 'l' }))
      }
      break
    case 'codeBlock':
      pushText(ctx, {
        type: 'paragraph',
        runs: [{ text: plainText(node), font: CODE_FONT, sizeHalfPoints: 19 }],
        format: mergeFormat(base, { shadingFill: CODE_FILL }),
      })
      break
    case 'horizontalRule':
      pushText(ctx, { type: 'paragraph', runs: [], format: mergeFormat(base, { borders: 'b' }) })
      break
    case 'image': {
      const src = String(node.attrs?.src ?? '')
      const alt = String(node.attrs?.alt ?? '')
      // placeholder now, replaced by the loaded image (or alt text) after the async pass
      ctx.pendingImages.push({ index: ctx.blocks.length, src, alt })
      ctx.blocks.push({ kind: 'text', block: { type: 'paragraph', runs: [] } })
      break
    }
    case 'table':
      ctx.blocks.push({ kind: 'table', model: mapTable(node) })
      break
    case 'blockMath': {
      const latex = String(node.attrs?.latex ?? '')
      pushText(ctx, {
        type: 'paragraph',
        runs: [{ text: `$$${latex}$$`, font: CODE_FONT }],
        format: base,
      })
      break
    }
    default: {
      // unknown block: keep its text so nothing silently disappears
      const text = plainText(node)
      if (text.trim()) pushText(ctx, { type: 'paragraph', runs: [{ text }], format: base })
    }
  }
}

/** Map a ProseMirror document to odt-engine SaveBlocks (pure except image loading) */
export async function mapDocToOdtSaveBlocks(
  doc: JSONContent,
  loadImage: ImageLoader,
): Promise<OdtSaveBlock[]> {
  const ctx: WalkContext = { blocks: [], loadImage, pendingImages: [] }
  for (const node of doc.content ?? []) walkBlock(ctx, node)

  for (const pending of ctx.pendingImages) {
    const image = await loadImage(pending.src).catch(() => null)
    ctx.blocks[pending.index] = image
      ? {
          kind: 'image',
          dataUrl: `data:${image.mime};base64,${image.base64}`,
          widthPx: image.widthPx,
          heightPx: image.heightPx,
        }
      : {
          kind: 'text',
          block: {
            type: 'paragraph',
            runs: [{ text: `[${pending.alt || pending.src}]`, italic: true, color: '888888' }],
          },
        }
  }

  return ctx.blocks
}

export async function exportOdtBytes(
  doc: JSONContent,
  loadImage: ImageLoader,
): Promise<Uint8Array> {
  return saveOdt(await mapDocToOdtSaveBlocks(doc, loadImage))
}
