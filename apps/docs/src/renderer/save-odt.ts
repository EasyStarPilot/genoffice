/**
 * ProseMirror doc -> OdtSaveBlock[] bridge, the .odt analog of file-actions'
 * buildDocBytes for .docx. docx-engine's own pmNodeToGeneratedBlock /
 * pmTableToModel already turn a top-level PM node into the same plain-data
 * shapes a .docx save uses (docHeading/docListItem/docParagraph -> a
 * GeneratedBlock, docTable -> a TableModel); this only adds the top-level
 * walk and the docProtected case, which those two functions don't handle
 * (docProtected is a catch-all "whole unit passed through unless it's an
 * image" node docx-engine invented for its own OOXML fidelity, so reading it
 * back out is odt-engine's problem alone).
 */
import type { OdtSaveBlock } from '@genoffice/odt-engine'
import { pmNodeToGeneratedBlock, pmTableToModel, type PmNode } from './editor/convert'

const TEXT_NODE_TYPES = new Set(['docHeading', 'docListItem', 'docParagraph'])

export function pmDocToOdtSaveBlocks(doc: PmNode): OdtSaveBlock[] {
  const blocks: OdtSaveBlock[] = []
  for (const node of doc.content ?? []) {
    if (node.type === 'docTable') {
      blocks.push({ kind: 'table', model: pmTableToModel(node) })
      continue
    }
    if (node.type === 'docProtected') {
      const dataUrl = node.attrs?.imageDataUrl as string | null | undefined
      if (dataUrl) {
        blocks.push({
          kind: 'image',
          dataUrl,
          widthPx: (node.attrs?.imageWidthPx as number | null) ?? undefined,
          heightPx: (node.attrs?.imageHeightPx as number | null) ?? undefined,
          align: (node.attrs?.imageAlign as 'left' | 'center' | 'right' | null) ?? undefined,
        })
      }
      // Anything else protected (chart, OLE object, textbox-only shape,
      // formula, diagram, broken/unrecognized drawing) has no ODF equivalent
      // this engine models yet: dropped rather than guessed at, same as an
      // image block whose data URL failed to resolve.
      continue
    }
    if (TEXT_NODE_TYPES.has(node.type)) {
      blocks.push({ kind: 'text', block: pmNodeToGeneratedBlock(node) })
    }
    // any other top-level node type has no odt-engine representation and is
    // silently dropped — the schema does not currently emit any (every block
    // becomes one of the four cases above), so this is a forward-compat guard
  }
  return blocks
}
