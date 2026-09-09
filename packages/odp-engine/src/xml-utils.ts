/** Shared fast-xml-parser node helpers + XML text/attr escaping for generation. */

export type XmlNode = Record<string, unknown>

export function asXmlNode(value: unknown): XmlNode {
  return (value ?? {}) as XmlNode
}

/** fast-xml-parser gives a bare object for a single child, an array for repeats; always work with an array. */
export function xmlArray(value: unknown): XmlNode[] {
  if (value == null) return []
  return Array.isArray(value) ? (value as XmlNode[]) : [value as XmlNode]
}

/** Text content of a node that may be a bare string, or an object with a '#text' key (mixed content). */
export function xmlText(value: unknown): string {
  if (value == null) return ''
  if (typeof value === 'string' || typeof value === 'number') return String(value)
  const node = asXmlNode(value)
  return typeof node['#text'] === 'string' ? node['#text'] : ''
}

export function escapeXmlText(s: string): string {
  return s.replace(/[&<>]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;'))
}

export function escapeXmlAttr(s: string): string {
  return s.replace(/[&<>"]/g, (c) =>
    c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;',
  )
}

/**
 * Serialize a fast-xml-parser node back to an XML string. Used to preserve
 * passthrough elements (tables, charts, groups) that this engine cannot edit
 * but should not silently drop on save.
 */
export function serializeXmlNode(node: XmlNode, indent = ''): string {
  const parts: string[] = []
  for (const [key, value] of Object.entries(node)) {
    if (key === '@_xmlns' || key.startsWith('@_xmlns:')) {
      // skip namespace declarations — they're already in the document root
      continue
    }
    if (key === '#text') {
      parts.push(indent + escapeXmlText(String(value)))
      continue
    }
    if (key.startsWith('@_')) {
      // attributes handled by the caller
      continue
    }
    if (Array.isArray(value)) {
      for (const child of value) {
        if (child && typeof child === 'object') {
          parts.push(serializeXmlNode(child as XmlNode, indent))
        }
      }
    } else if (value && typeof value === 'object') {
      parts.push(serializeXmlNode(value as XmlNode, indent))
    }
  }
  return parts.join('')
}

/**
 * Reconstruct the full XML string of an element node (tag + attributes + children).
 * This is a best-effort serializer for passthrough elements.
 */
export function reconstructElementXml(node: XmlNode): string {
  // Find the element tag (first non-attribute key that isn't #text)
  const entries = Object.entries(node)
  let tag = ''
  const attrs: string[] = []
  const children: unknown[] = []

  for (const [key, value] of entries) {
    if (key === '#text') {
      children.push(value)
    } else if (key.startsWith('@_')) {
      const attrName = key.slice(2)
      attrs.push(`${attrName}="${escapeXmlAttr(String(value))}"`)
    } else if (Array.isArray(value)) {
      tag = key
      for (const child of value) {
        children.push(child)
      }
    } else if (value && typeof value === 'object') {
      tag = key
      children.push(value)
    } else if (typeof value === 'string' || typeof value === 'number') {
      tag = key
    }
  }

  if (!tag) return ''

  const attrStr = attrs.length > 0 ? ' ' + attrs.join(' ') : ''
  if (children.length === 0) {
    return `<${tag}${attrStr}/>`
  }

  const childParts: string[] = []
  for (const child of children) {
    if (child == null) continue
    if (typeof child === 'string' || typeof child === 'number') {
      childParts.push(escapeXmlText(String(child)))
    } else if (typeof child === 'object') {
      if (Array.isArray(child)) {
        for (const item of child) {
          if (item && typeof item === 'object') {
            childParts.push(reconstructElementXml(item as XmlNode))
          }
        }
      } else {
        childParts.push(reconstructElementXml(child as XmlNode))
      }
    }
  }

  return `<${tag}${attrStr}>${childParts.join('')}</${tag}>`
}
