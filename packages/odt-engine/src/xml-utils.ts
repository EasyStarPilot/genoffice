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
