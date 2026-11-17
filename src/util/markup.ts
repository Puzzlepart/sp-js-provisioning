/**
 * CAML or HTML as SharePoint stores it, for comparing: without the whitespace between tags, around
 * an attribute's `=` and before a tag's end, and with an attribute in double quotes (a template's
 * `<FieldRef Name= "X"/>` or `<FieldRef Name='X' />` is stored as `<FieldRef Name="X" />`). A view's
 * `ViewQuery` compared this way is the same when only the layout differs.
 *
 * @param text - The CAML or HTML
 */
export function compactMarkup(text: any): string {
  return String(text ?? '')
    .replace(/>\s+</g, '><')
    .replace(/<[^<>]+>/g, (tag) =>
      tag
        .replace(/\s*=\s*'([^"']*)'/g, '="$1"')
        .replace(/\s*=\s*(["'])/g, '=$1')
        .replace(/\s+/g, ' ')
        .replace(/\s*(\/?)>$/, (_end, slash) => (slash ? ' />' : '>'))
    )
    .trim()
}
