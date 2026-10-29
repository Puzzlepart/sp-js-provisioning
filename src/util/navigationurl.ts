/**
 * The web a navigation node URL is resolved against.
 */
export interface INavigationUrlBase {
  /** The web's server-relative URL, e.g. `/sites/demo/sub` (`/` for a root site) */
  serverRelativeUrl: string
  /** The web's origin, e.g. `https://contoso.sharepoint.com`: an absolute URL on it is internal */
  origin?: string
}

/** `scheme:` at the start of a URL. */
const SCHEME = /^[a-z][\d+.a-z-]*:/i

/** `scheme://host[:port]` at the start of a URL. */
const ORIGIN = /^[a-z][\d+.a-z-]*:\/\/[^#/?]*/i

/**
 * A navigation node URL in the form SharePoint stores it, for comparison only: the result is never
 * sent. Both the template's URL (tokens replaced) and the stored one go through it, and two URLs are
 * the same when the results are equal. The rules follow what SharePoint stored on the tenant probe
 * (docs/specs/navigation-re-apply.md, N1):
 *
 * - an empty URL is the web, and `#` the web followed by `#`;
 * - a web-relative URL (`SitePages/Home.aspx`) is resolved against the web, and an absolute URL on
 *   the web's origin becomes server-relative;
 * - in a server-relative URL the path is decoded (`%20` and a space are the same), loses a trailing
 *   slash and is compared without case (SharePoint stores the target's own case); a query string or
 *   fragment is kept as is;
 * - any other absolute URL (another host, `http://linkless.header/`, `mailto:`) is external and
 *   kept verbatim, as SharePoint stores it.
 *
 * @param url - The URL, with its tokens replaced
 * @param base - The web it is resolved against
 */
export function normaliseNavigationUrl(
  url: string,
  base: INavigationUrlBase
): string {
  const text = `${url || ''}`.trim()
  const web = `${base.serverRelativeUrl || ''}`.replace(/\/+$/, '')
  let path: string
  const origin = ORIGIN.exec(text)
  if (origin) {
    if (
      !base.origin ||
      origin[0].toLowerCase() !== base.origin.replace(/\/+$/, '').toLowerCase()
    )
      return text
    path = text.slice(origin[0].length) || '/'
  } else if (SCHEME.test(text)) {
    return text
  } else if (text.charAt(0) === '/') {
    path = text
  } else {
    path = `${web}/${text}`
  }
  const cut = path.search(/[#?]/)
  const head = cut === -1 ? path : path.slice(0, cut)
  const tail = cut === -1 ? '' : path.slice(cut)
  // Segment by segment: decodeURI would leave reserved escapes (%2C, %26, ...) encoded, while
  // SharePoint stores the target's own path, decoded. A segment with a malformed escape is
  // compared as written.
  const decoded = head
    .split('/')
    .map((segment) => {
      try {
        return decodeURIComponent(segment)
      } catch {
        return segment
      }
    })
    .join('/')
  return `${decoded.replace(/\/+$/, '') || '/'}`.toLowerCase() + tail
}
