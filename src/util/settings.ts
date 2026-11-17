import { compactMarkup } from './markup'

/**
 * The settings whose value differs from what a list, view or web has, to write only those: an
 * unchanged re-run then writes nothing. A setting differs when the object lacks it (it is written,
 * as before), when a `ViewQuery` is laid out otherwise than SharePoint stores it (`compactMarkup`),
 * or when the value is neither the same nor the same text (an object compared as JSON).
 *
 * @param settings - The template's settings
 * @param stored - The object as REST returns it without `$select` (every property)
 */
export function differingSettings(
  settings: Record<string, any>,
  stored: Record<string, any>
): Record<string, any> {
  const changed: Record<string, any> = {}
  for (const key of Object.keys(settings)) {
    const want = settings[key]
    const have = stored ? stored[key] : undefined
    let same: boolean
    if (have === undefined) same = false
    else if (key === 'ViewQuery')
      same = compactMarkup(want) === compactMarkup(have)
    else if (want && typeof want === 'object')
      same = JSON.stringify(want) === JSON.stringify(have)
    else same = want === have || String(want) === String(have)
    if (!same) changed[key] = want
  }
  return changed
}
