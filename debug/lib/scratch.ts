import { IWeb, Web } from '@pnp/sp/presets/all'
import { connect, ISiteConnection } from './connect'
import { env } from './env'
import { describeError } from './redact'

/**
 * Scratch webs: a subweb of the reserved site per test suite (or per run), created with unique
 * permissions so it has associated groups of its own, and deleted afterwards. A template run
 * cannot be undone; a scratch web makes every first and second apply real and keeps the reserved
 * site itself untouched.
 *
 * A web created with unique permissions has no associated groups; `createDefaultAssociatedGroups`
 * creates them (`<title>-eiere`, `-medlemmer`, `-besøkende` on a Norwegian site). On a site with
 * custom script blocked (NoScript) it answers 403 although the groups are created and associated
 * (observed 2026-10-21 on a group-connected team site), so a 403 is accepted when all three groups
 * are there afterwards. Site groups belong to the site collection and outlive the web, so they are
 * deleted with it: those whose title starts with the web's title (not by the stamp alone, which two
 * suites started in the same second share).
 */

/** Prefix of every scratch web's URL segment. */
export const SCRATCH_PREFIX = 'spjs-itest-'

/**
 * A UTC stamp, `yyMMddHHmmss`, so a scratch web's age is readable from its name.
 *
 * @param date - The time
 */
export function stampOf(date: Date): string {
  return date.toISOString().replace(/[-:T]/g, '').slice(2, 14)
}

/**
 * The time a stamp stands for, or `undefined` when it is not one.
 *
 * @param stamp - `yyMMddHHmmss`
 */
export function dateOfStamp(stamp: string): Date | undefined {
  const match = /^(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})$/.exec(stamp)
  if (!match) return undefined
  const [, yy, MM, dd, HH, mm, ss] = match.map(Number)
  const date = new Date(Date.UTC(2000 + yy, MM - 1, dd, HH, mm, ss))
  // `Date.UTC` rolls over (month 13 is January): only a real time is a stamp.
  return stampOf(date) === stamp ? date : undefined
}

/**
 * The title and URL segment of a scratch web for a suite.
 *
 * @param suite - The suite, e.g. `lists`
 * @param stamp - The stamp
 */
export function scratchName(
  suite: string,
  stamp: string
): { title: string; url: string } {
  const slug = suite
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
  return {
    title: `SPJS Itest ${suite} ${stamp}`,
    url: `${SCRATCH_PREFIX}${slug}-${stamp}`
  }
}

/**
 * The stamp in a scratch web's URL segment or a scratch group's title, or `undefined` when the
 * name is not a scratch name.
 *
 * @param name - A URL segment (`spjs-itest-lists-261021151015`) or a title
 */
export function stampIn(name: string): string | undefined {
  const match = /(?:^|[\s-])(\d{12})(?:$|[\s-])/.exec(name)
  return match && dateOfStamp(match[1]) ? match[1] : undefined
}

/**
 * A scratch web and the connection to it.
 */
export interface IScratchWeb {
  title: string
  url: string
  stamp: string
  conn: ISiteConnection
}

/**
 * Creates a scratch web below the reserved site, in its language, with unique permissions.
 *
 * @param parent - The reserved site
 * @param suite - The suite the web is for
 * @param now - The time for the stamp
 * @param owner - Login name for the owners group: the signed-in user by default (cookies);
 * app-only, pass `INTEGRATION_USER`
 */
export async function createScratchWeb(
  parent: ISiteConnection,
  suite: string,
  now: Date = new Date(),
  owner?: string
): Promise<IScratchWeb> {
  const stamp = stampOf(now)
  const { title, url } = scratchName(suite, stamp)
  const ownerLogin = owner ?? (await defaultOwner(parent))
  try {
    // Inside the try: a webs.add that fails or times out can still have made the web.
    await parent.web.webs.add(
      title,
      url,
      'Created by the sp-js-provisioning integration tests; deleted afterwards.',
      'STS#3',
      parent.webInfo.Language,
      false
    )
    const conn = await connect(`${parent.url}/${url}`, parent.mode)
    await createAssociatedGroups(conn, title, ownerLogin)
    return { title, url, stamp, conn }
  } catch (error) {
    // Do not leave a half-made web (or the groups it already got) behind.
    await removeScratchWeb(parent, { title, url }).catch(() => undefined)
    throw error
  }
}

/**
 * The login name for a scratch web's owners group: the signed-in user (cookies), or
 * `INTEGRATION_USER` app-only, where there is no signed-in user.
 *
 * @param parent - The reserved site
 */
async function defaultOwner(parent: ISiteConnection): Promise<string> {
  if (parent.mode === 'cookies')
    return (
      await parent.web.currentUser.select('LoginName')<{ LoginName: string }>()
    ).LoginName
  if (!env.user)
    throw new Error(
      'Signed in as an app: set INTEGRATION_USER, the user who owns the scratch webs'
    )
  return (await parent.web.ensureUser(env.user)).LoginName
}

/**
 * Creates the web's associated owner, member and visitor groups, named after its title, and
 * checks they are there.
 *
 * @param conn - The scratch web
 * @param title - The web's title, the groups' name seed
 * @param owner - Login name of the user to put in the owners group
 */
async function createAssociatedGroups(
  conn: ISiteConnection,
  title: string,
  owner: string
): Promise<void> {
  let refused: any
  try {
    await conn.web.createDefaultAssociatedGroups(title, owner)
  } catch (error) {
    if (!error || error.status !== 403) throw error
    refused = error
  }
  const groups = await Promise.all(
    [
      conn.web.associatedOwnerGroup,
      conn.web.associatedMemberGroup,
      conn.web.associatedVisitorGroup
    ].map((group) =>
      group
        .select('Title')<{ Title?: string }>()
        .then((info) => info.Title)
        .catch(() => undefined)
    )
  )
  if (groups.some((group) => !group || !group.includes(title)))
    throw new Error(
      `The scratch web ${conn.webInfo.ServerRelativeUrl} has no associated groups of its own (${groups.join(
        ', '
      )})${refused ? `: ${describeError(refused)}` : ''}`
    )
}

/**
 * The site groups of a scratch web: its associated groups (`<title>-eiere` and so on) and the
 * fixture groups named after it through `{sitename}` (`<title> Readers`). Matched by the web's
 * title, which holds the suite and the stamp: two suites started in the same second share a stamp,
 * not a title.
 *
 * @param parent - The reserved site
 * @param title - The scratch web's title
 */
export async function groupsOfScratchWeb(
  parent: ISiteConnection,
  title: string
): Promise<Array<{ Id: number; Title: string }>> {
  if (!stampIn(title) || !title.startsWith('SPJS Itest '))
    throw new Error(`Not a scratch web title: ${title}`)
  const groups = await parent.web.siteGroups
    .filter(`substringof('${title.replace(/'/g, "''")}',Title)`)
    .select('Id', 'Title')<Array<{ Id: number; Title: string }>>()
  return groups.filter((group) => group.Title.startsWith(title))
}

/**
 * Deletes a web after its subwebs, deepest first: SharePoint refuses to delete a web that has one
 * (a test's nested web, left behind when the test was interrupted).
 *
 * @param web - The web
 */
export async function deleteWebTree(web: IWeb): Promise<void> {
  const children = await Web(web)
    .webs.select('Url')<Array<{ Url: string }>>()
  for (const child of children) await deleteWebTree(Web([web, child.Url]))
  await Web(web).delete()
}

/**
 * Deletes a scratch web by its name, and its site groups, whatever is left of them: for a web
 * whose creation failed half-way or never confirmed (a request that hung), where there is no
 * connection to it. A web that is not there is no error.
 *
 * @param parent - The reserved site
 * @param scratch - The scratch web's title and URL segment
 */
export async function removeScratchWeb(
  parent: ISiteConnection,
  scratch: Pick<IScratchWeb, 'title' | 'url'>
): Promise<{ web: boolean; groups: string[] }> {
  if (
    !scratch.url.startsWith(SCRATCH_PREFIX) ||
    !stampIn(scratch.url) ||
    stampIn(scratch.url) !== stampIn(scratch.title)
  )
    throw new Error(`Not a scratch web: ${scratch.url} (${scratch.title})`)
  const groups = await groupsOfScratchWeb(parent, scratch.title)
  // Looked up among the reserved site's webs, not by a request to its URL: that answers 404 for a
  // web that is not there, and for other reasons too.
  const expected = `${parent.webInfo.ServerRelativeUrl}/${scratch.url}`.toLowerCase()
  const found = (
    await parent.web.webs.select('ServerRelativeUrl')<
      Array<{ ServerRelativeUrl: string }>
    >()
  ).find((web) => web.ServerRelativeUrl.toLowerCase() === expected)
  if (found)
    await deleteWebTree(
      Web([
        parent.web,
        `${new URL(parent.url).origin}${found.ServerRelativeUrl}`
      ])
    )
  for (const group of groups) await parent.web.siteGroups.removeById(group.Id)
  return { web: !!found, groups: groups.map((group) => group.Title) }
}

/**
 * Deletes a scratch web and its site groups.
 *
 * @param parent - The reserved site
 * @param scratch - The scratch web
 */
export async function deleteScratchWeb(
  parent: ISiteConnection,
  scratch: Pick<IScratchWeb, 'title' | 'conn'>
): Promise<{ groups: string[] }> {
  const groups = await groupsOfScratchWeb(parent, scratch.title)
  await deleteWebTree(scratch.conn.web)
  for (const group of groups) await parent.web.siteGroups.removeById(group.Id)
  return { groups: groups.map((group) => group.Title) }
}

/**
 * A scratch web found below the reserved site.
 */
export interface IScratchWebInfo {
  title: string
  serverRelativeUrl: string
  stamp: string
  created: Date
}

/**
 * The scratch webs below the reserved site, oldest first.
 *
 * @param parent - The reserved site
 */
export async function listScratchWebs(
  parent: ISiteConnection
): Promise<IScratchWebInfo[]> {
  const webs = await parent.web.webs.select('Title', 'ServerRelativeUrl')<
    Array<{ Title: string; ServerRelativeUrl: string }>
  >()
  return webs
    .map((web) => {
      const segment = web.ServerRelativeUrl.split('/').pop() || ''
      const stamp = segment.startsWith(SCRATCH_PREFIX)
        ? stampIn(segment)
        : undefined
      return stamp
        ? {
            title: web.Title,
            serverRelativeUrl: web.ServerRelativeUrl,
            stamp,
            created: dateOfStamp(stamp)
          }
        : undefined
    })
    .filter(Boolean)
    .sort((a, b) => a.created.getTime() - b.created.getTime())
}
