/*
  SharePoint Content Type deterministic ID harness

  Usage:
  1) Open a modern page in the target site while logged in.
  2) Open browser devtools console.
  3) Paste this file.
  4) Set `template.siteUrl` and `template.contentTypes`.
  5) Run: await runContentTypeIdHarness(template)
*/

async function loadScript(url) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script')
    s.src = url
    s.onload = resolve
    s.onerror = reject
    document.head.appendChild(s)
  })
}

function getParentContentTypeId(contentTypeId) {
  const id = String(contentTypeId).toUpperCase()
  if (/00[0-9A-F]{32}$/.test(id)) return id.slice(0, -34)
  if (id.length > 4) return id.slice(0, -2)
  return '0x01'
}

async function ensureJsom(siteUrl) {
  const origin = new URL(siteUrl).origin
  const base = `${origin}/_layouts/15/`
  await loadScript(base + 'init.js')
  await loadScript(base + 'MicrosoftAjax.js')
  await loadScript(base + 'sp.runtime.js')
  await loadScript(base + 'sp.js')
}

function executeQuery(ctx) {
  return new Promise((resolve, reject) => {
    ctx.executeQueryAsync(resolve, (sender, args) => reject(new Error(args.get_message())))
  })
}

async function getOrCreateContentType(ctx, web, configuredCt) {
  const id = configuredCt.ID
  const cts = web.get_contentTypes()
  const existing = cts.getById(id)
  ctx.load(existing)
  try {
    await executeQuery(ctx)
    return { ct: existing, created: false }
  } catch {
    const info = new SP.ContentTypeCreationInformation()
    info.set_name(configuredCt.Name)
    info.set_id(id)
    if (configuredCt.Description) info.set_description(configuredCt.Description)
    if (configuredCt.Group) info.set_group(configuredCt.Group)

    const parentId = getParentContentTypeId(id)
    info.set_parentContentType(cts.getById(parentId))

    const createdCt = cts.add(info)
    createdCt.update(true)
    await executeQuery(ctx)
    return { ct: createdCt, created: true }
  }
}

async function runContentTypeIdHarness(template) {
  if (!template || !template.siteUrl || !Array.isArray(template.contentTypes)) {
    throw new Error('Expected template { siteUrl, contentTypes[] }')
  }

  await ensureJsom(template.siteUrl)
  const ctx = new SP.ClientContext(template.siteUrl)
  const web = ctx.get_web()

  const results = []
  for (const configuredCt of template.contentTypes.sort((a, b) => a.ID.localeCompare(b.ID))) {
    const { ct, created } = await getOrCreateContentType(ctx, web, configuredCt)

    ct.set_name(configuredCt.Name)
    if (configuredCt.Description) ct.set_description(configuredCt.Description)
    if (configuredCt.Group) ct.set_group(configuredCt.Group)
    ct.update(true)
    await executeQuery(ctx)

    const persisted = web.get_contentTypes().getById(configuredCt.ID)
    ctx.load(persisted)
    await executeQuery(ctx)

    const actualId = persisted.get_id().toString()
    const idMatches = actualId.toUpperCase() === configuredCt.ID.toUpperCase()
    results.push({
      name: configuredCt.Name,
      configuredId: configuredCt.ID,
      actualId,
      created,
      idMatches
    })

    if (!idMatches) {
      throw new Error(
        `ID mismatch for ${configuredCt.Name}: configured=${configuredCt.ID} actual=${actualId}`
      )
    }
  }

  console.table(results)
  console.log('✅ Deterministic content type ID verification passed')
  return results
}

// Example payload. Update before running.
const template = {
  siteUrl: 'https://<tenant>.sharepoint.com/sites/<fresh-site>',
  contentTypes: [
    {
      Name: 'Project Proposal',
      ID: '0x0100AAE66CFF1A3F488D843FF7CF96E41DD901',
      Group: 'Custom Content Types',
      Description: 'Provisioned deterministically with explicit ID'
    }
  ]
}
