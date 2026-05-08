/*
  SharePoint Content Type deterministic ID harness

  Usage:
  1) Open a modern page in the target site while logged in.
  2) Open browser devtools console.
  3) Paste this file.
  4) Set `template.siteUrl`, `template.libraries`, and `template.contentTypes`.
  5) Run: await runContentTypeIdHarness(template)

  The harness:
   - Creates each content type at site level with a deterministic ID.
   - Enables content type management on every library in `template.libraries`
     (defaults to the Norwegian default: "Delte dokumenter").
   - Attaches each site CT to every listed library.
   - Verifies both the site CT id and the list CT inheritance.
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
  const original = String(contentTypeId)
  const upper = original.toUpperCase()
  if (/00[0-9A-F]{32}$/.test(upper)) return original.slice(0, -34)
  if (original.length > 4) return original.slice(0, -2)
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

async function fetchExistingContentTypeIds(siteUrl) {
  const url = `${siteUrl.replace(/\/$/, '')}/_api/web/contenttypes?$select=StringId&$top=5000`
  const res = await fetch(url, { headers: { Accept: 'application/json;odata=nometadata' } })
  if (!res.ok) throw new Error(`Failed to fetch existing content types: ${res.status} ${res.statusText}`)
  const data = await res.json()
  return new Set((data.value || []).map((ct) => String(ct.StringId).toUpperCase()))
}

async function getOrCreateContentType(ctx, web, configuredCt, existingIds) {
  const id = configuredCt.ID
  const cts = web.get_contentTypes()

  if (existingIds.has(id.toUpperCase())) {
    return { ct: cts.getById(id), created: false }
  }

  const info = new SP.ContentTypeCreationInformation()
  info.set_name(configuredCt.Name)
  info.set_id(id)
  if (configuredCt.Description) info.set_description(configuredCt.Description)
  if (configuredCt.Group) info.set_group(configuredCt.Group)
  // NB: do NOT call info.set_parentContentType() when info.set_id() is used —
  // SharePoint rejects the combination with
  //   "parameters.Id, parameters.ParentContentType cannot be used together".
  // The parent is inferred from the explicit ID's structure.

  const createdCt = cts.add(info)
  createdCt.update(true)
  await executeQuery(ctx)
  existingIds.add(id.toUpperCase())
  return { ct: createdCt, created: true }
}

async function ensureLibraryAcceptsContentTypes(ctx, list) {
  ctx.load(list, 'ContentTypesEnabled', 'Title')
  await executeQuery(ctx)
  if (!list.get_contentTypesEnabled()) {
    list.set_contentTypesEnabled(true)
    list.update()
    await executeQuery(ctx)
  }
}

function resolveLibraryServerRelativeUrl(siteUrl, library) {
  if (library.startsWith('/')) return library
  const sitePath = new URL(siteUrl).pathname.replace(/\/$/, '')
  return `${sitePath}/${library}`
}

async function resolveLibrary(ctx, web, siteUrl, library) {
  const serverRelativeUrl = resolveLibraryServerRelativeUrl(siteUrl, library)
  const list = web.getList(serverRelativeUrl)
  ctx.load(list, 'Title', 'ContentTypesEnabled')
  await executeQuery(ctx)
  return { list, title: list.get_title(), serverRelativeUrl }
}

async function attachContentTypeToList(ctx, list, siteCt, siteCtId) {
  const listCts = list.get_contentTypes()
  ctx.load(listCts, 'Include(Id, Name, Parent)')
  await executeQuery(ctx)

  const upperSiteId = siteCtId.toUpperCase()
  const enumerator = listCts.getEnumerator()
  while (enumerator.moveNext()) {
    const existing = enumerator.get_current()
    const existingId = existing.get_id().toString().toUpperCase()
    if (existingId.startsWith(upperSiteId)) {
      return { listCtId: existing.get_id().toString(), attached: false }
    }
  }

  const added = listCts.addExistingContentType(siteCt)
  ctx.load(added, 'Id', 'Name')
  await executeQuery(ctx)
  return { listCtId: added.get_id().toString(), attached: true }
}

async function runContentTypeIdHarness(template) {
  if (!template || !template.siteUrl || !Array.isArray(template.contentTypes)) {
    throw new Error('Expected template { siteUrl, libraries?, contentTypes[] }')
  }

  const libraries = Array.isArray(template.libraries) && template.libraries.length > 0
    ? template.libraries
    : ['Delte dokumenter']

  const existingIds = await fetchExistingContentTypeIds(template.siteUrl)

  await ensureJsom(template.siteUrl)
  const ctx = new SP.ClientContext(template.siteUrl)
  const web = ctx.get_web()

  const lists = {}
  for (const library of libraries) {
    const { list, title, serverRelativeUrl } = await resolveLibrary(ctx, web, template.siteUrl, library)
    await ensureLibraryAcceptsContentTypes(ctx, list)
    lists[library] = { list, title, serverRelativeUrl }
    console.log(`Resolved library "${library}" → "${title}" (${serverRelativeUrl})`)
  }

  const siteResults = []
  const libraryResults = []
  for (const configuredCt of template.contentTypes.sort((a, b) => a.ID.localeCompare(b.ID))) {
    const { ct, created } = await getOrCreateContentType(ctx, web, configuredCt, existingIds)

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
    siteResults.push({
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

    for (const library of libraries) {
      const { list, title } = lists[library]
      const { listCtId, attached } = await attachContentTypeToList(
        ctx,
        list,
        persisted,
        configuredCt.ID
      )
      const inherits = listCtId.toUpperCase().startsWith(configuredCt.ID.toUpperCase())
      libraryResults.push({
        library: title,
        contentType: configuredCt.Name,
        siteCtId: configuredCt.ID,
        listCtId,
        attached,
        inherits
      })
      if (!inherits) {
        throw new Error(
          `List CT for ${configuredCt.Name} on ${title} does not inherit from ${configuredCt.ID} (got ${listCtId})`
        )
      }
    }
  }

  console.log('Site content types:')
  console.table(siteResults)
  console.log('Library attachments:')
  console.table(libraryResults)
  console.log('✅ Deterministic content type ID verification + library attachment passed')
  return { siteResults, libraryResults }
}

// Example payload. Update before running.
// `var` (not `const`) so re-pasting in the DevTools console doesn't throw
// "Identifier 'template' has already been declared".
var template = {
  siteUrl: 'https://pzlokms.sharepoint.com/sites/test-sp-js-prov',
  libraries: ['Delte dokumenter'],
  contentTypes: [
    {
      Name: 'Møtereferat',
      ID: '0x010100A1B2C3D4E5F67890123456789012ABCD',
      Group: 'Crayon dokumenter',
      Description: 'Referat fra møter'
    },
    {
      Name: 'Prosjektplan',
      ID: '0x010100B2C3D4E5F67890123456789012ABCDEF',
      Group: 'Crayon dokumenter',
      Description: 'Planleggingsdokument for prosjekter'
    },
    {
      Name: 'Tilbud',
      ID: '0x010100C3D4E5F67890123456789012ABCDEF01',
      Group: 'Crayon dokumenter',
      Description: 'Tilbudsdokument til kunde'
    },
    {
      Name: 'Kontrakt',
      ID: '0x010100D4E5F67890123456789012ABCDEF0123',
      Group: 'Crayon dokumenter',
      Description: 'Kontraktsdokument'
    },
    {
      Name: 'Rammeavtale',
      ID: '0x010100D4E5F67890123456789012ABCDEF012301',
      Group: 'Crayon dokumenter',
      Description: 'Rammeavtale (barn av Kontrakt for å teste ID-arv)'
    },
    {
      Name: 'Faktura',
      ID: '0x010100E5F67890123456789012ABCDEF012345',
      Group: 'Crayon dokumenter',
      Description: 'Fakturadokument'
    }
  ]
}
