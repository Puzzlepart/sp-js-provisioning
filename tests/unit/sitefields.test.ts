import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { LogLevel } from '@pnp/logging'
import { ISiteField, SiteFields } from '../../src/handlers/sitefields'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { fieldAttributes } from '../../src/util/fieldxml'
import { captureLogs, levelOf } from '../helpers/logs'

const WEB_ID = '583398dc-780a-4a9c-9156-ad5273c8c65e'
const LIST_A = '11111111-1111-4111-8111-111111111111'
const LIST_B = '22222222-2222-4222-8222-222222222222'
const STORE = '33333333-3333-4333-8333-333333333333'
const TERM_SET = '44444444-4444-4444-8444-444444444444'
const EMPTY = '00000000-0000-0000-0000-000000000000'

/**
 * A fake field collection read: calling it resolves `value` (or rejects with it, when it is an
 * Error), `select` returns itself, and `top` fails the test: `$top` truncates the collection
 * silently (probe SF4).
 *
 * @param value - What a call resolves with
 */
function collection(value: any): any {
  const instance: any = async () => {
    if (value instanceof Error) throw value
    return value
  }
  instance.select = () => instance
  instance.top = () => {
    throw new Error('the fields are read without top()')
  }
  return instance
}

/**
 * A fake web whose `fields` record what is sent: `createFieldAsXml` resolves with the XML's id,
 * `getById(id).update` records the id, the properties and the type.
 *
 * @param own - The web's own columns
 * @param options - The columns available to the web (its own by default), and requests that fail
 */
function createWeb(
  own: ISiteField[],
  options: {
    available?: ISiteField[]
    readError?: Error
    createError?: (xml: string) => Error | undefined
    updateError?: (id: string) => Error | undefined
  } = {}
) {
  const sent: any[] = []
  const fields = collection(options.readError ?? own)
  fields.getById = (id: string) => ({
    update: async (properties: any, type?: string) => {
      sent.push({ update: id, properties, type })
      const error = options.updateError && options.updateError(id)
      if (error) throw error
    }
  })
  fields.createFieldAsXml = async (xml: string) => {
    sent.push({ create: xml })
    const error = options.createError && options.createError(xml)
    if (error) throw error
    return { Id: fieldAttributes(xml).ID.replace(/[{}]/g, '') }
  }
  const web: any = {
    fields,
    availablefields: collection(options.available ?? own)
  }
  return { web, sent }
}

/**
 * Runs the handler on a fake web and returns what it logged.
 *
 * @param web - The fake web
 * @param siteFields - The template's columns
 * @param context - Provisioning context
 */
async function provision(
  web: any,
  siteFields: string[],
  context = new ProvisioningContext()
) {
  context.web = context.web || { Id: WEB_ID, Title: 'Prosjekt A' }
  const { entries } = await captureLogs(() =>
    new SiteFields({}).ProvisionObjects(web, siteFields, context)
  )
  return { entries, context }
}

/** Lines logged at a level. */
const at = (entries: any[], level: LogLevel) =>
  entries.filter((entry) => entry.level === level).map((entry) => entry.message)

const id = 'a1a1a1a1-0000-4000-8000-000000000001'

/** A Text column as a template writes it. */
const text = (extra = '') =>
  `<Field ID="{A1A1A1A1-0000-4000-8000-000000000001}" Name="GtCode" StaticName="GtCode" DisplayName="Kode" Type="Text" Group="Prosjekt"${extra} />`

/** The same column as SharePoint stores it: SourceID, Version, an admin's index. */
const storedText = (extra = '') =>
  `<Field ID="{${id}}" Name="GtCode" StaticName="GtCode" DisplayName="Kode" Type="Text" Group="Prosjekt" SourceID="{${WEB_ID}}" Indexed="TRUE" Version="3"${extra} />`

const textField = (type = 'Text', schemaXml = storedText()): ISiteField => ({
  Id: id,
  InternalName: 'GtCode',
  TypeAsString: type,
  SchemaXml: schemaXml
})

/** A lookup column pointing at a list. */
const lookup = (list: string, extra = '') =>
  `<Field ID="{a1a1a1a1-0000-4000-8000-000000000002}" Name="GtLink" StaticName="GtLink" DisplayName="Lenke" Type="Lookup" List="${list}" ShowField="Title"${extra} />`

const lookupField = (list: string): ISiteField => ({
  Id: 'a1a1a1a1-0000-4000-8000-000000000002',
  InternalName: 'GtLink',
  TypeAsString: 'Lookup',
  SchemaXml: lookup(list, ` SourceID="{${WEB_ID}}" Version="1"`)
})

describe('SiteFields: creating', () => {
  it('creates a missing column with its tokens replaced, and records it as made in this setup', async () => {
    const { web, sent } = createWeb([])

    const { entries, context } = await provision(web, [
      text(' Description="Kode for {sitename}"')
    ])

    assert.equal(sent.length, 1)
    assert.match(sent[0].create, /Description="Kode for Prosjekt A"/)
    assert.equal(context.siteFields.GtCode, id)
    assert.deepEqual(context.createdSiteFields, { GtCode: true })
    assert.equal(levelOf(entries, /scope ended$/), LogLevel.Info)
  })

  it('creates a column whose token cannot be resolved with the token as written, and says which', async () => {
    // Content types and list FieldRefs still find it by name; a later run heals the target.
    const { web, sent } = createWeb([])

    const { entries } = await provision(web, [lookup('{listid:Kilder}')])

    assert.match(sent[0].create, /List="\{listid:Kilder\}"/)
    assert.deepEqual(at(entries, LogLevel.Error).length, 1)
    assert.equal(
      levelOf(
        entries,
        /GtLink \(Lenke\): \{listid:Kilder\} could not be resolved/
      ),
      LogLevel.Error
    )
  })
})

describe('SiteFields: re-applying (the list-field rules, no push)', () => {
  it('writes nothing for an unchanged column', async () => {
    const { web, sent } = createWeb([textField()])

    const { entries } = await provision(web, [text()])

    assert.deepEqual(sent, [])
    assert.equal(
      levelOf(entries, /GtCode \(Kode\) is unchanged/),
      LogLevel.Info
    )
  })

  it('merges a change onto the stored column: one MERGE, what the template leaves out kept, no Version', async () => {
    // SharePoint replaces a site column's schema wholesale (probe SF2): sent as the template has
    // it, the admin's index and the column's Default would be gone.
    const { web, sent } = createWeb([
      textField(
        'Text',
        storedText().replace(' />', '><Default>A</Default></Field>')
      )
    ])

    await provision(web, [text(' MaxLength="100"')])

    assert.equal(sent.length, 1)
    assert.equal(sent[0].update, id)
    assert.equal(sent[0].type, 'SP.Field')
    const schemaXml: string = sent[0].properties.SchemaXml
    assert.match(schemaXml, /MaxLength="100"/)
    assert.match(schemaXml, /Indexed="TRUE"/)
    assert.match(schemaXml, /<Default>A<\/Default>/)
    assert.doesNotMatch(schemaXml, /Version=/)
  })

  it('applies a widening type change and warns about any other, changing nothing', async () => {
    const widening = createWeb([textField()])
    const narrowing = createWeb([
      textField('Note', storedText().replace('Type="Text"', 'Type="Note"'))
    ])

    await provision(widening.web, [
      text().replace('Type="Text"', 'Type="Note"')
    ])
    const { entries } = await provision(narrowing.web, [text()])

    assert.match(widening.sent[0].properties.SchemaXml, /Type="Note"/)
    assert.deepEqual(narrowing.sent, [])
    assert.equal(
      levelOf(entries, /GtCode \(Kode\) is Note, the template says Text/),
      LogLevel.Warning
    )
  })

  it('keeps the list a lookup points at when the template names another, with a warning', async () => {
    const { web, sent } = createWeb([lookupField(`{${LIST_A}}`)])

    const { entries } = await provision(web, [
      lookup(LIST_B, ' Description="Ny"')
    ])

    assert.equal(
      levelOf(entries, /points at another List \(\{11111111/),
      LogLevel.Warning
    )
    const schemaXml: string = sent[0].properties.SchemaXml
    assert.match(schemaXml, new RegExp(`List="\\{${LIST_A}\\}"`))
    assert.match(schemaXml, /Description="Ny"/)
  })

  it('points a lookup an earlier run stored with a literal token at the list the template names', async () => {
    // S1c: an old broken column heals; SharePoint re-targets it with a SchemaXml MERGE (probe SF6).
    const { web, sent } = createWeb([lookupField('{listid:Kilder}')])
    const context = new ProvisioningContext()
    context.lists['Kilder'] = LIST_B

    const { entries } = await provision(
      web,
      [lookup('{listid:Kilder}')],
      context
    )

    assert.match(sent[0].properties.SchemaXml, new RegExp(`List="${LIST_B}"`))
    assert.deepEqual(at(entries, LogLevel.Warning), [])
    assert.deepEqual(at(entries, LogLevel.Error), [])
  })

  it('points a column made in this setup at the target the template names', async () => {
    const { web, sent } = createWeb([lookupField(`{${LIST_A}}`)])
    const context = new ProvisioningContext()
    context.createdSiteFields['GtLink'] = true

    const { entries } = await provision(web, [lookup(LIST_B)], context)

    assert.match(sent[0].properties.SchemaXml, new RegExp(`List="${LIST_B}"`))
    assert.deepEqual(at(entries, LogLevel.Warning), [])
  })

  it('never writes an unresolved token over a stored target, and says which token', async () => {
    const { web, sent } = createWeb([lookupField(`{${LIST_A}}`)])

    const { entries } = await provision(web, [
      lookup('{listid:Omdøpt}', ' Description="Ny"')
    ])

    assert.match(
      sent[0].properties.SchemaXml,
      new RegExp(`List="\\{${LIST_A}\\}"`)
    )
    assert.deepEqual(at(entries, LogLevel.Warning), [])
    assert.equal(
      levelOf(entries, /\{listid:Omdøpt\} could not be resolved/),
      LogLevel.Error
    )
  })
})

describe('SiteFields: taxonomy columns', () => {
  const value = (q: string, text: string) =>
    `<Value xmlns:${q}="http://www.w3.org/2001/XMLSchema" p4:type="${q}:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">${text}</Value>`
  const property = (name: string, text: string, q = 'q1') =>
    `<Property><Name>${name}</Name>${value(q, text)}</Property>`
  /** As the vei hub template writes it: no TextField, the term store as a token, braces. */
  const template = (termSet = `{${TERM_SET.toUpperCase()}}`, extra = '') =>
    `<Field DisplayName="Fag" Description="" Type="TaxonomyFieldTypeMulti" Mult="TRUE" ShowField="Term1033" ID="{8f785c45-ca9f-444c-958a-89f74a1a7143}" StaticName="GtVeiSubject" Name="GtVeiSubject" Group="Prosjekt"${extra}><Customization><ArrayOfProperty>${property(
      'SspId',
      '{sitecollectiontermstoreid}'
    )}${property(
      'TermSetId',
      termSet,
      'q2'
    )}</ArrayOfProperty></Customization></Field>`
  /** As SharePoint stores it (probe SF3): its hidden list, the web, ShowField in the site's language, a TextField. */
  const stored = (sspId = STORE, termSet = TERM_SET) =>
    `<Field Type="TaxonomyFieldTypeMulti" DisplayName="Fag" Description="" Name="GtVeiSubject" StaticName="GtVeiSubject" ID="{8f785c45-ca9f-444c-958a-89f74a1a7143}" ShowField="Term1044" Mult="TRUE" Sortable="FALSE" Group="Prosjekt" SourceID="{${WEB_ID}}" List="{6cbafefd-a7ab-47fa-bbe9-f7f006f3b080}" WebId="${WEB_ID}" Version="1"><Default></Default><Customization><ArrayOfProperty>${property(
      'SspId',
      sspId
    )}<Property><Name>GroupId</Name></Property>${property(
      'TermSetId',
      termSet,
      'q2'
    )}${property(
      'TextField',
      '{e5f24e7a-ed15-4afb-9216-9dbfe3c63ceb}',
      'q3'
    )}</ArrayOfProperty></Customization></Field>`
  const taxonomyField = (schemaXml: string): ISiteField => ({
    Id: '8f785c45-ca9f-444c-958a-89f74a1a7143',
    InternalName: 'GtVeiSubject',
    TypeAsString: 'TaxonomyFieldTypeMulti',
    SchemaXml: schemaXml
  })
  const withStore = () => {
    const context = new ProvisioningContext()
    context.termStoreId = STORE
    return context
  }

  it('writes nothing for an unchanged column: its Customization, hidden list and ShowField stay as SharePoint keeps them', async () => {
    // A SchemaXml MERGE ignores the Customization (probe SF3): sent anyway, it would bump the
    // Version on every import and change nothing.
    const { web, sent } = createWeb([taxonomyField(stored())])

    const { entries } = await provision(web, [template()], withStore())

    assert.deepEqual(sent, [])
    assert.deepEqual(at(entries, LogLevel.Warning), [])
  })

  it('applies an attribute change with the stored Customization, TextField included', async () => {
    const { web, sent } = createWeb([taxonomyField(stored())])

    await provision(web, [template().replace('Fag', 'Fagområde')], withStore())

    assert.equal(sent.length, 1)
    const schemaXml: string = sent[0].properties.SchemaXml
    assert.match(schemaXml, /DisplayName="Fagområde"/)
    assert.match(schemaXml, /ShowField="Term1044"/)
    assert.match(schemaXml, /<Name>TextField<\/Name>/)
    assert.match(
      schemaXml,
      new RegExp(`<Name>SspId</Name><Value[^>]*>${STORE}<`)
    )
  })

  it('sets a term store an earlier run left empty as a TaxonomyField property', async () => {
    // A literal {sitecollectiontermstoreid} is stored as the empty GUID (probe SF3); only the
    // property MERGE sets it.
    const { web, sent } = createWeb([taxonomyField(stored(EMPTY))])

    await provision(web, [template()], withStore())

    assert.deepEqual(sent, [
      {
        update: '8f785c45-ca9f-444c-958a-89f74a1a7143',
        properties: { SspId: STORE },
        type: 'SP.Taxonomy.TaxonomyField'
      }
    ])
  })

  it('sets no term store that is no GUID, and says so, rather than reporting it healed', async () => {
    // A token TokenHelper does not know (the case differs) is no GUID: SharePoint would store the
    // empty GUID again, so nothing is sent and an Error names the value.
    const { web, sent } = createWeb([taxonomyField(stored(EMPTY))])

    const { entries } = await provision(
      web,
      [
        template().replace(
          '{sitecollectiontermstoreid}',
          '{SiteCollectionTermStoreId}'
        )
      ],
      withStore()
    )

    assert.deepEqual(sent, [])
    assert.equal(at(entries, LogLevel.Error).length, 1)
    assert.match(
      at(entries, LogLevel.Error)[0],
      /GtVeiSubject \(Fag\): its SspId \{SiteCollectionTermStoreId\} is not a GUID/
    )
    assert.equal(levelOf(entries, /updated in place/), undefined)
  })

  it('keeps the term set when the template names another, with a warning, unless the column was made in this setup', async () => {
    const other = '55555555-5555-4555-8555-555555555555'
    const kept = createWeb([taxonomyField(stored())])
    const followed = createWeb([taxonomyField(stored())])
    const made = withStore()
    made.createdSiteFields['GtVeiSubject'] = true

    const { entries } = await provision(
      kept.web,
      [template(other)],
      withStore()
    )
    await provision(followed.web, [template(other)], made)

    assert.deepEqual(kept.sent, [])
    assert.equal(
      levelOf(entries, /points at another TermSetId/),
      LogLevel.Warning
    )
    assert.deepEqual(followed.sent, [
      {
        update: '8f785c45-ca9f-444c-958a-89f74a1a7143',
        properties: { TermSetId: other },
        type: 'SP.Taxonomy.TaxonomyField'
      }
    ])
  })

  it('leaves the stored term store alone when the token cannot be resolved, with an error', async () => {
    const { web, sent } = createWeb([taxonomyField(stored())])

    const { entries } = await provision(web, [template()])

    assert.deepEqual(sent, [])
    assert.equal(
      levelOf(entries, /\{sitecollectiontermstoreid\} could not be resolved/),
      LogLevel.Error
    )
  })
})

describe('SiteFields: a column that fails', () => {
  it('logs the column and the operation, and carries on with the others', async () => {
    const { web, sent } = createWeb([textField()], {
      createError: (xml) =>
        xml.includes('GtBad') ? new Error('Ugyldig formel') : undefined,
      updateError: () => new Error('Throttled')
    })

    const { entries } = await provision(web, [
      text(' MaxLength="10"'),
      '<Field Name="GtBroken"',
      '<Field ID="{a1a1a1a1-0000-4000-8000-0000000000bb}" Name="GtBad" DisplayName="Feil" Type="Calculated" />',
      '<Field ID="{a1a1a1a1-0000-4000-8000-0000000000cc}" Name="GtGood" DisplayName="Bra" Type="Text" />'
    ])

    const errors = at(entries, LogLevel.Error)
    assert.equal(errors.length, 3)
    assert.match(
      errors[0],
      /Failed to update site column GtCode \(Kode\): Throttled/
    )
    assert.match(errors[1], /Failed to read a site column's XML, it is skipped/)
    assert.match(
      errors[2],
      /Failed to create site column GtBad \(Feil\): Ugyldig formel/
    )
    assert.match(sent[sent.length - 1].create, /Name="GtGood"/)
    assert.equal(levelOf(entries, /scope ended$/), LogLevel.Info)
  })

  it('creates and changes nothing when the web cannot be read, and names the columns', async () => {
    const { web, sent } = createWeb([], { readError: new Error('Throttled') })

    const { entries } = await provision(web, [text(), lookup(LIST_A)])

    assert.deepEqual(sent, [])
    assert.equal(
      levelOf(
        entries,
        /Failed to read the site columns of the web, so none is created or changed \(GtCode, GtLink\): Throttled/
      ),
      LogLevel.Error
    )
  })
})

describe('SiteFields: identity', () => {
  it('updates a column found by its internal name under another id, keeping its id, with a warning', async () => {
    // SharePoint ignores the ID in a SchemaXml MERGE (probe SF5); the warning is the engine's.
    const other = 'b2b2b2b2-0000-4000-8000-000000000001'
    const { web, sent } = createWeb([
      { ...textField(), Id: other, SchemaXml: storedText().replace(id, other) }
    ])

    const { entries } = await provision(web, [text(' MaxLength="10"')])

    assert.equal(levelOf(entries, /has the id b2b2b2b2/), LogLevel.Warning)
    assert.equal(sent[0].update, other)
    assert.match(
      sent[0].properties.SchemaXml,
      new RegExp(`ID="\\{${other}\\}"`)
    )
  })

  it('does not create a column whose id the web has under another name, with an error', async () => {
    // SharePoint would refuse it with "The local device name is already in use" (probe SF5).
    const { web, sent } = createWeb([
      { ...textField(), InternalName: 'GtOther' }
    ])

    const { entries } = await provision(web, [text()])

    assert.deepEqual(sent, [])
    assert.equal(
      levelOf(
        entries,
        /GtCode \(Kode\) is not created: the web has its id .* as the column GtOther/
      ),
      LogLevel.Error
    )
  })

  it('neither creates nor changes a column that only a parent web has, with a warning', async () => {
    // On a subweb `web.fields` lists only the web's own columns (probe SF4).
    const { web, sent } = createWeb([], {
      available: [
        // The parent's column by the same name has another id: the name decides.
        { Id: 'd4d4d4d4-0000-4000-8000-000000000001', InternalName: 'GtCode' },
        { Id: 'c3c3c3c3-0000-4000-8000-000000000001', InternalName: 'GtRoot' }
      ]
    })

    const { entries } = await provision(web, [
      text(),
      '<Field ID="{c3c3c3c3-0000-4000-8000-000000000001}" Name="GtMine" DisplayName="Min" Type="Text" />'
    ])

    assert.deepEqual(sent, [])
    assert.equal(
      levelOf(entries, /GtCode \(Kode\) is not created: a parent web has it/),
      LogLevel.Warning
    )
    assert.equal(
      levelOf(
        entries,
        /GtMine \(Min\) is not created: a parent web has its id .* as the column GtRoot/
      ),
      LogLevel.Error
    )
  })
})
