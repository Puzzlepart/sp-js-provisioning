import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import {
  declaredType,
  fieldAttributes,
  mergeFieldXml,
  sameFieldXml,
  setFieldAttributes,
  targetChanged,
  typeChange
} from '../../src/util/fieldxml'

const stored =
  '<Field Type="Text" DisplayName="Kode" Name="GtCode" StaticName="GtCode" ID="{a1a1a1a1-0000-4000-8000-000000000001}" MaxLength="255" Indexed="TRUE" SourceID="{583398dc-780a-4a9c-9156-ad5273c8c65e}" ColName="nvarchar7" RowOrdinal="0" Version="3"/>'

describe('fieldxml: merging a template onto a stored list field', () => {
  it("lets the template's attributes win and keeps identity, storage and what it does not mention", () => {
    const template =
      '<Field Type="Text" DisplayName="Prosjektkode" Name="GtCodeRenamed" ID="{A1A1A1A1-0000-4000-8000-000000000001}" MaxLength="100" Description="Koden" />'
    const merged = fieldAttributes(mergeFieldXml(stored, template))

    assert.deepEqual(merged, {
      Type: 'Text',
      DisplayName: 'Prosjektkode',
      MaxLength: '100',
      Description: 'Koden',
      // An admin's index survives; identity and storage stay SharePoint's.
      Indexed: 'TRUE',
      Name: 'GtCode',
      StaticName: 'GtCode',
      ID: '{a1a1a1a1-0000-4000-8000-000000000001}',
      SourceID: '{583398dc-780a-4a9c-9156-ad5273c8c65e}',
      ColName: 'nvarchar7',
      RowOrdinal: '0'
    })
    // Version is dropped: the update is not tied to the version read.
    assert.equal(merged.Version, undefined)
  })

  it('replaces a child element the template declares and keeps the others', () => {
    const choice =
      '<Field Type="Choice" Name="GtStatus" ID="{a1a1a1a1-0000-4000-8000-000000000002}" ColName="nvarchar8"><Default>A</Default><CHOICES><CHOICE>A</CHOICE><CHOICE>B</CHOICE><CHOICE>Lagt til av admin</CHOICE></CHOICES></Field>'
    const merged = mergeFieldXml(
      choice,
      '<Field Type="Choice" Name="GtStatus" ID="{a1a1a1a1-0000-4000-8000-000000000002}"><CHOICES><CHOICE>A</CHOICE><CHOICE>B</CHOICE><CHOICE>C</CHOICE></CHOICES></Field>'
    )

    assert.match(merged, /<Default>A<\/Default>/)
    assert.match(
      merged,
      /<CHOICES><CHOICE>A<\/CHOICE><CHOICE>B<\/CHOICE><CHOICE>C<\/CHOICE><\/CHOICES>/
    )
    assert.doesNotMatch(merged, /Lagt til av admin/)
  })

  it("merges a taxonomy field's Customization by property, keeping the stored TextField", () => {
    // Replacing Customization wholesale with a template that has no TextField would orphan the
    // hidden note field SharePoint made for the column.
    const property = (name: string, value: string) =>
      `<Property><Name>${name}</Name><Value>${value}</Value></Property>`
    const taxonomy = (properties: string) =>
      `<Field Type="TaxonomyFieldType" Name="GtProjectPhase" ID="{325543a5-815d-485d-a9a5-e0773ad762e9}"><Customization><ArrayOfProperty>${properties}</ArrayOfProperty></Customization></Field>`
    const merged = mergeFieldXml(
      taxonomy(
        property('SspId', 'store') +
          property('TermSetId', 'old-set') +
          property('TextField', 'note-field')
      ),
      taxonomy(property('SspId', 'store') + property('TermSetId', 'new-set'))
    )

    assert.match(merged, /<Name>TermSetId<\/Name><Value>new-set<\/Value>/)
    assert.match(merged, /<Name>TextField<\/Name><Value>note-field<\/Value>/)
    assert.equal((merged.match(/<Property>/g) || []).length, 3)
  })

  it('keeps the stored value of the attributes asked for, and leaves out one it has none of', () => {
    const lookup =
      '<Field Type="Lookup" Name="GtLink" ID="{a1a1a1a1-0000-4000-8000-000000000003}" List="{11111111-1111-4111-8111-111111111111}" ShowField="Title" />'
    const merged = fieldAttributes(
      mergeFieldXml(
        lookup,
        '<Field Type="Lookup" Name="GtLink" ID="{a1a1a1a1-0000-4000-8000-000000000003}" List="{22222222-2222-4222-8222-222222222222}" WebId="{33333333-3333-4333-8333-333333333333}" ShowField="Kode" />',
        ['List', 'WebId']
      )
    )

    assert.equal(merged.List, '{11111111-1111-4111-8111-111111111111}')
    assert.equal(merged.WebId, undefined)
    assert.equal(merged.ShowField, 'Kode')
  })
})

describe('fieldxml: comparing', () => {
  it('ignores attribute order, Version, GUID braces and case, and boolean case', () => {
    assert.equal(
      sameFieldXml(
        stored,
        '<Field ColName="nvarchar7" RowOrdinal="0" Version="9" SourceID="583398DC-780A-4A9C-9156-AD5273C8C65E" ID="a1a1a1a1-0000-4000-8000-000000000001" Indexed="true" MaxLength="255" StaticName="GtCode" Name="GtCode" DisplayName="Kode" Type="Text" />'
      ),
      true
    )
    assert.equal(
      sameFieldXml(
        stored,
        stored.replace('MaxLength="255"', 'MaxLength="100"')
      ),
      false
    )
  })

  it('sees an unchanged merge as the same, so a second apply writes nothing', () => {
    const template =
      '<Field Type="Text" DisplayName="Kode" Name="GtCode" ID="{A1A1A1A1-0000-4000-8000-000000000001}" MaxLength="255" />'
    assert.equal(sameFieldXml(mergeFieldXml(stored, template), stored), true)
  })
})

describe('fieldxml: type changes', () => {
  it('reads Mult on a Lookup, User or taxonomy field as the multi-value type', () => {
    assert.equal(declaredType({ Type: 'Lookup', Mult: 'TRUE' }), 'LookupMulti')
    assert.equal(declaredType({ Type: 'User', Mult: 'true' }), 'UserMulti')
    assert.equal(
      declaredType({ Type: 'TaxonomyFieldType', Mult: 'TRUE' }),
      'TaxonomyFieldTypeMulti'
    )
    assert.equal(
      declaredType({ Type: 'LookupMulti', Mult: 'TRUE' }),
      'LookupMulti'
    )
    assert.equal(declaredType({ Type: 'Note' }), 'Note')
  })

  it('applies widening in place and nothing else', () => {
    assert.equal(typeChange('Text', 'Text'), 'same')
    assert.equal(typeChange('Text', 'Note'), 'widening')
    assert.equal(typeChange('Lookup', 'LookupMulti'), 'widening')
    assert.equal(typeChange('User', 'UserMulti'), 'widening')
    // Multi to single kept one value of two in the probe.
    assert.equal(typeChange('LookupMulti', 'Lookup'), 'other')
    assert.equal(typeChange('Note', 'Text'), 'other')
    assert.equal(typeChange('Text', 'Number'), 'other')
  })

  it("sees a lookup's new target list, whatever the braces and case", () => {
    const lookup = (list: string) =>
      `<Field Type="Lookup" Name="GtLink" ID="{a1a1a1a1-0000-4000-8000-000000000003}" List="${list}" />`
    assert.equal(
      targetChanged(
        lookup('{11111111-1111-4111-8111-111111111111}'),
        lookup('11111111-1111-4111-8111-111111111111'.toUpperCase())
      ),
      false
    )
    assert.equal(
      targetChanged(
        lookup('{11111111-1111-4111-8111-111111111111}'),
        lookup('{22222222-2222-4222-8222-222222222222}')
      ),
      true
    )
    // An unresolved token is not the stored list either.
    assert.equal(
      targetChanged(
        lookup('{11111111-1111-4111-8111-111111111111}'),
        lookup('{listid:Kilder}')
      ),
      true
    )
  })
})

describe('fieldxml: a taxonomy field pointed at another term set', () => {
  const taxonomy = (termSet: string, textField = '') =>
    `<Field Type="TaxonomyFieldType" Name="GtProjectPhase" ID="{325543a5-815d-485d-a9a5-e0773ad762e9}"><Customization><ArrayOfProperty><Property><Name>SspId</Name><Value>store</Value></Property><Property><Name>TermSetId</Name><Value>${termSet}</Value></Property>${textField}</ArrayOfProperty></Customization></Field>`
  const note =
    '<Property><Name>TextField</Name><Value>note-field</Value></Property>'

  it('is a new target, whatever the case and braces of the id', () => {
    assert.equal(
      targetChanged(
        taxonomy('{AAAAAAAA-0000-4000-8000-000000000001}', note),
        taxonomy('aaaaaaaa-0000-4000-8000-000000000001')
      ),
      false
    )
    assert.equal(
      targetChanged(
        taxonomy('aaaaaaaa-0000-4000-8000-000000000001', note),
        taxonomy('bbbbbbbb-0000-4000-8000-000000000002')
      ),
      true
    )
  })

  it('sees the same term set written with braces as unchanged, so a re-import writes nothing', () => {
    // The hub templates write TermSetId with braces; SharePoint stores it without.
    const stored = taxonomy('aaaaaaaa-0000-4000-8000-000000000001', note)
    assert.equal(
      sameFieldXml(
        mergeFieldXml(
          stored,
          taxonomy('{AAAAAAAA-0000-4000-8000-000000000001}')
        ),
        stored
      ),
      true
    )
  })

  it('keeps the stored term set when asked to, and the note field always', () => {
    const merged = mergeFieldXml(
      taxonomy('aaaaaaaa-0000-4000-8000-000000000001', note),
      taxonomy('bbbbbbbb-0000-4000-8000-000000000002'),
      ['TermSetId', 'SspId']
    )
    assert.match(
      merged,
      /<Name>TermSetId<\/Name><Value>aaaaaaaa-0000-4000-8000-000000000001</
    )
    assert.match(merged, /<Name>TextField<\/Name><Value>note-field</)
  })
})

describe('fieldxml: characters SharePoint escapes', () => {
  // An admin's column formatting (`&&`, `<`) or a description with `&` and line breaks: the
  // parser decodes them, and xml-js wrote them back raw.
  const storedWithEntities =
    '<Field Type="Text" Name="GtCode" ID="{a1a1a1a1-0000-4000-8000-000000000001}" ColName="nvarchar7" Description="Plan &amp; tiltak&#xD;&#xA;linje 2" CustomFormatter="{&quot;operator&quot;:&quot;&amp;&amp;&quot;,&quot;operands&quot;:[&quot;@currentField &lt; 5&quot;]}" />'

  it('writes them back escaped, so an unchanged field is the same and the XML stays well-formed', () => {
    const template =
      '<Field Type="Text" Name="GtCode" ID="{a1a1a1a1-0000-4000-8000-000000000001}" />'
    const merged = mergeFieldXml(storedWithEntities, template)

    assert.equal(sameFieldXml(merged, storedWithEntities), true)
    assert.match(merged, /Plan &amp; tiltak&#xD;&#xA;linje 2/)
    assert.match(merged, /&amp;&amp;/)
    assert.match(merged, /@currentField &lt; 5/)
  })

  it('keeps them escaped when the template changes something else', () => {
    const merged = mergeFieldXml(
      storedWithEntities,
      '<Field Type="Text" Name="GtCode" ID="{a1a1a1a1-0000-4000-8000-000000000001}" MaxLength="40" />'
    )
    assert.equal(fieldAttributes(merged).MaxLength, '40')
    assert.equal(
      fieldAttributes(merged).Description,
      'Plan & tiltak\r\nlinje 2'
    )
    assert.match(merged, /@currentField &lt; 5/)
  })
})

describe('fieldxml: setting attributes', () => {
  it('sets the given attributes, leaves an undefined one as it is and drops Version', () => {
    const set = fieldAttributes(
      setFieldAttributes(stored, {
        Required: true,
        Hidden: undefined,
        Indexed: undefined
      })
    )
    assert.equal(set.Required, 'true')
    assert.equal(set.Indexed, 'TRUE')
    assert.equal(set.DisplayName, 'Kode')
    assert.equal(set.Version, undefined)
  })
})
