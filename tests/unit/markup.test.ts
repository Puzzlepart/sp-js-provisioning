import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { compactMarkup } from '../../src/util/markup'

describe('compactMarkup', () => {
  const stored = '<OrderBy><FieldRef Name="LinkFilename" /></OrderBy>'

  it('lays CAML out as SharePoint stores it', () => {
    // As the project template writes its queries: a line break, a space after `=`, single quotes.
    assert.equal(
      compactMarkup(
        '<OrderBy>\n  <FieldRef Name= "LinkFilename"/>\n</OrderBy>'
      ),
      stored
    )
    assert.equal(
      compactMarkup("<OrderBy><FieldRef Name='LinkFilename' /></OrderBy>"),
      stored
    )
    assert.equal(compactMarkup(stored), stored)
  })

  it('leaves text between tags, and a quote inside a value, as they are', () => {
    assert.notEqual(
      compactMarkup('<Value Type="Text">x = "y"</Value>'),
      compactMarkup('<Value Type="Text">x="y"</Value>')
    )
    assert.equal(
      compactMarkup(`<Value Type="Text" Note='a "b"'>c</Value>`),
      `<Value Type="Text" Note='a "b"'>c</Value>`
    )
  })
})
