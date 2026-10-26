import { ISiteConnection } from '../connect'

/**
 * A DateTime value as SharePoint stores it (UTC) and as the web shows it (the wall time in the
 * web's regional time zone, from `utcToLocalTime`).
 */
export interface IDateValue {
  utc: string
  local: string
}

/**
 * One list item as the DataRows of a template leave it.
 */
export interface IDataRowItem {
  Id: number
  /**
   * The raw REST values by internal name. Lookup and User columns are read as their ids
   * (`<name>Id`: a number, or an array for the multi-value types), DateTime as an
   * {@link IDateValue}; an empty column is `null`.
   */
  values: Record<string, any>
  /**
   * `FieldValuesAsText` by internal name, as the list shows the values (in the web's locale: a
   * Number `12,50` and a Boolean `Ja` on a Norwegian web).
   */
  text: Record<string, string>
}

/**
 * A list's items with the columns that were asked for.
 */
export interface IDataRowsList {
  /** `TypeAsString` by internal name. */
  types: Record<string, string>
  /** Ordered by id. */
  items: IDataRowItem[]
}

/** Columns that store an id: written and read as `<name>Id`. */
const ID_TYPES = /^(Lookup|User)/

/**
 * Reads back the items of the given lists with the given columns (`Title` is always read).
 *
 * @param conn - The web
 * @param lists - List title -> internal names of the columns to read
 */
export async function readDataRows(
  conn: ISiteConnection,
  lists: Record<string, string[]>
): Promise<Record<string, IDataRowsList>> {
  const state: Record<string, IDataRowsList> = {}
  for (const title of Object.keys(lists)) {
    const names = ['Title', ...lists[title].filter((name) => name !== 'Title')]
    const fields = await conn.web.lists
      .getByTitle(title)
      .fields.select('InternalName', 'TypeAsString')<
      Array<{ InternalName: string; TypeAsString: string }>
    >()
    const types: Record<string, string> = {}
    for (const name of names) {
      const field = fields.find((candidate) => candidate.InternalName === name)
      if (!field) throw new Error(`List ${title} has no column ${name}`)
      types[name] = field.TypeAsString
    }
    const column = (name: string) =>
      ID_TYPES.test(types[name]) ? `${name}Id` : name
    const rows = await conn.web.lists
      .getByTitle(title)
      .items.select('Id', ...names.map(column))
      .orderBy('Id')
      .top(5000)<any[]>()
    const items: IDataRowItem[] = []
    for (const row of rows) {
      const values: Record<string, any> = {}
      const text: Record<string, string> = {}
      // Per item: `$expand=FieldValuesAsText` on the collection shows a Lookup or User as its id
      // and drops the properties a `$select` names (observed 2026-10-21).
      const asText = await conn.web.lists
        .getByTitle(title)
        .items.getById(row.Id)
        .fieldValuesAsText<Record<string, string>>()
      for (const name of names) {
        const raw = row[column(name)]
        values[name] =
          types[name] === 'DateTime' && raw
            ? {
                utc: raw,
                local: await conn.web.regionalSettings.timeZone.utcToLocalTime(
                  raw
                )
              }
            : raw === undefined
            ? null
            : raw
        text[name] = asText[name]
      }
      items.push({ Id: row.Id, values, text })
    }
    state[title] = { types, items }
  }
  return state
}
