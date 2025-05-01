import { readFile, writeFile } from 'fs/promises'
import { parse as csvParse } from 'csv/sync'
import { decrypt as mysqlDecrypt } from 'mysql-aes'
import type { Simplify } from 'type-fest'
import xlsx from 'sheetjs-style'

export interface DBColumn<T = unknown> {
  name: string
  encrypted?: boolean
  processor?: (value: string) => T
}

export type DBRow<DBColumns extends readonly DBColumn[]> = Simplify<{
  [K in DBColumns[number] as K['name']]: K['processor'] extends (value: unknown) => infer R
    ? R
    : (string | null);
}>

export type ExcelColumn<DBColumns extends readonly DBColumn[]> = Simplify<{
  title: string
  from: (row: DBRow<DBColumns>) => string | null | undefined
  width?: number
  titleStyle?: CellStyle
  contentStyle?: CellStyle
}>

export type CellStyle = xlsx.CellObject['s']

export const SortMode = {
  Asc: 'asc',
  Desc: 'desc',
} as const

export type SortModeValue = typeof SortMode[keyof typeof SortMode]

interface ConvertParamsBase<DBColumns extends readonly DBColumn[]> {
  dbColumns: DBColumns
  excelColumns: readonly ExcelColumn<DBColumns>[]
  sortFns?: ((row: DBRow<DBColumns>) => number)[]
  sortMode?: SortModeValue
  titleStyle?: CellStyle
  contentStyle?: CellStyle
  csvHasColumns?: boolean
  csvDelimiter?: string
}

export type ConvertParams<DBColumns extends readonly DBColumn[]> =
  { [K in keyof DBColumns]: DBColumns[K]['encrypted'] extends true ? true : never }[number] extends never
    ? Simplify<ConvertParamsBase<DBColumns> & { aesKey?: string }>
    : Simplify<ConvertParamsBase<DBColumns> & { aesKey: string }>

export const convertDBCsvToExcel = async <
  const DBColumns extends readonly DBColumn[],
>(
  inputFilePath: string | URL,
  outputFilePath: string | URL,
  {
    aesKey = '',
    dbColumns,
    excelColumns,
    sortFns = [],
    sortMode = 'desc',
    titleStyle,
    contentStyle,
    csvHasColumns = false,
    csvDelimiter = ',',
  }: ConvertParams<DBColumns>,
) => {
  if (!aesKey && dbColumns.some(col => col.encrypted)) {
    throw new Error('aesKey is required when dbColumns has encrypted column')
  }

  const inputContent = await readFile(inputFilePath, 'utf-8')
  const decrypt = (str: string) => mysqlDecrypt(str, aesKey)

  const dbRows = csvParse(inputContent, {
    skipEmptyLines: true,
    columns: csvHasColumns ? true : dbColumns.map(col => col.name),
    delimiter: csvDelimiter,
    cast: (value, context) => {
      if (context.header) {
        return value
      }

      let parsedValue: unknown = value === 'NULL' ? null : value
      const options = dbColumns[context.index]
      if (options?.encrypted && parsedValue) {
        parsedValue = decrypt(parsedValue as string)
      }

      if (options?.processor) {
        parsedValue = options.processor(parsedValue as string)
      }

      return parsedValue
    },
  }) as DBRow<DBColumns>[]

  const excelTitleRow: xlsx.CellObject[] = excelColumns.map(col => ({
    v: col.title,
    t: 's',
    s: col.titleStyle ?? titleStyle,
  }))

  if (sortFns.length > 0) {
    const fnSortMap = new WeakMap(
      sortFns.map(fn => [
        fn,
        new WeakMap(dbRows.map(row => [row, fn(row)])),
      ]),
    )

    dbRows.sort((rowA, rowB) => {
      for (const sortFn of sortFns) {
        const sortMap = fnSortMap.get(sortFn)!
        const sortValueA = sortMap.get(rowA)!
        const sortValueB = sortMap.get(rowB)!
        const diff = sortValueA - sortValueB
        if (diff !== 0) {
          return sortMode === SortMode.Desc ? -diff : diff
        }
      }

      return 0
    })
  }

  const excelContentRows: xlsx.CellObject[][] = dbRows
    .map(row => excelColumns.map(col => ({ v: col.from(row) ?? '', t: 's', s: col.contentStyle ?? contentStyle })))

  const book = xlsx.utils.book_new()
  const sheet = xlsx.utils.aoa_to_sheet([excelTitleRow, ...excelContentRows])
  sheet['!cols'] = excelColumns.map(col => ({ wch: col.width }))
  xlsx.utils.book_append_sheet(book, sheet)

  const outputFileData = xlsx.write(book, {
    type: 'buffer',
    bookType: 'xlsx',
    compression: true,
  }) as Buffer
  await writeFile(outputFilePath, outputFileData)
}
