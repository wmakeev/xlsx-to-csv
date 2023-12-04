import ExcelJS, { Cell } from 'exceljs'
import assert from 'node:assert'
import { fromAsyncGenerator } from '@wmakeev/highland-tools'
import { AssertConfig, Config, HeadFieldConfig } from './types.js'

export * from './types.js'
export * as tools from './tools/index.js'

export const defaultGetCellValue = (cell: Cell | undefined): string => {
  return cell?.text ?? ''
}

async function* sheetRowsGen(
  this: XlsxToCsvParser,
  worksheetReader: ExcelJS.stream.xlsx.WorksheetReader,
  sheetIndex: number
) {
  // @ts-expect-error not typed
  const sheetName = worksheetReader.name

  const getCellValue = this.config.cellToString ?? defaultGetCellValue

  const configName = this.config.sheetConfigSelector?.(sheetName) ?? sheetName

  const sheetConfig =
    this.config?.sheetConfigs?.find((it, index) => {
      if (it.name === configName) return true

      if (it.sheetIndex === sheetIndex) return true

      if (it.name == null && it.sheetIndex == null && index === sheetIndex) {
        return true
      }

      return false
    }) ?? {}

  /**
   * Результирующее наименовение заголовка и его ключ.
   * Индекс массива соотв. идексу заголовка в конфигурации.
   */
  let headerColumns: (
    | { header: string; columnKey: number | string }
    | undefined
  )[]

  let headersConfigs = sheetConfig.headers

  const headFieldConfigsByRowNum = sheetConfig.headFields?.reduce((res, it) => {
    const items = res.get(it.rowNum)

    if (items !== undefined) items.push(it)
    else res.set(it.rowNum, [it])

    return res
  }, new Map<number, HeadFieldConfig[]>())

  const assetsConfigsByRowNum = sheetConfig.asserts?.reduce((res, it) => {
    const items = res.get(it.rowNum)

    if (items !== undefined) items.push(it)
    else res.set(it.rowNum, [it])

    return res
  }, new Map<number, AssertConfig[]>())

  const headerRow = sheetConfig.headerRow ?? 1

  const headFieldsMap = new Map<string, unknown>()

  for await (const row of worksheetReader) {
    const rowNum = row.number

    // assets
    if (rowNum <= headerRow) {
      // search for fields
      assetsConfigsByRowNum?.get(rowNum)?.forEach(c => {
        const cell = row.getCell(c.columnKey)

        if (cell == null) return

        const result = c.assert(cell)

        if (result === false) {
          throw new Error(
            `Assertion "${c.name}" at cell ${cell.address} failed`
          )
        }
      })
    }

    // If pre table space
    if (rowNum < headerRow) {
      // search for fields
      headFieldConfigsByRowNum?.get(rowNum)?.forEach(c => {
        const cell = row.getCell(c.columnKey)

        if (cell == null) return

        headFieldsMap.set(
          c.name,
          c.value ? c.value({ cell }) : getCellValue(cell)
        )
      })

      continue
    }

    /** Row cells (order from 1) */
    const rowCells: Cell[] = []

    row.eachCell((cell, colNumber) => {
      rowCells[colNumber] = cell
    })

    //#region Row is header
    if (rowNum === headerRow) {
      if (headersConfigs === undefined) {
        headersConfigs = []

        for (const [colNum, cell] of rowCells.entries()) {
          if (cell == null) continue

          headersConfigs!.push({
            type: 'actual',
            columnKey: colNum
          })
        }
      }

      headerColumns = headersConfigs.map(headerConf => {
        if (headerConf.type === 'virtual') return undefined

        let columnKey = headerConf.columnKey

        // #region Header auto mapping
        if (columnKey == null) {
          const headerTest =
            headerConf.headerNameTest ??
            (headerConf.name == null
              ? headerName => headerName === headerConf.name
              : undefined)

          if (!headerTest) {
            throw new Error(
              `[${sheetName}] headerNameTest or header name should be specified`
            )
          }

          columnKey = rowCells.findIndex(cell => headerTest(getCellValue(cell)))

          if (columnKey == null || columnKey === -1) {
            throw new Error(
              `[${sheetName}] Column "${headerConf.name}" not found in sheet "${sheetName}"`
            )
          }

          const header = getCellValue(rowCells[columnKey]!)

          return { header, columnKey }
        }
        //#endregion

        //#region Header implicit mapping
        const sourceCell = row.getCell(columnKey)

        const sourceHeader = getCellValue(sourceCell)

        const targetHeader = headerConf.name ?? sourceHeader

        if (targetHeader == null || targetHeader === '') {
          throw new Error(
            `[${sheetName}] Header cell at column=${columnKey} should not be empty or have implicit name in config`
          )
        }

        if (headerConf.headerNameTest) {
          const isCorrectHeader = headerConf.headerNameTest(sourceHeader)

          if (!isCorrectHeader) {
            throw new Error(
              `[${sheetName}] "${targetHeader}" header mapping is not correct`
            )
          }
        }

        columnKey = sourceCell?.col ?? headerConf.columnKey

        return { header: targetHeader, columnKey }
        //#endregion
      })

      const headers = headerColumns.map((col, index) => {
        const name = headersConfigs?.[index]?.name

        if (name != null) return name

        assert.ok(col)

        return col.header
      })

      yield headers

      continue
    }
    //#endregion

    //#region Regular row

    assert.ok(headersConfigs, 'No headers config')

    const targetRowValues = headersConfigs.map((h, index) => {
      if (h.type === 'virtual') {
        return h.value({ row, headFields: headFieldsMap, headerName: h.name })
      }

      const col = headerColumns[index]

      assert.ok(col)

      const cell = row.getCell(col.columnKey)

      return (
        h.value?.({
          cell,
          row,
          headFields: headFieldsMap,
          headerName: col.header
        }) ??
        this.config.cellToString?.(cell) ??
        getCellValue(cell)
      )
    })

    if (sheetConfig.rowsFilter?.(targetRowValues) ?? true) {
      yield targetRowValues
    }
    //#endregion
  }
}

export class XlsxToCsvParser {
  protected config: Config

  constructor(config?: Config) {
    this.config = config ?? {}
  }

  /**
   * Получить строки указанного листа
   *
   * @param file Путь к xlsx файлу
   * @param sheet Наименование листа или его индекс (если не указано, то первый лист)
   * @returns Highland покток строк таблицы листа
   */
  async getSheetRowsStream(file: string, sheet?: string | number) {
    const result$ = await this.getSheetsStream(file)
      .find(it => {
        if (sheet == null) return true

        return typeof sheet === 'number'
          ? it.sheetIndex === sheet
          : it.sheetName === sheet
      })
      .map(it => it.rows$)
      .last()
      .toPromise(Promise)

    return result$
  }

  /**
   * @param file Путь к XLSX файлу
   */
  getSheetsStream(file: string) {
    const sheetsGen = async function* () {
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(file, {})

      let sheetIndex = 0

      for await (const worksheetReader of workbookReader) {
        yield [worksheetReader, sheetIndex++] as const
      }
    }

    return fromAsyncGenerator(() => sheetsGen()).map(
      ([worksheetReader, sheetIndex]) => {
        const sheetName = (worksheetReader as any).name as string

        return {
          sheetName,
          sheetIndex,
          rows$: fromAsyncGenerator(() =>
            sheetRowsGen.call(this, worksheetReader, sheetIndex)
          )
        }
      }
    )
  }
}
