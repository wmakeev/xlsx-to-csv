import { Cell, CellValue, Row } from '@wmakeev/exceljs'

export type { Cell, CellValue, Row } from '@wmakeev/exceljs'

export type ActualRowArg = {
  value: CellValue
}

export interface HeadFieldConfig {
  /**
   * Наименование поля. Которое будет доступно для использования в подстановке
   * в процессе разбора табличной части.
   */
  name: string

  /**
   * Номер колонки в которой находится ячейка поля.
   *
   * Например: `A` или `1`
   */
  columnKey: number | string

  /** Номер строки в которой находится ячейка поля */
  rowNum: number

  /**
   * Возвращает преобразованное значение
   */
  value?: (ctx: { cell: Cell }) => string | undefined
}

export interface valueContext {
  headerName: string
  headFields: Map<string, unknown>
}

/**
 * Заголовок значения которого находятся в ячейках
 */
export interface ActualHeaderConfig {
  /** Актуальный заголовок */
  type: 'actual'

  /**
   * Наименование заголовка в результирующей таблице.
   *
   * Если наименование не указано, используется наименование из исходной таблицы.
   */
  name?: string

  /**
   * Номер колонки в которой находится заголовок.
   *
   * Например: `A` или `1`
   *
   * Если не указано, то заголовок будет найден по наименованию `name` или
   * по предикату `srcHeaderNameTest`.
   */
  columnKey?: number | string

  /**
   * Функция которая проверяет значение полученной ячейки заголовка.
   *
   * Если значение не соответствует, то возвращает `false`.
   *
   * Необходимо для дополнительного контроля, что расположение данных на
   * листе соответствует ожидаемому формату если указан `column`.
   *
   * Или для автоматического поиска индекса колонки по наименованию, если
   * `column` или `name`.
   */
  headerNameTest?: (name: string | undefined) => boolean

  /**
   * Возвращает преобразованное значение
   */
  value?: (ctx: { cell: Cell; row: Row } & valueContext) => string | undefined
}

export interface VirtualHeaderConfig {
  /** Виртуальный заголовок */
  type: 'virtual'

  /**
   * Наименование заголовка в результирующей таблице.
   *
   * Если наименование не указано, используется наименование из исходной таблицы.
   */
  name: string

  /** Значение в конкретной строке таблицы для данного заголовка */
  value: (ctx: { row: Row } & valueContext) => string | undefined
}

export type HeaderConfig = ActualHeaderConfig | VirtualHeaderConfig

export interface AssertConfig {
  /**
   * Наименование поля. Которое будет доступно для использования в подстановке
   * в процессе разбора табличной части.
   */
  name: string

  /**
   * Номер колонки в которой находится ячейка поля.
   *
   * Например: `A` или `1`
   */
  columnKey: number | string

  /** Номер строки в которой находится ячейка поля */
  rowNum: number

  /** Функция для проверки значения */
  assert: (cell: Cell) => boolean
}

/**
 * Конфигурация листа
 */
export interface SheetConfig {
  /**
   * The name of the sheet for which this configuration is intended.
   * Or custom string used with `sheetConfigSelector` for config binding.
   */
  name?: string

  /**
   * The index of the sheet for which this configuration is intended.
   *
   * Used for auto bind config if `sheetConfigSelector` not defined.
   */
  sheetIndex?: number

  /**
   * Номер строки в которой содержится заголовок
   *
   * default: `1`
   */
  headerRow?: number

  asserts?: AssertConfig[]

  headFields?: HeadFieldConfig[]

  /** Конфигурация для заголовков */
  headers?: HeaderConfig[]

  /** Предикат для пост-фильтрации результирующих строк файла */
  rowsFilter?: (row: (string | undefined)[]) => boolean
}

export interface Config {
  /**
   * Конфигурации листов по наименованию
   */
  sheetConfigs?: SheetConfig[]

  /**
   * Возвращает наименование конфигурации, которая будет использоваться
   * для обработки текущего листа в исходном файле.
   *
   * @param sheetName Наименование исходного листа в файле
   * @returns Наименование конфигурации
   */
  sheetConfigSelector?: (sheetName: string) => string

  cellToString?: (cell: Cell) => string
}

export interface SheetRowsStreamInfo {
  sheetName: string
  sheetIndex: number
  rows$: Highland.Stream<(string | undefined)[]>
}
