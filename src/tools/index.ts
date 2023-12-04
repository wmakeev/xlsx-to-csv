import assert from 'node:assert'
import * as url from 'node:url'

import { ActualHeaderConfig, Cell, Row, valueContext } from '../types.js'

/**
 * Returns `__dirname` like value for ESM
 *
 * ```js
 * const dirname = getDirname(import.meta.url)
 * ```
 * @link https://blog.logrocket.com/alternatives-dirname-node-js-es-modules/
 * @param fileUrl
 */
export const getDirname = (fileUrl: string) =>
  url.fileURLToPath(new URL('.', fileUrl))

/**
 * @param val
 */
export const parseMoney = (cell: Cell) => {
  const val = cell.value

  let _val

  if (typeof val === 'string') {
    _val = Number.parseFloat(val)
    if (Number.isNaN(_val)) return undefined
  } else if (typeof val === 'number') {
    _val = val
  } else {
    return undefined
  }

  return _val.toFixed(2)
}

/**
 * Заполнение пустых строк последним значением
 */
export const fillLastValue = () => {
  let lastValue: string | undefined = undefined

  const valueFn: NonNullable<ActualHeaderConfig['value']> = ({ cell }) => {
    const val = cell.text

    assert.ok(typeof val === 'string', 'expect value to be string')

    if (val === '') {
      return lastValue
    }

    lastValue = val

    return lastValue
  }

  return valueFn
}

/**
 * Fill row number if cell not empty
 */
export const fillNotEmptyRowNum: () => NonNullable<
  ActualHeaderConfig['value']
> =
  () =>
  ({ cell }) => {
    return cell.text !== '' ? cell.row : undefined
  }

/**
 * Fill col number if cell not empty
 */
export const fillNotEmptyColNum: () => NonNullable<
  ActualHeaderConfig['value']
> =
  () =>
  ({ cell }) => {
    return cell.text !== '' ? cell.col : undefined
  }

/**
 * Fill global head fild value
 *
 * @param fieldName head fild name
 */
export const fillHeadFieldValue: (
  fieldName?: string
) => (ctx: { row: Row } & valueContext) => string | undefined =
  fieldName =>
  ({ headFields, headerName }) => {
    const val =
      fieldName != null ? headFields.get(fieldName) : headFields.get(headerName)

    return typeof val === 'string' ? val : val != null ? String(val) : undefined
  }
