import assert from 'assert/strict'
import { stringify } from 'csv-stringify/sync'
import { mkdir, writeFile } from 'node:fs/promises'
import path from 'node:path'
import test from 'node:test'

import { createReadStream } from 'fs'
import { XlsxToCsvParser, tools } from '../src/index.js'

const { fillHeadFieldValue } = tools

/**
 * Get simple table on first sheet
 *
 * - Table from first sheet
 * - Table header is on first row
 */
test('Simple table', async () => {
  const XLSX_FILE = path.join(process.cwd(), 'test/cases/01_simple.xlsx')

  // Simple case without config
  const parser = new XlsxToCsvParser()

  const rows$ = await parser.getSheetRowsStream(XLSX_FILE)

  const rows = await rows$.collect().toPromise(Promise)

  assert.deepEqual(
    rows,
    [
      ['№', 'Info', 'Value', 'Header 1', 'Header 2', 'Header1'],
      ['1', 'Text', 'желтый', '1', '', ''],
      ['2', 'Number', '123', '', '2', 'a'],
      ['3', 'Formula', '42', '', '', 'b'],
      ['4', 'Error 1', '', '3.53', '', 'c'],
      ['5', 'Error 2', '', '', '', ''],
      ['6', 'Date', '45260', '452.89', '4', ''],
      ['7', 'Empty', '', '', '', ''],
      ['8', 'Image', '', '5', '', ''],
      ['9', 'QR-code', '', '', '', ''],
      ['10', 'Link', 'http://example.com/', '', '', '']
    ],
    'should return rows'
  )
})

/**
 *
 */
test('Simple table (formats)', async () => {
  const XLSX_FILE = path.join(process.cwd(), 'test/cases/01_simple.xlsx')

  const parser = new XlsxToCsvParser()

  const rows$ = await parser.getSheetRowsStream(XLSX_FILE, 1)

  const rows = await rows$.collect().toPromise(Promise)

  // FIXME Почему видит дату как число?
  assert.deepEqual(
    rows,
    [
      ['Float', 'Date'],
      ['45.258965', '45260']
    ],
    'should return rows'
  )
})

test('Complex report #1', async () => {
  const XLSX_FILE = path.join(process.cwd(), 'test/cases/02_report_1.xlsx')

  const parser = new XlsxToCsvParser({
    sheetConfigs: [
      {
        asserts: [
          {
            name: 'Заголовок',
            columnKey: 'A',
            rowNum: 2,
            assert: cell => cell?.text.startsWith('Финансовый отчет')
          },
          {
            name: 'Продавец',
            columnKey: 'A',
            rowNum: 4,
            assert: cell => cell?.text === 'Продавец:'
          },
          {
            name: 'Договор',
            columnKey: 'A',
            rowNum: 5,
            assert: cell => cell?.text === 'Договор:'
          },
          {
            name: 'Номер п/п',
            columnKey: 'A',
            rowNum: 6,
            assert: cell => cell?.text === 'Номер п/п:'
          },
          {
            name: 'Дата п/п',
            columnKey: 'A',
            rowNum: 7,
            assert: cell => cell?.text === 'Дата п/п:'
          },
          {
            name: 'Заголовок Отправление',
            columnKey: 'A',
            rowNum: 11,
            assert: cell => cell?.text.startsWith('Отправление')
          },
          {
            name: 'Заголовок Описание',
            columnKey: 'D',
            rowNum: 11,
            assert: cell => cell?.text.startsWith('Описание')
          },
          {
            name: 'Заголовок Заказ Продавца',
            columnKey: 'F',
            rowNum: 11,
            assert: cell => cell?.text.startsWith('Заказ Продавца')
          },
          {
            name: 'Заголовок Классификатор',
            columnKey: 'I',
            rowNum: 11,
            assert: cell => cell?.text.startsWith('Классификатор')
          },
          {
            name: 'Заголовок Долг компании',
            columnKey: 'L',
            rowNum: 12,
            assert: cell => cell?.text.startsWith('Долг компании')
          },
          {
            name: 'Заголовок Долг продавца',
            columnKey: 'M',
            rowNum: 12,
            assert: cell => cell?.text.startsWith('Долг продавца')
          }
        ],

        headFields: [
          {
            name: 'Финансовый отчет',
            columnKey: 'A',
            rowNum: 2,
            value: ({ cell }) => cell.text.split(' ')[3]
          },
          {
            name: 'Продавец',
            columnKey: 'C',
            rowNum: 4
          },

          {
            name: 'Договор',
            columnKey: 'C',
            rowNum: 5,
            value: ({ cell }) => cell.text.split('№')[1]?.split('от')[0]?.trim()
          },
          {
            name: 'Номер п/п',
            columnKey: 'C',
            rowNum: 6
          },
          {
            name: 'Дата п/п',
            columnKey: 'C',
            rowNum: 7
          }
        ],

        headerRow: 12,

        headers: [
          {
            type: 'virtual',
            name: 'Финансовый отчет',
            value: fillHeadFieldValue()
          },
          {
            type: 'virtual',
            name: 'Продавец',
            value: fillHeadFieldValue()
          },
          {
            type: 'virtual',
            name: 'Договор',
            value: fillHeadFieldValue()
          },
          {
            type: 'virtual',
            name: 'Номер п/п',
            value: fillHeadFieldValue()
          },
          {
            type: 'virtual',
            name: 'Дата п/п',
            value: fillHeadFieldValue()
          },
          {
            type: 'actual',
            columnKey: 'A',
            name: 'Отправление Маркетплейс / id задолженности'
          },
          {
            type: 'actual',
            columnKey: 'D',
            name: 'Описание (расшифровка)'
          },
          {
            type: 'actual',
            columnKey: 'F',
            name: 'Заказ Продавца'
          },
          {
            type: 'actual',
            columnKey: 'I',
            name: 'Классификатор'
          },
          {
            type: 'actual',
            columnKey: 'L',
            name: 'Долг компании'
          },
          {
            type: 'actual',
            columnKey: 'M',
            name: 'Долг продавца'
          }
        ],

        rowsFilter(row) {
          return row[5] !== '' && row[5] !== 'Итого'
        }
      }
    ]
  })

  const xlsxStream = createReadStream(XLSX_FILE)

  const rows$ = await parser.getSheetRowsStream(xlsxStream)

  assert.ok(rows$)

  const rows = await rows$.collect().toPromise(Promise)

  const sample = rows.slice(0, 3)

  assert.deepEqual(
    sample,
    [
      [
        'Финансовый отчет',
        'Продавец',
        'Договор',
        'Номер п/п',
        'Дата п/п',
        'Отправление Маркетплейс / id задолженности',
        'Описание (расшифровка)',
        'Заказ Продавца',
        'Классификатор',
        'Долг компании',
        'Долг продавца'
      ],
      [
        'МПБЛ-108601',
        'Иван Иванович Иванов',
        'К-4284-05-2020',
        '529999',
        '22.11.2023',
        '8010012619972',
        '',
        '8010012619972',
        'Вознаграждение оператора ПЛ',
        '',
        '76.6'
      ],
      [
        'МПБЛ-108601',
        'Иван Иванович Иванов',
        'К-4284-05-2020',
        '529999',
        '22.11.2023',
        '8010968940808',
        '',
        '8010968940808',
        'Вознаграждение за предоставление поощрения',
        '',
        '1396'
      ]
    ],
    'should return rows'
  )

  const reportPath = path.join(process.cwd(), '__temp/test-out/sm')

  await mkdir(reportPath, { recursive: true })

  await writeFile(
    path.join(reportPath, 'report-1.csv'),
    stringify(rows, { bom: true }),
    'utf-8'
  )
})

test('Complex report #2', async () => {
  const XLSX_FILE = path.join(process.cwd(), 'test/cases/02_report_1.xlsx')

  const parser = new XlsxToCsvParser({
    sheetConfigs: [
      {
        asserts: [
          {
            name: 'Заголовок',
            columnKey: 'A',
            rowNum: 2,
            assert: cell => cell?.text.startsWith('Финансовый отчет')
          },
          {
            name: 'Продавец',
            columnKey: 'A',
            rowNum: 4,
            assert: cell => cell?.text === 'Продавец:'
          },
          {
            name: 'Договор',
            columnKey: 'A',
            rowNum: 5,
            assert: cell => cell?.text === 'Договор:'
          },
          {
            name: 'Номер п/п',
            columnKey: 'A',
            rowNum: 6,
            assert: cell => cell?.text === 'Номер п/п:'
          },
          {
            name: 'Дата п/п',
            columnKey: 'A',
            rowNum: 7,
            assert: cell => cell?.text === 'Дата п/п:'
          }
        ],

        headFields: [
          {
            name: 'Финансовый отчет',
            columnKey: 'A',
            rowNum: 2,
            value: ({ cell }) => cell.text.split(' ')[3]
          },
          {
            name: 'Продавец',
            columnKey: 'C',
            rowNum: 4
          },

          {
            name: 'Договор',
            columnKey: 'C',
            rowNum: 5,
            value: ({ cell }) => cell.text.split('№')[1]?.split('от')[0]?.trim()
          },
          {
            name: 'Номер п/п',
            columnKey: 'C',
            rowNum: 6
          },
          {
            name: 'Дата п/п',
            columnKey: 'C',
            rowNum: 7
          }
        ],

        headerRow: 11,

        headers: [
          {
            type: 'virtual',
            name: 'Финансовый отчет',
            value: fillHeadFieldValue()
          },
          {
            type: 'virtual',
            name: 'Продавец',
            value: fillHeadFieldValue()
          },
          {
            type: 'virtual',
            name: 'Договор',
            value: fillHeadFieldValue()
          },
          {
            type: 'virtual',
            name: 'Номер п/п',
            value: fillHeadFieldValue()
          },
          {
            type: 'virtual',
            name: 'Дата п/п',
            value: fillHeadFieldValue()
          },
          {
            type: 'actual',
            name: 'Отправление Маркетплейс / id задолженности'
          },
          {
            type: 'actual',
            name: 'Описание (расшифровка)'
          },
          {
            type: 'actual',
            name: 'Заказ Продавца'
          },
          {
            type: 'actual',
            name: 'Классификатор'
          },
          {
            type: 'actual',
            name: 'Долг компании',
            headerNameTest: name => name?.startsWith('Итого') ?? false
          }
        ],

        rowsFilter(row) {
          return row[5] !== '' && row[5] !== 'Итого' && row[9] !== ''
        }
      }
    ]
  })

  const xlsxStream = createReadStream(XLSX_FILE)

  const rows$ = await parser.getSheetRowsStream(xlsxStream)

  assert.ok(rows$)

  const rows = await rows$.collect().toPromise(Promise)

  const sample = rows.slice(0, 3)

  assert.deepEqual(
    sample,
    [
      [
        'Финансовый отчет',
        'Продавец',
        'Договор',
        'Номер п/п',
        'Дата п/п',
        'Отправление Маркетплейс / id задолженности',
        'Описание (расшифровка)',
        'Заказ Продавца',
        'Классификатор',
        'Долг компании'
      ],
      [
        'МПБЛ-108601',
        'Иван Иванович Иванов',
        'К-4284-05-2020',
        '529999',
        '22.11.2023',
        '8010968940808',
        '',
        '8010968940808',
        'Вознаграждение оператора ПЛ',
        '25.15'
      ],
      [
        'МПБЛ-108601',
        'Иван Иванович Иванов',
        'К-4284-05-2020',
        '529999',
        '22.11.2023',
        '8011422267125',
        '',
        '8011422267125',
        'Вознаграждение за предоставление поощрения',
        '991'
      ]
    ],
    'should return rows'
  )

  const reportPath = path.join(process.cwd(), '__temp/test-out/sm')

  await mkdir(reportPath, { recursive: true })

  await writeFile(
    path.join(reportPath, 'report-2.csv'),
    stringify(rows, { bom: true }),
    'utf-8'
  )
})

test('Complex report #3', async () => {
  const xlsxParser = new XlsxToCsvParser({
    sheetConfigs: [
      {
        headers: [
          {
            type: 'actual',
            name: 'Кабинет поставщика'
          },
          {
            type: 'actual',
            name: 'Ид кабинета поставщика'
          },
          {
            type: 'actual',
            name: 'Артикул поставщика (uid)'
          },
          {
            type: 'actual',
            name: 'Название товара'
          },
          {
            type: 'actual',
            name: 'Код размера (chrt_id)'
          },
          {
            type: 'actual',
            name: 'Артикул WB'
          },
          {
            type: 'actual',
            name: 'Артикул ИМТ'
          },
          {
            type: 'actual',
            name: 'Размер'
          },
          {
            type: 'actual',
            name: 'Штрихкод'
          },
          {
            type: 'actual',
            name: 'Торговая марка'
          }
        ]
      }
    ]
  })

  const XLSX_FILE = path.join(process.cwd(), 'test/cases/1401887.xlsx')

  const xlsxStream = createReadStream(XLSX_FILE, {
    highWaterMark: 2000
  })

  const rows = await (await xlsxParser.getSheetRowsStream(xlsxStream))
    .collect()
    .toPromise(Promise)

  const csv = stringify(rows, {
    quoted_empty: true
  })

  assert.strictEqual(csv.indexOf('�'), -1)
})
