import assert from 'assert'
import ExcelJS from '@wmakeev/exceljs'
import path from 'path'

const workbook = new ExcelJS.Workbook()

const FILE = path.join(process.cwd(), '__temp/income/test-images.xlsx')

const wb = await workbook.xlsx.readFile(FILE)

const ws = wb.getWorksheet(1)

// const images = ws?.getImages()

const a2 = ws?.getCell('A2')

assert.ok(a2)

console.log('DONE.')
