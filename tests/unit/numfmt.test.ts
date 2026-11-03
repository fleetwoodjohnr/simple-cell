import { describe, expect, it } from 'vitest'
import { formatValue, generalNumber } from '@/core/numfmt/format'
import { dateSerial, parseDateTime, serialToDate } from '@/core/numfmt/dates'

const f = (v: number | string, code: string) => formatValue(v, code).text

describe('number formats', () => {
  it('formats numbers like Excel', () => {
    const cases: [number, string, string][] = [
      [1234.567, '#,##0.00', '1,234.57'],
      [1234567.891, '#,##0', '1,234,568'],
      [0, '#,##0', '0'],
      [0.25, '0%', '25%'],
      [0.12345, '0.00%', '12.35%'],
      [-5, '0.00', '-5.00'],
      [-5, '0.00;(0.00)', '(5.00)'],
      [0, '0.00;-0.00;"zero"', 'zero'],
      [12345.678, '0.00E+00', '1.23E+04'],
      [0.000123, '0.00E+00', '1.23E-04'],
      [12345, '##0.0E+0', '12.3E+3'],
      [5551234567, '(###) ###-####', '(555) 123-4567'],
      [1500000, '#,##0,,"M"', '2M'],
      [1500, '#,##0.0,"K"', '1.5K'],
      [0.5, '# ?/?', ' 1/2'],
      [1.25, '# ?/?', '1 1/4'],
      [3.14159, '# ??/??', '3 14/99'],
      [0.75, '?/8', '6/8'],
      [12.5, '.00', '12.50'],
      [7, '000', '007'],
      [2.5, '0.#', '2.5'],
      [2, '0.#', '2.'],
      [1234.5, '$#,##0.00', '$1,234.50'],
      [-1234.5, '"$"#,##0.00_);\\("$"#,##0.00\\)', '($1,234.50)'],
      [42, '[Red]0;[Blue]-0', '42'],
      [2.675, '0.00', '2.68']
    ]
    for (const [v, code, want] of cases) expect([v, code, f(v, code)]).toEqual([v, code, want])
    expect(formatValue(-1, '[Red]0;[Blue]-0').color).toBe('#0000FF')
  })
  it('formats dates and times', () => {
    const d = dateSerial(2024, 1, 5) + (13 * 3600 + 7 * 60 + 9) / 86400
    expect(f(d, 'yyyy-mm-dd')).toBe('2024-01-05')
    expect(f(d, 'm/d/yyyy')).toBe('1/5/2024')
    expect(f(d, 'dddd, mmmm d, yyyy')).toBe('Friday, January 5, 2024')
    expect(f(d, 'h:mm AM/PM')).toBe('1:07 PM')
    expect(f(d, 'hh:mm:ss')).toBe('13:07:09')
    expect(f(d, 'mmm-yy')).toBe('Jan-24')
    expect(f(1.5, '[h]:mm')).toBe('36:00')
    expect(f(0.000694444, 'mm:ss')).toBe('01:00')
    expect(f(dateSerial(2024, 1, 5) + 0.5 + 0.25 / 86400, 'h:mm:ss.00')).toBe('12:00:00.25')
  })
  it('text sections', () => {
    expect(f('abc', '0;-0;0;"[" @ "]"')).toBe('[ abc ]')
    expect(f('abc', '0.00')).toBe('abc')
    expect(f('abc', '@')).toBe('abc')
  })
  it('General fits the column', () => {
    expect(generalNumber(1234567.891)).toBe('1234567.891')
    expect(generalNumber(123456789012345)).toBe('1.23457E+14')
    expect(generalNumber(0.1 + 0.2)).toBe('0.3')
    expect(generalNumber(1 / 3)).toBe('0.333333333')
  })
})

describe('date serials', () => {
  it('match Excel, including the 1900 leap-year bug', () => {
    expect(dateSerial(1900, 1, 1)).toBe(1)
    expect(dateSerial(1900, 2, 28)).toBe(59)
    expect(dateSerial(1900, 3, 1)).toBe(61)
    expect(dateSerial(2024, 1, 1)).toBe(45292)
    expect(serialToDate(60)).toEqual({ y: 1900, m: 2, d: 29 })
    expect(serialToDate(45292)).toEqual({ y: 2024, m: 1, d: 1 })
    expect(dateSerial(2024, 14, 1)).toBe(dateSerial(2025, 2, 1))
    expect(dateSerial(2024, 1, 1, true)).toBe(45292 - 1462)
  })
  it('parse typed dates and times', () => {
    expect(parseDateTime('2024-01-15')).toBe(dateSerial(2024, 1, 15))
    expect(parseDateTime('1/15/2024')).toBe(dateSerial(2024, 1, 15))
    expect(parseDateTime('15-Jan-2024')).toBe(dateSerial(2024, 1, 15))
    expect(parseDateTime('January 15, 2024')).toBe(dateSerial(2024, 1, 15))
    expect(parseDateTime('13:30')).toBeCloseTo(13.5 / 24)
    expect(parseDateTime('1:30 PM')).toBeCloseTo(13.5 / 24)
    expect(parseDateTime('1/15/2024 6:00')).toBeCloseTo(dateSerial(2024, 1, 15) + 0.25)
    expect(parseDateTime('hello')).toBeNull()
    expect(parseDateTime('2/30/2024')).toBeNull()
  })
})
