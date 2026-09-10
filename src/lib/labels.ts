import * as XLSX from 'xlsx';
import JsBarcode from 'jsbarcode';
import { jsPDF } from 'jspdf';
import regularFont from 'dejavu-fonts-ttf/ttf/DejaVuSans.ttf?url';
import boldFont from 'dejavu-fonts-ttf/ttf/DejaVuSans-Bold.ttf?url';

export const MAX_FILE_MB = 10;
export const MAX_ROWS = 10000;
export type SourceRow = { rowNumber: number; cells: string[]; numeric: boolean[]; formulas: boolean[]; raw: unknown[] };
export type SheetData = { fileName: string; fileSize: number; headers: string[]; rows: SourceRow[]; fnskuColumn: number; mrpColumn: number };
export type LabelRow = { rowNumber: number; fnsku: string; mrpDisplay: string; errors: string[]; duplicate: boolean; bits: string };
export const sampleRecords = [['X001ABC123', 499], ['X001ABC124', 799], ['X001ABC125', 1299], ['X001ABC126', 349], ['X001ABC127', 599]];
export const demoLabel: LabelRow = { rowNumber: 2, fnsku: 'X001ABC123', mrpDisplay: '499', errors: [], duplicate: false, bits: barcodeBits('X001ABC123') };

export function barcodeBits(value: string): string {
  const output: { encodings?: { data: string }[] } = {};
  JsBarcode(output, value, { format: 'CODE128', displayValue: false, margin: 0 });
  return output.encodings?.map(e => e.data).join('') || '';
}
export function detect(headers: string[], kind: 'fnsku' | 'mrp') {
  const accepted = kind === 'fnsku' ? ['fnsku', 'fnsku code'] : ['mrp', 'mrp (₹)', 'maximum retail price'];
  const matches = headers.map((h, i) => accepted.includes(h.trim().toLowerCase()) ? i : -1).filter(i => i >= 0);
  return matches.length === 1 ? matches[0] : -1;
}
export async function parseFile(file: File): Promise<SheetData> {
  if (!/\.xlsx?$/i.test(file.name)) throw new Error('Please upload a valid Excel spreadsheet (.xlsx or .xls).');
  if (file.size > MAX_FILE_MB * 1024 * 1024) throw new Error(`This file exceeds the ${MAX_FILE_MB} MB limit. Please upload a smaller spreadsheet.`);
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellFormula: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet || !sheet['!ref']) throw new Error('This spreadsheet is empty. Add FNSKU and MRP headers and at least one data row.');
  const range = XLSX.utils.decode_range(sheet['!ref']);
  if (range.e.r > 100000 || range.e.c > 1000) throw new Error('This worksheet is too large. Remove unused rows and columns and try again.');
  const headers: string[] = [];
  for (let c = range.s.c; c <= range.e.c; c++) headers.push(String(XLSX.utils.format_cell(sheet[XLSX.utils.encode_cell({ r: range.s.r, c })]) || '').trim());
  const rows: SourceRow[] = [];
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const cells: string[] = [], numeric: boolean[] = [], formulas: boolean[] = [], raw: unknown[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      cells.push(cell ? String(XLSX.utils.format_cell(cell)).trim() : '');
      numeric.push(cell?.t === 'n'); formulas.push(Boolean(cell?.f)); raw.push(cell?.v);
    }
    if (cells.some(Boolean) || formulas.some(Boolean)) rows.push({ rowNumber: r + 1, cells, numeric, formulas, raw });
  }
  if (!rows.length) throw new Error('No data rows found. Add at least one FNSKU and MRP below the headers.');
  if (rows.length > MAX_ROWS) throw new Error(`Please use batches of ${MAX_ROWS.toLocaleString()} rows or fewer.`);
  return { fileName: file.name, fileSize: file.size, headers, rows, fnskuColumn: detect(headers, 'fnsku'), mrpColumn: detect(headers, 'mrp') };
}
export function validateRows(sheet: SheetData, f: number, m: number): LabelRow[] {
  const rows = sheet.rows.map(source => {
    const fnsku = source.cells[f] || '';
    const originalMrp = source.cells[m] || '';
    const errors: string[] = [];
    let bits = '';
    if (!fnsku) errors.push('Missing FNSKU');
    else if (source.numeric[f]) errors.push('FNSKU must be stored as text in Excel to preserve its exact value');
    else if (!/^[\x20-\x7E]+$/.test(fnsku)) errors.push('FNSKU contains characters not supported by Code 128');
    else {
      try { bits = barcodeBits(fnsku); if (bits.length > 200) errors.push('FNSKU is too long for a reliably scannable 2-inch label'); }
      catch { errors.push('Invalid Code 128 FNSKU'); }
    }
    if (source.formulas[f] || source.formulas[m]) errors.push('Replace formulas with their values before uploading');
    const mrpString = (source.numeric[m] ? String(source.raw[m]) : originalMrp).replace(/^(?:₹|INR|Rs\.?)\s*/i, '').trim();
    const currencyPattern = /^(?:\d+|\d{1,3}(?:,\d{3})+|\d{1,2}(?:,\d{2})*,\d{3})(?:\.\d{1,2})?$/;
    const mrp = Number(mrpString.replace(/,/g, ''));
    if (!originalMrp) errors.push('Missing MRP');
    else if (!currencyPattern.test(mrpString) || !Number.isFinite(mrp) || mrp < 0 || mrp > 999999999) errors.push('Invalid MRP: enter a non-negative amount with at most 2 decimal places');
    const mrpDisplay = Number.isFinite(mrp) ? mrp.toLocaleString('en-IN', { maximumFractionDigits: 2, minimumFractionDigits: mrpString.includes('.') ? 2 : 0 }) : originalMrp;
    return { rowNumber: source.rowNumber, fnsku, mrpDisplay, errors, duplicate: false, bits };
  });
  const counts = new Map<string, number>();
  rows.forEach(r => { if (r.fnsku) counts.set(r.fnsku, (counts.get(r.fnsku) || 0) + 1); });
  return rows.map(r => ({ ...r, duplicate: (counts.get(r.fnsku) || 0) > 1 }));
}
export function downloadTemplate() {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['FNSKU', 'MRP (₹)'], ...sampleRecords]), 'Labels');
  XLSX.writeFile(workbook, 'FNSKU_Label_Template.xlsx');
}
export function sampleSheet(): SheetData {
  return { fileName: 'FNSKU_MRP_Sample.xlsx', fileSize: 8640, headers: ['FNSKU', 'MRP (₹)'], fnskuColumn: 0, mrpColumn: 1, rows: sampleRecords.map((r, i) => ({ rowNumber: i + 2, cells: r.map(String), numeric: [false, true], formulas: [false, false], raw: r })) };
}
export const labelLayout = { width: 144, height: 72, x: 12, y: 10, barcodeWidth: 120, barcodeHeight: 30, fnskuY: 49, fnskuSize: 7.5, mrpY: 62, mrpSize: 9 };
export function barRects(bits: string) {
  const bars: { x: number; width: number }[] = [];
  const unit = labelLayout.barcodeWidth / bits.length;
  for (let i = 0; i < bits.length; i++) if (bits[i] === '1') {
    const start = i;
    while (i + 1 < bits.length && bits[i + 1] === '1') i++;
    bars.push({ x: labelLayout.x + start * unit, width: (i - start + 1) * unit });
  }
  return bars;
}
async function fontBase64(url: string) {
  const response = await fetch(url);
  if (!response.ok) throw new Error('Unable to load the label font. Please try again.');
  const bytes = new Uint8Array(await response.arrayBuffer());
  let binary = '';
  for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return btoa(binary);
}
let fonts: Promise<string[]> | undefined;
export async function generatePdf(rows: LabelRow[], progress: (n: number) => void) {
  if (!rows.length || rows.some(r => r.errors.length)) throw new Error('Resolve all invalid rows before generating labels.');
  fonts ??= Promise.all([fontBase64(regularFont), fontBase64(boldFont)]).catch(error => { fonts = undefined; throw error; });
  const [normal, bold] = await fonts;
  const pdf = new jsPDF({ orientation: 'landscape', unit: 'pt', format: [144, 72], compress: true, putOnlyUsedFonts: true });
  pdf.addFileToVFS('DejaVuSans.ttf', normal); pdf.addFont('DejaVuSans.ttf', 'Label', 'normal');
  pdf.addFileToVFS('DejaVuSans-Bold.ttf', bold); pdf.addFont('DejaVuSans-Bold.ttf', 'Label', 'bold');
  pdf.setProperties({ title: 'FNSKU Barcode Labels', creator: 'FNSKU Label Generator' });
  for (let i = 0; i < rows.length; i++) {
    if (i) pdf.addPage([144, 72], 'landscape');
    const row = rows[i];
    pdf.setFillColor(0, 0, 0);
    barRects(row.bits).forEach(bar => pdf.rect(bar.x, labelLayout.y, bar.width, labelLayout.barcodeHeight, 'F'));
    pdf.setFont('Label', 'normal'); pdf.setFontSize(labelLayout.fnskuSize);
    pdf.text(row.fnsku, 72, labelLayout.fnskuY, { align: 'center' });
    pdf.setFont('Label', 'bold'); pdf.setFontSize(labelLayout.mrpSize);
    pdf.text(`MRP: ₹${row.mrpDisplay}`, 72, labelLayout.mrpY, { align: 'center' });
    if (i % 25 === 0 || i === rows.length - 1) { progress(i + 1); await new Promise(resolve => setTimeout(resolve, 0)); }
  }
  return pdf.output('blob');
}
export function savePdf(blob: Blob) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a'); link.href = url; link.download = 'FNSKU_MRP_Barcodes.pdf'; link.click();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}
