import { useMemo, useRef, useState, useEffect } from 'react';
import { ArrowDownToLine, ArrowLeft, ArrowRight, Check, CheckCheck, ChevronDown, ChevronLeft, ChevronRight, CircleCheck, CircleHelp, Download, ExternalLink, FileCheck2, FileSpreadsheet, FileText, FolderOpen, Info, Layers, Lightbulb, LoaderCircle, LockKeyhole, Menu, Plus, Printer, ScanBarcode, ShieldCheck, Sparkles, Upload, X, AlertTriangle } from 'lucide-react';
import { MAX_FILE_MB, barRects, demoLabel, downloadTemplate, generatePdf, labelLayout, parseFile, sampleSheet, savePdf, validateRows } from './lib/labels';
import type { LabelRow, SheetData } from './lib/labels';

type Modal = 'help' | 'template' | 'printing' | null;
function LabelPreview({ row }: { row: LabelRow }) {
  return <svg viewBox="0 0 144 72" className="label-svg" role="img" aria-label={`Code 128 barcode for ${row.fnsku} ${row.title}, MRP ₹${row.mrpDisplay}`}>
    <rect width="144" height="72" fill="white" />
    {barRects(row.bits).map((bar, i) => <rect key={i} x={bar.x} y={labelLayout.y} width={bar.width} height={labelLayout.barcodeHeight} fill="#111" />)}
    <text x="72" y={labelLayout.fnskuY} textAnchor="middle" fontFamily="LabelFont" fontSize={labelLayout.fnskuSize}>{row.fnsku} {row.title}</text>
    <text x="72" y={labelLayout.mrpY} textAnchor="middle" fontFamily="LabelFont" fontSize={labelLayout.mrpSize} fontWeight="700">MRP: ₹{row.mrpDisplay}</text>
  </svg>;
}

export default function App() {
  const [sheet, setSheet] = useState<SheetData | null>(null);
  const [fCol, setFCol] = useState(-1), [tCol, setTCol] = useState(-1), [mCol, setMCol] = useState(-1);
  const [mapping, setMapping] = useState(false), [mapped, setMapped] = useState(false);
  const [error, setError] = useState(''), [loading, setLoading] = useState(false), [drag, setDrag] = useState(false);
  const [acknowledged, setAcknowledged] = useState(false), [preview, setPreview] = useState(0);
  const [generating, setGenerating] = useState(false), [progress, setProgress] = useState(0), [pdf, setPdf] = useState<Blob | null>(null);
  const [modal, setModal] = useState<Modal>(null), [mobileNav, setMobileNav] = useState(false);
  const [showData, setShowData] = useState(false), [dataPage, setDataPage] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const modalRef = useRef<HTMLDivElement>(null);
  const rows = useMemo(() => sheet && mapped ? validateRows(sheet, fCol, tCol, mCol) : [], [sheet, mapped, fCol, tCol, mCol]);
  const validRows = useMemo(() => rows.filter(r => !r.errors.length), [rows]);
  const invalidRows = rows.filter(r => r.errors.length);
  const duplicateGroups = useMemo(() => {
    const groups = new Map<string, number[]>();
    rows.filter(r => r.duplicate).forEach(r => groups.set(r.fnsku, [...(groups.get(r.fnsku) || []), r.rowNumber]));
    return [...groups.entries()];
  }, [rows]);
  const ready = Boolean(!loading && mapped && validRows.length && !invalidRows.length && (!duplicateGroups.length || acknowledged));
  const currentLabel = validRows[preview] || validRows[0] || demoLabel;
  const step = pdf ? 4 : ready ? 3 : sheet ? 2 : 1;
  useEffect(() => {
    if (!modal) return;
    const previous = document.activeElement as HTMLElement;
    modalRef.current?.focus();
    const keydown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setModal(null);
      if (e.key === 'Tab') {
        const focusables = modalRef.current?.querySelectorAll<HTMLElement>('button, a, input, select, [tabindex="0"]');
        if (!focusables?.length) return;
        const first = focusables[0], last = focusables[focusables.length - 1];
        if (e.shiftKey && (document.activeElement === first || document.activeElement === modalRef.current)) { e.preventDefault(); last.focus(); }
        else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
      }
    };
    document.addEventListener('keydown', keydown);
    return () => { document.removeEventListener('keydown', keydown); previous?.focus(); };
  }, [modal]);
  function reset() {
    setSheet(null); setMapped(false); setMapping(false); setFCol(-1); setTCol(-1); setMCol(-1); setPdf(null); setError(''); setPreview(0); setAcknowledged(false); setShowData(false); setDataPage(0);
    if (inputRef.current) inputRef.current.value = '';
  }
  function loadSheet(data: SheetData) {
    setSheet(data); setFCol(data.fnskuColumn); setTCol(data.titleColumn); setMCol(data.mrpColumn);
    const detected = data.fnskuColumn >= 0 && data.titleColumn >= 0 && data.mrpColumn >= 0;
    setMapped(detected); setMapping(!detected); setPreview(0); setAcknowledged(false); setPdf(null); setError(''); setDataPage(0);
  }
  async function upload(file?: File) {
    if (!file || generating || loading) return;
    setLoading(true); setError('');
    try { loadSheet(await parseFile(file)); }
    catch (e) { setError(e instanceof Error ? e.message : 'Unable to read this file. Please upload a valid Excel spreadsheet.'); }
    finally { setLoading(false); if (inputRef.current) inputRef.current.value = ''; }
  }
  async function generate() {
    if (!ready) return;
    setGenerating(true); setProgress(0); setError('');
    try { setPdf(await generatePdf(validRows, setProgress)); }
    catch (e) { setError(e instanceof Error ? e.message : 'PDF generation failed. Please try again.'); }
    finally { setGenerating(false); }
  }
  function openModal(value: Modal) { setModal(value); setMobileNav(false); }

  return <div className="app-shell">
    <aside className={`sidebar ${mobileNav ? 'nav-open' : ''}`}>
      <a className="brand" href="#" onClick={e => { e.preventDefault(); if (!generating) reset(); }} aria-label="FNSKU Label Generator home">
        <span className="brand-icon"><ScanBarcode size={25} strokeWidth={1.8} /></span><span><strong>FNSKU<span className="brand-dot">.</span></strong><small>LABEL GENERATOR</small></span>
      </a>
      <div className="workspace-label">WORKSPACE</div>
      <nav aria-label="Main navigation">
        <button className="nav-item active" onClick={() => { setMobileNav(false); setModal(null); }}><ScanBarcode size={19} /> Label generator <span className="nav-active-dot" /></button>
        <button className="nav-item" onClick={() => openModal('template')}><FileSpreadsheet size={19} /> Excel template <ArrowDownToLine size={14} className="nav-end" /></button>
        <button className="nav-item" onClick={() => openModal('printing')}><Printer size={19} /> Printing guide <ExternalLink size={13} className="nav-end" /></button>
      </nav>
      <div className="sidebar-bottom">
        <div className="sidebar-tip"><span className="tip-icon"><Lightbulb size={19} /></span><h3>Less prep. More shipping.</h3><p>From spreadsheet to ready-to-print labels in a few clicks.</p><button onClick={() => openModal('help')}>See how it works <ArrowRight size={14} /></button></div>
        <button className="help-nav" onClick={() => openModal('help')}><CircleHelp size={18} /> Help & resources <ChevronRight size={15} /></button>
        <div className="sidebar-version"><span className="tiny-brand"><ScanBarcode size={13} /> Built for Amazon sellers</span><span>v1.0</span></div>
      </div>
    </aside>

    <div className="main-shell">
      <header className="topbar"><div className="breadcrumb"><button className="mobile-menu icon-button" aria-label="Toggle navigation" onClick={() => setMobileNav(!mobileNav)}><Menu size={21} /></button><span>Workspace</span><ChevronRight size={14} /><strong>Label generator</strong></div><div className="topbar-right"><span className="local-status"><span /> All processing stays on your device</span><span className="topbar-divider" /><button onClick={() => openModal('help')}><CircleHelp size={17} /> Help</button></div></header>
      <main>
        <div className="page-heading"><div><div className="eyebrow">YOUR LABELS, SIMPLIFIED</div><h1>Generate FNSKU labels<span>.</span></h1><p>Turn your Excel sheet into accurate, print-ready Amazon labels.</p></div><button className="button secondary sample-button" disabled={generating || loading} onClick={() => loadSheet(sampleSheet())}><Sparkles size={16} /> Try a sample file <ArrowUpRightIcon /></button></div>

        <div className="steps" aria-label="Generation progress">{[{ title: 'Upload Excel', sub: 'Add your product data', icon: Upload }, { title: 'Validate data', sub: 'Check for any issues', icon: ShieldCheck }, { title: 'Preview labels', sub: 'Make sure it looks right', icon: ScanBarcode }, { title: 'Download PDF', sub: 'Ready for your printer', icon: Download }].map((item, i) => <div key={item.title} className={`step ${step === i + 1 ? 'current' : ''} ${step > i + 1 ? 'complete' : ''}`} aria-current={step === i + 1 ? 'step' : undefined}><span className="step-number">{step > i + 1 ? <Check size={17} /> : i + 1}</span><div><strong>{item.title}</strong><small>{item.sub}</small></div>{i < 3 && <ChevronRight className="step-chevron" size={16} />}</div>)}</div>

        {error && <div className="alert error" role="alert"><AlertTriangle size={18} /><span>{error}</span><button className="icon-button" aria-label="Dismiss error" onClick={() => setError('')}><X size={16} /></button></div>}
        <div className="content-grid">
          <div className="left-column">
            <section className="card upload-card">
              <div className="card-heading"><div className="heading-icon"><Upload size={18} /></div><div><h2>{sheet ? 'Your Excel file' : 'Upload your Excel file'}</h2><p>{sheet ? 'Your product data, ready for review.' : 'Start with the products you want to label.'}</p></div><span className="small-tag">STEP 01</span></div>
              <input type="file" ref={inputRef} accept=".xlsx,.xls" aria-label="Upload Excel spreadsheet" className="sr-only" onChange={e => upload(e.target.files?.[0])} />
              {!sheet ? <div className={`dropzone ${drag ? 'dragging' : ''}`} onDragOver={e => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)} onDrop={e => { e.preventDefault(); setDrag(false); upload(e.dataTransfer.files[0]); }}>
                <div className="upload-illustration"><div className="file-back" /><div className="file-front"><FileSpreadsheet size={30} strokeWidth={1.5} /><span className="upload-mini"><Plus size={12} strokeWidth={3} /></span></div></div>
                <h3>{loading ? 'Reading your spreadsheet…' : drag ? 'Drop it here. We’ll take it from there.' : 'Drag & drop your Excel file here'}</h3><p>or choose a file from your computer</p>
                <button className="button primary" disabled={loading} onClick={() => inputRef.current?.click()}>{loading ? <LoaderCircle className="spin" size={16} /> : <FolderOpen size={16} />} {loading ? 'Processing file…' : 'Choose Excel file'}</button>
                <small>.xlsx or .xls <span>•</span> Up to {MAX_FILE_MB} MB</small>
              </div> : <div className="uploaded-area"><div className="file-summary"><span className="excel-file-icon"><FileSpreadsheet size={25} /></span><div><strong>{sheet.fileName}</strong><span>{(sheet.fileSize / 1024).toFixed(1)} KB <i>·</i> {sheet.rows.length.toLocaleString()} rows detected <i>·</i> First worksheet</span></div><button className="icon-button" disabled={generating} onClick={reset} aria-label="Remove uploaded file"><X size={17} /></button></div>
                <div className="detected-columns"><div><span>FNSKU column</span><strong>{mapped ? <><Check size={14} />{sheet.headers[fCol]}</> : 'Selection required'}</strong></div><div><span>Title column</span><strong>{mapped ? <><Check size={14} />{sheet.headers[tCol]}</> : 'Selection required'}</strong></div><div><span>MRP column</span><strong>{mapped ? <><Check size={14} />{sheet.headers[mCol]}</> : 'Selection required'}</strong></div></div>
                <div className="file-actions"><span><ShieldCheck size={14} /> File processed locally</span><button disabled={generating} onClick={() => { setMapping(true); setMapped(false); setPdf(null); }}>Map columns <ChevronDown size={13} /></button><button disabled={generating || loading} onClick={() => inputRef.current?.click()}>Replace file</button></div>
                {mapping && <div className="mapping-panel"><h3>Column mapping</h3><p>Select the correct columns. Missing or ambiguous headers need your confirmation.</p><div className="mapping-fields"><label>FNSKU<select value={fCol} onChange={e => setFCol(Number(e.target.value))}><option value={-1}>Select column</option>{sheet.headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`} ({i + 1})</option>)}</select></label><label>Title<select value={tCol} onChange={e => setTCol(Number(e.target.value))}><option value={-1}>Select column</option>{sheet.headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`} ({i + 1})</option>)}</select></label><label>MRP<select value={mCol} onChange={e => setMCol(Number(e.target.value))}><option value={-1}>Select column</option>{sheet.headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`} ({i + 1})</option>)}</select></label></div><button className="button primary small" disabled={fCol < 0 || tCol < 0 || mCol < 0 || new Set([fCol, tCol, mCol]).size < 3} onClick={() => { setMapped(true); setMapping(false); setAcknowledged(false); setPreview(0); }}>Apply mapping <Check size={14} /></button></div>}
              </div>}
              <div className="upload-footer"><LockKeyhole size={13} /><span>Your files never leave your browser. Private by design.</span></div>
            </section>

            {sheet && mapped ? <section className="card validation-card"><div className="section-header"><h2><ShieldCheck size={17} /> Data validation</h2><span className={`pill ${invalidRows.length ? 'red' : duplicateGroups.length ? 'amber' : 'green'}`}>{invalidRows.length ? 'Needs attention' : duplicateGroups.length ? 'Review duplicates' : 'Checks passed'}</span></div><div className="validation-stats"><div><strong>{rows.length}</strong><span>Total rows</span></div><div><strong className="green-text">{validRows.length}</strong><span>Valid rows</span></div><div><strong className={invalidRows.length ? 'red-text' : ''}>{invalidRows.length}</strong><span>Invalid rows</span></div><div><strong>{duplicateGroups.length}</strong><span>Duplicates</span></div></div>
                {(invalidRows.length > 0 || duplicateGroups.length > 0) && <div className="issue-list">{invalidRows.map(r => <p key={r.rowNumber} className="row-error"><AlertTriangle size={14} /><span><strong>Excel row {r.rowNumber}:</strong> {r.errors.join('; ')}</span></p>)}{duplicateGroups.map(([sku, nums]) => <p key={sku} className="row-warning"><AlertTriangle size={14} /><span><strong>{sku}</strong> — duplicate in Excel rows {nums.join(', ')}</span></p>)}</div>}
                {duplicateGroups.length > 0 && <label className="duplicate-confirm"><input type="checkbox" checked={acknowledged} disabled={generating} onChange={e => { setAcknowledged(e.target.checked); setPdf(null); }} />I understand the duplicates are intentional. Include every row.</label>}
                <div className="validation-bottom"><span>{invalidRows.length ? 'Fix invalid rows in Excel, then replace your file.' : <><CircleCheck size={14} /> {validRows.length} labels {duplicateGroups.length && !acknowledged ? 'awaiting confirmation' : 'ready to generate'}</>}</span><button onClick={() => setShowData(!showData)}>{showData ? 'Hide data' : 'Review data'} <ChevronDown size={13} /></button></div>
              </section> : <section className="card requirements-card"><div className="section-header"><h2><FileSpreadsheet size={17} /> A simple sheet is all you need</h2><button className="text-button" onClick={downloadTemplate}><Download size={14} /> Get template</button></div><div className="requirements-content"><div className="requirements-copy"><p>Include these three columns in your first worksheet.</p><span><CircleCheck size={14} /> Headers are detected automatically</span><span><CircleCheck size={14} /> One row becomes one label</span></div><div className="mini-sheet"><div className="sheet-row sheet-head"><span /><span>A</span><span>B</span><span>C</span></div><div className="sheet-row"><span>1</span><strong>FNSKU</strong><strong>Title</strong><strong>MRP (₹)</strong></div><div className="sheet-row"><span>2</span><span>X001ABC123</span><span>Stainless Steel Bottle</span><span>499</span></div><div className="sheet-row"><span>3</span><span>X001ABC124</span><span>Organic Honey 500g</span><span>799</span></div></div></div></section>}
          </div>

          <div className="right-column">
            <section className="card preview-card"><div className="section-header"><h2><ScanBarcode size={18} /> Label preview</h2><span className={`pill ${validRows.length ? 'green' : 'neutral'}`}><span className="pill-dot" />{validRows.length ? 'Live preview' : 'Sample preview'}</span></div>
              <div className="preview-stage"><div className="dimension-width"><span /><span>2 inches</span><span /></div><div className="label-with-height"><div className="paper-label"><LabelPreview row={currentLabel} /></div><div className="dimension-height"><span /><span>1 inch</span><span /></div></div><div className="preview-caption"><span />Actual size: 2″ × 1″ <i>·</i> Shown enlarged</div></div>
              <div className="preview-bottom">{validRows.length ? <><span>Label <strong>{Math.min(preview + 1, validRows.length)}</strong> of <strong>{validRows.length}</strong></span><div className="preview-nav"><button aria-label="Previous label" className="icon-button" disabled={preview === 0} onClick={() => setPreview(p => p - 1)}><ChevronLeft size={17} /></button><button aria-label="Next label" className="icon-button" disabled={preview >= validRows.length - 1} onClick={() => setPreview(p => p + 1)}><ChevronRight size={17} /></button></div></> : <><Info size={14} /><span>Upload a file to preview your own labels</span></>}</div>
            </section>
            <section className="card format-card"><div className="section-header"><h2><Layers size={17} /> Label format</h2><span className="fixed-tag"><LockKeyhole size={11} /> Fixed format</span></div><div className="format-details"><div><span>Label size</span><strong>2″ × 1″ <small>(50.8 × 25.4 mm)</small></strong></div><div><span>Barcode type</span><strong>Code 128</strong></div><div><span>PDF layout</span><strong>One label per page</strong></div></div><div className="thermal-note"><Printer size={14} /><span>Made for thermal label printers</span><CircleCheck size={14} /></div></section>
          </div>
        </div>

        {showData && sheet && <section className="card data-card"><div className="section-header"><h2><FileText size={17} /> Spreadsheet review</h2><span>Original Excel row order</span></div><div className="data-scroll"><table><thead><tr><th>Excel row</th><th>FNSKU</th><th>Title</th><th>MRP</th><th>Status</th></tr></thead><tbody>{rows.slice(dataPage * 10, dataPage * 10 + 10).map(r => <tr key={r.rowNumber}><td>{r.rowNumber}</td><td className="mono">{r.fnsku || '—'}</td><td>{r.title || '—'}</td><td>₹{r.mrpDisplay}</td><td><span className={`pill ${r.errors.length ? 'red' : r.duplicate ? 'amber' : 'green'}`}>{r.errors.length ? 'Invalid' : r.duplicate ? 'Duplicate' : 'Valid'}</span></td></tr>)}</tbody></table></div><div className="table-pagination"><span>{dataPage * 10 + 1}–{Math.min(dataPage * 10 + 10, rows.length)} of {rows.length} rows</span><div><button className="icon-button" aria-label="Previous data page" disabled={!dataPage} onClick={() => setDataPage(p => p - 1)}><ChevronLeft size={17} /></button><button className="icon-button" aria-label="Next data page" disabled={(dataPage + 1) * 10 >= rows.length} onClick={() => setDataPage(p => p + 1)}><ChevronRight size={17} /></button></div></div></section>}

        <section className={`generate-bar ${pdf ? 'pdf-ready' : ''}`} aria-live="polite"><div className="generate-info"><span className="generate-icon">{pdf ? <CheckCheck size={22} /> : <FileCheck2 size={22} />}</span><div><h3>{pdf ? 'Your PDF is ready!' : generating ? 'Creating your print-ready labels…' : ready ? `${validRows.length} labels, ready to print.` : 'Your next batch starts here.'}</h3><p>{pdf ? `${validRows.length} labels generated · FNSKU_MRP_Barcodes.pdf` : generating ? `${progress} / ${validRows.length} labels generated` : ready ? 'Validated, correctly sized, and in your original row order.' : sheet ? 'Review your data and resolve any issues to continue.' : 'Upload an Excel file to generate your labels.'}</p></div></div><div className="generate-actions">{pdf && <button className="text-button" onClick={reset}><Plus size={15} /> New batch</button>}{!pdf && <span className="output-hint">2″ × 1″ PDF <span>·</span> Code 128</span>}<button className="button primary generate-button" disabled={generating || (!pdf && !ready)} onClick={() => pdf ? savePdf(pdf) : generate()}>{generating ? <LoaderCircle className="spin" size={17} /> : <Download size={17} />}{pdf ? 'Download PDF' : generating ? 'Generating…' : 'Generate PDF'}{!generating && <ArrowRight size={16} />}</button></div>{generating && <div className="generation-progress" style={{ width: `${(progress / validRows.length) * 100}%` }} />}</section>
        <footer className="page-footer"><span><ShieldCheck size={14} /> Private. Precise. Print-ready.</span><span>Built for your workflow, not your data. <button onClick={() => openModal('help')}>Learn more <ArrowRight size={12} /></button></span></footer>
      </main>
    </div>

    {modal && <div className="modal-backdrop" onClick={e => { if (e.target === e.currentTarget) setModal(null); }}><div className="modal" role="dialog" aria-modal="true" aria-labelledby="modal-title" tabIndex={-1} ref={modalRef}><div className="modal-top"><span className="modal-icon">{modal === 'template' ? <FileSpreadsheet size={23} /> : modal === 'printing' ? <Printer size={23} /> : <CircleHelp size={23} />}</span><button className="icon-button" aria-label="Close dialog" onClick={() => setModal(null)}><X size={21} /></button></div><h2 id="modal-title">{modal === 'template' ? 'A head start for your next batch.' : modal === 'printing' ? 'Small labels. Perfect prints.' : 'From spreadsheet to shipping.'}</h2><p className="modal-intro">{modal === 'template' ? 'Use our ready-to-go Excel template. Just replace the sample products with your own.' : modal === 'printing' ? 'A few quick settings for accurate 2″ × 1″ thermal labels.' : 'Everything you need to generate your Amazon FNSKU labels, without sending your data anywhere.'}</p>
      {modal === 'template' ? <><div className="modal-table"><table><thead><tr><th>FNSKU</th><th>Title</th><th>MRP (₹)</th></tr></thead><tbody><tr><td>X001ABC123</td><td>Stainless Steel Bottle</td><td>499</td></tr><tr><td>X001ABC124</td><td>Organic Honey 500g</td><td>799</td></tr><tr><td>X001ABC125</td><td>Cotton T-Shirt Large</td><td>1299</td></tr></tbody></table></div><div className="modal-note"><Info size={17} /><p>Keep FNSKU cells formatted as <strong>Text</strong>. Enter prices as numbers or currency amounts, with up to two decimal places. Only the first worksheet is read.</p></div><button className="button primary modal-cta" onClick={downloadTemplate}><Download size={17} /> Download Excel template</button></> : modal === 'printing' ? <><div className="guide-list"><div><span>1</span><div><h3>Load your 2″ × 1″ label roll</h3><p>Select your thermal printer and set the paper size to 50.8 × 25.4 mm, landscape.</p></div></div><div><span>2</span><div><h3>Print at actual size (100%)</h3><p>Turn off “Fit to page” and scaling. The PDF already contains one correctly sized label per page.</p></div></div><div><span>3</span><div><h3>Test one label before the full batch</h3><p>Check that the barcode is sharp, the quiet zones are clear, and the FNSKU, title, and ₹ price are readable. Test with a barcode scanner.</p></div></div></div><div className="modal-note"><Printer size={17} /><p>Use a 203 DPI or higher thermal printer. Adjust darkness and speed in your printer settings if bars appear faint or bleed together.</p></div></> : <><div className="guide-list"><div><span>1</span><div><h3>Upload & validate</h3><p>Choose an .xlsx or .xls file up to {MAX_FILE_MB} MB. We check missing values, invalid prices, and duplicate FNSKUs before you print.</p></div></div><div><span>2</span><div><h3>Preview every label</h3><p>Use the preview arrows to check your barcode and price. Duplicate rows are only included after your explicit confirmation.</p></div></div><div><span>3</span><div><h3>Generate & download</h3><p>Download vector Code 128 barcodes, one 2″ × 1″ label per PDF page, in the same order as your Excel rows.</p></div></div></div><details><summary>How are columns detected?</summary><p>Headers are trimmed and matched case-insensitively. FNSKU accepts “FNSKU” or “FNSKU Code”. Title accepts “Title”, “Product Title”, “Name”, or “Item Title”. MRP accepts “MRP”, “MRP (₹)”, or “Maximum Retail Price”. Multiple matches or missing headers require manual mapping. Empty rows are ignored. Formulas must be replaced with values.</p></details><div className="modal-note"><LockKeyhole size={17} /><p>Your spreadsheet is processed entirely in this browser. No uploads, accounts, or permanent storage. Closing the page clears your batch.</p></div></>}
      <button className="modal-close-link" onClick={() => setModal(null)}><ArrowLeft size={14} /> Back to label generator</button></div></div>}
  </div>;
}
function ArrowUpRightIcon() { return <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round"><path d="M7 17 17 7M7 7h10v10" /></svg>; }
