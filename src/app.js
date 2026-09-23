const FIELDS = [['studentName', 'Student name', true], ['guardian', 'Father / guardian', true], ['idNo', 'ID no.', true], ['className', 'Class', true], ['oldBalance', 'Old balance'], ['fees', 'Fees'], ['books', 'Books'], ['others', 'Others']]
const sample = { studentName: 'Aarav Sharma', guardian: 'Rajesh Sharma', idNo: 'S-10234', className: 'VI – A', oldBalance: 1200, fees: 8000, books: 1500, others: 300 }
const aliases = { studentName: ['student name', 'name', 'student'], guardian: ['father', 'guardian', 'parent'], idNo: ['id no', 'id', 'admission'], className: ['class', 'grade'], oldBalance: ['old balance', 'balance'], fees: ['fees', 'fee', 'tuition'], books: ['books', 'book'], others: ['others', 'other'] }
const state = { headers: [], rows: [], mapping: {}, fileName: '' }
const money = value => new Intl.NumberFormat('en-IN', { maximumFractionDigits: 2 }).format(Number(value) || 0)
const clean = value => String(value ?? '').trim().toLowerCase().replace(/[^a-z0-9]/g, '')
const notice = (record, compact = false) => { const total = ['oldBalance', 'fees', 'books', 'others'].reduce((sum, field) => sum + (Number(record[field]) || 0), 0); return `<article class="notice ${compact ? 'compact' : ''}"><div class="notice-title">ANNOOR FOUNDATION SCHOOL</div><div class="notice-heading">FEES - NOTICE</div><div class="identity"><div>STUDENT NAME <span>${record.studentName || '—'}</span></div><div>FATHER/GUARDIAN <span>${record.guardian || '—'}</span></div><div><b>ID NO</b> <span>${record.idNo || '—'}</span><b class="class-label">CLASS</b><span>${record.className || '—'}</span></div></div><div class="particulars">Particulars - Balance Amounts</div>${[['OLD BALANCE','oldBalance'],['FEES','fees'],['BOOKS','books'],['OTHERS','others']].map(([label, field]) => `<div class="fee-row"><span>${label}</span><strong>₹ ${money(record[field])}</strong></div>`).join('')}<div class="total"><span>TOTAL</span><strong>₹ ${money(total)}</strong></div><div class="notice-foot">PARENTS ARE REQUESTED TO PAY<br>THE AMOUNT BY 27/09/2026.</div><div class="contact">FOR DETAILS CONTACT :&nbsp; 7207506400</div></article>` }
const records = () => state.rows.map(row => Object.fromEntries(FIELDS.map(([key]) => [key, row[state.mapping[key]] ?? ''])))
function render() { const mapped = records(); const first = mapped[0] || sample; const options = header => `<option value="">Not mapped</option>${state.headers.map(h => `<option ${state.mapping[header] === h ? 'selected' : ''} value="${h}">${h}</option>`).join('')}`; document.querySelector('#root').innerHTML = `<main><header><div class="brand-mark">▰</div><h1>Fee Notice Studio</h1><span class="pipe"></span><span class="local">Local-only</span><span class="header-note">Your records stay on this computer</span></header><section class="workspace"><aside class="control-panel"><section><div class="step"><i>1</i><h2>Upload student data</h2></div><p>Upload an Excel file. Its first row should contain column names.</p><label class="drop-zone" id="drop-zone"><input id="file-input" type="file" accept=".xlsx,.xls,.csv"><span class="excel">X</span><strong>Drop your Excel file here</strong><em>or choose a file</em><button type="button" id="choose-file">Choose file</button><small>Supported: .xlsx, .xls, .csv</small></label>${state.fileName ? `<div class="file-row"><span>▧</span><div><strong>${state.fileName}</strong><small>${state.rows.length} student records detected</small></div><b>✓</b></div>` : ''}</section><section><div class="step"><i>2</i><h2>Map columns</h2></div><p>Match spreadsheet headers to the fields on each notice.</p><div class="mapping">${FIELDS.map(([key, label, required]) => `<label>${label}${required ? '<sup>*</sup>' : ''}<select data-field="${key}">${options(key)}</select></label>`).join('')}</div></section><section class="export"><div class="step"><i>3</i><h2>Export</h2></div><p>Creates a print-ready A4 PDF with four notices on every page.</p><button class="primary" id="export-pdf">Download PDF <span>↓</span></button></section></aside><section class="preview-area"><div class="preview-head"><div><h2>Notice preview</h2><p>${state.fileName ? 'First record from your upload' : 'Example notice — upload a file to personalise it'}</p></div><span>4 notices / A4</span></div>${notice(first)}<div class="sheet-preview"><div><h3>A4 print layout</h3><p>Four notices per page</p></div><div class="mini-grid">${[0,1,2,3].map(i => notice(mapped[i] || first, true)).join('')}</div></div></section></section><footer>All processing happens in your browser. No Excel data is uploaded.</footer></main>`; bind() }
async function loadFile(file) { if (!file) return; if (!window.XLSX) return alert('The spreadsheet reader is still loading. Please try again in a moment.'); const book = XLSX.read(await file.arrayBuffer(), { type: 'array' }); const data = XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]], { defval: '' }); state.headers = Object.keys(data[0] || {}); state.rows = data; state.mapping = Object.fromEntries(FIELDS.map(([key]) => [key, state.headers.find(header => aliases[key].some(a => clean(header).includes(clean(a)))) || ''])); state.fileName = file.name; render() }
function bind() { const input = document.querySelector('#file-input'); const zone = document.querySelector('#drop-zone'); document.querySelector('#choose-file').onclick = e => { e.preventDefault(); input.click() }; input.onchange = () => loadFile(input.files[0]); ['dragover','dragleave','drop'].forEach(type => zone.addEventListener(type, e => { e.preventDefault(); zone.classList.toggle('dragging', type === 'dragover') })); zone.addEventListener('drop', e => loadFile(e.dataTransfer.files[0])); document.querySelectorAll('[data-field]').forEach(select => select.onchange = () => { state.mapping[select.dataset.field] = select.value; render() }); document.querySelector('#export-pdf').onclick = exportPDF }
function drawNotice(doc, x, y, w, h, r) {
  const total = ['oldBalance', 'fees', 'books', 'others'].reduce((sum, field) => sum + (Number(r[field]) || 0), 0)
  const px = value => x + value * w
  const py = value => y + value * h
  const rule = value => doc.line(x, py(value), x + w, py(value))
  const amount = value => `Rs. ${money(value)}`

  doc.setTextColor(0)
  doc.setDrawColor(38)
  doc.setLineWidth(0.18)
  doc.rect(x, y, w, h)

  doc.setFont('times', 'bold')
  doc.setFontSize(h * 0.030)
  doc.text('ANNOOR FOUNDATION SCHOOL', px(0.5), py(0.060), { align: 'center' })
  rule(0.090)
  doc.setFontSize(h * 0.026)
  doc.text('FEES - NOTICE', px(0.5), py(0.140), { align: 'center' })
  rule(0.180)

  doc.setFont('times', 'normal')
  doc.setFontSize(h * 0.020)
  doc.text('STUDENT NAME', px(0.04), py(0.240))
  doc.text(String(r.studentName || '—'), px(0.48), py(0.240))
  doc.text('FATHER/GUARDIAN', px(0.04), py(0.315))
  doc.text(String(r.guardian || '—'), px(0.48), py(0.315))
  doc.text('ID NO', px(0.04), py(0.390))
  doc.text(String(r.idNo || '—'), px(0.25), py(0.390))
  doc.text('CLASS', px(0.58), py(0.390))
  doc.text(String(r.className || '—'), px(0.78), py(0.390))

  doc.setFont('times', 'bold')
  doc.setFontSize(h * 0.022)
  doc.text('Particulars - Balance Amounts', px(0.5), py(0.475), { align: 'center' })
  rule(0.505)

  const items = [['OLD BALANCE', r.oldBalance], ['FEES', r.fees], ['BOOKS', r.books], ['OTHERS', r.others]]
  doc.setFont('times', 'normal')
  doc.setFontSize(h * 0.019)
  items.forEach(([label, value], index) => {
    const baseline = 0.560 + index * 0.070
    doc.text(label, px(0.38), py(baseline), { align: 'right' })
    doc.text(amount(value), px(0.96), py(baseline), { align: 'right' })
    rule(baseline + 0.031)
  })

  doc.setFont('times', 'bold')
  doc.setFontSize(h * 0.022)
  doc.text('TOTAL', px(0.42), py(0.865), { align: 'right' })
  doc.text(amount(total), px(0.96), py(0.865), { align: 'right' })
  rule(0.890)
  doc.setFontSize(h * 0.016)
  doc.text(['PARENTS ARE REQUESTED TO PAY', 'THE AMOUNT BY 27/09/2026.'], px(0.5), py(0.925), { align: 'center', lineHeightFactor: 1.05 })
  doc.setFont('times', 'normal')
  doc.setFontSize(h * 0.015)
  doc.text('FOR DETAILS CONTACT : 7207506400', px(0.5), py(0.978), { align: 'center' })
}
function exportPDF() { if (!window.jspdf) return alert('The PDF generator is still loading. Please try again in a moment.'); const doc = new jspdf.jsPDF({orientation:'portrait',unit:'mm',format:'a4'}); const gap=5,margin=8,w=(210-margin*2-gap)/2,h=(297-margin*2-gap)/2, source=records().length ? records() : [sample]; source.forEach((r,i)=>{if(i&&i%4===0)doc.addPage();const slot=i%4;drawNotice(doc,margin+(slot%2)*(w+gap),margin+Math.floor(slot/2)*(h+gap),w,h,r)});doc.save('fee-notices.pdf') }
render()
