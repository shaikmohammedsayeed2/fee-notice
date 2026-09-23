import * as XLSX from 'xlsx'
import { jsPDF } from 'jspdf'
import './styles.css'

// The app uses these browser-local libraries only. No spreadsheet is uploaded.
window.XLSX = XLSX
window.jspdf = { jsPDF }

import('./app.js')
