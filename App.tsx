
import React, { useState, useCallback, useMemo, useEffect } from 'react';
import { FileUp, FileSpreadsheet, Trash2, AlertCircle, Printer, CheckCircle2, Settings2, Maximize2, Type, FileDown, Zap, Smartphone, Monitor, Loader2 } from 'lucide-react';
import { ExcelRow, PageGroup } from './types';

declare var XLSX: any;
declare var html2pdf: any;

const App: React.FC = () => {
  const [pages, setPages] = useState<PageGroup[]>([]);
  const [allDataRows, setAllDataRows] = useState<ExcelRow[]>([]);
  const [tableHeaders, setTableHeaders] = useState<ExcelRow | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Layout State with updated defaults: Font 11pt, Row 33px, Specific Col Widths
  const [orientation, setOrientation] = useState<'portrait' | 'landscape'>('portrait');
  const [globalRowHeight, setGlobalRowHeight] = useState<number>(33);
  const [globalFontSize, setGlobalFontSize] = useState<number>(11);
  const [colWidths, setColWidths] = useState<number[]>([40, 139, 316, 130, 78, 74, 80, 90, 190]); 

  // Constants
  const PRINT_COLS_START = 4;
  const COL_COUNT = 9;
  
  // Fixed Row Limit per Page
  const maxRowsPerPage = 41;

  // Calculate total table width
  const totalTableWidth = useMemo(() => colWidths.reduce((acc, curr) => acc + curr, 0), [colWidths]);

  // Grouping Logic - Refined to strictly follow the 4-column break rule
  useEffect(() => {
    if (allDataRows.length === 0) return;

    const groups: PageGroup[] = [];
    let currentChunk: ExcelRow[] = [];
    let activeSiteHeader: any = null;
    let prevKey = "";

    allDataRows.forEach((row, index) => {
      // The "4 Column Rule": Detect changes in Site Name, SDN ID, SDN Type, and Location
      const keyA = String(row[0] ?? "");
      const keyB = String(row[1] ?? "");
      const keyC = String(row[2] ?? "");
      const keyD = String(row[3] ?? "");
      const currentKey = `${keyA}|${keyB}|${keyC}|${keyD}`;

      const rowSiteHeader = {
        siteName: keyA,
        sdnId: keyB,
        sdnType: keyC,
        location: keyD,
        sdnDate: String(row[13] ?? ""),
        reqDate: String(row[14] ?? ""),
        szsPob: String(row[15] ?? ""),
        clientPob: String(row[16] ?? ""),
      };

      // Initialization for first row
      if (index === 0) {
        activeSiteHeader = rowSiteHeader;
        prevKey = currentKey;
      }

      const siteChanged = currentKey !== prevKey;
      const limitReached = currentChunk.length >= maxRowsPerPage;

      // Trigger break if 4-column key changes OR row limit is reached
      if (siteChanged || limitReached) {
        // Prevent pushing empty pages (common if site change happens exactly at row limit)
        if (currentChunk.length > 0) {
          groups.push({
            id: `page-${groups.length}`,
            rows: currentChunk,
            rowMetas: [], 
            siteHeader: { ...activeSiteHeader } 
          });
          currentChunk = [];
        }
        
        // If the site actually changed, update the active header for the next page
        if (siteChanged) {
          activeSiteHeader = rowSiteHeader;
          prevKey = currentKey;
        }
      }

      currentChunk.push(row);
    });

    // Final chunk push
    if (currentChunk.length > 0) {
      groups.push({
        id: `page-${groups.length}`,
        rows: currentChunk,
        rowMetas: [],
        siteHeader: { ...activeSiteHeader }
      });
    }

    setPages(groups);
  }, [allDataRows, maxRowsPerPage]);

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);
    setFileName(file.name);
    setPages([]);
    setAllDataRows([]);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buffer = e.target?.result;
        const workbook = XLSX.read(buffer, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data: ExcelRow[] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        if (data.length > 0) {
          setTableHeaders(data[0]);
          setAllDataRows(data.slice(1).filter(row => row && row.length > 0));
        } else {
          setError("The uploaded file is empty.");
        }
      } catch (err) {
        setError("Could not parse Excel file. Ensure it is a valid .xlsx file.");
      } finally {
        setLoading(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handlePrint = useCallback(() => {
    if (pages.length === 0) return;
    window.print();
  }, [pages.length]);

  /**
   * CORRECTED PDF GENERATION FUNCTION
   * Fixed: Blank pages, capture area, and added margins
   */
  const handleDownloadPDF = async () => {
    const element = document.getElementById('printArea');
    if (!element || pages.length === 0) return;
    
    if (typeof html2pdf === 'undefined') {
      setError("PDF library missing. Please check your internet connection.");
      return;
    }

    setIsExporting(true);
    
    // Construct dynamic filename
    let dynamicName = 'SDN_Report';
    if (pages.length > 0) {
      const { siteName, sdnId } = pages[0].siteHeader;
      const safeSiteName = (siteName || 'Site').replace(/[^a-z0-9]/gi, '_');
      dynamicName = `${safeSiteName}_${sdnId || 'SDN'}`;
    }

    // PDF configuration with increased Top Margin (20mm instead of 12mm)
    const opt = {
      margin: [20, 10, 10, 10], // [top, left, bottom, right] in mm
      filename: `${dynamicName}.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { 
        scale: 2, 
        useCORS: true, 
        backgroundColor: '#ffffff',
        width: element.scrollWidth,
        scrollY: 0,
        scrollX: 0,
        letterRendering: true
      },
      jsPDF: { 
        unit: 'mm', 
        format: 'a4', 
        orientation: orientation,
        compress: true
      },
      pagebreak: { 
        mode: ['css', 'legacy'], 
        before: '.site-group' // Matches the manual site-change page breaks
      }
    };

    try {
      // Use the Promise API to ensure rendering completes before generation
      const worker = html2pdf().set(opt).from(element);
      await worker.save();
    } catch (err) {
      console.error("PDF Export Error:", err);
      setError("PDF export failed. Please use 'Direct Print' as an alternative.");
    } finally {
      setIsExporting(false);
    }
  };

  const autoAdjustAllWidths = () => {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.font = `bold ${globalFontSize}pt Calibri, Arial, sans-serif`;

    const newWidths = colWidths.map((_, colIdx) => {
      const dataIdx = colIdx + PRINT_COLS_START;
      let maxWidth = tableHeaders ? ctx.measureText(String(tableHeaders[dataIdx])).width : 0;
      pages.forEach(p => p.rows.forEach(r => {
        const val = String(r[dataIdx] || "");
        maxWidth = Math.max(maxWidth, ctx.measureText(val).width);
      }));
      return Math.ceil(maxWidth + (globalFontSize * 2.5));
    });
    setColWidths(newWidths);
  };

  const clear = () => {
    setPages([]);
    setAllDataRows([]);
    setFileName(null);
    setError(null);
  };

  return (
    <div className="flex flex-col min-h-screen">
      <style>{`
        @media print {
          @page {
            size: A4 ${orientation};
            margin: 20mm 10mm 10mm 10mm; /* Increased top margin for direct browser print */
          }
        }
      `}</style>

      {isExporting && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[9999] flex flex-col items-center justify-center text-white no-print">
          <Loader2 className="w-12 h-12 animate-spin text-blue-400 mb-4" />
          <p className="font-bold uppercase tracking-widest text-sm">Generating PDF Document...</p>
        </div>
      )}

      <nav className="no-print bg-white border-b border-gray-200 px-6 py-4 sticky top-0 z-50 shadow-sm flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-xl">
            <FileSpreadsheet className="text-white w-5 h-5" />
          </div>
          <div>
            <h1 className="text-md font-black text-gray-900 leading-tight">SDN Print Pro</h1>
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest">
              Direct System Output
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          {!fileName ? (
            <label className="flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-5 py-2.5 rounded-xl cursor-pointer transition-all shadow-md active:scale-95 font-bold text-xs uppercase">
              <FileUp className="w-4 h-4" />
              Upload Excel
              <input type="file" accept=".xlsx" className="hidden" onChange={handleFileUpload} />
            </label>
          ) : (
            <div className="flex gap-2">
              <button 
                onClick={handlePrint}
                className="flex items-center gap-2 bg-gray-900 hover:bg-black text-white px-5 py-2.5 rounded-xl transition-all shadow-md active:scale-95 font-bold text-xs uppercase"
              >
                <Printer className="w-4 h-4" />
                Direct Print
              </button>
              <button 
                onClick={handleDownloadPDF}
                className="flex items-center gap-2 bg-blue-50 text-blue-700 border border-blue-100 px-5 py-2.5 rounded-xl transition-all active:scale-95 font-bold text-xs uppercase"
              >
                <FileDown className="w-4 h-4" />
                Save PDF
              </button>
              <button onClick={clear} className="p-2.5 text-gray-400 hover:text-red-500 transition-colors bg-gray-50 rounded-xl border border-gray-200">
                <Trash2 className="w-4 h-4" />
              </button>
            </div>
          )}
        </div>
      </nav>

      {pages.length > 0 && (
        <div className="no-print bg-white border-b border-gray-100 px-6 py-3 editor-controls">
          <div className="max-w-7xl mx-auto flex items-center gap-8 overflow-x-auto pb-1">
            <div className="flex items-center gap-4 border-r pr-8 shrink-0">
               <button 
                  onClick={() => setOrientation(prev => prev === 'portrait' ? 'landscape' : 'portrait')}
                  className="flex items-center gap-2 text-gray-600 hover:text-blue-600 transition-colors font-bold text-[10px] uppercase"
                >
                  {orientation === 'portrait' ? <Smartphone className="w-4 h-4" /> : <Monitor className="w-4 h-4" />}
                  {orientation}
                </button>
                <button 
                  onClick={autoAdjustAllWidths}
                  className="flex items-center gap-2 text-blue-600 font-bold text-[10px] uppercase"
                >
                  <Zap className="w-4 h-4" />
                  Auto-Fit
                </button>
            </div>

            <div className="flex items-center gap-6 shrink-0 border-r pr-8">
              <div className="flex flex-col gap-1 items-start">
                <label className="text-[9px] font-bold text-gray-400 uppercase flex items-center gap-1">
                  <Maximize2 className="w-3 h-3" /> Height (px)
                </label>
                <input 
                  type="number" 
                  min="1" 
                  value={globalRowHeight} 
                  onChange={(e) => setGlobalRowHeight(Math.max(1, parseInt(e.target.value) || 0))}
                  className="w-16 px-2 py-1.5 border border-gray-100 rounded-lg text-[10px] text-center font-bold focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>

              <div className="flex flex-col gap-1 items-start">
                <label className="text-[9px] font-bold text-gray-400 uppercase flex items-center gap-1">
                  <Type className="w-3 h-3" /> Font Size (pt)
                </label>
                <input 
                  type="number" 
                  min="1" 
                  step="0.5" 
                  value={globalFontSize} 
                  onChange={(e) => setGlobalFontSize(Math.max(1, parseFloat(e.target.value) || 0))}
                  className="w-16 px-2 py-1.5 border border-gray-100 rounded-lg text-[10px] text-center font-bold focus:ring-2 focus:ring-emerald-500 outline-none"
                />
              </div>
            </div>

            <div className="flex items-center gap-2 overflow-x-auto py-1">
              <div className="text-[9px] font-bold text-gray-400 uppercase shrink-0 mr-2">Col Widths (E-M):</div>
              {colWidths.map((w, idx) => (
                <div key={idx} className="flex flex-col gap-0.5 items-center shrink-0">
                  <span className="text-[8px] text-gray-300 font-black">{String.fromCharCode(69 + idx)}</span>
                  <input 
                    type="number" 
                    value={w} 
                    onChange={(e) => {
                      const nw = [...colWidths];
                      nw[idx] = Math.max(1, parseInt(e.target.value) || 0);
                      setColWidths(nw);
                    }}
                    className="w-12 px-1 py-1.5 border border-gray-100 rounded-lg text-[9px] text-center font-bold focus:ring-1 focus:ring-blue-500 outline-none"
                  />
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      <main className="flex-1 bg-gray-50 p-4 sm:p-8">
        {loading && (
          <div className="no-print flex flex-col items-center justify-center p-20 gap-4">
            <Loader2 className="animate-spin h-8 w-8 text-blue-600" />
            <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Reading Excel...</p>
          </div>
        )}

        {error && (
          <div className="no-print max-w-xl mx-auto bg-red-50 text-red-700 p-4 rounded-xl mb-8 flex gap-3 border border-red-100 text-xs font-medium">
            <AlertCircle className="w-4 h-4 shrink-0" />
            {error}
          </div>
        )}

        <div id="printArea" className={`${pages.length > 0 ? 'block' : 'hidden'} mx-auto bg-white shadow-xl print:shadow-none overflow-visible`}>
          {pages.map((page, pIdx) => (
            <div key={page.id} className="site-group">
              <table className="printable-site-table" style={{ width: `${totalTableWidth}px` }}>
                <colgroup>
                  {colWidths.map((w, i) => <col key={i} style={{ width: `${w}px` }} />)}
                </colgroup>
                <thead>
                  <tr>
                    <th colSpan={COL_COUNT} className="dynamic-header-cell">
                      <div className="site-header-box">
                        <div className="flex flex-col">
                          <div className="header-site-title">{page.siteHeader.siteName}</div>
                          <div className="flex gap-4 items-center mt-1">
                            <span className="flex gap-1">
                              <span className="header-label">SDN ID:</span>
                              <span className="text-[12.5pt] font-black">{page.siteHeader.sdnId}</span>
                            </span>
                            <span className="flex gap-1 border-l border-gray-300 pl-3">
                              <span className="header-label">TYPE:</span>
                              <span className="text-[10.5pt] font-black uppercase">{page.siteHeader.sdnType}</span>
                            </span>
                          </div>
                        </div>
                        <div className="flex flex-col items-end">
                          <div className="flex gap-3 text-[9.5pt] font-bold">
                            <span className="flex gap-1"><span className="header-label">SDN</span>{page.siteHeader.sdnDate}</span>
                            <span className="flex gap-1"><span className="header-label">REQ</span>{page.siteHeader.reqDate}</span>
                          </div>
                          <div className="flex gap-4 text-[9.5pt] font-black border-t border-black mt-1 pt-0.5">
                            <span className="flex gap-1"><span className="header-label">SZS</span>{page.siteHeader.szsPob}</span>
                            <span className="flex gap-1"><span className="header-label">CLIENT</span>{page.siteHeader.clientPob}</span>
                          </div>
                        </div>
                      </div>
                    </th>
                  </tr>
                  <tr className="col-header-row" style={{ height: `${globalRowHeight}px` }}>
                    {Array.from({ length: COL_COUNT }).map((_, i) => (
                      <th key={i} style={{ fontSize: `${globalFontSize}pt` }}>{tableHeaders ? String(tableHeaders[i + PRINT_COLS_START] || "") : ""}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {page.rows.map((row, rIdx) => (
                    <tr key={rIdx} style={{ height: `${globalRowHeight}px` }}>
                      {Array.from({ length: COL_COUNT }).map((_, i) => (
                        <td key={i} style={{ fontSize: `${globalFontSize}pt` }}>
                          {row[i + PRINT_COLS_START] !== null && row[i + PRINT_COLS_START] !== undefined 
                            ? String(row[i + PRINT_COLS_START]) 
                            : ""}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ))}
        </div>

        {!fileName && !loading && (
          <div className="no-print mt-12 text-center max-w-md mx-auto bg-white p-10 rounded-[2.5rem] shadow-sm border border-gray-100">
            <div className="bg-blue-50 w-16 h-16 rounded-3xl flex items-center justify-center mx-auto mb-6">
              <FileSpreadsheet className="w-8 h-8 text-blue-600" />
            </div>
            <h2 className="text-xl font-bold text-gray-900 mb-2">Ready to Print</h2>
            <p className="text-gray-400 text-xs mb-8">
              Upload an Excel file to start direct high-speed printing.
            </p>
            <label className="inline-flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-8 py-3 rounded-2xl cursor-pointer transition-all shadow-lg shadow-blue-500/20 active:scale-95 font-bold text-xs uppercase tracking-wider">
              <FileUp className="w-4 h-4" />
              Upload Source
              <input type="file" accept=".xlsx" className="hidden" onChange={handleFileUpload} />
            </label>
          </div>
        )}
      </main>

      {pages.length > 0 && !loading && (
        <div className="no-print fixed bottom-6 left-1/2 -translate-x-1/2 bg-gray-900 text-white px-6 py-3 rounded-2xl flex items-center gap-3 text-[10px] font-bold shadow-2xl">
          <CheckCircle2 className="w-4 h-4 text-emerald-400" /> 
          <span className="uppercase tracking-widest">{pages.length} PAGES READY FOR DIRECT PRINT</span>
        </div>
      )}
    </div>
  );
};

export default App;
