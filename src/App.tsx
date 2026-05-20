/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useMemo, useRef, ChangeEvent } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Search, 
  FileText, 
  Download, 
  Settings, 
  Database, 
  AlertCircle, 
  ChevronRight, 
  Table, 
  Columns,
  RefreshCw,
  Info,
  Volume2,
  VolumeX,
  Globe
} from 'lucide-react';
import Papa from 'papaparse';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { CIVILIZATIONS, GALAXIES } from './constants';

// Column configuration mapping
interface ColumnConfig {
  name: string;
  enabled: boolean;
  rawIndex: number;
}

type ReportType = 'simple' | 'detailed';

const CIV_ACRONYMS: Record<string, string> = {
  'Alliance of Galactic Travellers': 'AGT',
  'Intergalactic Travellers Foundation': 'IGTF',
  'Calypso Travellers Foundation': 'CTF',
  'Hyades Travellers Foundation': 'HTF',
  'Budullanger Travellers Foundation': 'BTF',
  'Budullangr Travellers Foundation': 'BTF',
  'Isdoraijung Travellers Foundation': 'ITF',
  'Kikolgallr Travellers Foundation': 'KTF',
  'Eissentam Travellers Foundation': 'ETF',
  'Ickjamatew Travellers Foundation': 'IJTF',
  'Rycempler Travellers Foundation': 'RTF',
  'Zavainlani Travellers Foundation': 'ZTF',
  'Animal Cracker Projects': 'ACP',
  'United Star Navy': 'USN',
  'CELAB Galactic Industries': 'CGI',
  'IVc Project': 'IVc',
  'AAAM Expeditionary': 'AAAM',
  'Riven Minerals and Exploration': 'RME',
  'Gravemind Expeditionary Force': 'GMEF'
};

const getDisplayValue = (val: any, colIdx?: number) => {
  const strVal = String(val || '').trim();
  if (colIdx === 18 && CIV_ACRONYMS[strVal]) {
    return CIV_ACRONYMS[strVal];
  }
  return strVal;
};

export default function App() {
  const [reportType, setReportType] = useState<ReportType>('simple');
  const [sheetUrl, setSheetUrl] = useState<string>(() => {
    const saved = localStorage.getItem('sheet_reporter_url');
    const oldDefault = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSWiJE26JMTHgjGeZfpfTrwT1HL2ZnXIqiOVkNs-V8wtDkGE7ey0Q9hnAM-bpMhy475q45qHa09o2vC/pub?gid=0&single=true&output=csv';
    const previousDefault = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ0jFq80ut0o5jtApdhRG8sR2CIufVn0FNcugR_7fdCIfrDRfgB9s-SvEhBAePrQCibr1RcxFVoXj7o/pub?gid=354119689&single=true&output=tsv';
    const newDefault = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ0jFq80ut0o5jtApdhRG8sR2CIufVn0FNcugR_7fdCIfrDRfgB9s-SvEhBAePrQCibr1RcxFVoXj7o/pub?gid=0&single=true&output=tsv';
    
    if (!saved || saved === oldDefault || saved === previousDefault) return newDefault;
    return saved;
  });
  const [showSettings, setShowSettings] = useState(false);
  const [audioEnabled, setAudioEnabled] = useState<boolean>(() => {
    const saved = localStorage.getItem('agt_audio_enabled');
    return saved === null ? true : saved === 'true';
  });
  
  const audioRef = useRef<HTMLAudioElement>(null);

  // Initial fetch
  useEffect(() => {
    if (sheetUrl) {
      fetchData();
    }
  }, []);

  // Background Audio Management
  useEffect(() => {
    const handleFirstInteraction = () => {
      if (audioEnabled && audioRef.current) {
        audioRef.current.volume = 0.4;
        audioRef.current.play().catch(() => {});
      }
      window.removeEventListener('mousedown', handleFirstInteraction);
      window.removeEventListener('keydown', handleFirstInteraction);
      window.removeEventListener('touchstart', handleFirstInteraction);
    };

    window.addEventListener('mousedown', handleFirstInteraction);
    window.addEventListener('keydown', handleFirstInteraction);
    window.addEventListener('touchstart', handleFirstInteraction);

    return () => {
      window.removeEventListener('mousedown', handleFirstInteraction);
      window.removeEventListener('keydown', handleFirstInteraction);
      window.removeEventListener('touchstart', handleFirstInteraction);
    };
  }, [audioEnabled]);

  useEffect(() => {
    if (audioRef.current) {
      audioRef.current.volume = 0.4;
      if (audioEnabled) {
        audioRef.current.play().catch(() => {});
      } else {
        audioRef.current.pause();
      }
    }
    localStorage.setItem('agt_audio_enabled', String(audioEnabled));
  }, [audioEnabled]);

  const handleManualPlay = () => {
    if (audioEnabled && audioRef.current && audioRef.current.paused) {
      audioRef.current.play().catch(() => {});
    }
  };

  const [logoUrl, setLogoUrl] = useState('/AGTicon.png');
  const [logoTriedCount, setLogoTriedCount] = useState(0);
  const [pageSize, setPageSize] = useState<number>(() => {
    const saved = localStorage.getItem('agt_page_size');
    return saved ? parseInt(saved, 10) : 15;
  });
  const [currentPage, setCurrentPage] = useState<number>(1);

  useEffect(() => {
    localStorage.setItem('agt_page_size', String(pageSize));
  }, [pageSize]);

  const [searchKey, setSearchKey] = useState('');
  const [selectedGalaxy, setSelectedGalaxy] = useState('All');
  const [data, setData] = useState<any[]>([]);
  const [columns, setColumns] = useState<ColumnConfig[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [matchedRecords, setMatchedRecords] = useState<any[]>([]);

  // Save sheet URL to localStorage
  useEffect(() => {
    if (sheetUrl) {
      localStorage.setItem('sheet_reporter_url', sheetUrl);
    }
  }, [sheetUrl]);

  useEffect(() => {
    setCurrentPage(1);
  }, [matchedRecords, pageSize]);

  const fetchData = async () => {
    if (!sheetUrl) {
      setError('Please provide a Google Sheet CSV URL in settings.');
      setShowSettings(true);
      return;
    }

    setLoading(true);
    setError(null);
    setMatchedRecords([]);

    try {
      // Handle the case where the user might paste a regular sheet URL instead of a pub link
      let fetchUrl = sheetUrl;
      if (sheetUrl.includes('docs.google.com/spreadsheets/') && !sheetUrl.includes('pub?')) {
        // Try to convert regular URL to CSV export if possible, 
        // though "Publish to Web" is the official way.
        if (sheetUrl.includes('/edit')) {
          fetchUrl = sheetUrl.replace(/\/edit.*$/, '/export?format=csv');
        }
      }

      const response = await fetch(fetchUrl);
      if (!response.ok) throw new Error('Failed to fetch sheet data. Is it published to the web?');
      
      const csvText = await response.text();
      
      Papa.parse(csvText, {
        header: false,
        skipEmptyLines: true,
        delimiter: fetchUrl.includes('output=tsv') ? '\t' : undefined,
        complete: (results) => {
          const rawRows = results.data as string[][];
          if (rawRows.length < 2) {
            setError('The source sheet data is insufficient (need at least 2 rows).');
            setLoading(false);
            return;
          }

          const headers = rawRows[1]; // Row 2 is headers
          
          const simpleIndices = [0, 1, 2, 6, 7, 10, 11, 14, 15, 18, 76];
          const detailedIndices = [0, 1, 2, 6, 7, 8, 10, 11, 13, 14, 15, 17, 18, 22, 23, 24, 25, 26, 27, 28, 31, 32, 33, 34, 35, 36, 37, 76];
          
          const targetIndexes = reportType === 'simple' ? simpleIndices : detailedIndices;
          
          const filteredColumns = targetIndexes.map(idx => ({
            name: headers[idx] || `Col ${String.fromCharCode(65 + (idx % 26))}${idx >= 26 ? String.fromCharCode(65 + Math.floor(idx / 26) - 1) : ''}`,
            enabled: true,
            rawIndex: idx
          }));
          
          setColumns(filteredColumns);
          
          const processedData = rawRows.slice(3) // Skip Rows 1, 2, 3 (index 0, 1, 2)
            .filter(row => {
              const colA = String(row[0] || '').trim();
              const colB = String(row[1] || '').trim();
              
              // Skip if:
              // - Col A is empty
              // - Col A has SKIPROW
              // - Col A has #N/A
              // - Col B is blank
              if (colA === '' || colA.includes('SKIPROW') || colA.includes('#N/A')) return false;
              if (colB === '') return false;
              
              return true;
            })
            .map(row => {
              const rowObj: any = {};
              targetIndexes.forEach((colIdx, listIdx) => {
                const headerName = filteredColumns[listIdx].name;
                rowObj[headerName] = row[colIdx] || '';
              });
              return rowObj;
            });
          
          setData(processedData);
          
          if (searchKey || selectedGalaxy !== 'All') {
            findRecord(processedData, filteredColumns, searchKey, selectedGalaxy);
          }
          setLoading(false);
        },
        error: (err: any) => {
          setError(`Parsing error: ${err.message}`);
          setLoading(false);
        }
      });
    } catch (err: any) {
      setError(err.message || 'Operation failed');
      setLoading(false);
    }
  };

  const handleSearch = () => {
    if (!data.length) {
      fetchData();
    } else {
      findRecord(data, columns);
    }
  };

  const findRecord = (sourceData: any[], sourceCols: ColumnConfig[], civTerm?: string, galTerm?: string) => {
    const currentCivTerm = (civTerm ?? searchKey).trim().toLowerCase();
    const currentGalTerm = (galTerm ?? selectedGalaxy).trim().toLowerCase();
    
    if (!currentCivTerm && currentGalTerm === 'all' && !sourceCols.length) return;

    // Civilization search matches on Column S (index 18)
    // Galaxy search matches on Column A (index 0)
    const galaxyFieldName = sourceCols.find(c => c.rawIndex === 0)?.name;
    const civFieldName = sourceCols.find(c => c.rawIndex === 18)?.name;
    
    if (!civFieldName || !galaxyFieldName) {
      setMatchedRecords([]);
      setError('Required matching fields (Column A or S) not found in the current report mode.');
      return;
    }

    const matches = sourceData.filter(row => {
      const civMatch = currentCivTerm === 'all' || !currentCivTerm || 
                      String(row[civFieldName] || '').toLowerCase().includes(currentCivTerm);
      const galMatch = currentGalTerm === 'all' || 
                      String(row[galaxyFieldName] || '').toLowerCase().includes(currentGalTerm);
      return civMatch && galMatch;
    });

    // Sort by Column A (Galaxy/Region logic) then Column B
    const nameFieldName = sourceCols.find(c => c.rawIndex === 1)?.name || sourceCols[1]?.name;
    const sortedMatches = [...matches].sort((a, b) => {
      const galA = String(a[galaxyFieldName] || '').toLowerCase();
      const galB = String(b[galaxyFieldName] || '').toLowerCase();
      
      if (galA !== galB) return galA.localeCompare(galB);
      
      const nameA = String(a[nameFieldName] || '').toLowerCase();
      const nameB = String(b[nameFieldName] || '').toLowerCase();
      return nameA.localeCompare(nameB);
    });

    if (sortedMatches.length > 0) {
      setMatchedRecords(sortedMatches);
      setError(null);
    } else {
      setMatchedRecords([]);
      setError(`No records found for the selected criteria.`);
    }
  };

  const downloadFullReportPdf = () => {
    if (matchedRecords.length === 0) return;

    const doc = new jsPDF('l', 'mm', 'a4'); // Landscape for tables
    const displayId = searchKey || 'Bulk';
    
    doc.setFontSize(22);
    doc.text(`AGT System Report`, 20, 20);
    
    doc.setFontSize(10);
    doc.text(`Extraction Ref: ${searchKey || 'All'} / ${selectedGalaxy}`, 20, 30);
    doc.text(`Generated on: ${new Date().toLocaleString()}`, 20, 35);
    doc.text(`Result Count: ${matchedRecords.length} Verified Entries`, 20, 40);

    // Inject AGT Icon logo on Page 1, right of report summary introduction text
    try {
      const imgElements = document.getElementsByTagName('img');
      let logoImg: HTMLImageElement | null = null;
      for (let i = 0; i < imgElements.length; i++) {
        if (imgElements[i].alt === 'AGT Logo') {
          logoImg = imgElements[i];
          break;
        }
      }
      if (logoImg && logoImg.style.display !== 'none' && logoImg.complete) {
        doc.addImage(logoImg, 'PNG', 240, 15, 25, 25);
      } else {
        const tempImg = new Image();
        tempImg.src = logoUrl;
        if (tempImg.complete) {
          doc.addImage(tempImg, 'PNG', 240, 15, 25, 25);
        }
      }
    } catch (err) {
      console.warn('PDF image logo inject failed:', err);
    }
    
    const activeCols = columns.filter(col => col.enabled);
    const urlMap = new Map<string, string>();
    
    const tableData = matchedRecords.map((record, rIdx) => 
      activeCols.map((col, cIdx) => {
        const rawVal = record[col.name];
        const val = getDisplayValue(rawVal, col.rawIndex);
        
        if (String(rawVal || '').startsWith('http')) {
          urlMap.set(`${rIdx}-${cIdx}`, String(rawVal));
          return 'LINK';
        }
        return val || '-';
      })
    );

    // Add total row to PDF
    const countFieldName = columns[0]?.name;
    const totalRow = activeCols.map(col => {
      if (col.name === countFieldName) return `Count: ${matchedRecords.length}`;
      return '';
    });
    tableData.push(totalRow);

    autoTable(doc, {
      startY: 50,
      head: [activeCols.map(col => col.name)],
      body: tableData,
      theme: 'grid',
      headStyles: { fillColor: [20, 20, 20], textColor: [255, 255, 255], fontSize: 8 },
      bodyStyles: { fontSize: 8 },
      footStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], fontStyle: 'bold' },
      margin: { top: 50, left: 20, right: 20 },
      didParseCell: (data) => {
        if (data.row.index === tableData.length - 1) {
          data.cell.styles.fillColor = [220, 220, 220];
          data.cell.styles.fontStyle = 'bold';
        }
        
        const key = `${data.row.index}-${data.column.index}`;
        if (urlMap.has(key)) {
          data.cell.styles.textColor = [0, 0, 255];
        }
      },
      didDrawCell: (data) => {
        const key = `${data.row.index}-${data.column.index}`;
        const url = urlMap.get(key);
        if (url && data.section === 'body') {
          doc.link(data.cell.x, data.cell.y, data.cell.width, data.cell.height, { url });
        }
      }
    });

    const filename = displayId.replace(/[^a-z0-9]/gi, '_');
    doc.save(`agt_full_report_${filename}.pdf`);
  };

  const downloadCsv = () => {
    if (matchedRecords.length === 0) return;
    
    const activeCols = columns.filter(col => col.enabled);
    const csvData = matchedRecords.map(record => {
      const row: any = {};
      activeCols.forEach(col => {
        row[col.name] = getDisplayValue(record[col.name], col.rawIndex);
      });
      return row;
    });
    
    const csv = Papa.unparse(csvData);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const displayId = (searchKey || 'Bulk').replace(/[^a-z0-9]/gi, '_');
    
    if (link.download !== undefined) {
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', `agt_${reportType}_report_${displayId}.csv`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  const toggleColumn = (name: string) => {
    setColumns(prev => prev.map(c => c.name === name ? { ...c, enabled: !c.enabled } : c));
  };

  const activeColumnsCount = useMemo(() => columns.filter(c => c.enabled).length, [columns]);

  const totalPoints = useMemo(() => {
    return matchedRecords.length;
  }, [matchedRecords]);

  const totalPages = useMemo(() => {
    return Math.ceil(matchedRecords.length / pageSize);
  }, [matchedRecords.length, pageSize]);

  const paginatedRecords = useMemo(() => {
    const startIdx = (currentPage - 1) * pageSize;
    return matchedRecords.slice(startIdx, startIdx + pageSize);
  }, [matchedRecords, currentPage, pageSize]);

  return (
    <div 
      onMouseDown={handleManualPlay}
      onTouchStart={handleManualPlay}
      className="min-h-screen bg-[#0a0a0a] text-agt-orange font-sans selection:bg-agt-orange selection:text-black"
    >
      {/* Header */}
      <header className="border-b border-agt-orange/5 bg-black/40 backdrop-blur-md sticky top-0 z-50">
        <div className="max-w-5xl mx-auto px-6 h-20 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="relative w-10 h-10 flex items-center justify-center shrink-0">
              <img 
                src={logoUrl} 
                alt="AGT Logo" 
                className={`w-10 h-10 object-contain opacity-90 ${logoTriedCount >= 3 ? 'hidden' : ''}`}
                onError={(e) => {
                  if (logoTriedCount === 0) {
                    setLogoTriedCount(1);
                    setLogoUrl('/AGTIcon.png');
                  } else if (logoTriedCount === 1) {
                    setLogoTriedCount(2);
                    setLogoUrl('/api/asset-proxy?id=1h9HvAGeru6Vo7PiWdLbXmGogD8TySnnz');
                  } else {
                    setLogoTriedCount(3);
                    const img = e.target as HTMLImageElement;
                    img.style.display = 'none';
                  }
                }}
              />
              {logoTriedCount >= 3 && (
                <div className="agt-fallback w-10 h-10 border border-agt-orange rounded-sm flex items-center justify-center shrink-0">
                  <span className="text-agt-orange font-bold text-[10px] tracking-tighter">AGT</span>
                </div>
              )}
            </div>
            <div className="flex flex-col">
              <h1 className="font-bold text-xs tracking-[0.2em] uppercase text-agt-orange">Alliance of Galactic Travellers</h1>
              <span className="text-[9px] text-agt-orange uppercase tracking-[0.3em] font-bold">AGT System Report Tool</span>
            </div>
          </div>
          
          <div className="flex items-center gap-6">
            <div className="hidden md:block text-[9px] text-agt-orange/30 tracking-widest font-mono">
              STATUS: <span className={
                loading ? 'text-yellow-500' :
                sheetUrl ? 'text-emerald-500' : 
                'text-red-500'
              }>
                {loading ? 'SYNCING' : sheetUrl ? 'CONNECTED' : 'DISCONNECTED'}
              </span>
            </div>
            <button 
              onClick={() => setShowSettings(!showSettings)}
              className="p-2 hover:bg-agt-orange/5 rounded-lg transition-colors relative group"
              title="Settings"
              id="settings-btn"
            >
              <Settings className={`w-5 h-5 transition-transform duration-300 ${showSettings ? 'text-agt-orange rotate-90' : 'text-agt-orange group-hover:text-agt-orange'}`} />
              {!sheetUrl && (
                <span className="absolute top-0 right-0 w-1.5 h-1.5 bg-agt-orange rounded-full shadow-[0_0_5px_rgba(255,180,81,0.5)]"></span>
              )}
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-16">
        <div className="flex flex-col gap-16">
          
          {/* Main Search Logic Container - centered aesthetic */}
          <div className="flex flex-col items-center space-y-12">
            <div className="w-full max-w-xl text-center space-y-4">
              <h2 className="text-4xl font-light tracking-tight text-agt-orange">AGT System Report Tool</h2>
              
              {/* Report Mode Selector */}
              <div className="flex justify-center mb-8">
                <div className="inline-flex p-1 bg-agt-orange/5 border border-agt-orange/20 rounded-xl">
                  <button
                    onClick={() => {
                      setReportType('simple');
                      setData([]);
                      setMatchedRecords([]);
                    }}
                    className={`px-6 py-2 rounded-lg text-[10px] uppercase tracking-widest font-bold transition-all ${
                      reportType === 'simple' 
                        ? 'bg-agt-orange text-black shadow-lg' 
                        : 'text-agt-orange hover:bg-agt-orange/10'
                    }`}
                  >
                    Simple Report
                  </button>
                  <button
                    onClick={() => {
                      setReportType('detailed');
                      setData([]);
                      setMatchedRecords([]);
                    }}
                    className={`px-6 py-2 rounded-lg text-[10px] uppercase tracking-widest font-bold transition-all ${
                      reportType === 'detailed' 
                        ? 'bg-agt-orange text-black shadow-lg' 
                        : 'text-agt-orange hover:bg-agt-orange/10'
                    }`}
                  >
                    Detailed Report
                  </button>
                </div>
              </div>

              <div className="flex flex-col md:flex-row items-center justify-center gap-4">
                <div className="space-y-1">
                  <p className="text-agt-orange text-[10px] font-bold tracking-widest uppercase">Criteria 1</p>
                  <p className="text-agt-orange text-xs font-bold tracking-widest uppercase">Select Civilization</p>
                </div>
                <div className="h-px w-8 bg-agt-orange/20 hidden md:block mt-4"></div>
                <div className="space-y-1">
                  <p className="text-agt-orange text-[10px] font-bold tracking-widest uppercase">Criteria 2</p>
                  <p className="text-agt-orange text-xs font-bold tracking-widest uppercase">Preferred Galaxy</p>
                </div>
              </div>
            </div>

            <div className="w-full max-w-4xl grid grid-cols-1 md:grid-cols-2 gap-6">
              {/* Civilization Dropdown */}
              <div className="relative group">
                <div className="absolute inset-y-0 left-0 pl-6 flex items-center pointer-events-none text-agt-orange group-focus-within:text-agt-orange transition-colors">
                  <Search className="h-5 w-5" />
                </div>
                <select
                  value={searchKey}
                  onChange={(e) => {
                    const val = e.target.value;
                    setSearchKey(val);
                    if (data.length) {
                      findRecord(data, columns, val, selectedGalaxy);
                    } else {
                      fetchData();
                    }
                  }}
                  className="block w-full pl-14 pr-12 py-5 bg-[#1d1d1d] border-2 border-agt-orange rounded-full text-lg font-mono tracking-wider focus:outline-none focus:ring-2 focus:ring-agt-orange focus:border-agt-orange transition-all input-glow text-agt-orange appearance-none shadow-[0_0_30px_rgba(255,180,81,0.05)]"
                  id="civilization-select"
                >
                  <option value="" disabled className="bg-[#1d1d1d]">-- Choose Civilization --</option>
                  <option value="All" className="bg-[#1d1d1d]">All</option>
                  {CIVILIZATIONS.map(civ => (
                    <option key={civ} value={civ} className="bg-[#1d1d1d]">{civ}</option>
                  ))}
                </select>
                <div className="absolute right-4 inset-y-0 flex items-center pointer-events-none text-agt-orange">
                  <ChevronRight className="w-5 h-5 rotate-90" />
                </div>
              </div>

              {/* Galaxy Search */}
              <div className="relative group">
                <div className="absolute inset-y-0 left-0 pl-6 flex items-center pointer-events-none text-agt-orange group-focus-within:text-agt-orange transition-colors">
                  <Globe className="h-5 w-5" />
                </div>
                <input
                  value={selectedGalaxy}
                  placeholder="Enter Galaxy..."
                  onChange={(e) => {
                    const val = e.target.value;
                    setSelectedGalaxy(val);
                    if (data.length) {
                      findRecord(data, columns, searchKey, val);
                    }
                  }}
                  className="block w-full pl-14 pr-12 py-5 bg-[#1d1d1d] border-2 border-agt-orange rounded-full text-lg font-mono tracking-wider focus:outline-none focus:ring-2 focus:ring-agt-orange focus:border-agt-orange transition-all input-glow text-agt-orange shadow-[0_0_30px_rgba(255,180,81,0.05)]"
                  id="galaxy-select"
                />
                <div className="mt-2 ml-6">
                  <a 
                    href="https://nomanssky.fandom.com/wiki/Galaxy" 
                    target="_blank" 
                    rel="noopener noreferrer"
                    className="text-[10px] text-agt-orange/60 hover:text-agt-orange underline transition-colors"
                  >
                    Find galaxy names here
                  </a>
                </div>
              </div>
            </div>

            <button
              onClick={handleSearch}
              disabled={loading || (!searchKey && selectedGalaxy === 'All')}
              className="px-20 py-5 border-2 border-agt-orange bg-transparent text-agt-orange rounded-full font-black text-sm uppercase tracking-[0.2em] hover:bg-agt-orange/10 active:scale-[0.96] disabled:opacity-25 disabled:pointer-events-none shadow-[0_4px_15px_rgba(0,0,0,0.3)] hover:shadow-[0_0_20px_rgba(255,180,81,0.4)] transition-all flex flex-col items-center gap-2"
              id="fetch-btn"
            >
              {loading ? (
                <>
                  <RefreshCw className="w-5 h-5 animate-spin" />
                  <span className="text-[10px] tracking-[0.1em] mt-1">Data Access in Process - Please Wait</span>
                </>
              ) : (
                <>
                  <RefreshCw className="w-5 h-5" />
                  <span>Extract Reports</span>
                </>
              )}
            </button>

            {error && (
              <motion.div 
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="flex items-center gap-3 px-6 py-3 bg-agt-orange/5 border border-agt-orange text-agt-orange rounded-full text-xs font-medium tracking-wide"
              >
                <AlertCircle className="w-3.5 h-3.5 shrink-0" />
                <p>{error}</p>
              </motion.div>
            )}
          </div>

          <div className="space-y-12">
            
            {/* Settings Area - Full Width Toggleable */}
            <AnimatePresence>
              {showSettings && (
                <motion.div 
                  initial={{ height: 0, opacity: 0 }}
                  animate={{ height: 'auto', opacity: 1 }}
                  exit={{ height: 0, opacity: 0 }}
                  className="overflow-hidden bg-[#161616] border border-agt-orange/5 rounded-2xl"
                >
                    <div className="p-8 grid grid-cols-1 md:grid-cols-2 gap-12">
                    {/* Data Section */}
                    <div className="space-y-4">
                      <h3 className="text-[10px] uppercase tracking-widest font-bold text-agt-orange flex items-center gap-2">
                        <Database className="w-3 h-3" />
                        Source Identity
                      </h3>
                      <div className="space-y-4">
                        <button 
                          onClick={fetchData}
                          className="w-full py-4 bg-agt-orange/5 border border-agt-orange text-agt-orange rounded-xl text-[10px] uppercase tracking-widest font-bold hover:bg-agt-orange/10 transition-colors"
                        >
                          Re-sync Point Log Source
                        </button>
                      </div>
                    </div>

                    {/* Pagination Config Section */}
                    <div className="space-y-4">
                      <h3 className="text-[10px] uppercase tracking-widest font-bold text-agt-orange flex items-center gap-2">
                        <FileText className="w-3 h-3" />
                        Pagination Config
                      </h3>
                      <div className="space-y-2">
                        <p className="text-[10px] text-agt-orange/60 uppercase tracking-widest font-bold">Default Records Displayed</p>
                        <div className="flex gap-2">
                          {[15, 30, 50, 100].map(size => {
                            const isSelected = pageSize === size;
                            return (
                              <button
                                key={size}
                                onClick={() => setPageSize(size)}
                                className={`flex-1 py-3 rounded-xl border text-xs font-mono font-bold transition-all ${
                                  isSelected 
                                    ? 'bg-agt-orange text-black border-agt-orange shadow-[0_0_15px_rgba(255,180,81,0.5)] scale-102' 
                                    : 'bg-black/40 border-white/5 text-agt-orange hover:bg-agt-orange/5'
                                }`}
                                id={`pagesize-${size}-btn`}
                              >
                                {size}
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    </div>

                    {/* Audio Section */}
                    <div className="col-span-1 md:col-span-2 pt-8 border-t border-white/5 space-y-4">
                      <div className="flex items-center justify-between">
                        <div className="space-y-1">
                          <h3 className="text-[10px] uppercase tracking-widest font-bold text-agt-orange flex items-center gap-2">
                            <Volume2 className="w-3 h-3" />
                            Ambience Module
                          </h3>
                          <p className="text-[10px] text-agt-orange uppercase tracking-wider">Atmospheric Background Loop</p>
                        </div>
                        <button 
                          onClick={() => setAudioEnabled(!audioEnabled)}
                          className={`flex items-center gap-3 px-6 py-3 rounded-xl border transition-all text-[10px] uppercase tracking-widest font-bold ${
                            audioEnabled 
                              ? 'bg-agt-orange/10 border-agt-orange text-agt-orange' 
                              : 'bg-black/40 border-white/5 text-agt-orange'
                          }`}
                        >
                          {audioEnabled ? <Volume2 className="w-3.5 h-3.5" /> : <VolumeX className="w-3.5 h-3.5" />}
                          {audioEnabled ? 'Active' : 'Muted'}
                        </button>
                      </div>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Results Section - Full Width for Table */}
            <div className="w-full">
              <AnimatePresence mode="wait">
                {matchedRecords.length > 0 ? (
                  <motion.section
                    key="results"
                    initial={{ opacity: 0, y: 30 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, scale: 0.98 }}
                    className="glass-card rounded-2xl overflow-hidden"
                  >
                    <div className="p-8 border-b border-agt-orange/5 flex flex-col md:flex-row md:items-center justify-between gap-6">
                      <div className="space-y-1">
                        <h3 className="text-xl font-medium text-agt-orange flex items-center gap-3">
                          AGT Galactic Archives Results
                          <span className="px-2 py-0.5 rounded-full bg-agt-orange/5 text-[10px] text-agt-orange border border-agt-orange font-mono">
                            {matchedRecords.length} FOUND
                          </span>
                        </h3>
                        <p className="text-[10px] text-agt-orange uppercase tracking-[0.2em]">Verified Galactic Ledger Matches</p>
                      </div>

                      <div className="flex flex-wrap items-center gap-3">
                        {reportType === 'simple' && (
                          <button
                            onClick={downloadFullReportPdf}
                            className="flex items-center gap-3 px-8 py-4 border-2 border-agt-orange bg-transparent text-agt-orange hover:bg-agt-orange/10 rounded-xl text-[10px] uppercase tracking-[0.2em] font-bold transition-all shadow-[0_4px_20px_rgba(255,180,81,0.1)] active:scale-[0.98]"
                          >
                            <Download className="w-4 h-4" />
                            <span>Export PDF</span>
                          </button>
                        )}
                        <button
                          onClick={downloadCsv}
                          className="flex items-center gap-3 px-8 py-4 border-2 border-agt-orange bg-transparent text-agt-orange hover:bg-agt-orange/10 rounded-xl text-[10px] uppercase tracking-[0.2em] font-bold transition-all shadow-[0_4px_20px_rgba(255,180,81,0.1)] active:scale-[0.98]"
                        >
                          <Table className="w-4 h-4" />
                          <span>Export CSV</span>
                        </button>
                      </div>
                    </div>

                    <div className="overflow-x-auto custom-scrollbar">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-agt-orange/[0.02] border-b border-agt-orange/5">
                            {columns.filter(col => col.enabled).map((col, idx) => (
                              <th key={idx} className="p-4 text-[9px] uppercase tracking-widest font-bold text-agt-orange whitespace-nowrap">
                                {col.name}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-agt-orange/5">
                          {paginatedRecords.map((record, rIdx) => (
                            <tr key={rIdx} className="hover:bg-agt-orange/[0.02] transition-colors group">
                              {columns.filter(col => col.enabled).map((col, cIdx) => (
                                <td key={cIdx} className="p-4 text-[11px] text-agt-orange font-mono whitespace-nowrap overflow-hidden text-ellipsis max-w-[200px]">
                                  {getDisplayValue(record[col.name], col.rawIndex) || <span className="text-agt-orange italic">-</span>}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                        <tfoot className="border-t-2 border-agt-orange/10 bg-agt-orange/[0.03]">
                          <tr>
                            {columns.filter(col => col.enabled).map((col, idx) => (
                              <td key={idx} className="p-4 text-[10px] font-bold text-agt-orange">
                                {col.name === columns[0]?.name ? (
                                  <div className="flex flex-col">
                                    <span className="text-[8px] text-agt-orange uppercase tracking-tighter">Total Matches</span>
                                    <span>{matchedRecords.length}</span>
                                  </div>
                                ) : null}
                              </td>
                            ))}
                          </tr>
                        </tfoot>
                      </table>
                    </div>

                    {/* Pagination Navigation Bar */}
                    {matchedRecords.length > pageSize && (
                      <div className="p-6 border-t border-agt-orange/5 bg-agt-orange/[0.02] flex flex-col md:flex-row items-center justify-between gap-4">
                        <div className="text-[11px] uppercase tracking-wider text-agt-orange/80 font-mono">
                          Showing Page <span className="text-agt-orange font-bold font-sans">{currentPage}</span> of <span className="text-agt-orange font-bold font-sans">{totalPages}</span> <span className="text-agt-orange/40">({matchedRecords.length} total rows)</span>
                        </div>

                        <div className="flex flex-wrap items-center gap-2">
                          <button
                            onClick={() => setCurrentPage(1)}
                            disabled={currentPage === 1}
                            className="px-3 py-1.5 rounded-lg border border-agt-orange/20 bg-black/40 text-[10px] uppercase font-bold tracking-widest text-agt-orange hover:bg-agt-orange/10 disabled:opacity-30 disabled:pointer-events-none transition-all"
                          >
                            First
                          </button>

                          <button
                            onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                            disabled={currentPage === 1}
                            className="px-3 py-1.5 rounded-lg border border-agt-orange/20 bg-black/40 text-[10px] uppercase font-bold tracking-widest text-agt-orange hover:bg-agt-orange/10 disabled:opacity-30 disabled:pointer-events-none transition-all"
                          >
                            Prev
                          </button>

                          <div className="flex items-center gap-1">
                            {Array.from({ length: totalPages }).map((_, idx) => {
                              const pageNum = idx + 1;
                              if (totalPages > 6) {
                                if (pageNum !== 1 && pageNum !== totalPages && Math.abs(pageNum - currentPage) > 1) {
                                  if (pageNum === 2 && currentPage > 3) {
                                    return <span key="ellipsis-start" className="text-agt-orange/40 font-mono text-[10px] px-1">...</span>;
                                  }
                                  if (pageNum === totalPages - 1 && currentPage < totalPages - 2) {
                                    return <span key="ellipsis-end" className="text-agt-orange/40 font-mono text-[10px] px-1">...</span>;
                                  }
                                  return null;
                                }
                              }

                              const isCurrent = pageNum === currentPage;
                              return (
                                <button
                                  key={pageNum}
                                  onClick={() => setCurrentPage(pageNum)}
                                  className={`w-8 h-8 rounded-lg flex items-center justify-center font-mono text-xs transition-all ${
                                    isCurrent
                                      ? 'bg-agt-orange text-black font-extrabold shadow-[0_0_15px_rgba(255,180,81,0.6)] ring-2 ring-agt-orange ring-offset-2 ring-offset-black scale-110'
                                      : 'border border-agt-orange/10 bg-black/20 text-agt-orange hover:bg-agt-orange/10'
                                  }`}
                                  id={`page-btn-${pageNum}`}
                                >
                                  {pageNum}
                                </button>
                              );
                            })}
                          </div>

                          <button
                            onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                            disabled={currentPage === totalPages}
                            className="px-3 py-1.5 rounded-lg border border-agt-orange/20 bg-black/40 text-[10px] uppercase font-bold tracking-widest text-agt-orange hover:bg-agt-orange/10 disabled:opacity-30 disabled:pointer-events-none transition-all"
                          >
                            Next
                          </button>

                          <button
                            onClick={() => setCurrentPage(totalPages)}
                            disabled={currentPage === totalPages}
                            className="px-3 py-1.5 rounded-lg border border-agt-orange/20 bg-black/40 text-[10px] uppercase font-bold tracking-widest text-agt-orange hover:bg-agt-orange/10 disabled:opacity-30 disabled:pointer-events-none transition-all"
                          >
                            Last
                          </button>
                        </div>
                      </div>
                    )}

                    <div className="p-6 border-t border-agt-orange/5 flex flex-col md:flex-row items-center justify-between gap-6 bg-agt-orange/[0.01]">
                      <div className="flex items-center gap-4">
                        <div className="flex items-center gap-2">
                          <div className="w-1.5 h-1.5 rounded-full bg-agt-orange shadow-[0_0_8px_rgba(255,180,81,0.4)]"></div>
                          <span className="text-[9px] uppercase tracking-widest text-agt-orange font-bold">Ledger Integrity: Verified</span>
                        </div>
                        <span className="text-[9px] font-mono text-agt-orange uppercase tracking-widest hidden md:inline">
                          Index Reference: {Math.random().toString(16).substring(2, 8).toUpperCase()}
                        </span>
                      </div>
                    </div>
                  </motion.section>
                ) : !loading && (
                  <motion.div 
                    key="empty"
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    className="py-32 flex flex-col items-center justify-center text-center space-y-6 opacity-20 border border-agt-orange/5 rounded-2xl bg-agt-orange/[0.01]"
                  >
                    <div className="w-16 h-16 rounded-full border border-agt-orange/10 flex items-center justify-center">
                      <Database className="w-6 h-6" />
                    </div>
                    <div className="space-y-1">
                      <p className="text-sm font-medium uppercase tracking-[0.2em]">Terminal Ready</p>
                      <p className="text-xs font-light">Report Generation Sequence Pending Civilization Selection</p>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </div>
        </div>
      </main>

      {/* Footer Area */}
      <footer className="bg-[#FFB451] mt-auto">
        <div className="max-w-5xl mx-auto px-6 py-12 flex flex-col items-center gap-6 text-black">
          <div className="flex flex-wrap justify-center items-center gap-y-2 text-[10px] uppercase tracking-[0.2em] font-bold">
            <a href="https://www.nms-agt.com" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Home</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/about-the-agt" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">About</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/team" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Team</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/contribute" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Contribute</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/agt-galactic-archives" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Galactic Archives</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/engage" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Engage</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/agt-navi" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">AGT NAVI</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/terms" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Terms</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/support" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Support</a>
            <span className="ml-1 mr-2 text-black/40">|</span>
            <a href="https://www.nms-agt.com/terms/copyright" target="_blank" rel="noopener noreferrer" className="hover:opacity-60 transition-opacity">Copyright</a>
          </div>
          <p className="text-[9px] font-mono uppercase tracking-[0.3em] font-bold">&copy; 2026 Alliance of Galactic Travellers</p>
        </div>
      </footer>

      {/* Background Audio */}
      <audio 
        ref={audioRef}
        src="/api/asset-proxy?id=1MLd7Vp0whtVXZF-KRxSoH2544q4TA5zD"
        loop
        preload="auto"
      />
    </div>
  );
}

