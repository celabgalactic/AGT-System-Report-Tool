/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useMemo, useRef, useCallback, ChangeEvent, ReactNode } from 'react';
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
import { LanguageCode, LANGUAGES, TRANSLATIONS, translateColumnHeader } from './translations';
// @ts-ignore
import starSystemsIcon from './star-systems-icon.png';
// @ts-ignore
import agtAnthem from './AGT Anthem (Instrumental).mp3';

// Column configuration mapping
interface ColumnConfig {
  name: string;
  enabled: boolean;
  rawIndex: number;
}

type ReportType = 'simple' | 'detailed' | 'custom';

const colLetterToIdx = (letter: string): number => {
  let col = 0;
  letter.toUpperCase().trim().split('').forEach(char => {
    col = col * 26 + (char.charCodeAt(0) - 64);
  });
  return col - 1;
};

const getColLetter = (idx: number): string => {
  let temp = idx;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
};

const parseExcelRange = (rangeStr: string): number[] => {
  const result: number[] = [];
  const segments = rangeStr.split(',');
  segments.forEach(seg => {
    const parts = seg.trim().split(/\s+to\s+|\s*-\s*/i);
    if (parts.length === 2) {
      const startIdx = colLetterToIdx(parts[0].trim());
      const endIdx = colLetterToIdx(parts[1].trim());
      for (let i = Math.min(startIdx, endIdx); i <= Math.max(startIdx, endIdx); i++) {
        result.push(i);
      }
    } else {
      const single = seg.trim();
      if (single) {
        result.push(colLetterToIdx(single));
      }
    }
  });
  return result;
};

const CUSTOM_COLUMN_DEFS = [
  { id: 'A', label: 'Galaxy', letters: 'A' },
  { id: 'B', label: 'Region', letters: 'B' },
  { id: 'C', label: 'System Name', letters: 'C' },
  { id: 'G', label: 'Original name', letters: 'G' },
  { id: 'H', label: 'Coordinates', letters: 'H' },
  { id: 'I', label: 'Glyphs', letters: 'I' },
  { id: 'K', label: 'Survey', letters: 'K' },
  { id: 'L', label: 'Discoverer', letters: 'L' },
  { id: 'P', label: 'Survey Date', letters: 'P' },
  { id: 'O', label: 'Discovery Date', letters: 'O' },
  { id: 'N', label: 'Giant?', letters: 'N' },
  { id: 'Q', label: 'BH/Atlas', letters: 'Q' },
  { id: 'R', label: 'Dissonant', letters: 'R' },
  { id: 'S', label: 'Civ', letters: 'S' },
  { id: 'V', label: 'Platform', letters: 'V' },
  { id: 'W', label: 'Mode', letters: 'W' },
  { id: 'X', label: 'Stars', letters: 'X' },
  { id: 'Y', label: 'Category', letters: 'Y' },
  { id: 'Z', label: 'Color', letters: 'Z' },
  { id: 'AA', label: 'Planets', letters: 'AA' },
  { id: 'AB', label: 'Moons', letters: 'AB' },
  { id: 'AC', label: 'Faction', letters: 'AC' },
  { id: 'AE', label: 'Distance', letters: 'AE' },
  { id: 'AF', label: 'Water', letters: 'AF' },
  { id: 'AG', label: 'Economy', letters: 'AG' },
  { id: 'AH', label: 'Wealth', letters: 'AH' },
  { id: 'AI', label: 'Ebuy', letters: 'AI' },
  { id: 'AJ', label: 'ESell', letters: 'AJ' },
  { id: 'AK', label: 'Conflict', letters: 'AK' },
  { id: 'AL', label: 'Release', letters: 'AL' },
  { id: 'CQ', label: 'Rel#', letters: 'CQ' },
  { id: 'AN_AR', label: 'Trade', letters: 'AN to AR' },
  { id: 'AS_BG', label: 'Upgrades', letters: 'AS to BG' },
  { id: 'BM_BU', label: 'Notes', letters: 'BM to BU' },
  { id: 'BV', label: 'Phantom', letters: 'BV' },
  { id: 'BW', label: 'CTR Access', letters: 'BW' },
  { id: 'BY', label: 'Wiki Link', letters: 'BY' },
  { id: 'BZ_CD', label: 'Other Links', letters: 'BZ to CD' },
  { id: 'CE_CJ', label: 'Legacy Info', letters: 'CE to CK, DB to DD' },
  { id: 'CL', label: 'Age', letters: 'CL' },
  { id: 'CM', label: 'Research', letters: 'CM' },
  { id: 'CN', label: 'Misc', letters: 'CN' },
  { id: 'CY', label: 'Wealth Lvl', letters: 'CY' },
  { id: 'CZ', label: 'Conflict Lvl', letters: 'CZ' }
];

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

const isTargetColumnWithUrl = (colIdx: number | undefined, val: any): boolean => {
  if (colIdx === undefined || colIdx === null) return false;
  let temp = colIdx;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  const targets = ['BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'DB', 'DC', 'DD'];
  if (!targets.includes(letter)) return false;
  
  const str = String(val || '').trim();
  return str.startsWith('http://') || str.startsWith('https://');
};

interface AutocompleteProps {
  value: string;
  onChange: (val: string) => void;
  suggestions: string[];
  placeholder: string;
  id: string;
  icon: ReactNode;
}

function AutocompleteInput({ value, onChange, suggestions, placeholder, id, icon }: AutocompleteProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [filterText, setFilterText] = useState(value);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setFilterText(value);
  }, [value]);

  const filteredSuggestions = useMemo(() => {
    const text = filterText.trim().toLowerCase();
    const pool = Array.from(new Set(['All', ...suggestions]));
    if (!text) return pool.slice(0, 10);
    return pool
      .filter(s => s.toLowerCase().includes(text) && s.toLowerCase() !== text)
      .slice(0, 10);
  }, [filterText, suggestions]);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div ref={containerRef} className="relative w-full group">
      <div className="absolute inset-y-0 left-0 pl-6 flex items-center pointer-events-none text-[#FFB451] transition-colors">
        {icon}
      </div>
      <input
        id={id}
        type="text"
        value={filterText}
        placeholder={placeholder}
        onFocus={() => setIsOpen(true)}
        onChange={(e) => {
          const val = e.target.value;
          setFilterText(val);
          onChange(val);
        }}
        className="block w-full pl-14 pr-12 py-5 bg-[#1d1d1d] border-2 border-[#FF0500] rounded-full text-lg font-mono tracking-wider focus:outline-none focus:ring-2 focus:ring-[#FF0500] focus:border-[#FF0500] transition-all input-glow text-[#FFB451] shadow-[0_0_30px_rgba(255,5,0,0.05)] placeholder-[#FFB451]/30"
      />
      
      <AnimatePresence>
        {isOpen && filteredSuggestions.length > 0 && (
          <motion.div
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            className="absolute z-[80] left-0 right-0 mt-2 bg-[#161616] border border-[#FF0500] rounded-2xl overflow-hidden shadow-[0_10px_30px_rgba(255,5,0,0.4)] max-h-60 overflow-y-auto settings-scrollbar"
          >
            {filteredSuggestions.map((suggestion) => (
              <button
                key={suggestion}
                type="button"
                onMouseDown={() => {
                  onChange(suggestion);
                  setFilterText(suggestion);
                  setIsOpen(false);
                }}
                className="w-full text-left px-6 py-3 text-sm font-mono text-white hover:bg-[#E25530] hover:text-white transition-colors border-b border-[#FF0500]/10 last:border-0"
              >
                {suggestion}
              </button>
            ))}
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

const setCookie = (name: string, value: string, days = 365) => {
  const expires = new Date();
  expires.setTime(expires.getTime() + days * 24 * 60 * 60 * 1000);
  document.cookie = `${name}=${encodeURIComponent(value)};expires=${expires.toUTCString()};path=/;SameSite=Lax;`;
};

const getCookie = (name: string): string | null => {
  const nameEQ = name + "=";
  const ca = document.cookie.split(';');
  for (let i = 0; i < ca.length; i++) {
    let c = ca[i];
    while (c.charAt(0) === ' ') c = c.substring(1, c.length);
    if (c.indexOf(nameEQ) === 0) return decodeURIComponent(c.substring(nameEQ.length, c.length));
  }
  return null;
};

const deleteCookie = (name: string) => {
  document.cookie = `${name}=;expires=Thu, 01 Jan 1970 00:00:00 UTC;path=/;SameSite=Lax;`;
};

const getSecurityLevelNum = (val: string): number => {
  const normalized = String(val || '').trim().toLowerCase();
  if (normalized === 'public' || normalized === '0') return 0;
  if (normalized === 'private' || normalized === '1') return 1;
  if (normalized === 'restricted' || normalized === '2') return 2;
  if (normalized === 'top secret' || normalized === '3') return 3;
  if (normalized === 'slt restricted' || normalized === '4') return 4;
  if (normalized === 'scc restricted' || normalized === '5') return 5;
  return 0; // Default to Public (0)
};

const getSecurityLevelLabel = (level: number | null): string => {
  if (level === 0) return 'Public';
  if (level === 1) return 'Private';
  if (level === 2) return 'Restricted';
  if (level === 3) return 'Top Secret';
  if (level === 4) return 'SLT Restricted';
  if (level === 5) return 'SCC Restricted';
  return 'Public';
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

  // Load Traveller credentials from cookies on mount
  useEffect(() => {
    const name = getCookie('agt_traveller_name');
    const id = getCookie('agt_traveller_id');
    const lvlStr = getCookie('agt_security_level');
    
    if (name && id && lvlStr !== null) {
      setSavedTravellerName(name);
      setSavedTravellerId(id);
      setSavedSecurityLevel(parseInt(lvlStr, 10));
      setEnteredName(name);
      setEnteredId(id);
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
  const [fontScale, setFontScale] = useState<string>(() => {
    return localStorage.getItem('agt_font_scale') || '1x';
  });
  const [sortColumn, setSortColumn] = useState<string | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc' | null>(null);

  useEffect(() => {
    localStorage.setItem('agt_page_size', String(pageSize));
  }, [pageSize]);

  useEffect(() => {
    localStorage.setItem('agt_font_scale', fontScale);
  }, [fontScale]);

  const [language, setLanguage] = useState<LanguageCode>(() => {
    return (localStorage.getItem('agt_app_language') as LanguageCode) || 'en';
  });

  useEffect(() => {
    localStorage.setItem('agt_app_language', language);
  }, [language]);

  const t = useCallback((key: string): string => {
    if (language === 'en') return key;
    const langSet = TRANSLATIONS[language];
    if (langSet && langSet[key]) {
      return langSet[key];
    }
    return key;
  }, [language]);

  const [searchKey, setSearchKey] = useState('');
  const [selectedGalaxy, setSelectedGalaxy] = useState('');
  const [selectedRegion, setSelectedRegion] = useState('');
  const [discovererName, setDiscovererName] = useState('');
  const [surveyorName, setSurveyorName] = useState('');
  const [generationSeconds, setGenerationSeconds] = useState(0);
  const [pdfErrorMsg, setPdfErrorMsg] = useState<string | null>(null);
  const [searchLoading, setSearchLoading] = useState(false);
  const [reportGeneratingLoading, setReportGeneratingLoading] = useState(false);

  // Traveller Authentication States
  const [savedTravellerName, setSavedTravellerName] = useState<string | null>(null);
  const [savedTravellerId, setSavedTravellerId] = useState<string | null>(null);
  const [savedSecurityLevel, setSavedSecurityLevel] = useState<number | null>(null);

  const [enteredName, setEnteredName] = useState('');
  const [enteredId, setEnteredId] = useState('');
  const [isVerifying, setIsVerifying] = useState(false);
  const [verificationError, setVerificationError] = useState(false);
  const [alertPopup, setAlertPopup] = useState<{
    show: boolean;
    message: string;
    onClose?: () => void;
  } | null>(null);

  const showAlert = (message: string, onClose?: () => void) => {
    setAlertPopup({
      show: true,
      message,
      onClose
    });
  };

  useEffect(() => {
    let interval: any = null;
    if (reportGeneratingLoading) {
      setGenerationSeconds(0);
      interval = setInterval(() => {
        setGenerationSeconds(prev => prev + 1);
      }, 1000);
    } else {
      setGenerationSeconds(0);
    }
    return () => {
      if (interval) clearInterval(interval);
    };
  }, [reportGeneratingLoading]);

  const [enabledCustomColumns, setEnabledCustomColumns] = useState<Record<string, boolean>>(() => {
    const saved = localStorage.getItem('agt_custom_columns_toggles');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        // ignore
      }
    }
    const defaults: Record<string, boolean> = {};
    CUSTOM_COLUMN_DEFS.forEach(def => {
      defaults[def.id] = true;
    });
    return defaults;
  });

  useEffect(() => {
    localStorage.setItem('agt_custom_columns_toggles', JSON.stringify(enabledCustomColumns));
  }, [enabledCustomColumns]);

  const [data, setData] = useState<any[]>([]);
  const [rawRows, setRawRows] = useState<string[][]>([]);

  const civSuggestions = useMemo(() => {
    const rawList = data.map(row => String(row._rawRow ? row._rawRow[18] : '').trim()).filter(Boolean);
    const unique = Array.from(new Set([...CIVILIZATIONS, ...rawList]));
    return unique.sort();
  }, [data]);

  const galaxySuggestions = useMemo(() => {
    const rawList = data.map(row => String(row._rawRow ? row._rawRow[0] : '').trim()).filter(Boolean);
    const unique = Array.from(new Set([...GALAXIES, ...rawList]));
    return unique.sort();
  }, [data]);

  const regionSuggestions = useMemo(() => {
    const rawList = data.map(row => String(row._rawRow ? row._rawRow[1] : '').trim()).filter(Boolean);
    const unique = Array.from(new Set(rawList));
    return unique.sort();
  }, [data]);

  const discovererSuggestions = useMemo(() => {
    const rawList = data.map(row => String(row._rawRow ? row._rawRow[11] : '').trim()).filter(Boolean);
    const unique = Array.from(new Set(rawList));
    return unique.sort();
  }, [data]);

  const surveyorSuggestions = useMemo(() => {
    const rawList = data.map(row => String(row._rawRow ? row._rawRow[10] : '').trim()).filter(Boolean);
    const unique = Array.from(new Set(rawList));
    return unique.sort();
  }, [data]);

  const [columns, setColumns] = useState<ColumnConfig[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [matchedRecords, setMatchedRecords] = useState<any[]>([]);

  // Synchronized horizontal scrollbars for displayed results table
  const topScrollRef = useRef<HTMLDivElement>(null);
  const bottomScrollRef = useRef<HTMLDivElement>(null);
  const [tableScrollWidth, setTableScrollWidth] = useState(0);
  const [tableClientWidth, setTableClientWidth] = useState(0);

  const isScrollingTopRef = useRef(false);
  const isScrollingBottomRef = useRef(false);

  const handleTopScroll = () => {
    if (isScrollingBottomRef.current) return;
    isScrollingTopRef.current = true;
    const topEl = topScrollRef.current;
    const bottomEl = bottomScrollRef.current;
    if (topEl && bottomEl) {
      if (bottomEl.scrollLeft !== topEl.scrollLeft) {
        bottomEl.scrollLeft = topEl.scrollLeft;
      }
    }
    requestAnimationFrame(() => {
      isScrollingTopRef.current = false;
    });
  };

  const handleBottomScroll = () => {
    if (isScrollingTopRef.current) return;
    isScrollingBottomRef.current = true;
    const topEl = topScrollRef.current;
    const bottomEl = bottomScrollRef.current;
    if (topEl && bottomEl) {
      if (topEl.scrollLeft !== bottomEl.scrollLeft) {
        topEl.scrollLeft = bottomEl.scrollLeft;
      }
    }
    requestAnimationFrame(() => {
      isScrollingBottomRef.current = false;
    });
  };

  useEffect(() => {
    const bottomEl = bottomScrollRef.current;
    if (!bottomEl) return;

    const updateWidths = () => {
      setTableScrollWidth(bottomEl.scrollWidth);
      setTableClientWidth(bottomEl.clientWidth);
    };

    updateWidths();

    const resizeObserver = new ResizeObserver(() => {
      updateWidths();
    });

    resizeObserver.observe(bottomEl);

    const tableEl = bottomEl.querySelector('table');
    if (tableEl) {
      resizeObserver.observe(tableEl);
    }

    return () => {
      resizeObserver.disconnect();
    };
  }, [data, columns]);

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
    setSearchLoading(true);
    const startTime = Date.now();
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
          const parsedRows = results.data as string[][];
          if (parsedRows.length < 2) {
            setError('The source sheet data is insufficient (need at least 2 rows).');
            const elapsed = Date.now() - startTime;
            setTimeout(() => {
              setLoading(false);
              setSearchLoading(false);
            }, Math.max(0, 1500 - elapsed));
            return;
          }

          setRawRows(parsedRows);
          
          const elapsed = Date.now() - startTime;
          setTimeout(() => {
            setLoading(false);
            setSearchLoading(false);
          }, Math.max(0, 1500 - elapsed));
        },
        error: (err: any) => {
          setError(`Parsing error: ${err.message}`);
          const elapsed = Date.now() - startTime;
          setTimeout(() => {
            setLoading(false);
            setSearchLoading(false);
          }, Math.max(0, 1500 - elapsed));
        }
      });
    } catch (err: any) {
      setError(err.message || 'Operation failed');
      const elapsed = Date.now() - startTime;
      setTimeout(() => {
        setLoading(false);
        setSearchLoading(false);
      }, Math.max(0, 1500 - elapsed));
    }
  };

  // Declarative effect to handle real-time Column toggles and formatting on state updates
  useEffect(() => {
    if (rawRows.length < 2) return;

    const headers = rawRows[1]; // Row 2 is headers
    
    const simpleIndices = [0, 1, 2, 6, 7, 11, 14, 15, 18, 76];
    const detailedIndices = [0, 1, 2, 6, 7, 8, 10, 11, 13, 14, 15, 17, 18, 22, 23, 24, 25, 26, 27, 28, 31, 32, 33, 34, 35, 36, 37, 76];
    
    let targetIndexes: number[];
    if (reportType === 'simple') {
      targetIndexes = simpleIndices;
    } else if (reportType === 'detailed') {
      targetIndexes = detailedIndices;
    } else {
      const indicesSet = new Set<number>();
      CUSTOM_COLUMN_DEFS.forEach(def => {
        if (enabledCustomColumns[def.id]) {
          parseExcelRange(def.letters).forEach(idx => indicesSet.add(idx));
        }
      });
      targetIndexes = Array.from(indicesSet).sort((a, b) => a - b);
      if (targetIndexes.length === 0) {
        targetIndexes = simpleIndices;
      }
    }
    
    const nameCounts = new Map<string, number>();
    const baseNames = targetIndexes.map(idx => {
      const hVal = headers[idx] ? headers[idx].trim() : '';
      const name = hVal || getColLetter(idx);
      nameCounts.set(name, (nameCounts.get(name) || 0) + 1);
      return { idx, name };
    });

    const filteredColumns = baseNames.map(({ idx, name }) => {
      const finalName = (nameCounts.get(name) || 0) > 1 
        ? `${name} (${getColLetter(idx)})` 
        : name;
      return {
        name: finalName,
        enabled: true,
        rawIndex: idx
      };
    });
    
    setColumns(filteredColumns);
    
    const processedData = rawRows.slice(3) // Skip Rows 1, 2, 3 (index 0, 1, 2)
      .map((row, idx) => ({ row, originalIndex: idx + 3 }))
      .filter(({ row }) => {
        const colA = String(row[0] || '').trim();
        const colB = String(row[1] || '').trim();
        
        // Skip if:
        // - Col A is empty
        // - Col A has SKIPROW
        // - Col A has #N/A
        // - Col B is blank
        if (colA === '' || colA.includes('SKIPROW') || colA.includes('#N/A')) return false;
        if (colB === '') return false;
        
        // Security level filtering based on Column DF (index 109)
        const securityVal = row[109] || '';
        const rowSecurityLvl = getSecurityLevelNum(securityVal);
        const allowedLvl = savedSecurityLevel !== null ? savedSecurityLevel : 0;
        if (rowSecurityLvl > allowedLvl) return false;
        
        return true;
      })
      .map(({ row, originalIndex }) => {
        const rowObj: any = { _rawRow: row, _originalIndex: originalIndex };
        targetIndexes.forEach((colIdx, listIdx) => {
          const headerName = filteredColumns[listIdx]?.name || getColLetter(colIdx);
          rowObj[headerName] = row[colIdx] || '';
        });
        return rowObj;
      });
    
    setData(processedData);
    
    findRecord(processedData, filteredColumns, searchKey, selectedGalaxy, selectedRegion, discovererName, surveyorName);
  }, [rawRows, reportType, enabledCustomColumns, searchKey, selectedGalaxy, selectedRegion, discovererName, surveyorName, savedSecurityLevel]);

  const handleSearch = async () => {
    setSearchLoading(true);
    await fetchData();
  };

  const findRecord = (
    sourceData: any[], 
    sourceCols: ColumnConfig[], 
    civTerm?: string, 
    galTerm?: string,
    regionTerm?: string,
    discTerm?: string,
    survTerm?: string
  ) => {
    const currentCivTerm = (civTerm ?? searchKey).trim().toLowerCase();
    const currentGalTerm = (galTerm ?? selectedGalaxy).trim().toLowerCase();
    const currentRegionTerm = (regionTerm ?? selectedRegion).trim().toLowerCase();
    const currentDiscTerm = (discTerm ?? discovererName).trim().toLowerCase();
    const currentSurvTerm = (survTerm ?? surveyorName).trim().toLowerCase();
    
    const getRawVal = (r: any, rawIdx: number) => {
      if (r._rawRow) return String(r._rawRow[rawIdx] || '').trim();
      const colMatch = sourceCols.find(c => c.rawIndex === rawIdx);
      return colMatch ? String(r[colMatch.name] || '').trim() : '';
    };

    const matches = sourceData.filter(row => {
      const cVal = getRawVal(row, 18); // Civ
      const gVal = getRawVal(row, 0);  // Galaxy
      const rVal = getRawVal(row, 1);  // Region
      const dVal = getRawVal(row, 11); // Discoverer
      const kVal = getRawVal(row, 10); // Surveyor

      // Civ matches: if currentCivTerm is ALL or empty, match all.
      // IF currentCivTerm is 'none', match only empty fields.
      // IF currentCivTerm is 'agt', treat as 'Alliance of Galactic Travellers'
      let civMatch = false;
      if (!currentCivTerm || currentCivTerm === 'all') {
        civMatch = true;
      } else if (currentCivTerm === 'none') {
        civMatch = cVal === '';
      } else {
        const mappedTerm = currentCivTerm === 'agt' ? 'alliance of galactic travellers' : currentCivTerm;
        civMatch = cVal.toLowerCase().includes(mappedTerm);
      }

      // Galaxy matches: if currentGalTerm is ALL or empty, match all
      let galMatch = false;
      if (!currentGalTerm || currentGalTerm === 'all') {
        galMatch = true;
      } else {
        galMatch = gVal.toLowerCase().includes(currentGalTerm);
      }

      // Region matches: if currentRegionTerm is ALL or empty, match all
      let regMatch = false;
      if (!currentRegionTerm || currentRegionTerm === 'all') {
        regMatch = true;
      } else {
        regMatch = rVal.toLowerCase().includes(currentRegionTerm);
      }

      // Discoverer matches: optional
      let discMatch = true;
      if (currentDiscTerm && currentDiscTerm !== 'all') {
        discMatch = dVal.toLowerCase().includes(currentDiscTerm);
      }

      // Surveyor matches: optional
      let survMatch = true;
      if (currentSurvTerm && currentSurvTerm !== 'all') {
        survMatch = kVal.toLowerCase().includes(currentSurvTerm);
      }

      return civMatch && galMatch && regMatch && discMatch && survMatch;
    });

    // Sort by Column A then Column B
    const sortedMatches = [...matches].sort((a, b) => {
      const galA = getRawVal(a, 0).toLowerCase();
      const galB = getRawVal(b, 0).toLowerCase();
      if (galA !== galB) return galA.localeCompare(galB);
      
      const regionA = getRawVal(a, 1).toLowerCase();
      const regionB = getRawVal(b, 1).toLowerCase();
      return regionA.localeCompare(regionB);
    });

    if (sortedMatches.length > 0) {
      setMatchedRecords(sortedMatches);
      setError(null);
    } else {
      setMatchedRecords([]);
      setError(`No records found for the selected criteria.`);
    }
  };

  const formatTimestamp = () => {
    const now = new Date();
    const pad = (num: number) => String(num).padStart(2, '0');
    return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}_${pad(now.getMinutes())}_${pad(now.getSeconds())}`;
  };

  const downloadFullReportPdf = () => {
    const name = getCookie('agt_traveller_name') || savedTravellerName;
    const id = getCookie('agt_traveller_id') || savedTravellerId;

    if (!name || !id) {
      showAlert("PDF Report and Export CSV is only available to registered AGT Travellers. Enter your credientials in the setting menu");
      return;
    }

    if (sortedMatchedRecords.length === 0) return;
    
    const activeCols = columns.filter(col => col.enabled);
    const URL_COL_INDICES = [76, 77, 78, 79, 80, 81, 105, 106, 107];

    const getPdfColumnWidths = (cols: ColumnConfig[]): number[] => {
      const baseWidths = cols.map(col => {
        const idx = col.rawIndex;
        if (idx === undefined) return 20;
        if (URL_COL_INDICES.includes(idx)) return 10;
        if (idx === 0) return 9;    // Galaxy (Index 0): compact
        if (idx === 1) return 18;   // Region Name (Index 1): Col 2 reduced in width, can word-wrap
        if (idx === 2) return 14;   // System Name (Index 2): Col 3
        if (idx === 6) return 18;   // Original Name (Index 6): Col 7 reduced in width, can word-wrap
        if (idx === 7) return 27;   // Galactic Coordinates (Index 7): Col 8 wide, guaranteed single line & fully visible!
        if (idx === 8) return 20;   // Glyph Code (Index 8): Col 9 wide, guaranteed single line & fully visible!
        if (idx === 10) return 11;  // Surveyor Name (Index 10)
        if (idx === 11) return 11;  // Discoverer Name (Index 11)
        if (idx === 14) return 16;  // Discovery Date (Index 14)
        if (idx === 15) return 16;  // Survey Date (Index 15)
        if (idx === 18) return 11;  // Civilized? (Index 18)
        return 20;                  // default
      });
      
      const sum = baseWidths.reduce((a, b) => a + b, 0) || 1;
      return baseWidths.map(w => (w / sum) * 257);
    };

    const getPdfCellValue = (rawVal: any, colIdx: number | undefined): string => {
      const isUrlCol = colIdx !== undefined && URL_COL_INDICES.includes(colIdx);
      const isTargetBlueUrl = isTargetColumnWithUrl(colIdx, rawVal);
      const isUrl = (colIdx !== undefined && colIdx >= 76 && colIdx <= 80) || String(rawVal || '').trim().startsWith('http');
      
      if (isUrlCol || isTargetBlueUrl || isUrl) {
        return 'LINK';
      }
      
      let val = getDisplayValue(rawVal, colIdx);
      if (!val) return '-';
      
      const truncatableIndices = [2, 10, 11, 18];
      if (colIdx !== undefined && truncatableIndices.includes(colIdx)) {
        const len = val.length;
        if (len > 10) {
          const keepLen = Math.ceil(len * 0.75); // Keeps at least 75% of characters (max 25% truncation)
          if (keepLen < len) {
            val = val.substring(0, keepLen) + '...';
          }
        }
      }
      return val;
    };
    
    // Check if any row in the PDF will wrap into more than 3 lines.
    let exceedsThreeLines = false;
    
    if (activeCols.length > 13) {
      exceedsThreeLines = true;
    } else {
      const widthsList = getPdfColumnWidths(activeCols);
      for (const record of sortedMatchedRecords) {
        for (let i = 0; i < activeCols.length; i++) {
          const col = activeCols[i];
          const rawVal = record[col.name];
          const val = getPdfCellValue(rawVal, col.rawIndex);
          const colWidth = widthsList[i];
          
          const isUrlCol = col.rawIndex !== undefined && URL_COL_INDICES.includes(col.rawIndex);
          const isSingleLineCol = isUrlCol || col.rawIndex === 7 || col.rawIndex === 8;
          
          let linesCount = 1;
          if (!isSingleLineCol) {
            const charLimit = Math.max(5, Math.floor(colWidth / 1.5));
            const len = val.length;
            linesCount = Math.ceil(len / charLimit) || 1;
          }
          
          if (linesCount > 3) {
            exceedsThreeLines = true;
            break;
          }
        }
        if (exceedsThreeLines) break;
      }
    }
    
    if (exceedsThreeLines) {
      setPdfErrorMsg("Too many columns for PDF report, reduce the column selections or choose CSV export");
      return;
    }

    setReportGeneratingLoading(true);

    setTimeout(() => {
      try {
        const doc = new jsPDF('l', 'mm', 'a4'); // Landscape for tables

        // --- 1. COVER PAGE (Page 1) ---
        // Title in hex color FF0500 (RGB: 255, 5, 0), horizontally centered under AGTIcon
        doc.setFont("Helvetica", "bold");
        doc.setFontSize(26);
        doc.setTextColor(255, 5, 0);
        doc.text("AGT System Report", 148.5, 87, { align: 'center' });

        // All other text on the cover page is black
        doc.setFont("Helvetica", "normal");
        doc.setTextColor(0, 0, 0);
        doc.setFontSize(14);
        
        const formattedReportType = reportType === 'simple' ? 'Simple' : reportType === 'detailed' ? 'Detailed' : 'Custom';
        doc.text(`Report Type: ${formattedReportType}`, 148.5, 102, { align: 'center' });

        // Civilization translation rule
        let civDisplay = searchKey || 'All';
        if (civDisplay.trim().toLowerCase() === 'agt') {
          civDisplay = 'Alliance of Galactic Travellers';
        }

        doc.setFontSize(11);
        doc.text(`Civilization: ${civDisplay}`, 148.5, 115, { align: 'center' });
        
        // Combined Galaxy & Region filter criteria line
        const galaxyVal = selectedGalaxy || 'All';
        const regionVal = selectedRegion || 'All';
        doc.text(`Galaxy: ${galaxyVal} / Region: ${regionVal}`, 148.5, 122, { align: 'center' });

        // Discoverer & Surveyor line
        const discVal = discovererName.trim() || 'All';
        const survVal = surveyorName.trim() || 'All';
        doc.text(`Discoverer: ${discVal} / Surveyor: ${survVal}`, 148.5, 129, { align: 'center' });

        // Date of Report line
        doc.text(`Date of Report: ${new Date().toLocaleString()}`, 148.5, 138, { align: 'center' });

        // Result Count line
        doc.text(`Result Count: ${matchedRecords.length} Verified Entries`, 148.5, 145, { align: 'center' });

        // Top logo center image placement
        try {
          const imgElements = document.getElementsByTagName('img');
          let logoImg: HTMLImageElement | null = null;
          for (let i = 0; i < imgElements.length; i++) {
            if (imgElements[i].alt === 'AGT Logo') {
              logoImg = imgElements[i];
              break;
            }
          }
          if (logoImg && logoImg.complete) {
            doc.addImage(logoImg, 'PNG', 133.5, 42, 30, 30);
          } else {
            const tempImg = new Image();
            tempImg.src = logoUrl;
            if (tempImg.complete) {
              doc.addImage(tempImg, 'PNG', 133.5, 42, 30, 30);
            }
          }
        } catch (err) {
          console.warn('Cover page logo inject failed:', err);
        }

        // Star Systems Icon centering at bottom segment midpoint y=162.5
        try {
          const imgElements = document.getElementsByTagName('img');
          let starImg: HTMLImageElement | null = null;
          for (let i = 0; i < imgElements.length; i++) {
            if (imgElements[i].alt === 'Star Systems') {
              starImg = imgElements[i];
              break;
            }
          }
          if (starImg && starImg.complete) {
            doc.addImage(starImg, 'PNG', 133.5, 162.5, 30, 30);
          } else {
            const tempImg = new Image();
            tempImg.src = starSystemsIcon;
            if (tempImg.complete) {
              doc.addImage(tempImg, 'PNG', 133.5, 162.5, 30, 30);
            }
          }
        } catch (err) {
          console.warn('Cover page star logo inject failed:', err);
        }

        // --- 2. ADD PAGE 2 (TABLE DATA) ---
        doc.addPage();

        const urlMap = new Map<string, string>();
        const tableData = sortedMatchedRecords.map((record, rIdx) => 
          activeCols.map((col, cIdx) => {
            const rawVal = record[col.name];
            const val = getPdfCellValue(rawVal, col.rawIndex);
            
            const isUrlCol = col.rawIndex !== undefined && URL_COL_INDICES.includes(col.rawIndex);
            const isTargetBlueUrl = isTargetColumnWithUrl(col.rawIndex, rawVal);
            const isUrl = (col.rawIndex !== undefined && col.rawIndex >= 76 && col.rawIndex <= 80) || String(rawVal || '').trim().startsWith('http');
            
            if ((isUrlCol || isTargetBlueUrl || isUrl) && rawVal) {
              urlMap.set(`${rIdx}-${cIdx}`, String(rawVal).trim());
            }
            return val;
          })
        );

        // Add total row to PDF
        const countFieldName = columns[0]?.name;
        const totalRow = activeCols.map(col => {
          if (col.name === countFieldName) return `Count: ${sortedMatchedRecords.length}`;
          return '';
        });
        tableData.push(totalRow);

        const widthsList = getPdfColumnWidths(activeCols);
        const columnStyles: Record<number, any> = {};
        
        activeCols.forEach((col, idx) => {
          const w = widthsList[idx];
          const isUrlCol = col.rawIndex !== undefined && URL_COL_INDICES.includes(col.rawIndex);
          const isNoTruncateSingleLineCol = col.rawIndex === 7 || col.rawIndex === 8;
          
          if (isNoTruncateSingleLineCol) {
            columnStyles[idx] = {
              cellWidth: w,
              overflow: 'visible'
            };
          } else {
            columnStyles[idx] = {
              cellWidth: w,
              overflow: isUrlCol ? 'ellipsize' : 'linebreak'
            };
          }
        });

        autoTable(doc, {
          startY: 25,
          head: [activeCols.map(col => {
            if (col.rawIndex === 76) return "Wiki";
            if (col.rawIndex === 10) return "Surveyor";
            return col.name;
          })],
          body: tableData,
          theme: 'grid',
          columnStyles: columnStyles,
          headStyles: { fillColor: [20, 20, 20], textColor: [255, 255, 255], fontSize: 8, overflow: 'linebreak' },
          bodyStyles: { fontSize: 8 },
          footStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], fontStyle: 'bold' },
          margin: { top: 25, bottom: 20, left: 20, right: 20 },
          didParseCell: (dataCell) => {
            if (dataCell.section === 'head') {
              dataCell.cell.styles.overflow = 'linebreak';
            }
            if (dataCell.row.index === tableData.length - 1) {
              dataCell.cell.styles.fillColor = [220, 220, 220];
              dataCell.cell.styles.fontStyle = 'bold';
            }
            
            const cellKey = `${dataCell.row.index}-${dataCell.column.index}`;
            if (urlMap.has(cellKey)) {
              dataCell.cell.styles.textColor = [0, 0, 255];
            }
          },
          didDrawCell: (dataCell) => {
            const cellKey = `${dataCell.row.index}-${dataCell.column.index}`;
            const url = urlMap.get(cellKey);
            if (url && dataCell.section === 'body') {
              doc.link(dataCell.cell.x, dataCell.cell.y, dataCell.cell.width, dataCell.cell.height, { url });
            }
          }
        });

        // --- 3. POST-PROCESS PAGES FOR HEADERS & FOOTERS (Page 2 onwards) ---
        const totalPagesCount = doc.getNumberOfPages();
        for (let i = 2; i <= totalPagesCount; i++) {
          doc.setPage(i);
          
          // Header left justified: Small logo (10x10) and text
          try {
            const imgElements = document.getElementsByTagName('img');
            let logoImg: HTMLImageElement | null = null;
            for (let j = 0; j < imgElements.length; j++) {
              if (imgElements[j].alt === 'AGT Logo') {
                logoImg = imgElements[j];
                break;
              }
            }
            if (logoImg && logoImg.complete) {
              doc.addImage(logoImg, 'PNG', 20, 10, 10, 10);
            } else {
              const tempImg = new Image();
              tempImg.src = logoUrl;
              if (tempImg.complete) {
                doc.addImage(tempImg, 'PNG', 20, 10, 10, 10);
              }
            }
          } catch (err) {
            console.warn('Page header logo inject failed:', err);
          }

          doc.setFont("Helvetica", "bold");
          doc.setFontSize(10);
          doc.setTextColor(0, 0, 0);
          doc.text("AGT System Report", 32, 16.5);

          // Header right justified: Page number starting at 1 following cover page
          doc.setFont("Helvetica", "normal");
          doc.setFontSize(9);
          doc.text(`Page ${i - 1}`, 277, 16.5, { align: 'right' });

          // Footer left justified: Date of Report
          doc.setFont("Helvetica", "normal");
          doc.setFontSize(8);
          doc.text(`Date of Report: ${new Date().toLocaleString()}`, 20, 202);
        }

        doc.save(`AGT  System Report - ${formatTimestamp()}.pdf`);
      } catch (err) {
        console.error('Error generating PDF:', err);
      } finally {
        setReportGeneratingLoading(false);
      }
    }, 400);
  };

  const downloadCsv = () => {
    const name = getCookie('agt_traveller_name') || savedTravellerName;
    const id = getCookie('agt_traveller_id') || savedTravellerId;

    if (!name || !id) {
      showAlert("PDF Report and Export CSV is only available to registered AGT Travellers. Enter your credientials in the setting menu");
      return;
    }

    if (sortedMatchedRecords.length === 0) return;
    setReportGeneratingLoading(true);

    setTimeout(() => {
      try {
        const activeCols = columns.filter(col => col.enabled);
        const csvData = sortedMatchedRecords.map(record => {
          const row: any = {};
          activeCols.forEach(col => {
            row[col.name] = getDisplayValue(record[col.name], col.rawIndex);
          });
          return row;
        });
        
        const csv = Papa.unparse(csvData);
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        
        if (link.download !== undefined) {
          const url = URL.createObjectURL(blob);
          link.setAttribute('href', url);
          link.setAttribute('download', `AGT  System CSV Export- ${formatTimestamp()}.csv`);
          link.style.visibility = 'hidden';
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
        }
      } catch (err) {
        console.error('Error generating CSV:', err);
      } finally {
        setReportGeneratingLoading(false);
      }
    }, 400);
  };

  const toggleColumn = (name: string) => {
    setColumns(prev => prev.map(c => c.name === name ? { ...c, enabled: !c.enabled } : c));
  };

  const handleVerifyTraveller = async () => {
    const trimmedName = enteredName.trim();
    const trimmedId = enteredId.trim();

    if (!trimmedName || !trimmedId) {
      showAlert("Please enter both your Traveller Name and Traveller ID.");
      return;
    }

    if (trimmedName.length > 42) {
      showAlert("Traveller Name cannot exceed 42 characters.");
      return;
    }

    // ID format validation: ########-????-#### (case-insensitive)
    const idRegex = /^\d{8}-[a-zA-Z0-9]{4}-\d{4}$/;
    if (!idRegex.test(trimmedId)) {
      showAlert("Invalid AGT Traveller ID format. Must follow ########-????-#### format.");
      return;
    }

    setIsVerifying(true);
    setVerificationError(false);

    try {
      const verifySheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSOZq3Cl2e0aNqzXdLRe63HuM7PlqGH3HnS_-0x6P_CYnGDJlK5QvI-YjU0lNaOgLyp3uoktS4WIXyK/pub?gid=505079663&single=true&output=tsv";
      const response = await fetch(verifySheetUrl);
      if (!response.ok) {
        throw new Error("Unable to connect to registry server.");
      }
      const tsvText = await response.text();
      const lines = tsvText.split(/\r?\n/);
      const rows = lines.map(line => line.split('\t'));

      let matchedRow: string[] | null = null;
      
      const decodeXOR = (encodedText: string): string => {
        const key = 969; 
        let decoded = ""; 
        for (let i = 0; i < encodedText.length; i++) { 
          let charCode = encodedText.charCodeAt(i); 
          let originalCharCode = charCode ^ key; 
          decoded += String.fromCharCode(originalCharCode); 
        } 
        return decoded; 
      };

      for (const row of rows) {
        if (row.length < 3) continue;
        const colA = row[0].trim();
        const colB = row[1].trim();

        if (colA.toLowerCase() === trimmedName.toLowerCase()) {
          const decodedB = decodeXOR(colB).trim();
          if (decodedB.toLowerCase() === trimmedId.toLowerCase()) {
            matchedRow = row;
            break;
          }
        }
      }

      if (matchedRow) {
        const colC = matchedRow[2].trim();
        const levelNum = getSecurityLevelNum(colC);

        // Save traveller name, AGT Traveller ID and security level to cookie
        setCookie('agt_traveller_name', trimmedName);
        setCookie('agt_traveller_id', trimmedId);
        setCookie('agt_security_level', String(levelNum));

        // Read and verify cookie save success
        const testName = getCookie('agt_traveller_name');
        const testId = getCookie('agt_traveller_id');
        const testLevel = getCookie('agt_security_level');

        if (testName === trimmedName && testId === trimmedId && testLevel === String(levelNum)) {
          setSavedTravellerName(trimmedName);
          setSavedTravellerId(trimmedId);
          setSavedSecurityLevel(levelNum);
          showAlert("Verification successful, setting saved");
        } else {
          // If cookie doesn't save properly, we still activate the session values so they are usable
          setSavedTravellerName(trimmedName);
          setSavedTravellerId(trimmedId);
          setSavedSecurityLevel(levelNum);
          showAlert("Verification successful, setting save error");
        }
      } else {
        setVerificationError(true);
        showAlert("Verification unsuccessful");
      }
    } catch (err) {
      console.error(err);
      showAlert("Verification failed due to a network or parsing issue.");
    } finally {
      setIsVerifying(false);
    }
  };

  const handleClearTraveller = () => {
    deleteCookie('agt_traveller_name');
    deleteCookie('agt_traveller_id');
    deleteCookie('agt_security_level');

    const testName = getCookie('agt_traveller_name');
    const testId = getCookie('agt_traveller_id');

    if (testName === null && testId === null) {
      setSavedTravellerName(null);
      setSavedTravellerId(null);
      setSavedSecurityLevel(null);
      setEnteredName('');
      setEnteredId('');
      setVerificationError(false);
      showAlert("Clearing successful");
    } else {
      showAlert("Clearing failed");
    }
  };

  const activeColumnsCount = useMemo(() => columns.filter(c => c.enabled).length, [columns]);

  const sortedMatchedRecords = useMemo(() => {
    if (!sortColumn || !sortDirection) return matchedRecords;
    
    return [...matchedRecords].sort((a, b) => {
      const valA = a[sortColumn];
      const valB = b[sortColumn];
      
      if (valA === undefined || valB === undefined) return 0;
      
      const strA = String(valA || '').trim();
      const strB = String(valB || '').trim();
      
      // Try numeric sort
      const cleanA = strA.replace(/[$,%]/g, '');
      const cleanB = strB.replace(/[$,%]/g, '');
      const numA = parseFloat(cleanA);
      const numB = parseFloat(cleanB);
      
      const isNumA = !isNaN(numA) && isFinite(numA) && cleanA !== '';
      const isNumB = !isNaN(numB) && isFinite(numB) && cleanB !== '';
      
      if (isNumA && isNumB) {
        return sortDirection === 'asc' ? numA - numB : numB - numA;
      }
      
      return sortDirection === 'asc'
        ? strA.localeCompare(strB, undefined, { numeric: true, sensitivity: 'base' })
        : strB.localeCompare(strA, undefined, { numeric: true, sensitivity: 'base' });
    });
  }, [matchedRecords, sortColumn, sortDirection]);

  const totalPoints = useMemo(() => {
    return sortedMatchedRecords.length;
  }, [sortedMatchedRecords]);

  const totalPages = useMemo(() => {
    return Math.ceil(sortedMatchedRecords.length / pageSize);
  }, [sortedMatchedRecords.length, pageSize]);

  const paginatedRecords = useMemo(() => {
    const startIdx = (currentPage - 1) * pageSize;
    return sortedMatchedRecords.slice(startIdx, startIdx + pageSize);
  }, [sortedMatchedRecords, currentPage, pageSize]);

  const omittedCount = useMemo(() => {
    if (!rawRows || rawRows.length < 4) return 0;
    
    const allowedLvl = savedSecurityLevel !== null ? savedSecurityLevel : 0;
    
    const currentCivTerm = searchKey.trim().toLowerCase();
    const currentGalTerm = selectedGalaxy.trim().toLowerCase();
    const currentRegionTerm = selectedRegion.trim().toLowerCase();
    const currentDiscTerm = discovererName.trim().toLowerCase();
    const currentSurvTerm = surveyorName.trim().toLowerCase();

    let count = 0;

    // Skip Rows 1, 2, 3 (index 0, 1, 2)
    const validRows = rawRows.slice(3);
    for (const row of validRows) {
      const colA = String(row[0] || '').trim();
      const colB = String(row[1] || '').trim();
      
      if (colA === '' || colA.includes('SKIPROW') || colA.includes('#N/A')) continue;
      if (colB === '') continue;

      // Check search match (following findRecord criteria exactly)
      const cVal = String(row[18] || '').trim(); // Civ
      const gVal = String(row[0] || '').trim();  // Galaxy
      const rVal = String(row[1] || '').trim();  // Region
      const dVal = String(row[11] || '').trim(); // Discoverer
      const kVal = String(row[10] || '').trim(); // Surveyor

      // Check Civ
      let civMatch = false;
      if (!currentCivTerm || currentCivTerm === 'all') {
        civMatch = true;
      } else if (currentCivTerm === 'none') {
        civMatch = cVal === '';
      } else {
        const mappedTerm = currentCivTerm === 'agt' ? 'alliance of galactic travellers' : currentCivTerm;
        civMatch = cVal.toLowerCase().includes(mappedTerm);
      }

      // Check Galaxy
      let galMatch = false;
      if (!currentGalTerm || currentGalTerm === 'all') {
        galMatch = true;
      } else {
        galMatch = gVal.toLowerCase().includes(currentGalTerm);
      }

      // Check Region
      let regMatch = false;
      if (!currentRegionTerm || currentRegionTerm === 'all') {
        regMatch = true;
      } else {
        regMatch = rVal.toLowerCase().includes(currentRegionTerm);
      }

      // Check Discoverer
      let discMatch = true;
      if (currentDiscTerm && currentDiscTerm !== 'all') {
        discMatch = dVal.toLowerCase().includes(currentDiscTerm);
      }

      // Check Surveyor
      let survMatch = true;
      if (currentSurvTerm && currentSurvTerm !== 'all') {
        survMatch = kVal.toLowerCase().includes(currentSurvTerm);
      }

      if (civMatch && galMatch && regMatch && discMatch && survMatch) {
        // It matches the search query! Now check if it is omitted solely due to security level
        const securityVal = row[109] || '';
        const rowSecurityLvl = getSecurityLevelNum(securityVal);
        if (rowSecurityLvl > allowedLvl) {
          count++;
        }
      }
    }
    return count;
  }, [rawRows, searchKey, selectedGalaxy, selectedRegion, discovererName, surveyorName, savedSecurityLevel]);

  const multiplier = parseFloat(fontScale.replace('x', '')) || 1.0;

  return (
    <div 
      onMouseDown={handleManualPlay}
      onTouchStart={handleManualPlay}
      className={`min-h-screen bg-[#0a0a0a] text-agt-orange font-sans selection:bg-agt-orange selection:text-black ${fontScale !== '1x' ? 'font-scale-active' : ''}`}
    >
      {/* Dynamic Font Scale Injector */}
      <style>{`
        .font-scale-active .text-\\[7px\\] { font-size: calc(7px * ${multiplier}) !important; }
        .font-scale-active .text-\\[8px\\] { font-size: calc(8px * ${multiplier}) !important; }
        .font-scale-active .text-\\[9px\\] { font-size: calc(9px * ${multiplier}) !important; }
        .font-scale-active .text-\\[10px\\] { font-size: calc(10px * ${multiplier}) !important; }
        .font-scale-active .text-\\[11px\\] { font-size: calc(11px * ${multiplier}) !important; }
        .font-scale-active .text-xs { font-size: calc(12px * ${multiplier}) !important; }
        .font-scale-active .text-sm { font-size: calc(14px * ${multiplier}) !important; }
        .font-scale-active .text-base { font-size: calc(16px * ${multiplier}) !important; }
        .font-scale-active .text-md { font-size: calc(16px * ${multiplier}) !important; }
        .font-scale-active .text-lg { font-size: calc(18px * ${multiplier}) !important; }
        .font-scale-active .text-xl { font-size: calc(20px * ${multiplier}) !important; }
        .font-scale-active .text-2xl { font-size: calc(24px * ${multiplier}) !important; }
        .font-scale-active .text-3xl { font-size: calc(30px * ${multiplier}) !important; }
        .font-scale-active .text-4xl { font-size: calc(36px * ${multiplier}) !important; }
        .font-scale-active .text-5xl { font-size: calc(48px * ${multiplier}) !important; }
        .font-scale-active .text-6xl { font-size: calc(60px * ${multiplier}) !important; }
        
        .font-scale-active input, 
        .font-scale-active select, 
        .font-scale-active textarea, 
        .font-scale-active button { 
          font-size: calc(100% * ${multiplier}) !important; 
        }
      `}</style>

      {/* Horizontally spinning logo processing overlay */}
      <AnimatePresence>
        {(searchLoading || reportGeneratingLoading) && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[200] bg-black/95 flex flex-col items-center justify-center p-6 text-center select-none"
          >
            <div className="relative w-48 h-48 flex items-center justify-center mb-8">
              {/* Outer orbit circle */}
              <div className="absolute inset-0 rounded-full border border-[#FF0500]/20 animate-pulse"></div>
              
              {/* Horizontally rotating logo img */}
              <motion.img 
                src="/AGTIcon.png" 
                alt="AGT Logo" 
                className="w-36 h-36 object-contain"
                animate={{ rotateY: 360 }}
                transition={{ repeat: Infinity, duration: 2, ease: "linear" }}
              />
            </div>
            <h3 className="text-[#FFB451] text-lg font-bold uppercase tracking-[0.25em] mb-2 animate-pulse">
              {reportGeneratingLoading ? t("Creating Report") : t("Searching AGT Galactic Archives")}
            </h3>
            <p className="text-[#FFB451]/50 text-xs font-mono uppercase tracking-[0.2em]">
              {reportGeneratingLoading ? t("Compiling intelligence packet...") : t("Establishing high-speed subspace connection...")}
            </p>
          </motion.div>
        )}
      </AnimatePresence>

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
              <h1 className="font-bold text-xs tracking-[0.2em] uppercase text-agt-orange">{t("Alliance of Galactic Travellers")}</h1>
              <span className="text-[9px] text-agt-orange uppercase tracking-[0.3em] font-bold">{t("AGT System Report Tool")}</span>
            </div>
          </div>
          
          <div className="flex items-center gap-6">
            <div className="hidden md:block text-[9px] text-agt-orange/30 tracking-widest font-mono">
              {t("STATUS: ")}<span className={
                loading ? 'text-yellow-500' :
                sheetUrl ? 'text-emerald-500' : 
                'text-red-500'
              }>
                {loading ? t('SYNCING') : sheetUrl ? t('CONNECTED') : t('DISCONNECTED')}
              </span>
            </div>
            {savedTravellerName && savedTravellerId ? (
              <div 
                className="border border-green-500 text-green-500 rounded-lg px-2.5 py-1 text-[11px] font-mono tracking-wider bg-green-950/20 select-none max-w-[150px] truncate"
                id="header-traveller-display"
                title={savedTravellerName}
              >
                {savedTravellerName.slice(0, 20)}
              </div>
            ) : (
              <div 
                className="border border-[#FF0500] text-[#FF0500] rounded-lg px-2.5 py-1 text-[11px] font-mono tracking-wider bg-red-950/20 select-none"
                id="header-traveller-display"
              >
                {t("Public User")}
              </div>
            )}
            <motion.button 
              onClick={() => setShowSettings(!showSettings)}
              className="p-2 hover:bg-[#FF0500]/5 rounded-lg transition-colors relative group text-[#FF0500]"
              title="Settings"
              id="settings-btn"
              whileHover={{ rotate: 180 }}
              transition={{ duration: 0.3 }}
              whileTap={{ scale: 0.9, rotate: -90 }}
            >
              <Settings className="w-5 h-5" />
              {!sheetUrl && (
                <span className="absolute top-0 right-0 w-1.5 h-1.5 bg-[#FF0500] rounded-full shadow-[0_0_5px_rgba(255,5,0,0.5)]"></span>
              )}
            </motion.button>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-16">
        <div className="flex flex-col gap-16">
          
          {/* Main Search Logic Container - centered aesthetic */}
          <div className="flex flex-col items-center space-y-12">
            <div className="w-full max-w-xl text-center space-y-6 flex flex-col items-center">
              <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
                <img 
                  src={starSystemsIcon} 
                  alt="Star Systems" 
                  className="w-12 h-12 object-contain" 
                  id="star-systems-icon-img"
                />
                <h2 className="text-4xl font-light tracking-tight text-[#FFB451]" id="main-title">
                  {t("AGT System Report Tool")}
                </h2>
              </div>
              
              {/* Report Mode Selector */}
              <div className="flex justify-center mb-4">
                <div className="inline-flex p-1.5 bg-[#161616] border-2 border-[#FF0500] rounded-xl gap-1">
                  {(['simple', 'detailed', 'custom'] as const).map((type) => {
                    const isActive = reportType === type;
                    const reportLabels: Record<typeof type, string> = {
                      simple: 'Simple Report',
                      detailed: 'Detailed Report',
                      custom: 'Custom Report'
                    };
                    return (
                      <button
                        key={type}
                        onClick={() => {
                          setReportType(type);
                          setData([]);
                          setMatchedRecords([]);
                        }}
                        className={`px-5 py-2.5 rounded-lg text-[10px] uppercase tracking-widest font-black transition-all ${
                          isActive 
                            ? 'bg-[#E25530] text-white shadow-[0_0_15px_rgba(226,85,48,0.6)] border border-white/20 scale-102' 
                            : 'bg-[#E25530]/30 text-white/70 hover:bg-[#E25530]/60 hover:text-white'
                        }`}
                        id={`report-mode-${type}`}
                      >
                        {t(reportLabels[type])}
                      </button>
                    );
                  })}
                </div>
              </div>

            </div>

            <div className="w-full max-w-4xl space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 items-start">
                {/* Civilization Autocomplete */}
                <div className="space-y-2">
                  <div className="space-y-1 text-left ml-1 select-none">
                    <p className="text-[#FFB451] text-xs font-bold tracking-widest uppercase">{t("Civilization")}</p>
                  </div>
                  <AutocompleteInput
                    id="civilization-select"
                    value={searchKey}
                    onChange={(val) => {
                      setSearchKey(val);
                      if (data.length) {
                        findRecord(data, columns, val, selectedGalaxy, selectedRegion, discovererName, surveyorName);
                      }
                    }}
                    suggestions={civSuggestions}
                    placeholder={t("Enter/Choose Civilization...")}
                    icon={<Search className="h-5 w-5" />}
                  />
                </div>

                {/* Galaxy Autocomplete */}
                <div className="space-y-2">
                  <div className="space-y-1 text-left ml-1 select-none">
                    <p className="text-[#FFB451] text-xs font-bold tracking-widest uppercase">{t("Galaxy")}</p>
                  </div>
                  <AutocompleteInput
                    id="galaxy-select"
                    value={selectedGalaxy}
                    onChange={(val) => {
                      setSelectedGalaxy(val);
                      if (data.length) {
                        findRecord(data, columns, searchKey, val, selectedRegion, discovererName, surveyorName);
                      }
                    }}
                    suggestions={galaxySuggestions}
                    placeholder={t("Enter/Choose Galaxy...")}
                    icon={<Globe className="h-5 w-5" />}
                  />
                </div>

                {/* Region Autocomplete */}
                <div className="space-y-2">
                  <div className="space-y-1 text-left ml-1 select-none">
                    <p className="text-[#FFB451] text-xs font-bold tracking-widest uppercase">{t("Region")}</p>
                  </div>
                  <AutocompleteInput
                    id="region-select"
                    value={selectedRegion}
                    onChange={(val) => {
                      setSelectedRegion(val);
                      if (data.length) {
                        findRecord(data, columns, searchKey, selectedGalaxy, val, discovererName, surveyorName);
                      }
                    }}
                    suggestions={regionSuggestions}
                    placeholder={t("Enter/Choose Region...")}
                    icon={<Search className="h-5 w-5" />}
                  />
                </div>
              </div>

              {/* Optional discoverer & surveyor rows */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-start w-full max-w-2xl mx-auto">
                {/* Discoverer Autocomplete (Optional) */}
                <div className="space-y-2">
                  <div className="space-y-1 text-left ml-1 select-none">
                    <p className="text-[#FFB451] text-xs font-bold tracking-widest uppercase opacity-75 font-bold">{t("Discoverer Name")}</p>
                  </div>
                  <AutocompleteInput
                    id="discoverer-select"
                    value={discovererName}
                    onChange={(val) => {
                      setDiscovererName(val);
                      if (data.length) {
                        findRecord(data, columns, searchKey, selectedGalaxy, selectedRegion, val, surveyorName);
                      }
                    }}
                    suggestions={discovererSuggestions}
                    placeholder={t("Discoverer Name")}
                    icon={<Search className="h-5 w-5" />}
                  />
                </div>

                {/* Surveyor Autocomplete (Optional) */}
                <div className="space-y-2">
                  <div className="space-y-1 text-left ml-1 select-none">
                    <p className="text-[#FFB451] text-xs font-bold tracking-widest uppercase opacity-75 font-bold">{t("Surveyor Name")}</p>
                  </div>
                  <AutocompleteInput
                    id="surveyor-select"
                    value={surveyorName}
                    onChange={(val) => {
                      setSurveyorName(val);
                      if (data.length) {
                        findRecord(data, columns, searchKey, selectedGalaxy, selectedRegion, discovererName, val);
                      }
                    }}
                    suggestions={surveyorSuggestions}
                    placeholder={t("Surveyor Name")}
                    icon={<Search className="h-5 w-5" />}
                  />
                </div>
              </div>
            </div>

            <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
              <button
                onClick={handleSearch}
                disabled={loading}
                className="px-20 py-5 bg-[#E25530] border-2 border-[#FF0500] text-white rounded-full font-black text-sm uppercase tracking-[0.2em] hover:bg-[#E25530]/90 active:scale-[0.96] disabled:opacity-20 disabled:pointer-events-none shadow-[0_4px_15px_rgba(0,0,0,0.3)] hover:shadow-[0_0_20px_rgba(255,5,0,0.3)] transition-all flex items-center justify-center gap-2 shrink-0"
                id="fetch-btn"
              >
                <RefreshCw className={`w-5 h-5 ${loading ? 'animate-spin' : ''}`} />
                <span>{t("Extract Reports")}</span>
              </button>

              <button
                onClick={() => {
                  setSearchKey('');
                  setSelectedGalaxy('');
                  setSelectedRegion('');
                  setDiscovererName('');
                  setSurveyorName('');
                  if (data.length) {
                    findRecord(data, columns, '', '', '', '', '');
                  }
                }}
                className="px-6 py-3.5 border border-[#FF0500]/40 bg-[#FF0500]/10 text-[#FFB451] hover:bg-[#FF0500]/20 rounded-full text-[10px] uppercase tracking-widest font-bold transition-all active:scale-[0.95] flex items-center gap-2 shadow-[0_4px_10px_rgba(255,5,0,0.05)]"
                id="clear-all-filters-btn"
                title={t("Clear All Filters")}
              >
                <span>{t("Clear All Filters")}</span>
              </button>
            </div>

            {error && (
              <motion.div 
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="flex items-center gap-3 px-6 py-3 bg-[#FF0500]/5 border border-[#FF0500] text-[#FFB451] rounded-full text-xs font-medium tracking-wide shadow-[0_0_15px_rgba(255,5,0,0.05)]"
              >
                <AlertCircle className="w-3.5 h-3.5 shrink-0 text-[#FF0500]" />
                <p>{error}</p>
              </motion.div>
            )}
          </div>

          <div className="space-y-12">
            
            {/* Settings Area - Beautiful pop-up window overlay */}
            <AnimatePresence>
              {showSettings && (
                <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/85 backdrop-blur-sm">
                  <div className="absolute inset-0" onClick={() => setShowSettings(false)} />
                  
                  <motion.div 
                    initial={{ scale: 0.95, opacity: 0 }}
                    animate={{ scale: 1, opacity: 1 }}
                    exit={{ scale: 0.95, opacity: 0 }}
                    className="w-full max-w-4xl max-h-[85vh] bg-[#111111] border-2 border-[#FF0500] rounded-2xl flex flex-col overflow-hidden relative shadow-[0_0_50px_rgba(255,5,0,0.25)]"
                    id="settings-popup-window"
                  >
                    {/* Header */}
                    <div className="p-6 border-b border-[#FF0500] bg-[#161616] flex items-center justify-between shrink-0">
                      <div className="flex items-center gap-2">
                        <Settings className="w-5 h-5 text-[#FF0500] animate-[spin_5s_linear_infinite]" />
                        <h2 className="text-md font-bold uppercase tracking-[0.2em] text-white">{t("System Settings Console")}</h2>
                      </div>
                      <button 
                        onClick={() => setShowSettings(false)}
                        className="px-5 py-2.5 bg-[#E25530] text-white border border-[#FF0500] rounded-xl text-[10px] font-bold uppercase tracking-widest hover:bg-[#E25530]/90 transition"
                        id="settings-header-close"
                      >
                        {t("Close Settings")}
                      </button>
                    </div>

                    {/* Scrollable Content */}
                    <div className="p-8 overflow-y-auto flex-1 settings-scrollbar space-y-12 bg-[#0c0c0c]">
                      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-12">
                        
                        {/* Records Displayed Per Page Section */}
                        <div className="space-y-4">
                          <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                            <FileText className="w-3 h-3" />
                            {t("Records Displayed Per Page")}
                          </h3>
                          <div className="space-y-2 relative">
                            <select
                              value={pageSize}
                              onChange={(e) => setPageSize(parseInt(e.target.value, 10))}
                              className="block w-full px-5 py-4 bg-[#1d1d1d] border border-[#FF0500] rounded-xl text-white font-mono text-xs focus:outline-none focus:ring-1 focus:ring-[#FF0500] appearance-none"
                              id="pagesize-select"
                            >
                              {[15, 30, 50, 100].map(size => (
                                <option key={size} value={size}>{size} {t("Records per Page")}</option>
                              ))}
                            </select>
                            <div className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-white/40">
                              <ChevronRight className="w-4 h-4 rotate-90" />
                            </div>
                          </div>
                        </div>

                        {/* Display Font Scale Section */}
                        <div className="space-y-4">
                          <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                            <Table className="w-3 h-3" />
                            {t("Display Font Scale")}
                          </h3>
                          <div className="space-y-2 relative">
                            <select
                              value={fontScale}
                              onChange={(e) => setFontScale(e.target.value)}
                              className="block w-full px-5 py-4 bg-[#1d1d1d] border border-[#FF0500] rounded-xl text-white font-mono text-xs focus:outline-none focus:ring-1 focus:ring-[#FF0500] appearance-none"
                              id="fontscale-select"
                            >
                              {['1x', '1.5x', '2x', '2.5x', '3x'].map(scale => (
                                <option key={scale} value={scale}>
                                  {scale === '1x' ? t('1x (Default)') : scale}
                                </option>
                              ))}
                            </select>
                            <div className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-white/40">
                              <ChevronRight className="w-4 h-4 rotate-90" />
                            </div>
                          </div>
                        </div>

                        {/* Language Selection Section */}
                        <div className="space-y-4">
                          <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                            <Globe className="w-3 h-3" />
                            {t("Language Selection")}
                          </h3>
                          <div className="space-y-2 relative">
                            <select
                              value={language}
                              onChange={(e) => setLanguage(e.target.value as LanguageCode)}
                              className="block w-full px-5 py-4 bg-[#1d1d1d] border border-[#FF0500] rounded-xl text-white font-mono text-xs focus:outline-none focus:ring-1 focus:ring-[#FF0500] appearance-none"
                              id="language-select"
                            >
                              {LANGUAGES.map(lang => (
                                <option key={lang.code} value={lang.code}>
                                  {lang.native}
                                </option>
                              ))}
                            </select>
                            <div className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-white/40">
                              <ChevronRight className="w-4 h-4 rotate-90" />
                            </div>
                          </div>
                        </div>

                        {/* AGT Anthem (Audio Section) */}
                        <div className="col-span-1 md:col-span-2 lg:col-span-3 pt-8 border-t border-[#FF0500]/20 space-y-4">
                          <div className="flex items-center justify-between">
                            <div className="space-y-1">
                              <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                                <Volume2 className="w-3 h-3" />
                                {t("AGT Anthem")}
                              </h3>
                            </div>
                            <button 
                              onClick={() => setAudioEnabled(!audioEnabled)}
                              className="flex items-center gap-3 px-6 py-3 rounded-xl border border-[#FF0500] bg-[#E25530] text-white transition-all text-[10px] uppercase tracking-widest font-bold hover:bg-[#E25530]/90"
                              id="audio-toggle-btn"
                            >
                              {audioEnabled ? <Volume2 className="w-3.5 h-3.5" /> : <VolumeX className="w-3.5 h-3.5" />}
                              {audioEnabled ? t('Active') : t('Muted')}
                            </button>
                          </div>
                        </div>

                        {/* Custom Report Columns Section */}
                        <div className="col-span-1 md:col-span-2 lg:col-span-3 pt-8 border-t border-[#FF0500]/20 space-y-4">
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                            <div className="space-y-1">
                              <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                                <Columns className="w-3 h-3" />
                                {t("Custom Report Columns")}
                              </h3>
                              <p className="text-[10px] text-[#FFB451]/60 font-mono tracking-wide">{t("Toggle columns to appear in the new custom report format")}</p>
                            </div>
                            <button
                              onClick={() => {
                                const cleared: Record<string, boolean> = {};
                                CUSTOM_COLUMN_DEFS.forEach(def => {
                                  cleared[def.id] = false;
                                });
                                setEnabledCustomColumns(cleared);
                              }}
                              className="px-4 py-2 border border-[#FF0500] bg-[#E25530] text-white hover:bg-[#E25530]/90 rounded-lg text-[10px] uppercase font-black tracking-widest transition shadow-[0_2px_10px_rgba(226,85,48,0.2)] active:scale-95 self-start sm:self-center"
                              id="clear-all-toggles-btn"
                            >
                              {t("Clear All")}
                            </button>
                          </div>
                          <div className="grid grid-cols-4 sm:grid-cols-5 md:grid-cols-6 lg:grid-cols-8 gap-1 max-h-52 overflow-y-auto p-2 bg-black/40 border border-[#FF0500]/50 rounded-xl settings-scrollbar">
                            {CUSTOM_COLUMN_DEFS.map(def => {
                              const isEnabled = enabledCustomColumns[def.id];
                              return (
                                <button
                                  key={def.id}
                                  onClick={() => {
                                    setEnabledCustomColumns(prev => ({
                                      ...prev,
                                      [def.id]: !prev[def.id]
                                    }));
                                  }}
                                  className={`py-1 px-1.5 rounded transition-all flex items-center justify-between text-white border text-[8px] font-mono leading-none ${
                                    isEnabled 
                                      ? 'bg-[#E25530] border-[#FF0500] shadow-[0_0_5px_rgba(226,85,48,0.3)] font-extrabold' 
                                      : 'bg-[#E25530]/5 border-[#FF0500]/15 opacity-40 hover:opacity-100'
                                  }`}
                                  id={`custom-col-${def.id}-btn`}
                                  title={translateColumnHeader(def.label, language)}
                                >
                                  <span className="truncate mr-0.5 text-[8px] font-medium font-sans uppercase">{translateColumnHeader(def.label, language)}</span>
                                  <span className="shrink-0 text-[7px] px-0.5 py-px rounded bg-black/40 text-white leading-none scale-90">
                                    {isEnabled ? 'ON' : 'OFF'}
                                  </span>
                                </button>
                              );
                            })}
                          </div>
                        </div>

                        {/* Traveller Registry Validation Section */}
                        <div className="col-span-1 md:col-span-2 lg:col-span-3 pt-8 border-t border-[#FF0500]/20 space-y-4">
                          <div className="space-y-1">
                            <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                              <Settings className="w-3 h-3" />
                              {t("Traveller Registry Validation")}
                            </h3>
                            <p className="text-[10px] text-[#FFB451]/60 font-mono tracking-wide">
                              {t("Enter your registered Traveller Name and ID to authenticate and access higher security classifications")}
                            </p>
                          </div>
                          
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div className="space-y-2">
                              <label className="block text-[9px] text-[#FFB451] font-bold tracking-widest uppercase ml-1">
                                {t("Traveller Name")}
                              </label>
                              <input
                                type="text"
                                maxLength={42}
                                value={enteredName}
                                onChange={(e) => setEnteredName(e.target.value)}
                                placeholder="Your AGT User Name"
                                className="block w-full px-5 py-4 bg-[#1d1d1d] border border-[#FF0500] rounded-xl text-white font-mono text-xs focus:outline-none focus:ring-1 focus:ring-[#FF0500]"
                                id="traveller-name-input"
                              />
                            </div>

                            <div className="space-y-2">
                              <label className="block text-[9px] text-[#FFB451] font-bold tracking-widest uppercase ml-1">
                                {t("AGT Traveller ID")}
                              </label>
                              <input
                                type="text"
                                value={enteredId}
                                onChange={(e) => setEnteredId(e.target.value)}
                                placeholder="37######-????-####"
                                className="block w-full px-5 py-4 bg-[#1d1d1d] border border-[#FF0500] rounded-xl text-white font-mono text-xs focus:outline-none focus:ring-1 focus:ring-[#FF0500]"
                                id="traveller-id-input"
                              />
                            </div>
                          </div>

                          {verificationError && (
                            <div className="text-left text-[11px] text-[#FF0500] font-bold tracking-wide mt-2 font-mono flex items-center gap-1.5 p-3 bg-red-950/20 rounded-lg border border-red-500/20" id="traveller-verification-error">
                              <AlertCircle className="w-4 h-4 shrink-0" />
                              <span>
                                {t("Traveller Name and ID and does not match, Please consult")}{" "}
                                <a 
                                  href="https://www.nms-agt.com/support/traveller-id" 
                                  target="_blank" 
                                  rel="noopener noreferrer"
                                  className="text-orange-400 hover:text-orange-300 underline transition font-extrabold uppercase text-[10px]"
                                  id="verification-error-support-link"
                                >
                                  {t("AGT Support")}
                                </a>
                              </span>
                            </div>
                          )}

                          <div className="flex flex-wrap items-center gap-3 pt-2">
                            <button
                              onClick={handleVerifyTraveller}
                              disabled={isVerifying}
                              className={`px-6 py-3.5 bg-[#E25530] border border-[#FF0500] text-white rounded-xl text-[10px] uppercase tracking-[0.1em] font-semibold hover:bg-[#E25530]/90 transition-colors shadow-[0_4px_20px_rgba(226,85,48,0.2)] active:scale-[0.98] flex items-center gap-2 ${
                                isVerifying ? 'opacity-50 cursor-not-allowed' : ''
                              }`}
                              id="verify-traveller-btn"
                            >
                              {isVerifying ? (
                                <>
                                  <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                                  {t("Authenticating...")}
                                </>
                              ) : (
                                t("Verify Credentials")
                              )}
                            </button>
                            
                            <button
                              onClick={handleClearTraveller}
                              className="px-6 py-3.5 bg-transparent border border-[#FF0500]/40 text-[#FFB451] hover:text-white hover:bg-[#FF0500]/10 rounded-xl text-[10px] uppercase tracking-[0.1em] font-semibold transition active:scale-[0.98]"
                              id="clear-traveller-btn"
                            >
                              {t("Clear Credentials")}
                            </button>

                            {savedTravellerName && (
                              <div className="ml-auto flex items-center gap-2 text-[10px] font-mono text-emerald-400 bg-emerald-950/20 border border-emerald-500/20 px-3 py-1.5 rounded-lg" id="saved-traveller-indicator">
                                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                                <span>{t("Verified:")} {savedTravellerName} ({getSecurityLevelLabel(savedSecurityLevel)})</span>
                              </div>
                            )}
                          </div>
                        </div>

                        {/* System Database Section at the bottom of the scrollable content */}
                        <div className="col-span-1 md:col-span-2 lg:col-span-3 pt-8 border-t border-[#FF0500]/20 space-y-4">
                          <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
                            <div className="space-y-1">
                              <h3 className="text-[10px] uppercase tracking-widest font-bold text-[#FFB451] flex items-center gap-2">
                                <Database className="w-3 h-3" />
                                {t("System Database Settings")}
                              </h3>
                              <p className="text-[10px] text-[#FFB451]/60 font-mono tracking-wide italic">
                                {t("System database sync may take up to 5 minutes")}
                              </p>
                            </div>
                            <button 
                              onClick={() => {
                                fetchData();
                              }}
                              className="w-full sm:w-auto px-8 py-3.5 bg-[#E25530] border border-[#FF0500] text-white rounded-xl text-[10px] uppercase tracking-[0.1em] font-semibold hover:bg-[#E25530]/90 transition-colors shadow-[0_4px_20px_rgba(226,85,48,0.2)] active:scale-[0.98]"
                              id="resync-db-btn"
                            >
                              {t("Re-Sync System DB")}
                            </button>
                          </div>
                        </div>

                      </div>
                    </div>

                  </motion.div>
                </div>
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
                    className="bg-[#111111] border-2 border-[#FF0500] rounded-2xl overflow-hidden shadow-[0_0_30px_rgba(255,5,0,0.15)]"
                  >
                    <div className="p-8 border-b border-[#FF0500]/20 flex flex-col md:flex-row md:items-center justify-between gap-6 bg-[#161616]">
                      <div className="space-y-1.5">
                        <h3 className="text-xl font-medium text-white flex items-center gap-3">
                          {t("AGT Galactic Archives Results")}
                          <span className="px-2 py-0.5 rounded-full bg-[#E25530]/10 text-white border border-[#FF0500]/40 font-mono text-[10px]">
                            {matchedRecords.length} {t("FOUND")}
                          </span>
                        </h3>
                        <div className="flex items-center">
                          <span className="px-2 py-0.5 rounded-full bg-[#E25530]/10 text-white border border-[#FF0500]/40 font-mono text-[10px]">
                            {t("Classified Records Omitted:")} {omittedCount}
                          </span>
                        </div>
                        <p className="text-[10px] text-white/60 uppercase tracking-[0.2em]">{t("Verified Galactic Ledger Matches")}</p>
                      </div>
 
                      <div className="flex flex-wrap items-center gap-3">
                        {(reportType === 'simple' || reportType === 'custom') && (
                          <button
                            onClick={downloadFullReportPdf}
                            className="flex items-center gap-3 px-8 py-4 border border-[#FF0500] bg-[#E25530] text-white hover:bg-[#E25530]/90 rounded-xl text-[10px] uppercase tracking-[0.2em] font-black transition-all shadow-[0_4px_20px_rgba(226,85,48,0.2)] active:scale-[0.98]"
                          >
                            <Download className="w-4 h-4 text-white" />
                            <span>{t("PDF Report")}</span>
                          </button>
                        )}
                        <button
                          onClick={downloadCsv}
                          className="flex items-center gap-3 px-8 py-4 border border-[#FF0500] bg-[#E25530] text-white hover:bg-[#E25530]/90 rounded-xl text-[10px] uppercase tracking-[0.2em] font-black transition-all shadow-[0_4px_20px_rgba(226,85,48,0.2)] active:scale-[0.98]"
                        >
                          <Table className="w-4 h-4 text-white" />
                          <span>{t("Export CSV")}</span>
                        </button>
                      </div>
                    </div>
 
                    {/* Top synchronized horizontal scrollbar */}
                    <div 
                      ref={topScrollRef} 
                      onScroll={handleTopScroll}
                      className={`overflow-x-auto overflow-y-hidden border-b border-[#FF0500]/20 bg-[#161616] settings-scrollbar ${
                        tableScrollWidth > tableClientWidth ? 'block' : 'hidden'
                      }`}
                      style={{ minHeight: '12px' }}
                      id="top-scrollbar-container"
                    >
                      <div style={{ width: `${tableScrollWidth}px`, height: '1px' }}></div>
                    </div>

                    <div 
                      ref={bottomScrollRef}
                      onScroll={handleBottomScroll}
                      className="overflow-x-auto max-h-[500px] overflow-y-auto settings-scrollbar"
                    >
                      <table className="w-full text-left border-collapse relative">
                        <thead className="sticky top-0 z-10 bg-[#1c1c1c] shadow-[0_1px_0_rgba(255,5,0,0.15)]">
                          <tr className="bg-[#1c1c1c] border-b border-[#FF0500]/20 sticky top-0 z-10">
                            {columns.filter(col => col.enabled).map((col, idx) => {
                              const isCurrentSort = sortColumn === col.name;
                              const displayName = col.rawIndex === 10 ? 'Surveyor' : col.name;
                              const translatedDisplayName = translateColumnHeader(displayName, language);
                              return (
                                <th 
                                  key={idx} 
                                  onClick={() => {
                                    if (sortColumn === col.name) {
                                      if (sortDirection === 'asc') {
                                        setSortDirection('desc');
                                      } else if (sortDirection === 'desc') {
                                        setSortColumn(null);
                                        setSortDirection(null);
                                      }
                                    } else {
                                      setSortColumn(col.name);
                                      setSortDirection('asc');
                                    }
                                    setCurrentPage(1);
                                  }}
                                  className="py-1 px-2 text-[8px] uppercase tracking-widest font-black text-[#FFB451] whitespace-normal sticky top-0 z-10 bg-[#1c1c1c] border-b border-[#FF0500]/20 cursor-pointer select-none hover:bg-black/40 group/th transition-all min-w-[70px]"
                                  id={`sort-th-${col.rawIndex}`}
                                  title={t("Click to sort")}
                                >
                                  <div className="flex items-start gap-1.5 justify-between">
                                    <span className="line-clamp-3 break-normal">{translatedDisplayName}</span>
                                    <span className="text-[#FF0500] font-mono text-[9px] select-none flex items-center shrink-0 mt-0.5">
                                      {sortColumn === col.name ? (
                                        sortDirection === 'asc' ? '▲' : '▼'
                                      ) : (
                                        <span className="opacity-0 group-hover/th:opacity-40 text-[7px] transition-all">▲▼</span>
                                      )}
                                    </span>
                                  </div>
                                </th>
                              );
                            })}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-[#FF0500]/10 bg-black/20">
                          {paginatedRecords.map((record, rIdx) => (
                            <motion.tr 
                              key={record._originalIndex !== undefined ? `row-${record._originalIndex}` : `row-${rIdx}`}
                              initial={{ opacity: 0 }}
                              animate={{ opacity: 1 }}
                              transition={{ duration: 0.25, ease: "easeOut" }}
                              className="hover:bg-[#E25530]/5 transition-colors group"
                            >
                              {columns.filter(col => col.enabled).map((col, cIdx) => {
                                const rawVal = record[col.name];
                                const displayVal = getDisplayValue(rawVal, col.rawIndex);
                                const isTargetBlueUrl = isTargetColumnWithUrl(col.rawIndex, rawVal);
                                const isUrl = (col.rawIndex >= 76 && col.rawIndex <= 80) || String(rawVal || '').trim().startsWith('http');
                                return (
                                  <td key={cIdx} className="py-0.5 px-2 text-[10px] text-[#FFB451] font-mono whitespace-nowrap overflow-hidden text-ellipsis max-w-[200px]">
                                    {isTargetBlueUrl ? (
                                      <a
                                        href={String(rawVal).trim()}
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        className="text-[#4da6ff] hover:text-[#80c2ff] font-bold underline transition-colors"
                                      >
                                        LINK
                                      </a>
                                    ) : isUrl && rawVal ? (
                                      <a
                                        href={String(rawVal).trim()}
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        className="text-[#FFB451] underline hover:text-[#ffca80] transition-colors"
                                      >
                                        {displayVal || 'LINK'}
                                      </a>
                                    ) : (
                                      displayVal || <span className="text-white/20 italic">-</span>
                                    )}
                                  </td>
                                );
                              })}
                            </motion.tr>
                          ))}
                        </tbody>
                        <tfoot className="border-t-2 border-[#FF0500]/20 bg-[#161616]">
                          <tr>
                            {columns.filter(col => col.enabled).map((col, idx) => (
                              <td key={idx} className="py-1 px-2 text-[9px] font-bold text-white">
                                {col.name === columns[0]?.name ? (
                                  <div className="flex flex-col">
                                    <span className="text-[7px] text-white/60 uppercase tracking-tighter">{t("Total Matches")}</span>
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
                      <div className="p-6 border-t border-[#FF0500]/20 bg-[#161616] flex flex-col md:flex-row items-center justify-between gap-4">
                        <div className="text-[11px] uppercase tracking-wider text-white/80 font-mono">
                          {t("Showing Page ")}<span className="text-[#FFB451] font-bold font-sans">{currentPage}</span>{t(" of ")}<span className="text-[#FFB451] font-bold font-sans">{totalPages}</span> <span className="text-white/40">({matchedRecords.length} {t(" total rows")})</span>
                        </div>
 
                        <div className="flex flex-wrap items-center gap-2">
                          <button
                            onClick={() => setCurrentPage(1)}
                            disabled={currentPage === 1}
                            className="px-3 py-1.5 rounded-lg border border-[#FF0500] bg-[#E25530] text-white text-[10px] uppercase font-bold tracking-widest hover:bg-[#E25530]/90 disabled:opacity-30 disabled:pointer-events-none transition-all"
                          >
                            {t("First")}
                          </button>
 
                          <button
                            onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                            disabled={currentPage === 1}
                            className="px-3 py-1.5 rounded-lg border border-[#FF0500] bg-[#E25530] text-white text-[10px] uppercase font-bold tracking-widest hover:bg-[#E25530]/90 disabled:opacity-30 disabled:pointer-events-none transition-all"
                          >
                            {t("Prev")}
                          </button>
 
                          <div className="flex items-center gap-1">
                            {Array.from({ length: totalPages }).map((_, idx) => {
                              const pageNum = idx + 1;
                              if (totalPages > 6) {
                                if (pageNum !== 1 && pageNum !== totalPages && Math.abs(pageNum - currentPage) > 1) {
                                  if (pageNum === 2 && currentPage > 3) {
                                    return <span key="ellipsis-start" className="text-white/40 font-mono text-[10px] px-1">...</span>;
                                  }
                                  if (pageNum === totalPages - 1 && currentPage < totalPages - 2) {
                                    return <span key="ellipsis-end" className="text-white/40 font-mono text-[10px] px-1">...</span>;
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
                                      ? 'bg-[#E25530] text-white font-extrabold shadow-[0_0_15px_rgba(226,85,48,0.8)] border-2 border-[#FF0500] scale-110'
                                      : 'border border-[#FF0500]/20 bg-[#E25530]/10 text-white hover:bg-[#E25530]/40'
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
                            className="px-3 py-1.5 rounded-lg border border-[#FF0500] bg-[#E25530] text-white text-[10px] uppercase font-bold tracking-widest hover:bg-[#E25530]/90 disabled:opacity-30 disabled:pointer-events-none transition-all"
                          >
                            {t("Next")}
                          </button>
 
                          <button
                            onClick={() => setCurrentPage(totalPages)}
                            disabled={currentPage === totalPages}
                            className="px-3 py-1.5 rounded-lg border border-[#FF0500] bg-[#E25530] text-white text-[10px] uppercase font-bold tracking-widest hover:bg-[#E25530]/90 disabled:opacity-30 disabled:pointer-events-none transition-all"
                          >
                            {t("Last")}
                          </button>
                        </div>
                      </div>
                    )}
 
                    <div className="p-6 border-t border-agt-orange/5 flex flex-col md:flex-row items-center justify-between gap-6 bg-agt-orange/[0.01]">
                      <div className="flex items-center gap-2">
                        <div className="relative flex h-2 w-2">
                          <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-[#00FF66] opacity-75"></span>
                          <span className="relative inline-flex rounded-full h-2 w-2 bg-[#00FF66] shadow-[0_0_8px_rgba(0,255,102,0.6)]"></span>
                        </div>
                        <span className="text-[9px] uppercase tracking-widest text-agt-orange font-bold">{t("Ledger Integrity: Verified")}</span>
                      </div>
                      <span className="text-[9px] font-mono text-agt-orange uppercase tracking-widest md:text-right select-none opacity-90">
                        {t("AGT SECURE ARCHIVE CLIENT")}
                      </span>
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
                      <p className="text-sm font-medium uppercase tracking-[0.2em]">{t("Terminal Ready")}</p>
                      <p className="text-xs font-light">{t("Report Generation Sequence Pending Civilization Selection")}</p>
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
          <p className="text-[9px] font-mono uppercase tracking-[0.3em] font-bold">&copy; 2026 {t("Alliance of Galactic Travellers")}</p>
        </div>
      </footer>

      {/* Background Audio */}
      <audio 
        ref={audioRef}
        src={agtAnthem}
        loop
        preload="auto"
      />

      {/* PDF Error Modal Pop-up */}
      <AnimatePresence>
        {pdfErrorMsg && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-black/90 backdrop-blur-sm">
            <div className="absolute inset-0" onClick={() => setPdfErrorMsg(null)} />
            
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="w-full max-w-lg bg-[#111111] border-2 border-[#FF0500] rounded-2xl p-8 flex flex-col relative shadow-[0_0_50px_rgba(255,5,0,0.4)] text-center space-y-6"
              id="pdf-error-popup"
            >
              <div className="mx-auto w-16 h-16 rounded-full bg-[#FF0500]/10 border-2 border-[#FF0500] flex items-center justify-center animate-pulse">
                <AlertCircle className="w-8 h-8 text-[#FF0500]" />
              </div>
              
              <div className="space-y-2">
                <h3 className="text-xl font-bold uppercase tracking-widest text-[#FF0500]" id="pdf-error-title">
                  Report Exceeds Limit
                </h3>
                <p className="text-xs text-[#FFB451] font-mono leading-relaxed animate-pulse" id="pdf-error-message">
                  {pdfErrorMsg}
                </p>
              </div>

              <div className="pt-2">
                <button 
                  onClick={() => setPdfErrorMsg(null)}
                  className="px-10 py-3.5 bg-[#E25530] border-2 border-[#FF0500] text-white rounded-full font-black text-[10px] uppercase tracking-widest hover:bg-[#E25530]/90 active:scale-[0.96] transition-all shadow-[0_4px_15px_rgba(255,5,0,0.2)]"
                  id="pdf-error-close-btn"
                >
                  Dismiss
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Custom Alert Modal Pop-up */}
      <AnimatePresence>
        {alertPopup && alertPopup.show && (
          <div className="fixed inset-0 z-[120] flex items-center justify-center p-4 bg-black/90 backdrop-blur-sm animate-[fadeIn_0.2s_ease-out]">
            <div className="absolute inset-0 max-w-full" onClick={() => {
              const cb = alertPopup.onClose;
              setAlertPopup(null);
              if (cb) cb();
            }} />
            
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="w-full max-w-lg bg-[#111111] border-2 border-[#FF0500] rounded-2xl p-8 flex flex-col relative shadow-[0_0_50px_rgba(255,5,0,0.4)] text-center space-y-6 z-[121]"
              id="custom-alert-popup"
            >
              <div className="mx-auto w-16 h-16 rounded-full bg-[#FF0500]/10 border-2 border-[#FF0500] flex items-center justify-center">
                <AlertCircle className="w-8 h-8 text-[#FFB451]" />
              </div>
              
              <div className="space-y-2">
                <h3 className="text-xl font-bold uppercase tracking-widest text-[#FFB451]" id="custom-alert-title">
                  System Notification
                </h3>
                <p className="text-xs text-[#FFB451] font-mono leading-relaxed" id="custom-alert-message">
                  {alertPopup.message}
                </p>
              </div>

              <div className="pt-2">
                <button 
                  onClick={() => {
                    const cb = alertPopup.onClose;
                    setAlertPopup(null);
                    if (cb) cb();
                  }}
                  className="px-10 py-3.5 bg-[#E25530] border-2 border-[#FF0500] text-white rounded-full font-black text-[10px] uppercase tracking-widest hover:bg-[#E25530]/90 active:scale-[0.96] transition-all shadow-[0_4px_15px_rgba(255,5,0,0.2)]"
                  id="custom-alert-close-btn"
                >
                  Dismiss
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

