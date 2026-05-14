/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { 
  Upload, 
  Search, 
  FileSpreadsheet, 
  X, 
  User, 
  Phone, 
  MapPin, 
  SearchSlash, 
  Clock, 
  BarChart3, 
  ChevronRight,
  Database,
  Download,
  Info,
  RotateCcw,
  Check
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface DataRow {
  [key: string]: string | number | null | undefined;
}

export default function App() {
  const [data, setData] = useState<DataRow[]>([]);
  const [originalData, setOriginalData] = useState<DataRow[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [paidRows, setPaidRows] = useState<Set<number>>(new Set());
  const [modifiedRows, setModifiedRows] = useState<Set<number>>(new Set());
  const [editingCell, setEditingCell] = useState<{ row: number; col: string } | null>(null);
  const [isModified, setIsModified] = useState(false);
  const [filterType, setFilterType] = useState<'all' | 'paid' | 'modified' | 'normal'>('all');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // 1. First, find matches just for the search query
  const searchMatches = useMemo(() => {
    if (!searchQuery.trim()) return data;
    const query = searchQuery.toLowerCase();
    return data.filter(row => {
      return Object.values(row).some(value => 
        String(value || '').toLowerCase().includes(query)
      );
    });
  }, [data, searchQuery]);

  // 2. Use those matches to figure out the counts for each filter button based on current search
  const filterCounts = useMemo(() => {
    const counts = { all: searchMatches.length, paid: 0, modified: 0, normal: 0 };
    searchMatches.forEach(row => {
      const originalIdx = data.indexOf(row);
      const isPaid = paidRows.has(originalIdx);
      const isModified = modifiedRows.has(originalIdx);
      if (isPaid) counts.paid++;
      else if (isModified) counts.modified++;
      else counts.normal++;
    });
    return counts;
  }, [searchMatches, data, paidRows, modifiedRows]);

  const filteredData = useMemo(() => {
    let results = searchMatches;
    
    // Apply status filter on top of search matches
    if (filterType !== 'all') {
      results = results.filter((row) => {
        const originalIdx = data.indexOf(row);
        if (filterType === 'paid') return paidRows.has(originalIdx);
        if (filterType === 'modified') return modifiedRows.has(originalIdx);
        if (filterType === 'normal') return !paidRows.has(originalIdx) && !modifiedRows.has(originalIdx);
        return true;
      });
    }
    
    return results;
  }, [searchMatches, filterType, data, paidRows, modifiedRows]);

  // Financial Summary Calculation
  const totals = useMemo(() => {
    // Global totals for the entire file
    const allData = data;
    // Filtered data for "settled currently"
    const currentOnScreenData = filteredData;
    
    // Attempt to find columns for Amount and Remaining
    const amountKeywords = ['مبلغ', 'المبلغ', 'قيمة', 'القيمة', 'amount', 'total', 'price', 'السعر'];
    const remainingKeywords = ['متبقي', 'المتبقي', 'باقي', 'الباقي', 'remaining', 'balance', 'left'];

    const amountCol = headers.find(h => amountKeywords.some(k => h.toLowerCase().includes(k)));
    const remainingCol = headers.find(h => remainingKeywords.some(k => h.toLowerCase().includes(k)));
    const notesKeywords = ['ملاحظات', 'ملاحظة', 'التفاصيل', 'notes', 'observation', 'comment'];
    const notesCol = headers.find(h => notesKeywords.some(k => h.toLowerCase().includes(k)));

    let totalAmount = 0;
    let totalRemaining = 0;
    let totalAmountFiltered = 0;
    let totalRemainingFiltered = 0;
    let totalSettledFiltered = 0;
    let totalSettledGlobal = 0;
    let totalLiquidationFiltered = 0;

    // 1. Calculate Global Totals (Entire File)
    allData.forEach(row => {
      if (amountCol) {
        const val = parseFloat(String(row[amountCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) totalAmount += val;
      }
      if (remainingCol) {
        const val = parseFloat(String(row[remainingCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) totalRemaining += val;
      }
    });

    // 2. Calculate Filtered Totals (Only rows matching current search/filter)
    currentOnScreenData.forEach(row => {
      const originalIdx = data.indexOf(row);
      const isPaid = paidRows.has(originalIdx);

      if (amountCol) {
        const val = parseFloat(String(row[amountCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) totalAmountFiltered += val;
      }

      if (remainingCol) {
        const val = parseFloat(String(row[remainingCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) {
          totalRemainingFiltered += val;
          // Only add to settled if this row is paid
          if (isPaid) {
            totalSettledFiltered += val;
          }
          
          // Calculate Liquidation (تصفية)
          // Search in notesCol or any available column for the word "تصفية"
          const hasTasfya = notesCol 
            ? String(row[notesCol] || '').includes('تصفية')
            : Object.values(row).some(v => String(v || '').includes('تصفية'));
            
          if (hasTasfya) {
            totalLiquidationFiltered += val;
          }
        }
      }
    });

    // 3. Calculate total settled GLOBAL (Entire file)
    paidRows.forEach(idx => {
      const row = data[idx];
      if (row && remainingCol) {
        const val = parseFloat(String(row[remainingCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) totalSettledGlobal += val;
      }
    });

    return { 
      totalAmount, 
      totalRemaining, 
      totalAmountFiltered,
      totalRemainingFiltered,
      totalSettledGlobal, 
      totalSettledFiltered, 
      totalLiquidationFiltered,
      amountCol, 
      remainingCol 
    };
  }, [data, filteredData, headers, paidRows]);

  // Initialize IndexedDB and load saved data
  useEffect(() => {
    const loadSavedData = async () => {
      setIsLoading(true);
      // Restore search query from localStorage
      const savedQuery = localStorage.getItem('last_search_query');
      if (savedQuery) setSearchQuery(savedQuery);

      try {
        const dbRequest = indexedDB.open('FileSearchDB', 1);
        
        dbRequest.onupgradeneeded = (e: any) => {
          const db = e.target.result;
          if (!db.objectStoreNames.contains('files')) {
            db.createObjectStore('files', { keyPath: 'id' });
          }
        };

        dbRequest.onsuccess = (e: any) => {
          const db = e.target.result;
          const transaction = db.transaction(['files'], 'readonly');
          const store = transaction.objectStore('files');
          const getRequest = store.get('current_file');

          getRequest.onsuccess = () => {
            if (getRequest.result) {
              const res = getRequest.result as { 
              data: DataRow[], 
              originalData?: DataRow[],
              headers: string[], 
              fileName: string, 
              paidIndices?: number[], 
              modifiedIndices?: number[],
              isModified?: boolean 
            };
            const { data: savedData, originalData: savedOriginal, headers: savedHeaders, fileName: savedName, paidIndices, modifiedIndices, isModified: savedModified } = res;
            setData(savedData);
            // Fallback: if no original data was saved, treat current data as original
            if (savedOriginal && savedOriginal.length > 0) {
              setOriginalData(savedOriginal);
            } else if (savedData && savedData.length > 0) {
              setOriginalData(JSON.parse(JSON.stringify(savedData)));
            }
            setHeaders(savedHeaders);
            setFileName(savedName);
            if (paidIndices) setPaidRows(new Set(paidIndices));
            if (modifiedIndices) setModifiedRows(new Set(modifiedIndices));
            if (savedModified) setIsModified(savedModified);
            }
            setIsLoading(false);
          };
          
          getRequest.onerror = () => {
            console.error('Failed to get data from store');
            setIsLoading(false);
          };
        };

        dbRequest.onerror = (e) => {
          console.error('DB Open Error:', e);
          setIsLoading(false);
        };
      } catch (error) {
        console.error('IndexedDB Error:', error);
        setIsLoading(false);
      }
    };

    loadSavedData();
  }, []);

  // Sync search query to localStorage
  useEffect(() => {
    if (searchQuery) {
      localStorage.setItem('last_search_query', searchQuery);
    } else {
      localStorage.removeItem('last_search_query');
    }
  }, [searchQuery]);

  // Save to IndexedDB helper
  const saveToDB = (
    fileData: DataRow[], 
    fileHeaders: string[], 
    name: string, 
    paidIndices: number[] = [], 
    modifiedIndices: number[] = [], 
    modified = false, 
    original: DataRow[] | null = null
  ) => {
    try {
      const dbRequest = indexedDB.open('FileSearchDB', 1);
      dbRequest.onsuccess = (e: any) => {
        const db = e.target.result;
        const transaction = db.transaction(['files'], 'readwrite');
        const store = transaction.objectStore('files');
        
        // Ensure we always have some original data to save
        const dataToSaveAsOriginal = (original && original.length > 0) 
          ? original 
          : (originalData && originalData.length > 0 ? originalData : fileData);

        store.put({
          id: 'current_file',
          data: fileData,
          originalData: dataToSaveAsOriginal,
          headers: fileHeaders,
          fileName: name,
          paidIndices: paidIndices,
          modifiedIndices: modifiedIndices,
          isModified: modified,
          updatedAt: new Date().toISOString()
        });
      };
    } catch (e) {
      console.error('Failed to save to DB:', e);
    }
  };

  // Clear IndexedDB helper
  const removeFromDB = () => {
    localStorage.removeItem('last_search_query');
    try {
      const dbRequest = indexedDB.open('FileSearchDB', 1);
      dbRequest.onsuccess = (e: any) => {
        const db = e.target.result;
        const transaction = db.transaction(['files'], 'readwrite');
        const store = transaction.objectStore('files');
        store.delete('current_file');
      };
    } catch (e) {
      console.error('Failed to delete from DB:', e);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    processFile(file);
  };

  const togglePaid = (originalIdx: number) => {
    setPaidRows(prev => {
      const next = new Set(prev);
      if (next.has(originalIdx)) {
        next.delete(originalIdx);
      } else {
        next.add(originalIdx);
      }
      setIsModified(true);
      // Persist immediately
      saveToDB(data, headers, fileName || '', Array.from(next) as number[], Array.from(modifiedRows) as number[], true, originalData);
      return next;
    });
  };

  const updateNote = (originalIdx: number, header: string, newValue: string) => {
    // 1. Update local state for immediate feedback
    const nextData = [...data];
    nextData[originalIdx] = { ...nextData[originalIdx], [header]: newValue };
    setData(nextData);
    
    // 2. Track modification
    const newModifiedRows = new Set(modifiedRows);
    newModifiedRows.add(originalIdx);
    setModifiedRows(newModifiedRows);
    setIsModified(true);

    // 3. Persist to DB
    saveToDB(
      nextData, 
      headers, 
      fileName || '', 
      Array.from(paidRows) as number[], 
      Array.from(newModifiedRows) as number[], 
      true, 
      originalData
    );
  };

  const undoRowChanges = (originalIdx: number) => {
    if (originalData[originalIdx]) {
      // 1. Restore from original
      const nextData = [...data];
      nextData[originalIdx] = { ...originalData[originalIdx] };
      setData(nextData);
      
      // 2. Remove from modified
      const newModifiedRows = new Set(modifiedRows);
      newModifiedRows.delete(originalIdx);
      setModifiedRows(newModifiedRows);
      
      const stillModified = newModifiedRows.size > 0 || paidRows.size > 0;
      setIsModified(stillModified);

      // 3. Persist back
      saveToDB(
        nextData, 
        headers, 
        fileName || '', 
        Array.from(paidRows) as number[], 
        Array.from(newModifiedRows) as number[], 
        stillModified, 
        originalData
      );
    }
  };

  const processFile = (file: File) => {
    setIsLoading(true);
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = XLSX.utils.sheet_to_json<DataRow>(ws);
        
        if (jsonData.length > 0) {
          const newHeaders = Object.keys(jsonData[0]);
          setHeaders(newHeaders);
          setData(jsonData);
          setOriginalData(JSON.parse(JSON.stringify(jsonData)));
          setPaidRows(new Set());
          setModifiedRows(new Set());
          setIsModified(false);
          saveToDB(jsonData, newHeaders, file.name, [], [], false, JSON.parse(JSON.stringify(jsonData)));
        } else {
          alert('الملف فارغ أو لا يحتوي على بيانات صحيحة.');
        }
      } catch (error) {
        console.error('Error parsing file:', error);
        alert('حدث خطأ أثناء قراءة الملف. تأكد من أنه ملف إكسل صالح.');
      } finally {
        setIsLoading(false);
      }
    };
    reader.readAsBinaryString(file);
  };

  const clearFile = () => {
    setData([]);
    setHeaders([]);
    setSearchQuery('');
    setFileName(null);
    setPaidRows(new Set());
    setModifiedRows(new Set());
    setIsModified(false);
    removeFromDB();
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const getHeaderTooltip = (header: string) => {
    const h = header.toLowerCase();
    if (h.includes('اسم') || h.includes('name')) return 'الاسم الكامل للعميل أو المستأجر';
    if (h.includes('عقار') || h.includes('وحدة') || h.includes('address')) return 'وصف العقار أو الوحدة السكنية';
    if (h.includes('مبلغ') || h.includes('إجمالي') || h.includes('amount')) return 'إجمالي القيمة المالية المستحقة';
    if (h.includes('متبقي') || h.includes('باقي') || h.includes('remaining')) return 'المبلغ المتبقي الذي لم يتم سداده بعد';
    if (h.includes('دفع') || h.includes('مدفوع') || h.includes('payment')) return 'المبالغ التي تم تحصيلها بالفعل';
    if (h.includes('تاريخ') || h.includes('date')) return 'تاريخ العملية أو موعد الاستحقاق';
    if (h.includes('ملاحظات') || h.includes('notes')) return 'أي تفاصيل إضافية أو تنبيهات متعلقة بالسجل';
    if (h.includes('هاتف') || h.includes('جوال') || h.includes('phone')) return 'رقم التواصل الخاص بالسجل';
    return 'بيانات إضافية متعلقة بهذا العمود';
  };

  const exportToExcel = (forceAll = false) => {
    // 1. Decide which data set to export
    // Export all if forced OR (no search AND no specific filter active)
    const isExportAll = forceAll || (!searchQuery.trim() && filterType === 'all');
    const dataToExport = isExportAll ? data : filteredData;
    
    if (dataToExport.length === 0) return;

    // 2. Create worksheet from data
    const worksheet = XLSX.utils.json_to_sheet(dataToExport);

    // 3. Define Styles
    const headerStyle = {
      fill: { fgColor: { rgb: "4F46E5" } }, // Indigo 600
      font: { color: { rgb: "FFFFFF" }, bold: true, sz: 12 },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: {
        top: { style: "thin", color: { rgb: "E2E8F0" } },
        bottom: { style: "thin", color: { rgb: "E2E8F0" } },
        left: { style: "thin", color: { rgb: "E2E8F0" } },
        right: { style: "thin", color: { rgb: "E2E8F0" } }
      }
    };

    const dataStyle = {
      font: { sz: 11, color: { rgb: "334155" } },
      alignment: { horizontal: "right", vertical: "center", wrapText: true },
      border: {
        bottom: { style: "thin", color: { rgb: "F1F5F9" } }
      }
    };

    const summaryStyle = {
      fill: { fgColor: { rgb: "F1F5F9" } }, 
      font: { bold: true, sz: 12, color: { rgb: "1E293B" } },
      alignment: { horizontal: "right", vertical: "center" },
      border: {
        top: { style: "medium", color: { rgb: "4F46E5" } }
      }
    };

    const highlightStyle = (colorCode: string) => ({
      ...summaryStyle,
      font: { ...summaryStyle.font, color: { rgb: colorCode } }
    });

    // 4. Apply Styles to Header
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const address = XLSX.utils.encode_col(C) + "1";
      if (worksheet[address]) {
        worksheet[address].s = headerStyle;
      }
    }

    // 5. Apply Styles to Data Rows
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
      const rowData = dataToExport[R - 1];
      const originalIdx = data.findIndex(r => r === rowData);
      const isPaid = paidRows.has(originalIdx);
      const isModifiedRow = modifiedRows.has(originalIdx);
      
      let rowBgColor = "FFFFFF";
      if (isPaid) {
        rowBgColor = "D1FAE5"; // Emerald 100
      } else if (isModifiedRow) {
        // Check if notes column contains "تصفية"
        const rowString = JSON.stringify(rowData);
        if (rowString.includes('تصفية')) {
          rowBgColor = "FECACA"; // Red 200 for Tasfya
        } else {
          rowBgColor = "DBEAFE"; // Blue 100 for general edit
        }
      }

      for (let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_cell({ r: R, c: C });
        if (worksheet[address]) {
          worksheet[address].s = {
            ...dataStyle,
            fill: { fgColor: { rgb: rowBgColor } }
          };
        }
      }
    }

    // 6. Add Summary Row
    const summaryRowData = {};
    const settledColIndex = headers.indexOf(totals.remainingCol || '') + 1;
    const hasNextCol = settledColIndex < headers.length;
    const nextColHeader = hasNextCol ? headers[settledColIndex] : null;

    const settledToDisplay = isExportAll ? totals.totalSettledGlobal : totals.totalSettledFiltered;
    const liquidationToDisplay = isExportAll ? 0 : totals.totalLiquidationFiltered; // Usually meaningful only for current results
    const amountToDisplay = isExportAll ? totals.totalAmount : totals.totalAmountFiltered;
    const remainingToDisplay = isExportAll ? totals.totalRemaining : totals.totalRemainingFiltered;
    
    const settledLabel = isExportAll ? "المسدد الإجمالي" : "المسدد حالياً";
    const liquidationLabel = "إجمالي التصفية";
    const amountLabel = isExportAll ? "إجمالي المبلغ" : "إجمالي المبالغ المعروضة";
    const remainingLabel = isExportAll ? "إجمالي المتبقي" : "المتبقي (للبحث الحالي)";

    headers.forEach((h) => {
      if (h === totals.amountCol) {
        summaryRowData[h] = `${amountLabel}: ${amountToDisplay.toLocaleString('ar-SA')} ر.س`;
        if (liquidationToDisplay > 0) {
          summaryRowData[h] += ` | ${liquidationLabel}: ${liquidationToDisplay.toLocaleString('ar-SA')} ر.س`;
        }
      } else if (h === totals.remainingCol) {
        summaryRowData[h] = `${remainingLabel}: ${remainingToDisplay.toLocaleString('ar-SA')} ر.س`;
        // If we can't use next column, append it here
        if (!nextColHeader && settledToDisplay > 0) {
          summaryRowData[h] += ` | ${settledLabel}: ${settledToDisplay.toLocaleString('ar-SA')} ر.س`;
        }
      } else if (nextColHeader && h === nextColHeader && settledToDisplay > 0) {
        // If the column after 'remaining' is available, use it for settled amount to avoid overlap
        summaryRowData[h] = `${settledLabel}: ${settledToDisplay.toLocaleString('ar-SA')} ر.س`;
      } else if (h === headers[0]) {
        summaryRowData[h] = isExportAll ? '--- الملخص الإجمالي (للملف كامل) ---' : '--- ملخص نتائج البحث ---';
      } else {
        summaryRowData[h] = '';
      }
    });

    XLSX.utils.sheet_add_json(worksheet, [summaryRowData], {
      skipHeader: true,
      origin: -1
    });

    // Style the summary row
    const lastRowIndex = range.e.r + 2; 
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const address = XLSX.utils.encode_cell({ r: lastRowIndex - 1, c: C });
      if (worksheet[address]) {
        const header = headers[C];
        if (header === totals.amountCol) {
          worksheet[address].s = highlightStyle("4F46E5"); 
        } else if (header === totals.remainingCol) {
          worksheet[address].s = highlightStyle("EA580C"); 
        } else if (nextColHeader && header === nextColHeader && settledToDisplay > 0) {
          worksheet[address].s = highlightStyle("10B981"); // Emerald 500 for settled
        } else {
          worksheet[address].s = summaryStyle;
        }
      }
    }

    // 7. Config widths & RTL
    const wscols = headers.map(h => {
      const hLower = h.toLowerCase();
      let w = 18;
      if (hLower.includes('ملاحظات')) w = 50;
      else if (hLower.includes('اسم') || hLower.includes('عقار')) w = 35;
      else if (hLower.includes('دفع') || hLower.includes('تاريخ')) w = 25;
      
      // Ensure summary row text fits
      const cellValue = summaryRowData[h] || '';
      if (cellValue.length > w) {
        w = Math.min(cellValue.length + 5, 80); // Max 80 width
      }
      
      return { wch: w };
    });
    worksheet['!cols'] = wscols;
    worksheet['!views'] = [{ RTL: true }];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "البيانات المعدلة");
    const dateStr = new Date().toISOString().split('T')[0];
    const fileNameSafe = (fileName || 'بيانات').replace(/[\\/:*?"<>|]/g, '_');
    const suffix = isExportAll ? 'الكل_معدل' : 'نتائج_البحث';
    XLSX.writeFile(workbook, `سجلات_${fileNameSafe}_${suffix}_${dateStr}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans" dir="rtl">
      {/* Top Navigation Bar */}
      <nav className="h-16 bg-white border-b border-slate-200 px-6 md:px-8 flex items-center justify-between shadow-sm shrink-0 sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <span className="text-xl font-bold text-slate-800 hidden sm:block">نظام البحث الذكي</span>
        </div>
        
        <div className="flex items-center gap-4">
          <AnimatePresence>
            {fileName && (
              <motion.div 
                initial={{ opacity: 0, x: -10 }}
                animate={{ opacity: 1, x: 0 }}
                className="text-left flex flex-col items-end"
              >
                <p className="text-[10px] uppercase font-bold text-slate-400 leading-tight">الملف النشط</p>
                <p className="text-sm font-medium text-slate-700 max-w-[150px] truncate">{fileName}</p>
              </motion.div>
            )}
          </AnimatePresence>
          <div className="w-10 h-10 rounded-xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-600">
            <User size={20} />
          </div>
        </div>
      </nav>

      <main className="flex-1 p-4 md:p-8 grid grid-cols-12 gap-8 max-w-[1400px] mx-auto w-full">
        {/* Sidebar / Action Panel */}
        <aside className="col-span-12 lg:col-span-4 flex flex-col gap-6">
          {/* Upload Area */}
          <section className="bg-white rounded-3xl p-6 border border-slate-200 shadow-sm">
            <h2 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
              <Upload size={18} className="text-indigo-600" />
              رفع الملفات
            </h2>
            <div 
              onClick={() => fileInputRef.current?.click()}
              className="border-2 border-dashed border-indigo-100 rounded-2xl bg-indigo-50/30 p-8 flex flex-col items-center justify-center text-center group cursor-pointer hover:bg-indigo-50/50 hover:border-indigo-300 transition-all"
            >
              <div className="w-12 h-12 bg-white rounded-2xl shadow-sm flex items-center justify-center mb-3 group-hover:scale-110 transition-transform">
                <FileSpreadsheet size={24} className="text-indigo-600" />
              </div>
              <p className="text-sm font-semibold text-slate-700">اضغط لرفع ملف Excel</p>
              <p className="text-xs text-slate-400 mt-1">يدعم XLSX, XLS, CSV</p>
            </div>
            <input
              ref={fileInputRef}
              type="file"
              className="hidden"
              accept=".xlsx, .xls, .csv"
              onChange={handleFileUpload}
            />
            
            {fileName && (
              <button 
                onClick={clearFile}
                className="w-full mt-4 py-2 text-sm font-medium text-red-500 bg-red-50 rounded-xl hover:bg-red-100 transition-colors flex items-center justify-center gap-2"
              >
                <X size={16} />
                إلغاء الملف
              </button>
            )}
          </section>

          {/* Statistics Cards */}
          <div className="grid grid-cols-2 gap-4">
            <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm group hover:border-indigo-200 transition-colors">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-normal mb-1">إجمالي السجلات</p>
              <p className="text-2xl font-bold text-slate-800 tracking-normal">{data.length.toLocaleString()}</p>
            </div>
            <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm group hover:border-indigo-200 transition-colors">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-normal mb-1">عدد الأعمدة</p>
              <p className="text-2xl font-bold text-slate-800 tracking-normal">{headers.length}</p>
            </div>
          </div>

          {/* Financial Summary */}
          <section className="bg-white rounded-3xl p-6 border border-slate-200 shadow-sm overflow-hidden relative">
            <div className="absolute top-0 left-0 w-1 h-full bg-indigo-600"></div>
            <h2 className="text-[10px] font-bold text-slate-400 mb-6 uppercase tracking-normal flex items-center gap-2">
              <BarChart3 size={14} className="text-indigo-600" />
              الملخص المالي {(filterType !== 'all' || searchQuery) ? '(للنتائج المعروضة)' : '(للملف بالكامل)'}
            </h2>
            
            <div className="space-y-6">
              <div className="flex flex-col">
                <div className="flex items-center justify-between mb-1">
                  <span className="text-sm font-medium text-slate-500">
                    {(filterType !== 'all' || searchQuery) ? 'إجمالي المبالغ بالبحث' : 'إجمالي المبلغ (كلي)'}
                  </span>
                  <span className="text-xs bg-indigo-50 text-indigo-600 px-2 py-0.5 rounded-full font-bold">
                    {totals.amountCol || 'لم يحدد'}
                  </span>
                </div>
                <p className="text-3xl font-black text-slate-800 tracking-normal">
                  {totals.totalAmountFiltered.toLocaleString('ar-SA')} <span className="text-xs font-normal text-slate-400">ر.س</span>
                </p>
              </div>

              <div className="h-px bg-slate-100 w-full"></div>

              <div className="grid grid-cols-2 gap-4">
                <div className="flex flex-col bg-orange-50/50 p-3 rounded-2xl border border-orange-100">
                  <div className="flex items-center justify-between mb-1">
                    <span className="text-[10px] font-medium text-orange-700">
                      {(filterType !== 'all' || searchQuery) ? 'المتبقي (للبحث)' : 'إجمالي المتبقي (كلي)'}
                    </span>
                    <span className="text-[9px] bg-orange-100 text-orange-700 px-1.5 py-0.5 rounded-full font-bold">
                      {totals.remainingCol || 'لم يحدد'}
                    </span>
                  </div>
                  <p className="text-xl font-black text-orange-600 tracking-tight">
                    {totals.totalRemainingFiltered.toLocaleString('ar-SA')} <span className="text-[10px] font-normal text-orange-400">ر.س</span>
                  </p>
                </div>

                <div className="flex flex-col bg-emerald-50/50 p-3 rounded-2xl border border-emerald-100">
                  <div className="flex items-center justify-between mb-1">
                    <span className="text-[10px] font-medium text-emerald-700">
                      {(filterType !== 'all' || searchQuery) ? 'المسدد (للبحث)' : 'إجمالي المسدد (كلي)'}
                    </span>
                    <span className="text-[9px] bg-emerald-100 text-emerald-700 px-1.5 py-0.5 rounded-full font-bold">
                      تم سدادها
                    </span>
                  </div>
                  <p className="text-xl font-black text-emerald-600 tracking-tight">
                    {totals.totalSettledFiltered.toLocaleString('ar-SA')} <span className="text-[10px] font-normal text-emerald-400">ر.س</span>
                  </p>
                </div>

                <div className="flex flex-col bg-red-50/50 p-3 rounded-2xl border border-red-100 col-span-2">
                  <div className="flex items-center justify-between mb-1">
                    <span className="text-[10px] font-medium text-red-700">إجمالي مبالغ التصفية (للتصفية حالياً)</span>
                    <span className="text-[9px] bg-red-100 text-red-700 px-1.5 py-0.5 rounded-full font-bold">
                      تصفية
                    </span>
                  </div>
                  <p className="text-xl font-black text-red-600 tracking-tight">
                    {totals.totalLiquidationFiltered.toLocaleString('ar-SA')} <span className="text-[10px] font-normal text-red-400">ر.س</span>
                  </p>
                </div>
              </div>

              <div className="bg-slate-50 p-3 rounded-xl border border-slate-100">
                <p className="text-[10px] text-slate-500 font-bold flex items-center gap-2">
                  <Info size={12} className="text-indigo-500" />
                  المسدد اﻹجمالي: للفترة بالكامل | المسدد حالياً: للبحث الحالي
                </p>
              </div>
            </div>

            {!totals.amountCol && !totals.remainingCol && fileName && (
              <p className="mt-6 text-[10px] text-slate-400 italic text-center leading-relaxed">
                * لم يتم العثور على أعمدة مخصصة للمبالغ تلقائياً. تأكد من تسمية الأعمدة بكلمات مثل (المبلغ) أو (المتبقي).
              </p>
            )}

            {/* Final Save/Download Button */}
            <AnimatePresence>
              {isModified && (
                <motion.div
                  initial={{ opacity: 0, height: 0, marginTop: 0 }}
                  animate={{ opacity: 1, height: 'auto', marginTop: 24 }}
                  exit={{ opacity: 0, height: 0, marginTop: 0 }}
                  className="overflow-hidden"
                >
                  <button 
                    onClick={() => exportToExcel(true)}
                    className="w-full py-4 bg-emerald-600 text-white rounded-2xl font-black text-sm hover:bg-emerald-700 transition-all shadow-xl shadow-emerald-100 flex items-center justify-center gap-3 group"
                  >
                    <Database size={20} className="group-hover:scale-110 transition-transform" />
                    <span>حفظ وتحميل الملف المعدّل (كامل)</span>
                  </button>
                  <p className="text-[10px] text-emerald-600 font-bold mt-2 text-center flex items-center justify-center gap-1">
                    <Info size={10} />
                    تم إجراء تعديلات على البيانات
                  </p>
                </motion.div>
              )}
            </AnimatePresence>
          </section>
        </aside>

        {/* Main Search Area */}
        <div className="col-span-12 lg:col-span-8 flex flex-col gap-6">
          <div className="flex flex-col gap-4">
            {/* Search Bar */}
            <div className="relative group">
              <input 
                type="text" 
                placeholder="ابحث في كافة الحقول والأعمدة (مثلاً: اسم، رقم، عقار، أو أي قيمة أخرى)..." 
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                disabled={!fileName || isLoading}
                className="w-full h-16 px-6 rounded-2xl border-none shadow-xl focus:ring-2 focus:ring-indigo-500 text-lg text-slate-700 placeholder-slate-400 bg-white transition-all disabled:bg-slate-50 disabled:cursor-not-allowed"
              />
              {searchQuery && (
                <button 
                  onClick={() => setSearchQuery('')}
                  className="absolute left-4 top-1/2 -translate-y-1/2 p-2 hover:bg-slate-100 rounded-lg text-slate-400 transition-colors"
                  title="مسح البحث"
                >
                  <X size={20} />
                </button>
              )}
            </div>

            {/* Filter Bar */}
            {fileName && (
              <div className="flex items-center gap-2 overflow-x-auto pb-2 scrollbar-none px-2">
                {[
                  { id: 'all', label: 'الكل', icon: Database, count: filterCounts.all },
                  { id: 'paid', label: 'المسددة', icon: Check, count: filterCounts.paid, color: 'text-emerald-600', bg: 'bg-emerald-50', activeBg: 'bg-emerald-600' },
                  { id: 'modified', label: 'المعدلة', icon: Clock, count: filterCounts.modified, color: 'text-blue-600', bg: 'bg-blue-50', activeBg: 'bg-blue-600' },
                  { id: 'normal', label: 'العادية', icon: FileSpreadsheet, count: filterCounts.normal }
                ].map((filter) => {
                  const isActive = filterType === filter.id;
                  const Icon = filter.icon;
                  return (
                    <button
                      key={filter.id}
                      onClick={() => setFilterType(filter.id as any)}
                      className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold whitespace-nowrap transition-all border ${
                        isActive 
                          ? `${filter.activeBg || 'bg-indigo-600'} text-white border-transparent shadow-lg shadow-indigo-100` 
                          : `${filter.bg || 'bg-white'} ${filter.color || 'text-slate-600'} border-slate-200 hover:border-indigo-300`
                      }`}
                    >
                      <Icon size={14} />
                      <span>{filter.label}</span>
                      <span className={`px-1.5 py-0.5 rounded-lg text-[10px] ${isActive ? 'bg-white/20' : 'bg-slate-100/50'}`}>
                        {filter.count}
                      </span>
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          {/* Results Container */}
          <div className="flex-1 flex flex-col gap-6 overflow-hidden min-h-[500px]">
            {isLoading ? (
              <div className="flex-1 flex flex-col items-center justify-center bg-white rounded-3xl border border-slate-200 shadow-sm transition-all animate-pulse">
                <div className="w-12 h-12 border-4 border-indigo-600 border-t-transparent rounded-full animate-spin mb-4"></div>
                <p className="text-slate-500 font-medium">جاري معالجة البيانات...</p>
              </div>
            ) : !fileName ? (
              <div className="flex-1 flex flex-col items-center justify-center bg-white rounded-3xl border border-slate-200 shadow-sm text-center p-8">
                <div className="w-20 h-20 bg-slate-50 rounded-full flex items-center justify-center mb-6 text-slate-200">
                  <Database size={48} strokeWidth={1} />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">ابدأ برفع ملف للبحث</h3>
                <p className="text-slate-500 max-w-sm">ارفع ملف إكسل يحتوي على بياناتك، وسنساعدك في البحث عنها بسرعة فائقة.</p>
              </div>
            ) : filteredData.length === 0 ? (
              <div className="flex-1 flex flex-col items-center justify-center bg-white rounded-3xl border border-slate-200 shadow-sm text-center p-8">
                <div className="w-20 h-20 bg-red-50 rounded-full flex items-center justify-center mb-6 text-red-300">
                  <SearchSlash size={48} strokeWidth={1} />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">لا توجد نتائج مطابقة</h3>
                <p className="text-slate-500 max-w-sm">
                  {searchQuery ? `لم نجد أي سجل يحتوي على "${searchQuery}".` : 'لا توجد بيانات لهذا الفلتر حالياً.'} 
                  حاول تجربة خيارات بحث أو فلاتر أخرى.
                </p>
              </div>
            ) : (
              <div className="flex flex-col gap-6 flex-1 overflow-hidden">
                {/* Search Results Header */}
                <div className="flex justify-between items-end px-2">
                  <div>
                    <h1 className="text-2xl font-bold text-slate-800">نتائج البحث</h1>
                    <p className="text-slate-500 font-medium">تم العثور على {filteredData.length} سجل مطابق</p>
                  </div>
                  <div className="flex gap-2">
                    <button 
                      onClick={() => exportToExcel(false)}
                      className="px-6 py-3 bg-indigo-600 text-white rounded-xl hover:bg-indigo-700 transition-all shadow-md flex items-center gap-2 group"
                      title={filterType !== 'all' ? `تحميل ${filterCounts[filterType]} سجل من ${filterType === 'paid' ? 'المسددة' : filterType === 'modified' ? 'المعدلة' : 'العادية'}` : "تحميل نتائج البحث الحالية"}
                    >
                      <Download size={18} className="group-hover:translate-y-0.5 transition-transform" />
                      <span className="text-sm font-bold">
                        {filterType === 'all' ? 'تحميل النتائج' : `تحميل ${filterType === 'paid' ? 'المسددة' : filterType === 'modified' ? 'المعدلة' : 'العادية'}`}
                      </span>
                    </button>
                  </div>
                </div>

                {/* Table View */}
                <div className="flex-1 bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
                  <div className="flex-1 overflow-auto custom-scrollbar">
                    <table className="w-full text-right border-collapse">
                      <thead className="sticky top-0 bg-slate-50 z-20 border-b border-slate-200">
                        <tr>
                          <th className="px-6 py-4 text-[10px] font-bold text-slate-400 uppercase text-center w-12 hover:text-indigo-600 transition-colors">#</th>
                          <th className="px-6 py-4 text-[10px] font-bold text-slate-400 uppercase text-center w-24">إجراء</th>
                          {headers.map((header) => (
                            <th 
                              key={header} 
                              className={`px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-wider hover:text-indigo-600 transition-colors whitespace-nowrap relative group/tooltip ${
                                header.includes('الملاحظات') || header.includes('ملاحظات') || header.includes('notes') 
                                  ? 'min-w-[300px]' 
                                  : header.includes('الدفع') || header.includes('payment')
                                  ? 'min-w-[150px]'
                                  : 'min-w-[140px]'
                              }`}
                            >
                              <div className="flex items-center justify-center gap-1.5">
                                {header}
                                <Info size={12} className="text-slate-300 group-hover/tooltip:text-indigo-400 transition-colors" />
                              </div>
                              
                              {/* Tooltip */}
                              <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 px-3 py-2 bg-slate-800 text-white text-[10px] rounded-lg opacity-0 invisible group-hover/tooltip:opacity-100 group-hover/tooltip:visible transition-all duration-200 z-50 w-48 pointer-events-none shadow-xl border border-slate-700">
                                <div className="relative z-10 text-center font-medium leading-relaxed">
                                  {getHeaderTooltip(header)}
                                </div>
                                <div className="absolute top-full left-1/2 -translate-x-1/2 -mt-1 border-l-4 border-r-4 border-t-4 border-transparent border-t-slate-800"></div>
                              </div>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredData.map((row, idx) => {
                          // Find original index in 'data' to maintain status correctly
                          const originalIdx = data.findIndex(r => r === row);
                          const isPaid = paidRows.has(originalIdx);
                          const isModifiedRow = modifiedRows.has(originalIdx);
                          
                          // Styling classes based on status
                          let rowBgClass = "hover:bg-indigo-50/30";
                          let textClass = "text-slate-600";
                          let numBgClass = "text-slate-400 bg-slate-50/30";
                          
                          if (isPaid) {
                            rowBgClass = "bg-emerald-50 hover:bg-emerald-100/80";
                            textClass = "text-emerald-700 font-bold";
                            numBgClass = "text-emerald-500 bg-emerald-50/50";
                          } else if (isModifiedRow) {
                            // Find any note-like field and check for "تصفية"
                            const rowString = JSON.stringify(row);
                            if (rowString.includes('تصفية')) {
                              rowBgClass = "bg-red-50/70 hover:bg-red-100/70 border-r-4 border-r-red-400";
                              textClass = "text-red-700 font-bold";
                              numBgClass = "text-red-500 bg-red-100/50";
                            } else {
                              rowBgClass = "bg-blue-50/70 hover:bg-blue-100/70 border-r-4 border-r-blue-400";
                              textClass = "text-blue-700 font-bold";
                              numBgClass = "text-blue-500 bg-blue-100/50";
                            }
                          }
                          
                          return (
                            <motion.tr
                              key={idx}
                              initial={{ opacity: 0 }}
                              animate={{ opacity: 1 }}
                              transition={{ delay: Math.min(idx * 0.01, 0.2) }}
                              className={`transition-colors group ${rowBgClass}`}
                            >
                              <td className={`px-6 py-4 text-xs font-medium text-center ${numBgClass}`}>
                                {idx + 1}
                              </td>
                               <td className="px-4 py-2 text-center">
                                <div className="flex flex-col items-center gap-2">
                                  <button
                                    onClick={() => togglePaid(originalIdx)}
                                    className={`px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all w-full max-w-[60px] ${
                                      isPaid 
                                        ? 'bg-emerald-500 text-white shadow-sm' 
                                        : 'bg-slate-100 text-slate-500 hover:bg-emerald-500 hover:text-white'
                                    }`}
                                  >
                                    {isPaid ? 'تم السداد' : 'تسوية'}
                                  </button>
                                  {modifiedRows.has(originalIdx) && (
                                    <button 
                                      onClick={() => undoRowChanges(originalIdx)}
                                      className="flex items-center gap-1 text-[9px] text-orange-500 hover:text-orange-600 font-bold bg-orange-50 px-2 py-0.5 rounded transition-colors"
                                      title="تراجع عن التعديل"
                                    >
                                      <RotateCcw size={10} />
                                      <span>تراجع</span>
                                    </button>
                                  )}
                                </div>
                              </td>
                              {headers.map((header) => {
                                const isNoteColumn = header.includes('الملاحظات') || header.includes('ملاحظات') || header.includes('notes');
                                const isEditing = editingCell?.row === originalIdx && editingCell?.col === header;

                                const isAmount = header === totals.amountCol;
                                const isRemaining = header === totals.remainingCol;

                                return (
                                  <td 
                                    key={header} 
                                    onClick={() => isNoteColumn && setEditingCell({ row: originalIdx, col: header })}
                                    className={`px-6 py-4 text-sm font-medium transition-all ${textClass} ${
                                      isNoteColumn 
                                        ? 'whitespace-normal min-w-[300px] break-words cursor-pointer' 
                                        : header.includes('الدفع') || header.includes('payment')
                                        ? 'whitespace-normal min-w-[150px]'
                                        : 'whitespace-nowrap max-w-[250px] overflow-hidden text-ellipsis'
                                    }`}
                                  >
                                    {isEditing ? (
                                      <input 
                                        autoFocus
                                        className="w-full bg-white border border-indigo-300 rounded px-2 py-1 outline-none ring-2 ring-indigo-100"
                                        value={String(row[header] || '')}
                                        onChange={(e) => updateNote(originalIdx, header, e.target.value)}
                                        onBlur={() => setEditingCell(null)}
                                        onKeyDown={(e) => e.key === 'Enter' && setEditingCell(null)}
                                      />
                                    ) : (
                                      <div className="flex items-start justify-between gap-2">
                                        <span>{highlightText(String(row[header] || '-'), searchQuery)}</span>
                                        {isNoteColumn && (
                                          <Info size={12} className="text-slate-300 mt-1 flex-shrink-0 opacity-0 group-hover:opacity-100 transition-opacity" />
                                        )}
                                      </div>
                                    )}
                                  </td>
                                );
                              })}
                            </motion.tr>
                          );
                        })}
                      </tbody>

                      {/* PDF/Print Financial Summary Row */}
                      <tfoot className="bg-slate-50 font-bold border-t-2 border-slate-200">
                        <tr>
                          <td className="px-6 py-4 text-xs text-slate-400 text-center">Σ</td>
                          <td className="px-6 py-4 text-xs text-slate-400 text-center">-</td>
                          {headers.map((header) => {
                            const isAmount = header === totals.amountCol;
                            const isRemaining = header === totals.remainingCol;
                            
                            return (
                              <td key={header} className={`px-6 py-4 text-sm ${isAmount ? 'text-indigo-600' : isRemaining ? 'text-orange-600' : 'text-slate-500'}`}>
                                {isAmount ? (
                                  <div className="flex items-center gap-2 whitespace-nowrap">
                                    <span className="text-[10px] text-slate-400 font-normal">إجمالي المبلغ:</span>
                                    <span>{totals.totalAmountFiltered.toLocaleString('ar-SA')} ر.س</span>
                                  </div>
                                ) : isRemaining ? (
                                  <div className="flex items-center gap-4 whitespace-nowrap">
                                    <div className="flex items-center gap-2">
                                      <span className="text-[10px] text-slate-400 font-normal">المتبقي:</span>
                                      <span>{totals.totalRemainingFiltered.toLocaleString('ar-SA')} ر.س</span>
                                    </div>
                                    <div className="flex items-center gap-2 bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded-lg border border-emerald-100">
                                      <span className="text-[10px] opacity-70">المسدد حالياً:</span>
                                      <span className="font-bold">{totals.totalSettledFiltered.toLocaleString('ar-SA')} ر.س</span>
                                    </div>
                                    <div className="flex items-center gap-2 bg-red-50 text-red-600 px-2 py-0.5 rounded-lg border border-red-100">
                                      <span className="text-[10px] opacity-70">التصفية:</span>
                                      <span className="font-bold">{totals.totalLiquidationFiltered.toLocaleString('ar-SA')} ر.س</span>
                                    </div>
                                  </div>
                                ) : header === headers[0] ? (
                                  <div className="flex items-center gap-2 text-indigo-600">
                                    <BarChart3 size={14} />
                                    <span>ملخص مالي</span>
                                  </div>
                                ) : ''}
                              </td>
                            );
                          })}
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                  
                  {/* Table Footer / Summary */}
                  <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-[10px] font-bold text-slate-400 uppercase tracking-normal">
                    <span>نهاية النتائج المطابقة</span>
                    <span>إجمالي الصفوف المعروضة: {filteredData.length}</span>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </main>

      {/* Bottom Status Bar */}
      <footer className="h-10 bg-slate-800 px-6 flex items-center justify-between text-[10px] text-slate-400 uppercase tracking-normal shrink-0 mt-auto">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2">
            <span className="w-1.5 h-1.5 bg-green-400 rounded-full shadow-[0_0_8px_rgba(74,222,128,0.5)]"></span>
            <span>نظام البحث: متصل</span>
          </div>
          <span className="w-1 h-1 bg-slate-600 rounded-full hidden sm:block"></span>
          <span className="hidden sm:block">آخر تحديث: {new Date().toLocaleTimeString('ar-SA')}</span>
        </div>
        <div className="flex items-center gap-2">
          <span className="text-indigo-400 font-bold">SMART ENGINE V1.0</span>
        </div>
      </footer>
    </div>
  );
}

function headerLabel(header: string) {
  return header.replace(/([A-Z])/g, ' $1').trim();
}

function highlightText(text: string, highlight: string) {
  if (!highlight.trim()) return text;
  // Escape special characters for RegExp
  const escapedHighlight = highlight.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const parts = text.split(new RegExp(`(${escapedHighlight})`, 'gi'));
  return (
    <span>
      {parts.map((part, i) => (
        <span
          key={i}
          className={part.toLowerCase() === highlight.toLowerCase() ? 'bg-indigo-100 text-indigo-700 px-0.5 rounded' : ''}
        >
          {part}
        </span>
      ))}
    </span>
  );
}

