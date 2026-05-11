/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
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
  Printer,
  Download
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface DataRow {
  [key: string]: string | number | null | undefined;
}

export default function App() {
  const [data, setData] = useState<DataRow[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const filteredData = useMemo(() => {
    if (!searchQuery.trim()) return [];
    const query = searchQuery.toLowerCase();
    const results = data.filter(row => {
      return Object.values(row).some(value => 
        String(value || '').toLowerCase().includes(query)
      );
    });
    return results;
  }, [data, searchQuery]);

  // Financial Summary Calculation
  const totals = useMemo(() => {
    const currentData = searchQuery.trim() ? filteredData : data;
    
    // Attempt to find columns for Amount and Remaining
    const amountKeywords = ['مبلغ', 'المبلغ', 'قيمة', 'القيمة', 'amount', 'total', 'price', 'السعر'];
    const remainingKeywords = ['متبقي', 'المتبقي', 'باقي', 'الباقي', 'remaining', 'balance', 'left'];

    const amountCol = headers.find(h => amountKeywords.some(k => h.toLowerCase().includes(k)));
    const remainingCol = headers.find(h => remainingKeywords.some(k => h.toLowerCase().includes(k)));

    let totalAmount = 0;
    let totalRemaining = 0;

    currentData.forEach(row => {
      if (amountCol) {
        const val = parseFloat(String(row[amountCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) totalAmount += val;
      }
      if (remainingCol) {
        const val = parseFloat(String(row[remainingCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) totalRemaining += val;
      }
    });

    return { totalAmount, totalRemaining, amountCol, remainingCol };
  }, [data, filteredData, headers, searchQuery]);

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
              const { data: savedData, headers: savedHeaders, fileName: savedName } = getRequest.result;
              setData(savedData);
              setHeaders(savedHeaders);
              setFileName(savedName);
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
  const saveToDB = (fileData: DataRow[], fileHeaders: string[], name: string) => {
    try {
      const dbRequest = indexedDB.open('FileSearchDB', 1);
      dbRequest.onsuccess = (e: any) => {
        const db = e.target.result;
        const transaction = db.transaction(['files'], 'readwrite');
        const store = transaction.objectStore('files');
        store.put({
          id: 'current_file',
          data: fileData,
          headers: fileHeaders,
          fileName: name,
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
          saveToDB(jsonData, newHeaders, file.name);
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
    removeFromDB();
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans" dir="rtl">
      {/* Top Navigation Bar */}
      <nav className="h-16 bg-white border-b border-slate-200 px-6 md:px-8 flex items-center justify-between shadow-sm shrink-0 sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white">
            <Search size={22} strokeWidth={2.5} />
          </div>
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
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">إجمالي السجلات</p>
              <p className="text-2xl font-bold text-slate-800 tracking-tight">{data.length.toLocaleString()}</p>
            </div>
            <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm group hover:border-indigo-200 transition-colors">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">عدد الأعمدة</p>
              <p className="text-2xl font-bold text-slate-800 tracking-tight">{headers.length}</p>
            </div>
          </div>

          {/* Financial Summary */}
          <section className="bg-white rounded-3xl p-6 border border-slate-200 shadow-sm overflow-hidden relative">
            <div className="absolute top-0 left-0 w-1 h-full bg-indigo-600"></div>
            <h2 className="text-[10px] font-bold text-slate-400 mb-6 uppercase tracking-wider flex items-center gap-2">
              <BarChart3 size={14} className="text-indigo-600" />
              الملخص المالي {searchQuery.trim() ? '(للنتائج الحالية)' : '(للملف كاملاً)'}
            </h2>
            
            <div className="space-y-6">
              <div className="flex flex-col">
                <div className="flex items-center justify-between mb-1">
                  <span className="text-sm font-medium text-slate-500">إجمالي المبلغ</span>
                  <span className="text-xs bg-indigo-50 text-indigo-600 px-2 py-0.5 rounded-full font-bold">
                    {totals.amountCol || 'لم يحدد'}
                  </span>
                </div>
                <p className="text-3xl font-black text-slate-800 tracking-tight">
                  {totals.totalAmount.toLocaleString('ar-SA')} <span className="text-xs font-normal text-slate-400">ر.س</span>
                </p>
              </div>

              <div className="h-px bg-slate-100 w-full"></div>

              <div className="flex flex-col">
                <div className="flex items-center justify-between mb-1">
                  <span className="text-sm font-medium text-slate-500">إجمالي المتبقي</span>
                  <span className="text-xs bg-orange-50 text-orange-600 px-2 py-0.5 rounded-full font-bold">
                    {totals.remainingCol || 'لم يحدد'}
                  </span>
                </div>
                <p className="text-3xl font-black text-orange-600 tracking-tight">
                  {totals.totalRemaining.toLocaleString('ar-SA')} <span className="text-xs font-normal text-slate-400">ر.س</span>
                </p>
              </div>
            </div>

            {!totals.amountCol && !totals.remainingCol && fileName && (
              <p className="mt-6 text-[10px] text-slate-400 italic text-center leading-relaxed">
                * لم يتم العثور على أعمدة مخصصة للمبالغ تلقائياً. تأكد من تسمية الأعمدة بكلمات مثل (المبلغ) أو (المتبقي).
              </p>
            )}
          </section>
        </aside>

        {/* Main Search Area */}
        <div className="col-span-12 lg:col-span-8 flex flex-col gap-6">
          {/* Search Bar */}
          <div className="relative group">
            <div className="absolute inset-y-0 right-6 flex items-center pointer-events-none text-slate-400 group-focus-within:text-indigo-600 transition-colors">
              <Search size={24} />
            </div>
            <input 
              type="text" 
              placeholder="ابحث في كافة الحقول والأعمدة (مثلاً: اسم، رقم، عقار، أو أي قيمة أخرى)..." 
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              disabled={!fileName || isLoading}
              className="w-full h-16 pr-16 pl-6 rounded-2xl border-none shadow-xl focus:ring-2 focus:ring-indigo-500 text-lg text-slate-700 placeholder-slate-400 bg-white transition-all disabled:bg-slate-50 disabled:cursor-not-allowed"
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
            ) : !searchQuery ? (
              <div className="flex-1 flex flex-col items-center justify-center bg-white rounded-3xl border border-slate-200 shadow-sm text-center p-8">
                <div className="w-20 h-20 bg-indigo-50 rounded-full flex items-center justify-center mb-6 text-indigo-300">
                  <Search size={48} strokeWidth={1} />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">اكتب شيئاً للبحث عنه</h3>
                <p className="text-slate-500 max-w-sm">أدخل أي معلومة للبحث عنها بداخل كافة بيانات الملف المرفوع (اسم الشخص، العقار، رقم الجوال، وغيرها).</p>
              </div>
            ) : filteredData.length === 0 ? (
              <div className="flex-1 flex flex-col items-center justify-center bg-white rounded-3xl border border-slate-200 shadow-sm text-center p-8">
                <div className="w-20 h-20 bg-red-50 rounded-full flex items-center justify-center mb-6 text-red-300">
                  <SearchSlash size={48} strokeWidth={1} />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">لا توجد نتائج مطابقة</h3>
                <p className="text-slate-500 max-w-sm">لم نجد أي سجل يحتوي على "{searchQuery}". حاول تجربة كلمات بحث أخرى.</p>
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
                    <button className="p-2.5 bg-white border border-slate-200 rounded-xl text-slate-600 hover:bg-slate-50 hover:text-indigo-600 transition-all shadow-sm">
                      <Printer size={18} />
                    </button>
                    <button className="p-2.5 bg-white border border-slate-200 rounded-xl text-slate-600 hover:bg-slate-50 hover:text-indigo-600 transition-all shadow-sm">
                      <Download size={18} />
                    </button>
                  </div>
                </div>

                {/* Detailed Record List */}
                <div className="flex-1 overflow-y-auto pr-1 space-y-4 pb-8">
                  {filteredData.map((row, idx) => (
                    <motion.div
                      key={idx}
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      transition={{ delay: idx * 0.05 }}
                      className="bg-white rounded-3xl border border-slate-200 shadow-sm hover:shadow-md transition-shadow flex flex-col overflow-hidden"
                    >
                      <div className="p-6 md:p-8 flex items-start gap-4 md:gap-6 border-b border-slate-100 bg-slate-50/30">
                        <div className="w-16 h-16 md:w-20 md:h-20 bg-gradient-to-tr from-indigo-500 to-purple-500 rounded-2xl md:rounded-3xl shadow-lg flex items-center justify-center text-2xl md:text-3xl font-bold text-white uppercase shrink-0">
                          {String(Object.values(row)[0] || '?')[0]}
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center flex-wrap gap-2 md:gap-3 mb-1">
                            <h2 className="text-xl md:text-2xl font-bold text-slate-800 truncate">{String(Object.values(row).find(v => typeof v === 'string' && v.length > 3) || 'سجل رقم ' + (idx + 1))}</h2>
                            <span className="px-3 py-0.5 bg-green-100 text-green-700 rounded-full text-[10px] font-bold uppercase tracking-wider">سجل موثق</span>
                          </div>
                          <p className="text-slate-400 text-sm italic">سجل معرف بـ: {headerLabel(headers[0])} #{idx + 1000}</p>
                        </div>
                      </div>
                      
                      <div className="p-6 md:p-8 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-y-8 gap-x-12">
                        {headers.map((header) => (
                          <div key={header} className="space-y-1 group">
                            <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest block transition-colors group-hover:text-indigo-400">
                              {header}
                            </label>
                            <p className="text-base md:text-lg font-semibold text-slate-700 break-all">
                              {highlightText(String(row[header] || '-'), searchQuery)}
                            </p>
                          </div>
                        ))}
                      </div>

                      {/* Action Footer */}
                      <div className="mt-auto p-4 md:p-6 bg-slate-50 border-t border-slate-100 flex justify-end gap-3 translate-y-0 shadow-[0_-1px_3px_rgba(0,0,0,0.02)]">
                        <button className="px-5 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-600 text-sm hover:bg-slate-50 transition-colors">عرض التفاصيل</button>
                        <button className="px-5 py-2 bg-indigo-600 text-white rounded-xl font-bold text-sm shadow-sm shadow-indigo-200 hover:bg-indigo-700 transition-colors">تعديل</button>
                      </div>
                    </motion.div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      </main>

      {/* Bottom Status Bar */}
      <footer className="h-10 bg-slate-800 px-6 flex items-center justify-between text-[10px] text-slate-400 uppercase tracking-widest shrink-0 mt-auto">
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
  const parts = text.split(new RegExp(`(${highlight})`, 'gi'));
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

