import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Users, 
  Settings2, 
  RotateCcw, 
  Copy, 
  Check, 
  Trash2, 
  Shuffle, 
  Plus,
  ArrowRightLeft,
  LayoutGrid,
  ClipboardList,
  Sparkles,
  Download,
  Upload,
  FileText,
  Link as LinkIcon,
  Globe,
  History,
  Calendar,
  UserPlus,
  User,
  LogOut
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

type Mode = 'count' | 'size';

export default function App() {
  const [inputText, setInputText] = useState('');
  const [mode, setMode] = useState<Mode>('count');
  const [value, setValue] = useState(2);
  const [results, setResults] = useState<string[][]>([]);
  const [history, setHistory] = useState<{
    id: string;
    timestamp: number;
    mode: Mode;
    value: number;
    groups: string[][];
    names: string[];
    createdBy?: string;
  }[]>([]);
  const [userName, setUserName] = useState<string>('');
  const [isEditingUser, setIsEditingUser] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [copiedIndex, setCopiedIndex] = useState<number | null>(null);
  const [copiedHistoryId, setCopiedHistoryId] = useState<string | null>(null);
  const [exportHistoryId, setExportHistoryId] = useState<string | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [sheetUrl, setSheetUrl] = useState('');
  const [isFetchingSheet, setIsFetchingSheet] = useState(false);
  const [showSheetInput, setShowSheetInput] = useState(false);
  const resultsRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [activeTab, setActiveTab] = useState<'tool' | 'history'>('tool');
  const [showSaveToast, setShowSaveToast] = useState(false);
  const [currentUser, setCurrentUser] = useState<string | null>(null);

  const authorizedUsers = ['林大維', '陳大明', '王小明'];

  // Load history and user from localStorage on mount
  React.useEffect(() => {
    const savedHistory = localStorage.getItem('grouping_history');
    if (savedHistory) {
      try {
        setHistory(JSON.parse(savedHistory));
      } catch (e) {
        console.error('Failed to parse history', e);
      }
    }
    const savedUser = localStorage.getItem('grouping_current_user');
    if (savedUser && authorizedUsers.includes(savedUser)) {
      setCurrentUser(savedUser);
    }
  }, []);

  const handleLogin = (user: string) => {
    setCurrentUser(user);
    localStorage.setItem('grouping_current_user', user);
  };

  const handleLogout = () => {
    setCurrentUser(null);
    localStorage.removeItem('grouping_current_user');
    setResults([]); // Clear results on logout for security
    setActiveTab('tool'); // Reset to tool tab
  };

  // Save history to localStorage
  const saveToHistory = (newGroups: string[][], currentMode: Mode, currentValue: number, currentNames: string[]) => {
    const newItem = {
      id: Math.random().toString(36).substr(2, 9),
      timestamp: Date.now(),
      mode: currentMode,
      value: currentValue,
      groups: newGroups,
      names: currentNames,
      createdBy: currentUser || '未知使用者'
    };
    const updatedHistory = [newItem, ...history].slice(0, 50); // Keep last 50
    setHistory(updatedHistory);
    localStorage.setItem('grouping_history', JSON.stringify(updatedHistory));
    
    // Visual feedback
    setShowSaveToast(true);
    setTimeout(() => setShowSaveToast(false), 3000);
  };

  const handleExportHistoryItem = (item: any, format: 'txt' | 'csv' | 'xlsx') => {
    const groups = item.groups;
    if (format === 'txt') {
      const text = groups.map((group: string[], i: number) => `第 ${i + 1} 組：\n${group.join('\n')}`).join('\n\n');
      const blob = new Blob([text], { type: 'text/plain' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `分組結果_${new Date(item.timestamp).toLocaleDateString()}_${item.id}.txt`;
      link.click();
      URL.revokeObjectURL(url);
    } else if (format === 'csv') {
      const maxLen = Math.max(...groups.map((g: string[]) => g.length));
      let csvContent = "";
      csvContent += groups.map((_: any, i: number) => `第 ${i + 1} 組`).join(",") + "\n";
      for (let i = 0; i < maxLen; i++) {
        csvContent += groups.map((g: string[]) => g[i] || "").join(",") + "\n";
      }
      const blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `分組結果_${new Date(item.timestamp).toLocaleDateString()}_${item.id}.csv`;
      link.click();
      URL.revokeObjectURL(url);
    } else if (format === 'xlsx') {
      const maxLen = Math.max(...groups.map((g: string[]) => g.length));
      const data = [];
      data.push(groups.map((_: any, i: number) => `第 ${i + 1} 組`));
      for (let i = 0; i < maxLen; i++) {
        data.push(groups.map((g: string[]) => g[i] || ""));
      }
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "分組結果");
      XLSX.writeFile(wb, `分組結果_${new Date(item.timestamp).toLocaleDateString()}_${item.id}.xlsx`);
    }
    setExportHistoryId(null);
  };

  const handleExportAllHistory = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(history, null, 2));
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", `分組全歷程備份_${new Date().toLocaleDateString()}.json`);
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
  };

  const deleteHistoryItem = (id: string) => {
    const updatedHistory = history.filter(item => item.id !== id);
    setHistory(updatedHistory);
    localStorage.setItem('grouping_history', JSON.stringify(updatedHistory));
  };

  const clearHistory = () => {
    if (confirm('確定要清空所有分組歷程嗎？')) {
      setHistory([]);
      localStorage.removeItem('grouping_history');
    }
  };

  const names = useMemo(() => {
    return inputText
      .split(/[\n,，、\s\t]+/)
      .map(name => name.trim())
      .filter(name => name.length > 0);
  }, [inputText]);

  const handleGenerate = () => {
    if (names.length === 0) return;
    
    setIsGenerating(true);
    
    // Artificial delay for "cute animation" feel
    setTimeout(() => {
      const shuffled = [...names].sort(() => Math.random() - 0.5);
      let newResults: string[][] = [];

      if (mode === 'count') {
        const groupsCount = Math.max(1, Math.min(value, names.length));
        for (let i = 0; i < groupsCount; i++) {
          newResults.push([]);
        }
        shuffled.forEach((name, index) => {
          newResults[index % groupsCount].push(name);
        });
      } else {
        const groupSize = Math.max(1, value);
        for (let i = 0; i < shuffled.length; i += groupSize) {
          newResults.push(shuffled.slice(i, i + groupSize));
        }
      }

      // Sort within each group by surname (first character)
      newResults = newResults.map(group => 
        [...group].sort((a, b) => {
          const surnameA = a[0] || '';
          const surnameB = b[0] || '';
          return surnameA.localeCompare(surnameB, 'zh-Hant');
        })
      );

      setResults(newResults);
      saveToHistory(newResults, mode, value, names);
      setIsGenerating(false);
      
      setTimeout(() => {
        resultsRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }, 100);
    }, 1800); // Longer delay for animation visibility
  };

  const handleCopy = (groupIndex: number, group: string[]) => {
    const text = `第 ${groupIndex + 1} 組：\n${group.join('\n')}`;
    navigator.clipboard.writeText(text);
    setCopiedIndex(groupIndex);
    setTimeout(() => setCopiedIndex(null), 2000);
  };

  const handleCopyAll = () => {
    const text = results.map((group, i) => `第 ${i + 1} 組：\n${group.join('\n')}`).join('\n\n');
    navigator.clipboard.writeText(text);
  };

  const loadSample = () => {
    const sample = [
      '林大維', '陳曉明', '張美玲', '王建國', '李靜宜', 
      '趙自強', '黃雅婷', '周思潔', '吳志明', '徐若瑄', 
      '劉德華', '郭富城', '黎明', '張學友', '王菲'
    ].join('\n');
    setInputText(sample);
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const fileReader = new FileReader();

    fileReader.onload = (event) => {
      try {
        const data = event.target?.result;
        if (!data) return;

        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert sheet to array of arrays
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        
        // Extract all non-empty cell values from all rows
        const flatNames: string[] = [];
        rows.forEach(row => {
          row.forEach(cell => {
            if (cell !== null && cell !== undefined) {
              const str = String(cell).trim();
              if (str) {
                // If the cell contains multiple names separated by commas or spaces, flatten them
                const splitNames = str.split(/[\n,，、\s\t]+/).filter(n => n.trim().length > 0);
                flatNames.push(...splitNames);
              }
            }
          });
        });

        if (flatNames.length > 0) {
          setInputText(flatNames.join('\n'));
        }
      } catch (error) {
        console.error('Error parsing file:', error);
        // Fallback for simple txt if xlsx fails (though xlsx should handle csv/txt too)
        const reader = new FileReader();
        reader.onload = (txtEvent) => {
          const content = txtEvent.target?.result as string;
          if (content) setInputText(content);
        };
        reader.readAsText(file);
      }
    };

    fileReader.readAsBinaryString(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const [showExportMenu, setShowExportMenu] = useState<boolean>(false);
  const [showResultExportMenu, setShowResultExportMenu] = useState<boolean>(false);

  const handleExportResults = (format: 'txt' | 'csv' | 'xlsx') => {
    if (results.length === 0) return;
    
    if (format === 'txt') {
      const text = results.map((group, i) => `第 ${i + 1} 組：\n${group.join('\n')}`).join('\n\n');
      const blob = new Blob([text], { type: 'text/plain' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `分組結果_${new Date().toLocaleDateString()}.txt`;
      link.click();
      URL.revokeObjectURL(url);
    } else if (format === 'csv') {
      const maxLen = Math.max(...results.map(g => g.length));
      let csvContent = "";
      // Headers
      csvContent += results.map((_, i) => `第 ${i + 1} 組`).join(",") + "\n";
      // Rows
      for (let i = 0; i < maxLen; i++) {
        csvContent += results.map(g => g[i] || "").join(",") + "\n";
      }
      const blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `分組結果_${new Date().toLocaleDateString()}.csv`;
      link.click();
      URL.revokeObjectURL(url);
    } else if (format === 'xlsx') {
      const maxLen = Math.max(...results.map(g => g.length));
      const data = [];
      // Headers
      const headerRow = results.map((_, i) => `第 ${i + 1} 組`);
      data.push(headerRow);
      // Rows
      for (let i = 0; i < maxLen; i++) {
        data.push(results.map(g => g[i] || ""));
      }
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "分組結果");
      XLSX.writeFile(wb, `分組結果_${new Date().toLocaleDateString()}.xlsx`);
    }
    setShowResultExportMenu(false);
  };

  const handleExportNames = (format: 'txt' | 'csv' | 'xlsx') => {
    if (names.length === 0) return;

    if (format === 'txt') {
      const blob = new Blob([inputText], { type: 'text/plain' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `成員名單_${new Date().toLocaleDateString()}.txt`;
      link.click();
      URL.revokeObjectURL(url);
    } else if (format === 'csv' || format === 'xlsx') {
      const data = names.map(n => [n]);
      const ws = XLSX.utils.aoa_to_sheet([["姓名"], ...data]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "名單");
      
      if (format === 'csv') {
        XLSX.writeFile(wb, `成員名單_${new Date().toLocaleDateString()}.csv`, { bookType: 'csv' });
      } else {
        XLSX.writeFile(wb, `成員名單_${new Date().toLocaleDateString()}.xlsx`);
      }
    }
    setShowExportMenu(false);
  };

  const handleFetchSheet = async () => {
    if (!sheetUrl) return;
    setIsFetchingSheet(true);
    try {
      const response = await fetch(`/api/fetch-gsheet?url=${encodeURIComponent(sheetUrl)}`);
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || '無法抓取試算表');
      }
      const csvData = await response.text();
      
      // Basic CSV parsing for the proxy result
      const rows = csvData.split('\n').map(row => row.split(','));
      const flatNames: string[] = [];
      rows.forEach(row => {
        row.forEach(cell => {
          const name = cell.replace(/^"|"$/g, '').trim();
          if (name && name !== '姓名') { // Filter out header if simple
            flatNames.push(name);
          }
        });
      });

      if (flatNames.length > 0) {
        setInputText(flatNames.join('\n'));
        setShowSheetInput(false);
        setSheetUrl('');
      }
    } catch (error: any) {
      alert(error.message);
    } finally {
      setIsFetchingSheet(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#07070a] text-slate-100 font-sans selection:bg-violet-500/30">
      {/* Login Overlay */}
      <AnimatePresence>
        {!currentUser && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[100] bg-[#07070a] flex items-center justify-center p-6"
          >
            <div className="max-w-md w-full space-y-8 text-center">
              <div className="space-y-4">
                <motion.div 
                  initial={{ scale: 0.5, opacity: 0 }}
                  animate={{ scale: 1, opacity: 1 }}
                  transition={{ type: "spring", damping: 15 }}
                  className="mx-auto w-20 h-20 bg-gradient-to-br from-violet-500 to-indigo-600 rounded-[2rem] shadow-2xl shadow-violet-500/20 flex items-center justify-center"
                >
                  <Users className="w-10 h-10 text-white" />
                </motion.div>
                <h2 className="text-3xl font-black font-display tracking-tight text-white">請選擇您的身分</h2>
                <p className="text-slate-500 text-sm font-medium">存取分組小幫手需要先識別您的身分</p>
              </div>

              <div className="grid gap-3">
                {authorizedUsers.map((user) => (
                  <button
                    key={user}
                    onClick={() => handleLogin(user)}
                    className="w-full py-4 bg-slate-900 border border-white/5 rounded-2xl text-lg font-bold hover:border-violet-500/50 hover:bg-violet-500/5 transition-all group flex items-center justify-center gap-3"
                  >
                    <div className="w-8 h-8 bg-slate-800 rounded-full flex items-center justify-center group-hover:bg-violet-500/20">
                      <User className="w-4 h-4 text-slate-400 group-hover:text-violet-400" />
                    </div>
                    {user}
                  </button>
                ))}
              </div>
              
              <p className="text-[10px] text-slate-700 uppercase font-black tracking-[0.2em] pt-8">
                SECURE ACCESS • PRIVATE SESSION
              </p>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Loading Overlay */}
      <AnimatePresence>
        {isGenerating && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[100] flex flex-col items-center justify-center bg-black/80 backdrop-blur-xl"
          >
            <motion.div
              animate={{ 
                rotate: [0, 10, -10, 0],
                scale: [1, 1.1, 1]
              }}
              transition={{ repeat: Infinity, duration: 1 }}
              className="relative mb-8"
            >
              <div className="absolute inset-0 bg-violet-500 blur-2xl opacity-40 animate-pulse" />
              <div className="relative p-6 bg-gradient-to-br from-violet-400 to-indigo-600 rounded-[2.5rem] shadow-2xl">
                <Shuffle className="w-16 h-16 text-white animate-bounce" />
              </div>
            </motion.div>
            
            <motion.div 
              initial={{ y: 20, opacity: 0 }}
              animate={{ y: 0, opacity: 1 }}
              className="text-center space-y-2"
            >
              <h3 className="text-2xl font-black text-white font-display tracking-widest uppercase">
                神奇魔法分組中...
              </h3>
              <p className="text-slate-400 font-medium animate-pulse">正在為成員尋找最佳拍檔 ✨</p>
            </motion.div>

            {/* Floating particles */}
            {[...Array(6)].map((_, i) => (
              <motion.div
                key={i}
                animate={{ 
                  y: [0, -100, 0],
                  x: [0, Math.sin(i) * 50, 0],
                  opacity: [0, 1, 0]
                }}
                transition={{ 
                  repeat: Infinity, 
                  duration: 2 + Math.random(),
                  delay: i * 0.2
                }}
                className="absolute text-violet-400/40 pointer-events-none"
                style={{ 
                  left: `${20 + i * 12}%`,
                  bottom: '20%'
                }}
              >
                <Sparkles className="w-8 h-8" />
              </motion.div>
            ))}
          </motion.div>
        )}
      </AnimatePresence>

      {/* Background Glow */}
      <div className="fixed top-0 left-1/2 -translate-x-1/2 w-full max-w-4xl h-screen pointer-events-none">
        <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-violet-600/10 blur-[120px] rounded-full" />
        <div className="absolute bottom-[10%] right-[-5%] w-[30%] h-[30%] bg-indigo-600/10 blur-[100px] rounded-full" />
      </div>

      {/* Header */}
      <header className="max-w-4xl mx-auto pt-16 px-8 pb-4 relative z-10">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
          <div className="flex items-center gap-4">
            <motion.div 
              initial={{ scale: 0.8, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              className="p-3 bg-gradient-to-br from-violet-500 to-indigo-600 rounded-2xl shadow-lg shadow-violet-500/20"
            >
              <Users className="w-7 h-7 text-white" />
            </motion.div>
            <div>
              <h1 className="text-4xl font-extrabold tracking-tight font-display bg-clip-text text-transparent bg-gradient-to-r from-white to-slate-400">
                分組棒棒小幫手
              </h1>
              <div className="flex items-center gap-2 mt-1">
                <p className="text-slate-400 text-sm font-medium">
                  專為高效團隊設計的隨機分組工具
                </p>
                <div className="w-1 h-1 bg-slate-700 rounded-full" />
                <button 
                  onClick={handleLogout}
                  className="text-[10px] font-bold text-violet-400/80 hover:text-violet-400 border border-violet-500/20 px-2 py-1 rounded-lg flex items-center gap-1.5 transition-all hover:bg-violet-500/5 group/logout"
                  title="切換使用者"
                >
                  <User className="w-2.5 h-2.5" />
                  操作者：{currentUser}
                  <div className="w-px h-2 bg-violet-500/20 mx-0.5" />
                  <LogOut className="w-2.5 h-2.5 group-hover/logout:translate-x-0.5 transition-transform" />
                  <span className="opacity-60">登出</span>
                </button>
              </div>
            </div>
          </div>

          <div className="flex p-1.5 bg-slate-900/60 backdrop-blur-md rounded-2xl border border-white/5 shadow-xl">
            <button
              onClick={() => setActiveTab('tool')}
              className={`flex items-center gap-2 px-6 py-2.5 rounded-xl text-sm font-bold transition-all cursor-pointer ${
                activeTab === 'tool' 
                  ? 'bg-white text-black shadow-lg' 
                  : 'text-slate-500 hover:text-slate-300'
              }`}
            >
              <Shuffle className="w-4 h-4" />
              分組工具
            </button>
            <button
              onClick={() => setActiveTab('history')}
              className={`flex items-center gap-2 px-6 py-2.5 rounded-xl text-sm font-bold transition-all relative cursor-pointer ${
                activeTab === 'history' 
                  ? 'bg-white text-black shadow-lg' 
                  : 'text-slate-500 hover:text-slate-300'
              }`}
            >
              <History className="w-4 h-4" />
              歷程記錄
              {history.length > 0 && (
                <span className="absolute -top-1 -right-1 w-5 h-5 bg-violet-600 text-white text-[10px] flex items-center justify-center rounded-full border-2 border-[#07070a]">
                  {history.length}
                </span>
              )}
            </button>
          </div>
        </div>
      </header>

      <AnimatePresence>
        {showSaveToast && (
          <motion.div 
            initial={{ opacity: 0, y: -20, x: '-50%' }}
            animate={{ opacity: 1, y: 0, x: '-50%' }}
            exit={{ opacity: 0, y: -20, x: '-50%' }}
            className="fixed top-8 left-1/2 z-[110] px-6 py-3 bg-indigo-600 text-white rounded-full font-bold shadow-2xl flex items-center gap-3 border border-indigo-400"
          >
            <Check className="w-4 h-4" />
            已自動儲存至歷程記錄
          </motion.div>
        )}
      </AnimatePresence>

      <main className="max-w-4xl mx-auto p-8 space-y-12 relative z-10">
        {activeTab === 'tool' ? (
          <>
            <section className="grid md:grid-cols-2 gap-10 items-start">
          {/* Input Box */}
          <div className="space-y-4">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <label className="flex items-center gap-2 text-xs font-bold uppercase tracking-[0.2em] text-slate-500">
                <ClipboardList className="w-3.5 h-3.5" />
                成員名單 <span className="text-violet-400 ml-1">{names.length}</span>
              </label>
              <div className="flex gap-2">
                <input 
                  type="file" 
                  ref={fileInputRef} 
                  onChange={handleImport} 
                  className="hidden" 
                  accept=".txt,.csv,.xlsx,.xls"
                />
                <button 
                  onClick={() => fileInputRef.current?.click()}
                  className="p-1 px-2.5 rounded-lg bg-white/5 border border-white/10 text-[10px] font-bold uppercase tracking-widest text-slate-400 hover:text-white hover:bg-white/10 transition-all cursor-pointer flex items-center gap-1.5"
                >
                  <Upload className="w-3 h-3" />
                  匯入
                </button>
                <button 
                  onClick={() => setShowSheetInput(!showSheetInput)}
                  className={`p-1 px-2.5 rounded-lg border text-[10px] font-bold uppercase tracking-widest transition-all cursor-pointer flex items-center gap-1.5 ${
                    showSheetInput 
                      ? 'bg-violet-600 border-violet-500 text-white shadow-lg shadow-violet-600/20' 
                      : 'bg-white/5 border-white/10 text-slate-400 hover:text-white hover:bg-white/10'
                  }`}
                >
                  <Globe className="w-3 h-3" />
                  雲端
                </button>
                <div className="relative">
                  <button 
                    onClick={() => setShowExportMenu(!showExportMenu)}
                    disabled={names.length === 0}
                    className="p-1 px-2.5 rounded-lg bg-white/5 border border-white/10 text-[10px] font-bold uppercase tracking-widest text-slate-400 hover:text-white hover:bg-white/10 transition-all cursor-pointer flex items-center gap-1.5 disabled:opacity-30 disabled:cursor-not-allowed"
                  >
                    <Download className="w-3 h-3" />
                    匯出
                  </button>
                  <AnimatePresence>
                    {showExportMenu && (
                      <>
                        <div className="fixed inset-0 z-40" onClick={() => setShowExportMenu(false)} />
                        <motion.div 
                          initial={{ opacity: 0, y: 10, scale: 0.95 }}
                          animate={{ opacity: 1, y: 0, scale: 1 }}
                          exit={{ opacity: 0, y: 10, scale: 0.95 }}
                          className="absolute right-0 mt-2 w-32 bg-slate-900 border border-white/10 rounded-xl shadow-2xl z-50 overflow-hidden"
                        >
                          {(['txt', 'csv', 'xlsx'] as const).map((fmt) => (
                            <button
                              key={fmt}
                              onClick={() => handleExportNames(fmt)}
                              className="w-full text-left px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 hover:bg-white/5 hover:text-violet-400 transition-colors border-b border-white/5 last:border-0"
                            >
                              .{fmt} 檔案
                            </button>
                          ))}
                        </motion.div>
                      </>
                    )}
                  </AnimatePresence>
                </div>
              </div>
            </div>

            <AnimatePresence>
              {showSheetInput && (
                <motion.div 
                  initial={{ height: 0, opacity: 0 }}
                  animate={{ height: 'auto', opacity: 1 }}
                  exit={{ height: 0, opacity: 0 }}
                  className="overflow-hidden"
                >
                  <div className="p-5 bg-violet-600/5 rounded-2xl border border-violet-500/20 space-y-4 mb-4">
                    <div className="flex items-center justify-between">
                      <span className="text-[10px] font-black text-violet-400 uppercase tracking-widest">Google 試算表連結</span>
                      <a 
                        href="https://support.google.com/docs/answer/183965" 
                        target="_blank" 
                        rel="noreferrer"
                        className="text-[10px] text-slate-500 hover:text-violet-400 underline"
                      >
                        如何設定公開？
                      </a>
                    </div>
                    <div className="flex gap-2">
                      <input 
                        type="text" 
                        value={sheetUrl}
                        onChange={(e) => setSheetUrl(e.target.value)}
                        placeholder="請貼上試算表網址..."
                        className="flex-1 bg-black/40 border border-white/5 rounded-xl px-4 py-2 text-sm focus:border-violet-500/50 outline-none transition-all"
                      />
                      <button 
                        onClick={handleFetchSheet}
                        disabled={!sheetUrl || isFetchingSheet}
                        className="px-4 py-2 bg-violet-600 hover:bg-violet-500 disabled:opacity-50 disabled:cursor-not-allowed rounded-xl text-xs font-bold transition-all flex items-center gap-2 whitespace-nowrap"
                      >
                        {isFetchingSheet ? <RotateCcw className="w-3 h-3 animate-spin" /> : '抓取資料'}
                      </button>
                    </div>
                    <p className="text-[10px] text-slate-600 italic">
                      請確保試算表已設定為「知道連結的使用者即可查看」。
                    </p>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
            
            <div className="relative group">
              <textarea
                value={inputText}
                onChange={(e) => setInputText(e.target.value)}
                placeholder="請輸入成員。可以用換行、逗號、空格隔開..."
                className="w-full h-96 p-8 bg-slate-900/40 backdrop-blur-sm border border-white/5 rounded-[2.5rem] shadow-2xl focus:ring-1 focus:ring-violet-500/30 focus:border-violet-500/50 transition-all resize-none outline-none text-lg leading-relaxed placeholder:text-slate-700"
              />
              <div className="absolute top-4 right-8 text-[10px] uppercase font-black tracking-widest text-white/5 pointer-events-none">
                INPUT AREA
              </div>
            </div>

            <div className="flex gap-4">
              <button 
                onClick={loadSample}
                className="flex-1 flex items-center justify-center gap-1.5 px-4 py-2.5 rounded-xl bg-violet-500/10 border border-violet-500/20 text-xs font-bold uppercase tracking-widest text-violet-400 hover:bg-violet-500/20 transition-all cursor-pointer group/sample"
              >
                <Sparkles className="w-3 h-3 group-hover/sample:animate-pulse" />
                載入預設範例
              </button>
              <button 
                onClick={() => setInputText('')}
                className="flex items-center justify-center gap-1.5 px-4 py-2.5 rounded-xl bg-slate-800/20 border border-white/5 text-xs font-bold uppercase tracking-widest text-slate-500 hover:text-slate-300 hover:bg-slate-800/40 transition-all cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                清空名單
              </button>
            </div>
          </div>

          {/* Controls */}
          <div className="space-y-8 h-full flex flex-col justify-between">
            <div className="space-y-8 bg-slate-900/40 backdrop-blur-sm p-10 rounded-[2.5rem] border border-white/5 shadow-2xl relative overflow-hidden">
              <div className="absolute top-0 right-0 p-6 opacity-10 pointer-events-none">
                <Settings2 className="w-24 h-24 text-violet-500" />
              </div>
              
              <h2 className="flex items-center gap-2 text-xs font-bold uppercase tracking-[0.2em] text-slate-500">
                <Settings2 className="w-4 h-4" />
                分組模式與設定
              </h2>

              <div className="space-y-6">
                <div className="flex p-1.5 bg-black/40 rounded-2xl border border-white/5">
                  <button
                    onClick={() => setMode('count')}
                    className={`flex-1 py-3.5 px-4 rounded-xl text-sm font-bold transition-all ${
                      mode === 'count' 
                        ? 'bg-gradient-to-br from-violet-500 to-indigo-600 text-white shadow-lg' 
                        : 'text-slate-500 hover:text-slate-300'
                    }`}
                  >
                    指定組數
                  </button>
                  <button
                    onClick={() => setMode('size')}
                    className={`flex-1 py-3.5 px-4 rounded-xl text-sm font-bold transition-all ${
                      mode === 'size' 
                        ? 'bg-gradient-to-br from-violet-500 to-indigo-600 text-white shadow-lg' 
                        : 'text-slate-500 hover:text-slate-300'
                    }`}
                  >
                    每組人數
                  </button>
                </div>

                <div className="pt-2">
                  <div className="flex justify-between items-end mb-6">
                    <span className="text-sm font-semibold text-slate-400">
                      {mode === 'count' ? '要把成員分成幾組？' : '每組理想人數？'}
                    </span>
                    <span className="text-4xl font-black text-white font-display">
                      {value}
                    </span>
                  </div>
                  <div className="relative h-2 w-full bg-slate-800 rounded-full overflow-hidden">
                    <div 
                      className="absolute top-0 left-0 h-full bg-gradient-to-r from-violet-500 to-indigo-600 rounded-full"
                      style={{ width: `${(value / Math.max(10, names.length)) * 100}%` }}
                    />
                    <input
                      type="range"
                      min="1"
                      max={Math.max(10, names.length)}
                      value={value}
                      onChange={(e) => setValue(parseInt(e.target.value))}
                      className="absolute top-0 left-0 w-full h-full opacity-0 cursor-pointer z-20"
                    />
                  </div>
                  <div className="flex justify-between mt-3 text-[10px] font-black text-slate-700 uppercase tracking-widest">
                    <span>MIN</span>
                    <span>MAX ({Math.max(10, names.length)})</span>
                  </div>
                </div>
              </div>

              <motion.button
                whileHover={{ scale: 1.02 }}
                whileTap={{ scale: 0.98 }}
                onClick={handleGenerate}
                disabled={names.length === 0 || isGenerating}
                className={`w-full py-5 rounded-2xl flex items-center justify-center gap-3 font-extrabold text-xl transition-all shadow-xl
                  ${names.length === 0 
                    ? 'bg-slate-800 text-slate-600 cursor-not-allowed border border-white/5' 
                    : 'bg-white text-black hover:bg-slate-100 hover:shadow-white/10'
                  }`}
              >
                {isGenerating ? (
                  <RotateCcw className="w-6 h-6 animate-spin" />
                ) : (
                  <>
                    <Shuffle className="w-6 h-6" />
                    開始分組
                  </>
                )}
              </motion.button>
            </div>

            {results.length > 0 && (
              <motion.div 
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="p-8 bg-violet-600/5 rounded-[2.5rem] border border-violet-500/20 text-center relative group overflow-hidden"
              >
                <div className="absolute inset-0 bg-gradient-to-br from-violet-500/5 to-transparent pointer-events-none" />
                <div className="relative z-10">
                  <div className="text-violet-400 text-xs font-black uppercase tracking-[0.3em] mb-4 flex items-center justify-center gap-2">
                    <Sparkles className="w-4 h-4" />
                    產生了 {results.length} 組結果
                  </div>
                  <div className="flex gap-3 justify-center">
                    <button 
                      onClick={handleCopyAll}
                      className="py-3 px-6 bg-white/10 border border-white/10 backdrop-blur-md rounded-xl text-sm font-bold hover:bg-white/20 transition-all flex items-center gap-2 cursor-pointer"
                    >
                      <Copy className="w-4 h-4 text-violet-400" />
                      複製
                    </button>
                    <div className="relative">
                      <button 
                        onClick={() => setShowResultExportMenu(!showResultExportMenu)}
                        className="py-3 px-6 bg-violet-600 border border-violet-500 rounded-xl text-sm font-bold hover:bg-violet-500 transition-all flex items-center gap-2 cursor-pointer text-white shadow-lg shadow-violet-600/20"
                      >
                        <Download className="w-4 h-4" />
                        匯出結果
                      </button>
                      <AnimatePresence>
                        {showResultExportMenu && (
                          <>
                            <div className="fixed inset-0 z-40" onClick={() => setShowResultExportMenu(false)} />
                            <motion.div 
                              initial={{ opacity: 0, y: -10, scale: 0.95 }}
                              animate={{ opacity: 1, y: 0, scale: 1 }}
                              exit={{ opacity: 0, y: -10, scale: 0.95 }}
                              className="absolute left-0 bottom-full mb-3 w-40 bg-slate-900 border border-violet-500/30 rounded-2xl shadow-2xl z-50 overflow-hidden"
                            >
                              <div className="p-3 border-b border-white/5 bg-violet-500/5 text-center">
                                <span className="text-[10px] font-black text-violet-400 uppercase tracking-widest">選擇格式</span>
                              </div>
                              {(['txt', 'csv', 'xlsx'] as const).map((fmt) => (
                                <button
                                  key={fmt}
                                  onClick={() => handleExportResults(fmt)}
                                  className="w-full text-center px-4 py-3 text-xs font-bold uppercase tracking-widest text-slate-300 hover:bg-violet-500 hover:text-white transition-all border-b border-white/5 last:border-0"
                                >
                                  {fmt.toUpperCase()} 格式
                                </button>
                              ))}
                            </motion.div>
                          </>
                        )}
                      </AnimatePresence>
                    </div>
                  </div>
                </div>
              </motion.div>
            )}
          </div>
        </section>

        {/* Results */}
        {results.length > 0 && (
          <div ref={resultsRef} className="pt-12">
            <div className="flex items-center gap-4 mb-10">
              <div className="h-px flex-1 bg-gradient-to-r from-transparent to-slate-800" />
              <h2 className="text-3xl font-black font-display tracking-tight flex items-center gap-3 decoration-violet-500 decoration-4 underline-offset-8">
                <LayoutGrid className="w-8 h-8 text-violet-500" />
                分區顯示結果
              </h2>
              <div className="h-px flex-1 bg-gradient-to-l from-transparent to-slate-800" />
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-8">
              <AnimatePresence mode="popLayout">
                {results.map((group, index) => (
                  <motion.div
                    key={`group-${index}-${group.join('-')}`}
                    layout
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, scale: 0.95 }}
                    transition={{ 
                      type: "spring", 
                      stiffness: 400, 
                      damping: 35,
                      delay: index * 0.04 
                    }}
                    className="bg-slate-900/40 backdrop-blur-sm rounded-[2rem] border border-white/5 shadow-2xl overflow-hidden group hover:border-violet-500/30 transition-colors"
                  >
                    <div className="bg-white/5 px-6 py-4 border-b border-white/5 flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <span className="flex items-center justify-center w-8 h-8 rounded-full bg-violet-600 text-white font-black text-sm">
                          {index + 1}
                        </span>
                        <span className="text-xs font-black text-slate-500 uppercase tracking-widest">
                           MEMBER: {group.length}
                        </span>
                      </div>
                      <button
                        onClick={() => handleCopy(index, group)}
                        className="p-2 bg-slate-800/50 hover:bg-violet-600 rounded-xl transition-all text-slate-400 hover:text-white cursor-pointer"
                        title="複製"
                      >
                        {copiedIndex === index ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                      </button>
                    </div>
                    <div className="p-7 space-y-3 max-h-72 overflow-y-auto custom-scrollbar">
                      {group.map((name, nameIdx) => (
                        <div 
                          key={nameIdx}
                          className="flex items-center justify-between p-3 rounded-xl hover:bg-white/5 transition-all group/name border border-transparent hover:border-white/5"
                        >
                          <div className="flex items-center gap-4">
                            <div className="w-1.5 h-1.5 rounded-full bg-violet-500 shadow-[0_0_8px_rgba(139,92,246,0.8)]" />
                            <span className="text-slate-200 font-bold tracking-wide">{name}</span>
                          </div>
                          <span className="text-[10px] font-black text-slate-700 opacity-0 group-hover/name:opacity-100 transition-opacity">
                            USER
                          </span>
                        </div>
                      ))}
                    </div>
                  </motion.div>
                ))}
              </AnimatePresence>
            </div>
          </div>
        )}
      </>
    ) : (
      /* History Section */
      <section className="pt-4">
            <div className="flex items-center justify-between mb-8">
              <div className="flex flex-col gap-1">
                <h2 className="text-3xl font-black font-display flex items-center gap-3">
                  <History className="w-7 h-7 text-indigo-500" />
                  分組歷程
                </h2>
                <p className="text-slate-500 text-sm font-medium">
                  所有數據皆加密存儲於您的瀏覽器 LocalStorage
                </p>
              </div>
              {history.length > 0 && (
                <div className="flex items-center gap-4">
                  <button 
                    onClick={handleExportAllHistory}
                    className="text-xs font-bold uppercase tracking-widest text-violet-400 hover:text-violet-300 transition-colors flex items-center gap-2 cursor-pointer p-2 rounded-lg hover:bg-violet-500/5 border border-violet-500/20"
                  >
                    <Download className="w-3 h-3" />
                    備份全部歷程
                  </button>
                  <button 
                    onClick={clearHistory}
                    className="text-xs font-bold uppercase tracking-widest text-slate-600 hover:text-red-400 transition-colors flex items-center gap-2 cursor-pointer p-2 rounded-lg hover:bg-red-500/5"
                  >
                    <Trash2 className="w-3 h-3" />
                    清空歷程
                  </button>
                </div>
              )}
            </div>

          <div className="space-y-4">
            {history.length === 0 ? (
              <div className="p-12 bg-slate-900/20 border border-white/5 border-dashed rounded-[2.5rem] text-center">
                <div className="p-4 bg-slate-800/50 rounded-2xl inline-flex mb-4">
                  <History className="w-8 h-8 text-slate-700" />
                </div>
                <p className="text-slate-500 font-medium tracking-wide">目前還沒有分組紀錄</p>
                <p className="text-slate-700 text-xs mt-1 uppercase font-bold tracking-widest italic">NO HISTORY FOUND</p>
              </div>
            ) : (
              history.map((item) => (
                <motion.div
                  key={item.id}
                  initial={{ opacity: 0, x: -10 }}
                  whileInView={{ opacity: 1, x: 0 }}
                  viewport={{ once: true }}
                  className="bg-slate-900/40 backdrop-blur-sm border border-white/5 rounded-3xl p-6 flex flex-col md:flex-row md:items-center justify-between gap-6 group hover:border-violet-500/20 transition-all shadow-xl hover:shadow-indigo-500/5"
                >
                  <div className="flex items-start gap-5">
                    <div className="p-4 bg-violet-600/10 rounded-2xl group-hover:bg-violet-600/20 transition-colors">
                      <Calendar className="w-6 h-6 text-violet-400" />
                    </div>
                    <div>
                      <div className="flex items-center gap-3 mb-1">
                        <span className="text-white font-bold text-lg">
                          {new Date(item.timestamp).toLocaleString('zh-TW', { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })}
                        </span>
                        <span className="px-2 py-0.5 bg-indigo-500/10 border border-indigo-500/20 rounded text-[10px] font-black text-indigo-400 uppercase tracking-widest">
                          {item.mode === 'count' ? `分成 ${item.value} 組` : `每組 ${item.value} 人`}
                        </span>
                      </div>
                      <div className="flex flex-wrap gap-x-3 gap-y-1">
                        <span className="text-slate-500 text-xs font-medium flex items-center gap-1">
                           執行者: <b className="text-violet-400 font-bold">{item.createdBy || '未知'}</b>
                        </span>
                        <span className="text-slate-500 text-xs font-medium flex items-center gap-1">
                           成員數: <b className="text-slate-300 font-bold">{item.names.length}</b>
                        </span>
                        <span className="text-slate-500 text-xs font-medium flex items-center gap-1">
                           結果組數: <b className="text-slate-300 font-bold">{item.groups.length}</b>
                        </span>
                      </div>
                    </div>
                  </div>

                  <div className="flex items-center gap-3">
                    <button
                      onClick={() => {
                        setInputText(item.names.join('\n'));
                        window.scrollTo({ top: 0, behavior: 'smooth' });
                      }}
                      className="flex-1 md:flex-none py-2.5 px-4 bg-white/5 border border-white/10 rounded-xl text-xs font-bold text-slate-300 hover:text-white hover:bg-white/10 transition-all flex items-center justify-center gap-2 cursor-pointer group/reuse"
                    >
                      <UserPlus className="w-3.5 h-3.5 group-hover/reuse:animate-bounce" />
                      沿用名單
                    </button>
                    <button
                      onClick={() => {
                        const text = item.groups.map((group, i) => `第 ${i + 1} 組：\n${group.join('\n')}`).join('\n\n');
                        navigator.clipboard.writeText(text);
                        setCopiedHistoryId(item.id);
                        setTimeout(() => setCopiedHistoryId(null), 2000);
                      }}
                      className="flex-1 md:flex-none py-2.5 px-4 bg-indigo-500/10 border border-indigo-500/20 rounded-xl text-xs font-bold text-indigo-400 hover:bg-indigo-500/20 transition-all flex items-center justify-center gap-2 cursor-pointer"
                    >
                      {copiedHistoryId === item.id ? (
                        <Check className="w-3.5 h-3.5" />
                      ) : (
                        <Copy className="w-3.5 h-3.5" />
                      )}
                      複製結果
                    </button>
                    
                    <div className="relative">
                      <button
                        onClick={() => setExportHistoryId(exportHistoryId === item.id ? null : item.id)}
                        className="flex-1 md:flex-none py-2.5 px-4 bg-slate-800/30 border border-white/5 rounded-xl text-xs font-bold text-slate-400 hover:text-white hover:bg-slate-800 transition-all flex items-center justify-center gap-2 cursor-pointer"
                      >
                        <Download className="w-3.5 h-3.5" />
                        匯出
                      </button>
                      <AnimatePresence>
                        {exportHistoryId === item.id && (
                          <>
                            <div className="fixed inset-0 z-40" onClick={() => setExportHistoryId(null)} />
                            <motion.div 
                              initial={{ opacity: 0, y: 10, scale: 0.95 }}
                              animate={{ opacity: 1, y: 0, scale: 1 }}
                              exit={{ opacity: 0, y: 10, scale: 0.95 }}
                              className="absolute right-0 mt-2 w-32 bg-slate-900 border border-white/10 rounded-xl shadow-2xl z-50 overflow-hidden"
                            >
                              {(['txt', 'csv', 'xlsx'] as const).map((fmt) => (
                                <button
                                  key={fmt}
                                  onClick={() => handleExportHistoryItem(item, fmt)}
                                  className="w-full text-left px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 hover:bg-white/5 hover:text-violet-400 transition-colors border-b border-white/5 last:border-0"
                                >
                                  .{fmt} 檔案
                                </button>
                              ))}
                            </motion.div>
                          </>
                        )}
                      </AnimatePresence>
                    </div>

                    <button
                      onClick={() => deleteHistoryItem(item.id)}
                      className="p-2.5 bg-slate-800/30 text-slate-600 hover:text-red-400 hover:bg-slate-800/50 rounded-xl transition-all cursor-pointer"
                      title="刪除紀錄"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </motion.div>
              ))
            )}
          </div>
        </section>
      )}
    </main>

      <footer className="max-w-4xl mx-auto py-24 px-8 text-center relative z-10">
        <div className="text-slate-600 text-sm font-bold uppercase tracking-[0.4em] flex items-center justify-center gap-3">
          <div className="h-px w-8 bg-slate-800" />
          Division Helper v3.0
          <div className="h-px w-8 bg-slate-800" />
        </div>
      </footer>
      
      <style>{`
        .custom-scrollbar::-webkit-scrollbar {
          width: 5px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: rgba(255,255,255,0.02);
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: rgba(255,255,255,0.1);
          border-radius: 10px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: rgba(139,92,246,0.3);
        }
        input[type="range"]::-webkit-slider-thumb {
          appearance: none;
          width: 20px;
          height: 20px;
          background: white;
          border-radius: 50%;
          cursor: pointer;
          border: 4px solid #8b5cf6;
          box-shadow: 0 0 15px rgba(139,92,246,0.5);
        }
      `}</style>
    </div>
  );
}

