import React, { useState, useEffect, useRef } from 'react';
import { 
  UploadCloud, 
  Sparkles, 
  Download, 
  Settings, 
  BookOpen, 
  FileText, 
  Layers, 
  CheckCircle2, 
  AlertCircle, 
  RefreshCw, 
  Copy, 
  Check, 
  Search, 
  Sliders, 
  Eye, 
  Columns, 
  Maximize2, 
  Languages, 
  Key, 
  Plus, 
  Trash2, 
  ArrowRight,
  HelpCircle,
  Zap,
  Cpu,
  Type,
  ShieldCheck,
  Layout,
  ExternalLink,
  ChevronLeft,
  ChevronRight,
  Maximize,
  Minimize2,
  SlidersHorizontal,
  FolderOpen
} from 'lucide-react';

const API_BASE = '/api';

export default function App() {
  // Session State
  const [session, setSession] = useState(null);
  const [activeSlideIdx, setActiveSlideIdx] = useState(1);
  const [loading, setLoading] = useState(false);
  const [translating, setTranslating] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [progress, setProgress] = useState(0);
  const [statusMsg, setStatusMsg] = useState('');

  // Settings State
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('gemini_api_key') || '');
  const [targetScript, setTargetScript] = useState('latin'); // 'latin' or 'cyrillic'
  const [domain, setDomain] = useState('IT & Dasturlash');
  const [autoFit, setAutoFit] = useState(true);
  const [viewMode, setViewMode] = useState('split'); // 'split', 'translated', 'original'

  // Modals & Search
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [showGlossaryModal, setShowGlossaryModal] = useState(false);
  const [showFontVaultModal, setShowFontVaultModal] = useState(false);
  const [fontsList, setFontsList] = useState({});
  const [fontSearch, setFontSearch] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState('all');
  const [copiedId, setCopiedId] = useState(null);

  // Custom Glossary
  const [glossary, setGlossary] = useState(() => {
    try {
      const saved = localStorage.getItem('slide_glossary');
      return saved ? JSON.parse(saved) : {
        "Cloud Computing": "Bulutli hisoblash",
        "Artificial Intelligence": "Sun'iy intellekt",
        "Machine Learning": "Mashinaviy o'rganish",
        "Framework": "Freymvork",
        "Pipeline": "Konveyer tizimi",
        "Roadmap": "Yo'l xaritasi",
        "Milestone": "Muhim bosqich",
        "SWOT Analysis": "SWOT tahlili"
      };
    } catch {
      return {};
    }
  });
  const [newTermKey, setNewTermKey] = useState('');
  const [newTermVal, setNewTermVal] = useState('');

  // Drag & drop
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef(null);

  useEffect(() => {
    localStorage.setItem('slide_glossary', JSON.stringify(glossary));
  }, [glossary]);

  useEffect(() => {
    // Dynamically load Google Fonts & Custom Cached Fonts CSS
    const linkId = 'slide-custom-fonts-css';
    if (!document.getElementById(linkId)) {
      const link = document.createElement('link');
      link.id = linkId;
      link.rel = 'stylesheet';
      link.href = '/api/fonts/css';
      document.head.appendChild(link);
    }
  }, []);

  useEffect(() => {
    if (apiKey) {
      localStorage.setItem('gemini_api_key', apiKey);
    }
  }, [apiKey]);

  // Fetch fonts list
  const loadFonts = async () => {
    try {
      const res = await fetch('/api/fonts');
      if (res.ok) {
        const data = await res.json();
        setFontsList(data.fonts || {});
      }
    } catch (e) {
      console.error("Fontlarni yuklashda xatolik:", e);
    }
  };

  useEffect(() => {
    loadFonts();
  }, []);

  // Handle File Upload
  const handleFileUpload = async (file) => {
    if (!file || !file.name.endsWith('.pptx')) {
      alert("Iltimos, faqat PowerPoint (.pptx) fayllarini yuklang!");
      return;
    }

    setLoading(true);
    setStatusMsg("PPTX fayli tahlil qilinmoqda...");
    const formData = new FormData();
    formData.append('file', file);

    try {
      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Faylni yuklashda xatolik yuz berdi.");
      }

      const data = await res.json();
      setSession(data);
      setActiveSlideIdx(1);
      setStatusMsg(`"${data.filename}" muvaffaqiyatli yuklandi (${data.slides_count} ta slayd, ${data.total_items} ta matn bloki).`);
      loadFonts();
    } catch (e) {
      alert("Xatolik: " + e.message);
      setStatusMsg("Yuklashda xatolik.");
    } finally {
      setLoading(false);
    }
  };

  // Handle Translate All
  const handleTranslateAll = async () => {
    if (!session) return;
    setTranslating(true);
    setProgress(10);
    setStatusMsg("Gemini AI orqali slaydlar tarjima qilinmoqda...");

    try {
      const res = await fetch('/api/translate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          session_id: session.session_id,
          target_script: targetScript,
          domain: domain,
          api_key: apiKey || undefined,
          glossary: Object.keys(glossary).length > 0 ? glossary : undefined
        })
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Tarjima jarayonida xatolik.");
      }

      const data = await res.json();
      if (data.session) {
        setSession(data.session);
      } else if (data.slides) {
        setSession(prev => ({
          ...prev,
          slides: data.slides
        }));
      }
      setProgress(100);
      setStatusMsg("Barcha slaydlar muvaffaqiyatli tarjima qilindi! Endi tahrirlashingiz yoki yuklab olishingiz mumkin.");
    } catch (e) {
      alert("Tarjima xatosi: " + e.message);
      setStatusMsg("Tarjimada xatolik yuz berdi.");
    } finally {
      setTranslating(false);
      setTimeout(() => setProgress(0), 1500);
    }
  };

  // Handle Item Text Change (Inline Edit)
  const handleItemTextChange = async (itemId, newText) => {
    if (!session) return;

    // Local state update immediately
    const updatedSlides = session.slides.map(slide => ({
      ...slide,
      items: slide.items.map(item => 
        item.id === itemId ? { ...item, translated_text: newText } : item
      )
    }));

    setSession(prev => ({
      ...prev,
      slides: updatedSlides
    }));

    // Debounced sync to server
    try {
      await fetch('/api/session/update_item', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          session_id: session.session_id,
          item_id: itemId,
          translated_text: newText
        })
      });
    } catch (e) {
      console.error("Yangilash xatosi:", e);
    }
  };

  // Toggle Script (Latin <-> Cyrillic)
  const handleToggleScript = (newScript) => {
    setTargetScript(newScript);
    if (!session) return;

    const updatedSlides = session.slides.map(slide => ({
      ...slide,
      items: slide.items.map(item => ({
        ...item,
        translated_text: transliterateUzbek(item.translated_text || item.original_text, newScript === 'latin')
      }))
    }));

    setSession(prev => ({
      ...prev,
      slides: updatedSlides
    }));
  };

  // Transliteration logic
  const transliterateUzbek = (text, toLatin = true) => {
    if (!text) return "";
    if (!toLatin) {
      const comp = [
        ["Sh", "Ш"], ["sh", "ш"], ["Ch", "Ч"], ["ch", "ч"],
        ["Yo", "Ё"], ["yo", "ё"], ["Yu", "Ю"], ["yu", "ю"],
        ["Ya", "Я"], ["ya", "я"], ["Ye", "Е"], ["ye", "е"],
        ["O'", "Ў"], ["o'", "ў"], ["Oʻ", "Ў"], ["oʻ", "ў"],
        ["G'", "Ғ"], ["g'", "ғ"], ["Gʻ", "Ғ"], ["gʻ", "ғ"],
        ["Ts", "Ц"], ["ts", "ц"]
      ];
      let res = text;
      for (const [lat, cyr] of comp) res = res.replaceAll(lat, cyr);
      const single = {
        'A':'А','a':'а','B':'Б','b':'б','D':'Д','d':'д','E':'Е','e':'е','F':'Ф','f':'ф',
        'G':'Г','g':'г','H':'Ҳ','h':'ҳ','I':'И','i':'и','J':'Ж','j':'ж','K':'К','k':'к',
        'L':'Л','l':'л','M':'М','m':'м','N':'Н','n':'н','O':'О','o':'о','P':'П','p':'п',
        'Q':'Қ','q':'қ','R':'Р','r':'р','S':'С','s':'с','T':'Т','t':'т','U':'У','u':'у',
        'V':'В','v':'в','X':'Х','x':'х','Y':'Й','y':'й','Z':'З','z':'з'
      };
      return res.split('').map(c => single[c] || c).join('');
    } else {
      const comp = [
        ["Ё", "Yo"], ["ё", "yo"], ["Ю", "Yu"], ["ю", "yu"], ["Ya", "Ya"], ["ya", "ya"],
        ["Ш", "Sh"], ["ш", "sh"], ["Ч", "Ch"], ["ch", "ch"],
        ["Ў", "Oʻ"], ["ў", "oʻ"], ["Ғ", "Gʻ"], ["ғ", "gʻ"], ["Ц", "Ts"], ["ц", "ts"]
      ];
      let res = text;
      for (const [cyr, lat] of comp) res = res.replaceAll(cyr, lat);
      const single = {
        'А':'A','а':'a','Б':'B','б':'b','В':'V','в':'v','Г':'G','г':'g','Д':'D','d':'d',
        'Е':'E','е':'e','Ж':'J','ж':'j','З':'Z','z':'z','И':'I','i':'i','Й':'Y','й':'y',
        'К':'K','к':'k','Л':'L','л':'l','М':'M','m':'m','Н':'N','н':'n','О':'O','о':'o',
        'П':'P','п':'p','Р':'R','r':'r','С':'S','с':'s','Т':'T','t':'t','У':'U','у':'u',
        'Ф':'F','ф':'f','Х':'X','х':'x','Ҳ':'H','ҳ':'h','Қ':'Q','қ':'q','Э':'E','э':'e'
      };
      return res.split('').map(c => single[c] || c).join('');
    }
  };

  // Export & Download PPTX
  const handleExport = async () => {
    if (!session) return;
    setExporting(true);
    setStatusMsg("Tarjima qilingan PowerPoint (.pptx) yaratilmoqda...");

    try {
      const res = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          session_id: session.session_id,
          auto_fit: autoFit,
          target_script: targetScript,
          api_key: apiKey || undefined
        })
      });

      if (!res.ok) {
        let errorMsg = "Eksport qilishda xatolik yuz berdi.";
        try {
          const errJson = await res.json();
          errorMsg = errJson.detail || errorMsg;
        } catch (_) {
          const errText = await res.text();
          errorMsg = errText || errorMsg;
        }
        throw new Error(errorMsg);
      }

      const data = await res.json();
      let dlName = data.download_filename || 'Taqdimot.pptx';
      if (!dlName.toLowerCase().endsWith('.pptx')) {
        dlName += '.pptx';
      }

      // Download via Blob to guarantee Windows extension
      const blobRes = await fetch(data.download_url);
      if (!blobRes.ok) throw new Error("Faylni serverdan yuklab olishda xatolik.");
      const blobData = await blobRes.blob();
      const blobUrl = window.URL.createObjectURL(blobData);

      const link = document.createElement('a');
      link.href = blobUrl;
      link.download = dlName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(blobUrl);

      setStatusMsg(`"${dlName}" muvaffaqiyatli yuklab olindi!`);
    } catch (e) {
      alert("Eksport xatosi: " + e.message);
    } finally {
      setExporting(false);
    }
  };

  // Copy to clipboard
  const handleCopy = (id, text) => {
    navigator.clipboard.writeText(text);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  // Add term to glossary
  const handleAddGlossaryTerm = () => {
    if (!newTermKey.trim() || !newTermVal.trim()) return;
    setGlossary(prev => ({
      ...prev,
      [newTermKey.trim()]: newTermVal.trim()
    }));
    setNewTermKey('');
    setNewTermVal('');
  };

  const handleRemoveGlossaryTerm = (key) => {
    setGlossary(prev => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  };

  const currentSlide = session?.slides?.find(s => s.slide_index === activeSlideIdx) || session?.slides?.[0];

  // Filter items for current slide
  const filteredItems = currentSlide?.items?.filter(item => {
    const matchSearch = searchTerm === '' || 
      item.original_text.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.translated_text.toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchType = filterType === 'all' || 
      (filterType === 'title' && item.item_type === 'title') ||
      (filterType === 'body' && item.item_type === 'body') ||
      (filterType === 'table' && item.item_type === 'table_cell');

    return matchSearch && matchType;
  }) || [];

  const filteredFonts = Object.values(fontsList).filter(f => 
    !fontSearch || f.name.toLowerCase().includes(fontSearch.toLowerCase())
  );

  return (
    <div className="flex flex-col min-h-screen bg-slate-950 text-slate-100 font-sans selection:bg-brand-500 selection:text-white">
      
      {/* 1. TOP NAVBAR */}
      <header className="sticky top-0 z-40 border-b border-slate-800/80 bg-slate-900/90 backdrop-blur-xl px-6 py-3 flex items-center justify-between shadow-2xl">
        <div className="flex items-center gap-3.5">
          <div className="h-10 w-10 rounded-2xl bg-gradient-to-tr from-brand-600 via-indigo-600 to-cyan-400 flex items-center justify-center shadow-lg shadow-brand-500/25 text-white font-bold ring-2 ring-brand-400/30">
            <Zap className="h-5 w-5 text-white" />
          </div>
          <div>
            <div className="flex items-center gap-2.5">
              <h1 className="font-extrabold text-lg tracking-tight bg-gradient-to-r from-white via-slate-100 to-brand-300 bg-clip-text text-transparent">
                SlideTranslate AI
              </h1>
              <span className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-full bg-gradient-to-r from-brand-500/20 to-indigo-500/20 text-brand-300 border border-brand-500/30">
                PRO Studio
              </span>
            </div>
            <div className="flex items-center gap-2 text-xs text-slate-400">
              <span>PowerPoint Slaydlarini O'zbek Tiliga 100% Formatda O'girish</span>
              <span className="text-slate-600">•</span>
              <span className="text-emerald-400 flex items-center gap-1 text-[11px]">
                <span className="h-1.5 w-1.5 rounded-full bg-emerald-400 animate-pulse" />
                Gemini 3.6 Flash Active
              </span>
            </div>
          </div>
        </div>

        {/* Center Control Group */}
        <div className="flex items-center gap-3">
          {/* Script Switcher */}
          <div className="bg-slate-800/90 border border-slate-700/80 p-1 rounded-xl flex items-center gap-1 shadow-inner">
            <button
              onClick={() => handleToggleScript('latin')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                targetScript === 'latin'
                  ? 'bg-gradient-to-r from-brand-600 to-indigo-600 text-white shadow-md shadow-brand-500/30'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-700/50'
              }`}
            >
              <span>Lotin</span>
            </button>
            <button
              onClick={() => handleToggleScript('cyrillic')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                targetScript === 'cyrillic'
                  ? 'bg-gradient-to-r from-brand-600 to-indigo-600 text-white shadow-md shadow-brand-500/30'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-700/50'
              }`}
            >
              <span>Кирилл</span>
            </button>
          </div>

          {/* Domain Selector */}
          <select
            value={domain}
            onChange={(e) => setDomain(e.target.value)}
            className="bg-slate-800/90 border border-slate-700/80 rounded-xl px-3 py-1.5 text-xs font-semibold text-slate-200 focus:outline-none focus:ring-1 focus:ring-brand-500 cursor-pointer"
          >
            <option value="IT & Dasturlash">💻 IT & Dasturlash</option>
            <option value="Biznes & Moliya">📈 Biznes & Moliya</option>
            <option value="Tibbiyot & Fan">🔬 Tibbiyot & Fan</option>
            <option value="Ta'lim & Dars">🎓 Ta'lim & Dars</option>
            <option value="Marketing & Savdo">🎯 Marketing & Savdo</option>
            <option value="Umumiy">🌐 Umumiy soha</option>
          </select>
        </div>

        {/* Right Action Icons */}
        <div className="flex items-center gap-2">
          {/* Font Vault Button */}
          <button
            onClick={() => setShowFontVaultModal(true)}
            className="p-2 rounded-xl bg-slate-800/80 hover:bg-slate-700 text-slate-300 hover:text-white transition-all text-xs flex items-center gap-1.5 border border-slate-700/60"
            title="Shriftlar ombori (Font Vault)"
          >
            <Type className="h-4 w-4 text-cyan-400" />
            <span className="font-semibold text-[11px] hidden sm:inline">Shriftlar ({Object.keys(fontsList).length})</span>
          </button>

          {/* Glossary Button */}
          <button
            onClick={() => setShowGlossaryModal(true)}
            className="p-2 rounded-xl bg-slate-800/80 hover:bg-slate-700 text-slate-300 hover:text-white transition-all text-xs flex items-center gap-1.5 border border-slate-700/60"
            title="Lug'at (Glossary)"
          >
            <BookOpen className="h-4 w-4 text-amber-400" />
            <span className="font-semibold text-[11px] hidden sm:inline">Lug'at ({Object.keys(glossary).length})</span>
          </button>

          {/* Settings Button */}
          <button
            onClick={() => setShowSettingsModal(true)}
            className="p-2 rounded-xl bg-slate-800/80 hover:bg-slate-700 text-slate-300 hover:text-white transition-all text-xs flex items-center gap-1.5 border border-slate-700/60"
            title="Sozlamalar (API Key)"
          >
            <Settings className="h-4 w-4 text-indigo-400" />
            <span className="font-semibold text-[11px] hidden sm:inline">Sozlamalar</span>
          </button>
        </div>
      </header>

      {/* STATUS BANNER */}
      {statusMsg && (
        <div className="bg-gradient-to-r from-brand-950 via-slate-900 to-indigo-950 border-b border-brand-800/40 px-6 py-2 flex items-center justify-between text-xs text-brand-200">
          <div className="flex items-center gap-2">
            <Sparkles className="h-4 w-4 text-brand-400 animate-spin" />
            <span className="font-medium">{statusMsg}</span>
          </div>
          <button 
            onClick={() => setStatusMsg('')}
            className="text-slate-400 hover:text-white text-xs px-2 py-0.5 rounded hover:bg-slate-800"
          >
            ✕
          </button>
        </div>
      )}

      {/* MAIN VIEW AREA */}
      {!session ? (
        /* 2. HERO & UPLOAD SCREEN */
        <main className="flex-1 flex flex-col items-center justify-center p-6 md:p-12 relative overflow-hidden">
          {/* Subtle Background Glows */}
          <div className="absolute top-1/4 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[700px] h-[350px] bg-gradient-to-tr from-brand-600/15 via-indigo-600/15 to-cyan-500/15 blur-[120px] pointer-events-none -z-10 rounded-full" />

          <div className="max-w-4xl w-full flex flex-col items-center text-center space-y-8">
            
            {/* Hero Title & Badges */}
            <div className="space-y-4">
              <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full bg-brand-500/10 border border-brand-500/30 text-brand-300 text-xs font-semibold shadow-inner">
                <Sparkles className="h-3.5 w-3.5 text-brand-400" />
                <span>AI bilan jihozlangan PowerPoint tarjima studiyasi</span>
              </div>

              <h2 className="text-4xl md:text-5xl font-black tracking-tight leading-tight">
                Taqdimotlarni <span className="bg-gradient-to-r from-brand-400 via-indigo-300 to-cyan-400 bg-clip-text text-transparent">Dizayni va Shriftini Buzmasdan</span> O'zbekchaga O'giring
              </h2>

              <p className="text-sm md:text-base text-slate-400 max-w-2xl mx-auto leading-relaxed">
                Diagrammalar, jadvallar, qulflangan Slide Master reklamalari va murakkab shakllar 100% tahlil qilinadi va professional O'zbek tiliga o'giriladi.
              </p>
            </div>

            {/* Drag & Drop Upload Card */}
            <div
              onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={(e) => {
                e.preventDefault();
                setIsDragging(false);
                if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                  handleFileUpload(e.dataTransfer.files[0]);
                }
              }}
              onClick={() => fileInputRef.current?.click()}
              className={`w-full max-w-2xl p-10 md:p-14 rounded-3xl border-2 border-dashed transition-all cursor-pointer flex flex-col items-center justify-center space-y-4 group relative ${
                isDragging
                  ? 'border-brand-400 bg-brand-500/10 scale-[1.01]'
                  : 'border-slate-700/80 bg-slate-900/60 hover:border-brand-500/80 hover:bg-slate-900/90 shadow-2xl'
              }`}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".pptx"
                className="hidden"
                onChange={(e) => {
                  if (e.target.files && e.target.files[0]) {
                    handleFileUpload(e.target.files[0]);
                  }
                }}
              />

              <div className="h-20 w-20 rounded-3xl bg-gradient-to-tr from-brand-600/30 to-indigo-600/30 border border-brand-500/40 flex items-center justify-center group-hover:scale-110 transition-transform shadow-xl shadow-brand-500/10">
                {loading ? (
                  <RefreshCw className="h-10 w-10 text-brand-400 animate-spin" />
                ) : (
                  <UploadCloud className="h-10 w-10 text-brand-400 group-hover:text-brand-300 transition-colors" />
                )}
              </div>

              <div className="space-y-1 text-center">
                <h3 className="text-lg font-bold text-slate-100 group-hover:text-brand-300 transition-colors">
                  {loading ? "Fayl tahlil qilinmoqda..." : "PowerPoint (.pptx) faylini shu yerga tashlang"}
                </h3>
                <p className="text-xs text-slate-400">
                  yoki kompyuterdan tanlash uchun <span className="text-brand-400 font-semibold underline">bosing</span>
                </p>
              </div>

              <div className="flex items-center gap-2 pt-2">
                <span className="text-[11px] font-mono px-2.5 py-1 rounded-lg bg-slate-800 border border-slate-700 text-slate-300">
                  .PPTX
                </span>
                <span className="text-[11px] font-mono px-2.5 py-1 rounded-lg bg-slate-800 border border-slate-700 text-slate-300">
                  Maksimal 100+ slayd
                </span>
                <span className="text-[11px] font-mono px-2.5 py-1 rounded-lg bg-slate-800 border border-slate-700 text-slate-300">
                  100% Maxfiy & Xavfsiz
                </span>
              </div>
            </div>

            {/* 4 Feature Cards */}
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 w-full text-left pt-6">
              <div className="p-5 rounded-2xl bg-slate-900/60 border border-slate-800/80 space-y-2 hover:border-slate-700 transition-all">
                <div className="h-8 w-8 rounded-xl bg-brand-500/20 text-brand-400 flex items-center justify-center font-bold">
                  <Layout className="h-4 w-4" />
                </div>
                <h4 className="text-xs font-bold text-slate-200">100% Dizayn Saqlanadi</h4>
                <p className="text-[11px] text-slate-400 leading-relaxed">
                  Ikonkalar, doiralar, ranglar va jadval kataklari joylashuvi o'zgarmaydi.
                </p>
              </div>

              <div className="p-5 rounded-2xl bg-slate-900/60 border border-slate-800/80 space-y-2 hover:border-slate-700 transition-all">
                <div className="h-8 w-8 rounded-xl bg-cyan-500/20 text-cyan-400 flex items-center justify-center font-bold">
                  <Type className="h-4 w-4" />
                </div>
                <h4 className="text-xs font-bold text-slate-200">Aqlli Shrift Tizimi</h4>
                <p className="text-[11px] text-slate-400 leading-relaxed">
                  129+ shriftlar bazasi va yangi shriftlarni internetdan avtomatik yuklab olish.
                </p>
              </div>

              <div className="p-5 rounded-2xl bg-slate-900/60 border border-slate-800/80 space-y-2 hover:border-slate-700 transition-all">
                <div className="h-8 w-8 rounded-xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center font-bold">
                  <ShieldCheck className="h-4 w-4" />
                </div>
                <h4 className="text-xs font-bold text-slate-200">Reklama va Watermark Tozalash</h4>
                <p className="text-[11px] text-slate-400 leading-relaxed">
                  Slide Master foni va shablon mualliflarining barcha logotiplarini tozalaydi.
                </p>
              </div>

              <div className="p-5 rounded-2xl bg-slate-900/60 border border-slate-800/80 space-y-2 hover:border-slate-700 transition-all">
                <div className="h-8 w-8 rounded-xl bg-amber-500/20 text-amber-400 flex items-center justify-center font-bold">
                  <Sparkles className="h-4 w-4" />
                </div>
                <h4 className="text-xs font-bold text-slate-200">Anti-Overflow & Auto-Fit</h4>
                <p className="text-[11px] text-slate-400 leading-relaxed">
                  So'zlar qutilardan toshib ketmaydi, qatorlar to'qnashuvi avtomatik oldi olinadi.
                </p>
              </div>
            </div>

          </div>
        </main>
      ) : (
        /* 3. WORKSPACE: EDITOR & LIVE SLIDE STUDIO */
        <div className="flex-1 flex flex-col overflow-hidden">
          
          {/* Workspace Sub-Header Toolbar */}
          <div className="border-b border-slate-800 bg-slate-900/70 px-6 py-2.5 flex items-center justify-between gap-4">
            
            {/* File Info */}
            <div className="flex items-center gap-3 min-w-0">
              <span className="p-1.5 rounded-lg bg-brand-500/20 text-brand-400 border border-brand-500/30">
                <FileText className="h-4 w-4" />
              </span>
              <div className="min-w-0">
                <span className="text-xs font-bold text-slate-200 truncate block max-w-xs md:max-w-md" title={session.filename}>
                  {session.filename}
                </span>
                <span className="text-[10px] text-slate-500 font-medium">
                  {session.slides_count} ta slayd • {session.total_items} ta matn bloki
                </span>
              </div>
            </div>

            {/* Quick Settings & Mode Buttons */}
            <div className="flex items-center gap-3">
              
              {/* Auto-Fit Toggle */}
              <label className="flex items-center gap-2 text-xs text-slate-300 cursor-pointer bg-slate-800/80 px-3 py-1.5 rounded-xl border border-slate-700 hover:bg-slate-700/60 transition-all">
                <input 
                  type="checkbox" 
                  checked={autoFit} 
                  onChange={(e) => setAutoFit(e.target.checked)}
                  className="rounded border-slate-700 text-brand-600 focus:ring-0 cursor-pointer"
                />
                <span className="font-semibold text-[11px]">⚡ Auto-fit (Shrift moslash)</span>
              </label>

              {/* View Mode Toggle */}
              <div className="bg-slate-800 p-1 rounded-xl flex items-center gap-1 border border-slate-700 text-xs">
                <button
                  onClick={() => setViewMode('split')}
                  className={`px-2.5 py-1 rounded-lg transition-all flex items-center gap-1 ${
                    viewMode === 'split' ? 'bg-slate-700 text-white font-bold' : 'text-slate-400'
                  }`}
                  title="Yonma-yon ko'rish"
                >
                  <Columns className="h-3.5 w-3.5" />
                  <span className="hidden sm:inline">Split</span>
                </button>
                <button
                  onClick={() => setViewMode('translated')}
                  className={`px-2.5 py-1 rounded-lg transition-all flex items-center gap-1 ${
                    viewMode === 'translated' ? 'bg-slate-700 text-white font-bold' : 'text-slate-400'
                  }`}
                  title="Faqat O'zbekcha"
                >
                  <Eye className="h-3.5 w-3.5" />
                  <span className="hidden sm:inline">O'zbek</span>
                </button>
              </div>
            </div>

            {/* Right Action Buttons */}
            <div className="flex items-center gap-2.5">
              {/* Translate All Button */}
              <button
                onClick={handleTranslateAll}
                disabled={translating}
                className="px-4 py-2 rounded-xl bg-gradient-to-r from-brand-600 via-indigo-600 to-cyan-500 hover:from-brand-500 hover:to-indigo-500 text-white text-xs font-bold transition-all shadow-md shadow-brand-500/20 flex items-center gap-2 disabled:opacity-50"
              >
                {translating ? (
                  <>
                    <RefreshCw className="h-4 w-4 animate-spin" />
                    <span>AI Tarjima qilinmoqda...</span>
                  </>
                ) : (
                  <>
                    <Sparkles className="h-4 w-4" />
                    <span>✨ Gemini AI Bilan Tarjima Qilish</span>
                  </>
                )}
              </button>

              {/* Export Button */}
              <button
                onClick={handleExport}
                disabled={exporting}
                className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold transition-all shadow-md shadow-emerald-500/20 flex items-center gap-2 disabled:opacity-50"
              >
                {exporting ? (
                  <>
                    <RefreshCw className="h-4 w-4 animate-spin" />
                    <span>Tayyorlanmoqda...</span>
                  </>
                ) : (
                  <>
                    <Download className="h-4 w-4" />
                    <span>📥 PPTX Yuklab Olish</span>
                  </>
                )}
              </button>

              {/* Reset / New File */}
              <button
                onClick={() => setSession(null)}
                className="p-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white transition-all text-xs border border-slate-700"
                title="Yangi fayl yuklash"
              >
                <RefreshCw className="h-4 w-4" />
              </button>
            </div>

          </div>

          {/* Progress Bar (if translating) */}
          {translating && (
            <div className="w-full bg-slate-800 h-1.5 overflow-hidden">
              <div 
                className="bg-gradient-to-r from-brand-500 via-indigo-400 to-emerald-400 h-full transition-all duration-300"
                style={{ width: `${progress}%` }}
              />
            </div>
          )}

          {/* Three-Column Workspace Body */}
          <div className="flex-1 flex overflow-hidden">
            
            {/* 1. LEFT SIDEBAR: Slide Navigator */}
            <aside className="w-64 border-r border-slate-800 bg-slate-950/70 flex flex-col overflow-y-auto p-4 space-y-2.5">
              <div className="flex items-center justify-between px-2 text-xs font-bold text-slate-400 uppercase tracking-wider">
                <span>Slaydlar ({session.slides_count})</span>
                <Layers className="h-3.5 w-3.5" />
              </div>

              {session.slides.map((slide) => {
                const isActive = slide.slide_index === activeSlideIdx;
                const firstTitle = slide.items?.find(i => i.item_type === 'title')?.translated_text || 
                                   slide.items?.find(i => i.item_type === 'title')?.original_text || 
                                   slide.items?.[0]?.translated_text || 
                                   `Slayd ${slide.slide_index}`;

                return (
                  <button
                    key={slide.slide_index}
                    onClick={() => setActiveSlideIdx(slide.slide_index)}
                    className={`w-full text-left p-3 rounded-2xl border transition-all flex flex-col gap-1.5 relative group ${
                      isActive 
                        ? 'bg-slate-800/95 border-brand-500/80 shadow-lg shadow-brand-500/10 ring-1 ring-brand-500/30' 
                        : 'bg-slate-900/40 border-slate-800/80 hover:bg-slate-900 hover:border-slate-700'
                    }`}
                  >
                    <div className="flex items-center justify-between">
                      <span className={`text-[10px] font-bold px-2 py-0.5 rounded-md ${
                        isActive ? 'bg-brand-500 text-white' : 'bg-slate-800 text-slate-400'
                      }`}>
                        Slayd {slide.slide_index}
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium">
                        {slide.items_count} ta matn
                      </span>
                    </div>

                    <p className={`text-xs font-semibold line-clamp-2 ${isActive ? 'text-white' : 'text-slate-300'}`}>
                      {firstTitle}
                    </p>
                  </button>
                );
              })}
            </aside>

            {/* 2. CENTER PANEL: Visual Slide Canvas Preview */}
            <div className="flex-1 bg-slate-950 p-4 md:p-8 flex flex-col items-center justify-start overflow-y-auto min-h-0">
              
              <div className="w-full max-w-5xl flex flex-col gap-4">
                
                {/* Slide Top Navigation & Meta Bar */}
                <div className="flex items-center justify-between bg-slate-900/90 border border-slate-800 px-4 py-2.5 rounded-2xl shadow-md">
                  <div className="flex items-center gap-3">
                    <span className="h-2.5 w-2.5 rounded-full bg-emerald-400 animate-pulse" />
                    <span className="text-xs font-bold text-white uppercase tracking-wider">
                      Slayd {currentSlide?.slide_index} / {session.slides_count}
                    </span>
                    <span className="text-[11px] text-slate-400 font-medium">
                      ({currentSlide?.items?.length || 0} ta matn bloki)
                    </span>
                  </div>

                  {/* Slide Stepper Controls */}
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => setActiveSlideIdx(Math.max(1, activeSlideIdx - 1))}
                      disabled={activeSlideIdx <= 1}
                      className="px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 disabled:opacity-30 transition-all flex items-center gap-1 border border-slate-700"
                    >
                      <ChevronLeft className="h-4 w-4" /> Oldingi
                    </button>

                    <div className="flex items-center gap-1 overflow-x-auto max-w-[240px] px-1">
                      {session.slides.map(s => (
                        <button
                          key={s.slide_index}
                          onClick={() => setActiveSlideIdx(s.slide_index)}
                          className={`h-7 w-7 rounded-lg text-xs font-bold transition-all flex-shrink-0 ${
                            s.slide_index === activeSlideIdx
                              ? 'bg-brand-600 text-white shadow-md shadow-brand-500/40 border border-brand-400'
                              : 'bg-slate-800/80 text-slate-400 hover:bg-slate-700 hover:text-white border border-slate-700/60'
                          }`}
                        >
                          {s.slide_index}
                        </button>
                      ))}
                    </div>

                    <button
                      onClick={() => setActiveSlideIdx(Math.min(session.slides_count, activeSlideIdx + 1))}
                      disabled={activeSlideIdx >= session.slides_count}
                      className="px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 disabled:opacity-30 transition-all flex items-center gap-1 border border-slate-700"
                    >
                      Keyingi <ChevronRight className="h-4 w-4" />
                    </button>
                  </div>
                </div>

                {/* The Aspect-Ratio Slide Card */}
                <div className="relative w-full aspect-video bg-white text-slate-900 rounded-3xl border-4 border-slate-700/80 shadow-2xl overflow-hidden flex flex-col justify-between p-6 md:p-10 transition-all">
                  
                  {/* Decorative slide header banner */}
                  <div className="absolute top-0 left-0 right-0 h-2 bg-gradient-to-r from-brand-600 via-indigo-600 to-cyan-400" />
                  
                  {/* Slide Content Area */}
                  <div className="flex-1 overflow-y-auto space-y-4 pr-1 mt-1">
                    
                    {/* View Mode: Split Comparison vs Full Translated */}
                    {viewMode === 'split' ? (
                      <div className="grid grid-cols-2 gap-4 h-full">
                        
                        {/* Original Slide Column */}
                        <div className="p-4 bg-slate-50 rounded-2xl border border-slate-200 overflow-y-auto space-y-3">
                          <span className="text-[10px] font-bold uppercase tracking-wider text-slate-500 bg-slate-200/80 px-2 py-0.5 rounded-md">
                            Asl Matn (Original)
                          </span>
                          <div className="space-y-2.5">
                            {currentSlide?.items?.map(item => (
                              <div key={item.id} className="p-2.5 bg-white rounded-xl border border-slate-200 text-xs text-slate-800 leading-relaxed shadow-sm">
                                <span className="text-[9px] text-slate-400 font-mono block mb-0.5">{item.shape_name}</span>
                                {item.original_text}
                              </div>
                            ))}
                          </div>
                        </div>

                        {/* Translated Slide Column */}
                        <div className="p-4 bg-brand-50/40 rounded-2xl border border-brand-200/80 overflow-y-auto space-y-3">
                          <span className="text-[10px] font-bold uppercase tracking-wider text-brand-700 bg-brand-100 px-2 py-0.5 rounded-md">
                            O'zbekcha Tarjima
                          </span>
                          <div className="space-y-2.5">
                            {currentSlide?.items?.map(item => (
                              <div 
                                key={item.id} 
                                className="p-2.5 bg-white rounded-xl border border-brand-200 text-xs text-slate-900 leading-relaxed shadow-sm"
                                style={{ fontFamily: item.font_name || 'inherit' }}
                              >
                                <span className="text-[9px] text-brand-600 font-semibold block mb-0.5">{item.shape_name}</span>
                                {item.translated_text || item.original_text}
                              </div>
                            ))}
                          </div>
                        </div>

                      </div>
                    ) : (
                      /* Full Translated Preview */
                      <div className="space-y-4">
                        
                        {/* Slide Title */}
                        {(() => {
                          const titleItem = currentSlide?.items?.find(i => i.item_type === 'title') || currentSlide?.items?.[0];
                          if (!titleItem) return null;
                          return (
                            <div className="border-b border-slate-200 pb-3">
                              <h2 
                                className="text-xl md:text-2xl font-black text-slate-900 tracking-tight"
                                style={{ fontFamily: titleItem.font_name || 'inherit' }}
                              >
                                {titleItem.translated_text || titleItem.original_text}
                              </h2>
                            </div>
                          );
                        })()}

                        {/* Slide Secondary / Body Items */}
                        {(() => {
                          const items = currentSlide?.items || [];
                          const nonTable = items.filter(i => i.item_type !== 'table_cell');
                          if (nonTable.length <= 1) return null;
                          const rest = nonTable.slice(1);
                          return (
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                              {rest.map(item => (
                                <div 
                                  key={item.id} 
                                  className="p-3.5 bg-slate-50 rounded-2xl border border-slate-200 text-xs md:text-sm text-slate-800 leading-relaxed shadow-sm"
                                  style={{ fontFamily: item.font_name || 'inherit' }}
                                >
                                  <span className="text-[9px] text-slate-400 font-mono block mb-1">{item.shape_name}</span>
                                  {item.translated_text || item.original_text}
                                </div>
                              ))}
                            </div>
                          );
                        })()}

                        {/* Table Matrix Preview (Clean Grid) */}
                        {currentSlide?.items?.some(i => i.item_type === 'table_cell') && (
                          <div className="pt-2">
                            <div className="border border-slate-300 rounded-2xl overflow-hidden shadow-sm">
                              <div className="bg-slate-100 px-3 py-1.5 border-b border-slate-300 text-[11px] font-bold text-slate-700 flex items-center justify-between">
                                <span>📊 Jadval Elementlari</span>
                                <span className="text-[10px] text-slate-500">{currentSlide.items.filter(i => i.item_type === 'table_cell').length} ta katak</span>
                              </div>
                              <div className="grid grid-cols-2 md:grid-cols-3 gap-1.5 p-2.5 bg-slate-50 max-h-56 overflow-y-auto">
                                {currentSlide.items.filter(i => i.item_type === 'table_cell').map(item => (
                                  <div key={item.id} className="p-2 bg-white rounded-lg border border-slate-200 text-xs text-slate-900 shadow-sm">
                                    <span className="text-[9px] text-indigo-600 font-semibold block mb-0.5">{item.shape_name}</span>
                                    <span className="font-medium">{item.translated_text || item.original_text}</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          </div>
                        )}

                      </div>
                    )}

                  </div>

                  {/* Slide Footer */}
                  <div className="pt-3 mt-2 border-t border-slate-200 flex items-center justify-between text-[11px] text-slate-500 font-medium">
                    <span className="flex items-center gap-1.5 text-brand-700 font-bold">
                      <Sparkles className="h-3.5 w-3.5" />
                      SlideTranslate AI — {targetScript === 'latin' ? "O'zbekcha (Lotin)" : "Ўзбекча (Кирилл)"}
                    </span>
                    <span className="bg-slate-100 text-slate-700 px-3 py-1 rounded-full border border-slate-300 font-bold">
                      {currentSlide?.slide_index} / {session.slides_count}
                    </span>
                  </div>

                </div>

              </div>

            </div>

            {/* 3. RIGHT PANEL: Inline Text Editor & Translation Cards */}
            <div className="w-[460px] border-l border-slate-800 bg-slate-950/90 flex flex-col overflow-hidden shadow-2xl">
              
              {/* Search & Filter Header */}
              <div className="p-4 border-b border-slate-800 space-y-3 bg-slate-900/50">
                <div className="flex items-center justify-between">
                  <h3 className="font-bold text-xs uppercase tracking-wider text-slate-300 flex items-center gap-2">
                    <span>Matnlarni Tahrirlash</span>
                    <span className="px-2 py-0.5 rounded-full bg-brand-500/20 text-brand-400 text-[10px] font-bold">
                      {filteredItems.length}
                    </span>
                  </h3>
                  <span className="text-[10px] text-emerald-400 font-semibold flex items-center gap-1">
                    <span className="h-1.5 w-1.5 rounded-full bg-emerald-400" />
                    Avto-saqlash
                  </span>
                </div>

                {/* Search Bar */}
                <div className="relative">
                  <Search className="h-3.5 w-3.5 absolute left-3 top-3 text-slate-500" />
                  <input
                    type="text"
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    placeholder="Slayd matnlarini izlash..."
                    className="w-full bg-slate-900 border border-slate-700/80 rounded-xl pl-9 pr-3 py-2 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:ring-1 focus:ring-brand-500"
                  />
                </div>

                {/* Type Filter Pills */}
                <div className="flex items-center gap-1.5 text-[10px]">
                  {['all', 'title', 'body', 'table'].map(type => (
                    <button
                      key={type}
                      onClick={() => setFilterType(type)}
                      className={`px-3 py-1 rounded-lg font-bold transition-all ${
                        filterType === type 
                          ? 'bg-brand-600 text-white shadow-md shadow-brand-500/20' 
                          : 'bg-slate-800/80 text-slate-400 hover:text-slate-200 hover:bg-slate-700'
                      }`}
                    >
                      {type === 'all' && 'Barchasi'}
                      {type === 'title' && 'Sarlavhalar'}
                      {type === 'body' && 'Matnlar'}
                      {type === 'table' && 'Jadvallar'}
                    </button>
                  ))}
                </div>
              </div>

              {/* Editable Cards List */}
              <div className="flex-1 overflow-y-auto p-4 space-y-4">
                {filteredItems.length === 0 ? (
                  <div className="text-center py-12 text-slate-500 text-xs">
                    Mos keluvchi matn bloklari topilmadi.
                  </div>
                ) : (
                  filteredItems.map((item) => (
                    <div 
                      key={item.id}
                      className="p-4 rounded-2xl bg-slate-900/90 border border-slate-800 hover:border-slate-700 transition-all space-y-3 shadow-lg"
                    >
                      {/* Item Header with Font Badge */}
                      <div className="flex items-center justify-between text-[11px]">
                        <span className="font-semibold text-brand-400 flex items-center gap-1.5 truncate max-w-[180px]">
                          <span className="h-2 w-2 rounded-full bg-brand-400 inline-block flex-shrink-0" />
                          <span className="truncate">{item.shape_name}</span>
                        </span>
                        
                        <div className="flex items-center gap-2">
                          <div className="flex items-center gap-1.5 bg-slate-800/90 px-2.5 py-0.5 rounded-lg border border-slate-700/70 text-[10px] text-cyan-300 font-medium shadow-inner">
                            <span className="truncate max-w-[100px]" title={item.font_name || 'Calibri'}>
                              {item.font_name || 'Calibri'}
                            </span>
                            <span className="text-slate-500">•</span>
                            <span>{item.font_size_pt}pt</span>
                            {item.is_bold && <span className="font-bold text-amber-400">B</span>}
                            {item.is_italic && <span className="italic text-indigo-400">I</span>}
                          </div>
                          <button
                            onClick={() => handleCopy(item.id, item.translated_text || item.original_text)}
                            className="text-slate-400 hover:text-white transition-colors p-1 hover:bg-slate-800 rounded-md"
                            title="Nusxa olish"
                          >
                            {copiedId === item.id ? <Check className="h-3.5 w-3.5 text-emerald-400" /> : <Copy className="h-3.5 w-3.5" />}
                          </button>
                        </div>
                      </div>

                      {/* Original Box */}
                      <div className="p-2.5 rounded-xl bg-slate-950/70 border border-slate-800/80 text-xs text-slate-400 leading-relaxed font-mono select-all">
                        <span className="text-[9px] uppercase font-bold text-slate-500 block mb-0.5">Asl matn:</span>
                        {item.original_text}
                      </div>

                      {/* Translated Editable Box */}
                      <div className="space-y-1">
                        <span className="text-[9px] uppercase font-bold text-emerald-400 flex items-center gap-1">
                          <span>O'zbekcha Tarjima:</span>
                        </span>
                        <textarea
                          rows={3}
                          value={item.translated_text || ''}
                          onChange={(e) => handleItemTextChange(item.id, e.target.value)}
                          placeholder="Tarjimani kiriting..."
                          className="w-full bg-slate-950 border border-slate-700/90 focus:border-brand-500 rounded-xl p-3 text-xs text-slate-100 placeholder-slate-600 focus:outline-none focus:ring-1 focus:ring-brand-500 leading-relaxed resize-y font-sans shadow-inner"
                          style={{ fontFamily: item.font_name || 'inherit' }}
                        />
                      </div>

                      {/* Item Quick Actions */}
                      <div className="flex items-center justify-between text-[11px] pt-1 text-slate-400">
                        <button
                          onClick={() => handleItemTextChange(item.id, item.original_text)}
                          className="hover:text-amber-400 transition-colors text-[10px]"
                        >
                          ↺ Asliga qaytarish
                        </button>
                        
                        <button
                          onClick={() => {
                            const toggled = transliterateUzbek(item.translated_text, targetScript === 'latin');
                            handleItemTextChange(item.id, toggled);
                          }}
                          className="hover:text-brand-400 transition-colors text-[10px]"
                        >
                          🔀 Lotin ↔ Кирилл
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>

            </div>

          </div>

        </div>
      )}

      {/* 4. FONT VAULT MODAL */}
      {showFontVaultModal && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-md flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-2xl w-full p-6 space-y-5 shadow-2xl">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="h-10 w-10 rounded-2xl bg-cyan-500/20 text-cyan-400 flex items-center justify-center font-bold">
                  <Type className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="font-extrabold text-base text-white">Shriftlar Ombori (Font Vault)</h3>
                  <p className="text-xs text-slate-400">Bazada {Object.keys(fontsList).length} ta shrift faol. Yangi shriftlar avtomatik yuklanadi.</p>
                </div>
              </div>
              <button 
                onClick={() => setShowFontVaultModal(false)}
                className="p-1.5 rounded-xl hover:bg-slate-800 text-slate-400 hover:text-white transition-colors"
              >
                ✕
              </button>
            </div>

            {/* Font Search Bar */}
            <div className="relative">
              <Search className="h-4 w-4 absolute left-3 top-3 text-slate-500" />
              <input
                type="text"
                value={fontSearch}
                onChange={(e) => setFontSearch(e.target.value)}
                placeholder="Shrift nomini izlash (masalan: Montserrat, Roboto, Poppins)..."
                className="w-full bg-slate-950 border border-slate-700 rounded-xl pl-9 pr-3 py-2 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:ring-1 focus:ring-cyan-500"
              />
            </div>

            {/* Fonts Grid */}
            <div className="max-h-72 overflow-y-auto grid grid-cols-1 md:grid-cols-2 gap-2 p-1">
              {filteredFonts.map((f, i) => (
                <div key={i} className="p-3 rounded-xl bg-slate-950/70 border border-slate-800/80 flex items-center justify-between">
                  <div>
                    <span className="text-xs font-bold text-slate-200 block" style={{ fontFamily: f.name }}>
                      {f.name}
                    </span>
                    <span className="text-[10px] text-slate-500">
                      {f.type === 'custom' ? `Cached TrueType (${f.size_kb || 'TTF'} KB)` : 'Office / Windows System Font'}
                    </span>
                  </div>
                  <span className="text-[9px] px-2 py-0.5 rounded-md bg-cyan-500/10 text-cyan-400 border border-cyan-500/20 font-mono">
                    Ready
                  </span>
                </div>
              ))}
            </div>

            <div className="pt-2 border-t border-slate-800 flex justify-end">
              <button
                onClick={() => setShowFontVaultModal(false)}
                className="px-5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-bold text-white transition-colors"
              >
                Yopish
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 5. GLOSSARY MODAL */}
      {showGlossaryModal && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-md flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-xl w-full p-6 space-y-5 shadow-2xl">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="h-10 w-10 rounded-2xl bg-amber-500/20 text-amber-400 flex items-center justify-center font-bold">
                  <BookOpen className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="font-extrabold text-base text-white">Maxsus Lug'at (Glossary)</h3>
                  <p className="text-xs text-slate-400">Atamalarning aniq tarjimasini belgilab qo'ying</p>
                </div>
              </div>
              <button 
                onClick={() => setShowGlossaryModal(false)}
                className="p-1.5 rounded-xl hover:bg-slate-800 text-slate-400 hover:text-white transition-colors"
              >
                ✕
              </button>
            </div>

            {/* Add New Term Inputs */}
            <div className="grid grid-cols-2 gap-2 pt-1">
              <input
                type="text"
                value={newTermKey}
                onChange={(e) => setNewTermKey(e.target.value)}
                placeholder="Inglizcha / Xitoycha atama..."
                className="bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:ring-1 focus:ring-amber-500"
              />
              <div className="flex gap-2">
                <input
                  type="text"
                  value={newTermVal}
                  onChange={(e) => setNewTermVal(e.target.value)}
                  placeholder="O'zbekcha tarjimasi..."
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:ring-1 focus:ring-amber-500"
                />
                <button
                  onClick={handleAddGlossaryTerm}
                  className="px-3.5 py-2 rounded-xl bg-amber-600 hover:bg-amber-500 text-white font-bold text-xs transition-colors flex items-center justify-center"
                >
                  <Plus className="h-4 w-4" />
                </button>
              </div>
            </div>

            {/* Terms List */}
            <div className="max-h-60 overflow-y-auto space-y-2 pr-1">
              {Object.entries(glossary).map(([k, v]) => (
                <div key={k} className="flex items-center justify-between p-2.5 rounded-xl bg-slate-950/70 border border-slate-800 text-xs">
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-slate-200">{k}</span>
                    <ArrowRight className="h-3 w-3 text-slate-600" />
                    <span className="text-amber-400 font-bold">{v}</span>
                  </div>
                  <button
                    onClick={() => handleRemoveGlossaryTerm(k)}
                    className="text-slate-500 hover:text-rose-400 p-1 transition-colors"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                  </button>
                </div>
              ))}
            </div>

            <div className="pt-2 border-t border-slate-800 flex justify-end">
              <button
                onClick={() => setShowGlossaryModal(false)}
                className="px-5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-bold text-white transition-colors"
              >
                Saqlash va Yopish
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 6. SETTINGS MODAL */}
      {showSettingsModal && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-md flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-lg w-full p-6 space-y-5 shadow-2xl">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="h-10 w-10 rounded-2xl bg-indigo-500/20 text-indigo-400 flex items-center justify-center font-bold">
                  <Settings className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="font-extrabold text-base text-white">AI Sozlamalari</h3>
                  <p className="text-xs text-slate-400">Gemini API kaliti va parametrlarini boshqarish</p>
                </div>
              </div>
              <button 
                onClick={() => setShowSettingsModal(false)}
                className="p-1.5 rounded-xl hover:bg-slate-800 text-slate-400 hover:text-white transition-colors"
              >
                ✕
              </button>
            </div>

            <div className="space-y-4 text-xs">
              <div className="space-y-1.5">
                <label className="font-bold text-slate-300 flex items-center gap-1.5">
                  <Key className="h-3.5 w-3.5 text-indigo-400" />
                  Google Gemini API Kaliti
                </label>
                <input
                  type="password"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder="AIzaSy..."
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl p-3 text-xs text-slate-100 placeholder-slate-600 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono"
                />
                <p className="text-[10px] text-slate-500">
                  Kalit faqat sizning brauzeringizda xavfsiz saqlanadi. Agar kiritilmasa, tizimdagi standart server kaliti ishlatiladi.
                </p>
              </div>

              <div className="p-3.5 rounded-2xl bg-slate-950/80 border border-slate-800 space-y-1.5 text-[11px] text-slate-400">
                <div className="flex justify-between">
                  <span className="font-medium">Faol AI modeli:</span>
                  <span className="font-bold text-indigo-300">Gemini 3.6 Flash / 3.5 Flash</span>
                </div>
                <div className="flex justify-between">
                  <span className="font-medium">Tarjima tezligi:</span>
                  <span className="font-bold text-emerald-400">Parallel Multi-Threaded</span>
                </div>
              </div>
            </div>

            <div className="pt-2 border-t border-slate-800 flex justify-end">
              <button
                onClick={() => setShowSettingsModal(false)}
                className="px-5 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-bold text-white transition-colors shadow-md shadow-indigo-500/20"
              >
                Tayyor
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
