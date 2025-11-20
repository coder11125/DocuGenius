import React, { useState, useRef, useEffect } from 'react';
import { 
  Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight, 
  List, ListOrdered, Type, Highlighter, Sparkles, X, 
  MessageSquarePlus, FileText, Download, Share2, 
  Undo, Redo, Save, Check, Minus, Plus, ZoomIn,
  Image as ImageIcon, Link as LinkIcon, Table, MinusSquare, Calendar,
  Printer, FilePlus, FolderOpen, Layout, Settings, Eye, Grid,
  Maximize, Minimize, Globe, Book, Quote, CheckSquare,
  Palette, Stamp, Mail, Users, Play, MousePointerClick,
  ChevronDown, Search, Menu, ChevronUp, FileType, File, Cloud, 
  MoreVertical, Trash2, Edit2, ArrowLeft, PlusSquare
} from 'lucide-react';

// --- Configuration ---
const API_KEY = "AIzaSyAjxNqUdB_8i3SHWtvX5l5ukaA8eFNlk30"; // Injected by environment
const MODEL_NAME = "gemini-2.5-flash-preview-09-2025";

// --- Helper: Retry Logic for API ---
const fetchWithRetry = async (url, options, retries = 3, delay = 1000) => {
  for (let i = 0; i < retries; i++) {
    try {
      const response = await fetch(url, options);
      if (!response.ok) {
        if (response.status === 429) throw new Error("Rate limited");
        throw new Error(`HTTP Error: ${response.status}`);
      }
      return await response.json();
    } catch (err) {
      if (i === retries - 1) throw err;
      await new Promise(resolve => setTimeout(resolve, delay * Math.pow(2, i)));
    }
  }
};

// --- Helper: Load Script Dynamically ---
const loadScript = (src) => {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) {
      resolve();
      return;
    }
    const script = document.createElement('script');
    script.src = src;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
};

// --- Helper: Load Mammoth.js for DOCX Import ---
const loadMammoth = () => {
  return new Promise((resolve, reject) => {
    if (window.mammoth) {
      resolve(window.mammoth);
      return;
    }
    loadScript("https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js")
      .then(() => resolve(window.mammoth))
      .catch(() => reject(new Error("Failed to load DOCX parser")));
  });
};

// --- Helper: Basic RTF Parser ---
const parseRTF = (rtf) => {
  let text = rtf.replace(/{\\fonttbl.*?}/g, '');
  text = text.replace(/{\\colortbl.*?}/g, '');
  text = text.replace(/{\\stylesheet.*?}/g, '');
  text = text.replace(/{\\u[0-9]+.*?}/g, '');
  text = text.replace(/\\par[d]?/g, '<br/>');
  text = text.replace(/\\b\s/g, '<b>').replace(/\\b0\s?/g, '</b>');
  text = text.replace(/\\i\s/g, '<i>').replace(/\\i0\s?/g, '</i>');
  text = text.replace(/\\ul\s/g, '<u>').replace(/\\ulnone\s?/g, '</u>');
  text = text.replace(/\\[a-z]+\d*[\s]?/g, ''); 
  text = text.replace(/[{}]/g, '');
  return text;
};

// ==========================================
// DASHBOARD COMPONENT
// ==========================================
const Dashboard = ({ documents, onCreate, onOpen, onDelete }) => {
  return (
    <div className="min-h-screen bg-gray-50 font-sans text-[#323130]">
      {/* Header */}
      <div className="h-16 bg-white border-b border-gray-200 flex items-center px-4 justify-between sticky top-0 z-10">
        <div className="flex items-center gap-3">
           <div className="p-2 bg-[#0078d4] rounded-lg text-white">
             <FileText size={20} />
           </div>
           <span className="text-xl font-normal text-gray-600">DocuGenius</span>
        </div>
        
        <div className="flex-1 max-w-2xl mx-4 hidden md:block">
           <div className="relative group">
              <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                  <Search size={16} className="text-gray-400 group-focus-within:text-[#0078d4]" />
              </div>
              <input 
                  type="text" 
                  className="block w-full pl-10 pr-3 py-2.5 border-none rounded-lg leading-5 bg-gray-100 text-gray-900 placeholder-gray-500 focus:outline-none focus:bg-white focus:ring-1 focus:ring-gray-200 focus:shadow-md transition-all"
                  placeholder="Search" 
              />
           </div>
        </div>

        <div className="w-10 h-10 rounded-full bg-[#0078d4] text-white flex items-center justify-center font-bold text-sm">JD</div>
      </div>

      {/* Start New Section */}
      <div className="bg-gray-100 py-8 border-b border-gray-200">
         <div className="max-w-5xl mx-auto px-4">
            <div className="flex items-center justify-between mb-4">
               <span className="text-sm font-medium text-gray-600">Start a new document</span>
            </div>
            <div className="flex gap-4">
               <div className="group cursor-pointer" onClick={onCreate}>
                  <div className="w-36 h-48 bg-white border border-gray-200 hover:border-[#0078d4] rounded-lg flex items-center justify-center shadow-sm hover:shadow-md transition-all relative overflow-hidden">
                      <div className="absolute inset-0 flex items-center justify-center">
                          <Plus size={48} className="text-[#0078d4] opacity-20 group-hover:opacity-100 transition-opacity" />
                      </div>
                      <img src="https://ssl.gstatic.com/docs/templates/thumbnails/docs-blank-googlecolors.png" className="w-full h-full object-cover opacity-0" alt="Blank" />
                  </div>
                  <p className="mt-2 text-sm font-medium text-gray-700 group-hover:text-[#0078d4]">Blank document</p>
               </div>
            </div>
         </div>
      </div>

      {/* Recent Documents Section */}
      <div className="max-w-5xl mx-auto px-4 py-8">
         <div className="flex items-center justify-between mb-4 pb-2 border-b border-gray-200">
            <span className="text-sm font-semibold text-gray-700">Recent documents</span>
            <div className="flex items-center gap-4 text-gray-500">
               <button className="p-1 hover:bg-gray-200 rounded"><ListOrdered size={18}/></button>
               <button className="p-1 hover:bg-gray-200 rounded"><ArrowLeft size={18} className="rotate-90"/></button>
            </div>
         </div>

         {documents.length === 0 ? (
             <div className="text-center py-20 text-gray-400">
                 <FileText size={48} className="mx-auto mb-4 opacity-20" />
                 <p>No documents yet. Start a new one above!</p>
             </div>
         ) : (
             <div className="grid grid-cols-1 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-4">
                {documents.map(doc => (
                    <div key={doc.id} className="group cursor-pointer" onClick={() => onOpen(doc.id)}>
                        <div className="aspect-[3/4] bg-white border border-gray-200 rounded-sm hover:border-[#0078d4] relative shadow-sm hover:shadow-md transition-all p-4 overflow-hidden">
                            <div className="text-[4px] text-gray-300 leading-relaxed select-none overflow-hidden h-full" dangerouslySetInnerHTML={{__html: doc.content}}></div>
                            <div className="absolute inset-x-0 bottom-0 h-20 bg-gradient-to-t from-white to-transparent"></div>
                        </div>
                        <div className="mt-3 px-1">
                            <h3 className="text-sm font-medium text-gray-800 truncate">{doc.name || 'Untitled Document'}</h3>
                            <div className="flex items-center justify-between mt-1">
                                <div className="flex items-center gap-1">
                                    <div className="w-4 h-4 bg-[#0078d4] rounded-sm flex items-center justify-center text-white text-[8px]"><FileText size={10}/></div>
                                    <span className="text-xs text-gray-500">Opened {new Date(doc.lastModified).toLocaleDateString()}</span>
                                </div>
                                <button 
                                    onClick={(e) => { e.stopPropagation(); onDelete(doc.id); }}
                                    className="p-1 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded opacity-0 group-hover:opacity-100 transition-opacity"
                                >
                                    <Trash2 size={14} />
                                </button>
                            </div>
                        </div>
                    </div>
                ))}
             </div>
         )}
      </div>
    </div>
  );
};

// ==========================================
// MAIN APP & EDITOR
// ==========================================
const App = () => {
  // --- Global State ---
  const [view, setView] = useState('dashboard'); // 'dashboard' | 'editor'
  const [documents, setDocuments] = useState([]);
  const [currentDocId, setCurrentDocId] = useState(null);

  // --- Editor State ---
  const [activeTab, setActiveTab] = useState('Home');
  const [fileName, setFileName] = useState('Untitled Document');
  const [language, setLanguage] = useState('English (U.S.)');
  const [showAIPanel, setShowAIPanel] = useState(false);
  const [aiPrompt, setAiPrompt] = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [aiResponse, setAiResponse] = useState(null);
  const [selectionRange, setSelectionRange] = useState(null);
  const [saveStatus, setSaveStatus] = useState('Saved');
  const [selectedText, setSelectedText] = useState('');
  const [zoomLevel, setZoomLevel] = useState(100);
  const [totalWords, setTotalWords] = useState(0);
  const [isRibbonOpen, setIsRibbonOpen] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  
  // Layout States
  const [orientation, setOrientation] = useState('portrait');
  const [margins, setMargins] = useState('normal');
  const [viewMode, setViewMode] = useState('print');
  const [pageColor, setPageColor] = useState('#ffffff');
  const [watermark, setWatermark] = useState(null);

  const editorRef = useRef(null);
  const fileInputRef = useRef(null);
  const imageInputRef = useRef(null);

  // --- Initialization & Data Migration ---
  useEffect(() => {
    const rawDocs = localStorage.getItem('docuGeniusDocs');
    const legacyData = localStorage.getItem('docuGeniusData');
    
    let docs = [];
    if (rawDocs) {
        try {
            docs = JSON.parse(rawDocs);
        } catch (e) { console.error("Corrupt storage", e); }
    } 
    
    // Migration Logic: If old format exists but no new format, migrate it
    if (legacyData && docs.length === 0) {
        try {
            const old = JSON.parse(legacyData);
            const migratedDoc = {
                id: crypto.randomUUID(),
                name: old.name || 'Migrated Document',
                content: old.content,
                lastModified: old.lastModified || new Date().toISOString()
            };
            docs.push(migratedDoc);
            localStorage.setItem('docuGeniusDocs', JSON.stringify(docs));
            localStorage.removeItem('docuGeniusData'); // Cleanup
        } catch (e) {}
    }
    
    setDocuments(docs.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified)));
  }, []);

  // --- Dashboard Actions ---
  const createNewDocument = () => {
      const newDoc = {
          id: crypto.randomUUID(),
          name: 'Untitled Document',
          content: '<p>Start writing here...</p>',
          lastModified: new Date().toISOString()
      };
      const updatedDocs = [newDoc, ...documents];
      setDocuments(updatedDocs);
      localStorage.setItem('docuGeniusDocs', JSON.stringify(updatedDocs));
      openDocument(newDoc.id);
  };

  const openDocument = (id) => {
      const doc = documents.find(d => d.id === id);
      if (doc) {
          setCurrentDocId(id);
          // Reset Editor State
          setFileName(doc.name);
          setPageColor('#ffffff');
          setWatermark(null);
          setOrientation('portrait');
          setMargins('normal');
          setView('editor');
          // Content loading happens in Editor Effect via currentDocId
      }
  };

  const deleteDocument = (id) => {
      if (window.confirm("Are you sure you want to delete this document?")) {
          const updated = documents.filter(d => d.id !== id);
          setDocuments(updated);
          localStorage.setItem('docuGeniusDocs', JSON.stringify(updated));
      }
  };

  const goToDashboard = () => {
      saveDocumentImmediate();
      setView('dashboard');
      // Refresh list
      const rawDocs = localStorage.getItem('docuGeniusDocs');
      if(rawDocs) setDocuments(JSON.parse(rawDocs).sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified)));
  };

  // --- Editor Logic ---

  // Load Content when entering Editor
  useEffect(() => {
    if (view === 'editor' && currentDocId && editorRef.current) {
        const doc = documents.find(d => d.id === currentDocId);
        if (doc) {
            editorRef.current.innerHTML = doc.content;
            updateWordCount();
        }
    }
  }, [view, currentDocId]);

  // Auto-save Loop
  useEffect(() => {
    if (view !== 'editor') return;
    const interval = setInterval(() => {
       saveDocumentImmediate();
    }, 3000); 
    return () => clearInterval(interval);
  }, [view, currentDocId, fileName]); // Include fileName to capture title changes

  const saveDocumentImmediate = () => {
      if (!editorRef.current || !currentDocId) return;
      
      const content = editorRef.current.innerHTML;
      const updatedDocs = documents.map(d => {
          if (d.id === currentDocId) {
              return { ...d, content, name: fileName, lastModified: new Date().toISOString() };
          }
          return d;
      });
      
      setDocuments(updatedDocs);
      localStorage.setItem('docuGeniusDocs', JSON.stringify(updatedDocs));
      setSaveStatus('Saved to device');
  };

  const handleEditorChange = () => {
      setSaveStatus('Saving...');
      updateWordCount();
  };

  // --- Helper Functions ---
  const updateWordCount = () => {
    if (editorRef.current) {
      const text = editorRef.current.innerText || "";
      const words = text.trim().split(/\s+/).filter(w => w.length > 0).length;
      setTotalWords(words);
    }
  };

  const formatDoc = (cmd, value = null) => {
    document.execCommand(cmd, false, value);
    if (editorRef.current) editorRef.current.focus();
    handleEditorChange();
  };

  const handlePaste = (e) => {
      e.preventDefault();
      const text = e.clipboardData.getData('text/plain');
      document.execCommand('insertText', false, text);
      handleEditorChange();
  };

  // --- Missing Functions Restored ---
  const handleImageUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            formatDoc('insertImage', e.target.result);
        };
        reader.readAsDataURL(file);
    }
  };

  const insertLink = () => {
    const url = prompt("Enter URL:", "https://");
    if (url) formatDoc('createLink', url);
  };

  const insertTable = () => {
    const html = `
      <table style="border-collapse: collapse; width: 100%; margin: 10px 0;" border="1">
        <tbody>
          <tr><td style="padding: 8px; border: 1px solid #ccc;">Cell 1</td><td style="padding: 8px; border: 1px solid #ccc;">Cell 2</td></tr>
          <tr><td style="padding: 8px; border: 1px solid #ccc;">Cell 3</td><td style="padding: 8px; border: 1px solid #ccc;">Cell 4</td></tr>
        </tbody>
      </table><br/>`;
    document.execCommand('insertHTML', false, html);
    handleEditorChange();
  };

  const insertDate = () => {
    document.execCommand('insertText', false, new Date().toLocaleDateString());
    handleEditorChange();
  };

  const toggleWatermark = (text) => {
    setWatermark(current => current === text ? null : text);
  };

  const insertMergeField = (field) => {
    document.execCommand('insertText', false, `{{${field}}}`);
    handleEditorChange();
  };
  // -------------------------------

  const handlePrint = () => window.print();

  const handleExport = async (format) => {
    if (!editorRef.current) return;
    const contentHTML = editorRef.current.innerHTML;
    const contentText = editorRef.current.innerText;
    const safeFileName = fileName || 'Document';

    const downloadBlob = (blob, name) => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    };

    switch (format) {
        case 'txt': {
            downloadBlob(new Blob([contentText], { type: 'text/plain' }), `${safeFileName}.txt`);
            break;
        }
        case 'docx': {
            const preHtml = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>${safeFileName}</title></head><body>`;
            const html = preHtml + contentHTML + "</body></html>";
            downloadBlob(new Blob(['\ufeff', html], { type: 'application/msword' }), `${safeFileName}.doc`);
            break;
        }
        case 'pdf': {
             try {
                await loadScript("https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js");
                const element = editorRef.current;
                const opt = { margin: 1, filename: `${safeFileName}.pdf`, image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 2 }, jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' } };
                window.html2pdf().set(opt).from(element).save();
             } catch (e) { alert("Failed to load PDF generator."); }
             break;
        }
    }
  };

  // --- AI Functions ---
  const callGemini = async (prompt, contextText = "") => {
    setAiLoading(true);
    setAiResponse(null);
    try {
      const fullPrompt = contextText 
        ? `Context: "${contextText}"\n\nTask: ${prompt}\n\nTarget Language: ${language}\nProvide only the result text.`
        : `${prompt}\n\nTarget Language: ${language}\nProvide only the result text.`;

      const data = await fetchWithRetry(
        `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${API_KEY}`,
        { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ contents: [{ parts: [{ text: fullPrompt }] }] }) }
      );
      setAiResponse(data.candidates?.[0]?.content?.parts?.[0]?.text || "No response.");
    } catch (error) { setAiResponse("Error connecting to AI services."); } 
    finally { setAiLoading(false); }
  };

  const insertAIContent = () => {
      if (!aiResponse || !editorRef.current) return;
      editorRef.current.focus();
      if (selectionRange) {
          const sel = window.getSelection();
          sel.removeAllRanges();
          sel.addRange(selectionRange);
          if (selectedText) document.execCommand('delete');
      }
      document.execCommand('insertText', false, aiResponse);
      setAiResponse(null);
      setAiPrompt('');
      setShowAIPanel(false);
      handleEditorChange();
  };

  // --- Toolbar & Style Logic ---
  const getPageStyle = () => {
    const base = {
        fontFamily: 'Calibri, sans-serif',
        fontSize: '12pt',
        lineHeight: '1.5',
        transform: `scale(${zoomLevel / 100})`,
        transformOrigin: 'top center',
        backgroundColor: pageColor, 
        boxShadow: '0 8px 24px rgba(0,0,0,0.12)', 
        marginBottom: '2rem',
        position: 'relative',
        minHeight: '1056px'
    };
    if (orientation === 'landscape') { base.width = '1056px'; base.minHeight = '816px'; } 
    else { base.width = '816px'; base.minHeight = '1056px'; }
    if (viewMode === 'web') { base.width = '100%'; base.minHeight = '100vh'; base.marginBottom = '0'; base.transform = 'none'; base.boxShadow = 'none'; }
    base.padding = margins === 'wide' ? '144px' : margins === 'narrow' ? '48px' : '96px';
    return base;
  };

  // Toolbar JSX extraction for cleaner main render
  const renderToolbar = () => {
    switch (activeTab) {
      case 'File':
        return (
          <div className="flex items-center space-x-6 px-6 h-24 animate-fadeIn overflow-x-auto">
             <div className="flex flex-col space-y-2 border-r border-gray-200 pr-6 h-20 justify-center min-w-fit">
                 <div className="flex space-x-3">
                    <ToolbarAction icon={FilePlus} label="New" onClick={createNewDocument} />
                    <ToolbarAction icon={FolderOpen} label="Open" onClick={() => goToDashboard()} />
                 </div>
                 <span className="text-[10px] text-gray-400 font-medium text-center uppercase tracking-wider">Document</span>
             </div>

             <div className="flex flex-col space-y-2 border-r border-gray-200 pr-6 h-20 justify-center min-w-fit">
                 <div className="flex space-x-3">
                    <button onClick={() => handleExport('docx')} className="flex flex-col items-center justify-center px-4 py-2 rounded hover:bg-blue-50 text-gray-600 min-w-[70px] transition-colors group" title="Download as Word">
                        <div className="relative">
                            <FileText size={24} strokeWidth={1.5} className="text-blue-600" />
                        </div>
                        <span className="text-[11px] font-medium group-hover:text-blue-700 mt-1">Word</span>
                    </button>
                    
                    <button onClick={() => handleExport('pdf')} className="flex flex-col items-center justify-center px-4 py-2 rounded hover:bg-red-50 text-gray-600 min-w-[70px] transition-colors group" title="Download as PDF">
                        <FileType size={24} strokeWidth={1.5} className="text-red-600" />
                        <span className="text-[11px] font-medium group-hover:text-red-700 mt-1">PDF</span>
                    </button>
                 </div>
                 <span className="text-[10px] text-gray-400 font-medium text-center uppercase tracking-wider">Export</span>
             </div>

             <ToolbarAction icon={Printer} label="Print" onClick={handlePrint} />
          </div>
        );
      case 'Insert':
        return (
          <div className="flex items-center space-x-6 px-6 h-24 animate-fadeIn overflow-x-auto">
             <div className="flex space-x-2 border-r border-gray-200 pr-6 h-16 items-center min-w-fit">
                <ToolbarAction icon={ImageIcon} label="Picture" onClick={() => imageInputRef.current.click()} />
                <ToolbarAction icon={Table} label="Table" onClick={insertTable} />
             </div>
             <div className="flex space-x-2 border-r border-gray-200 pr-6 h-16 items-center min-w-fit">
                <ToolbarAction icon={LinkIcon} label="Link" onClick={insertLink} />
                <ToolbarAction icon={MinusSquare} label="Line" onClick={() => document.execCommand('insertHorizontalRule')} />
             </div>
             <ToolbarAction icon={Calendar} label="Date" onClick={insertDate} />
          </div>
        );
      case 'Design':
        return (
          <div className="flex items-center space-x-6 px-6 h-24 animate-fadeIn overflow-x-auto">
             <div className="flex flex-col space-y-2 border-r border-gray-200 pr-6 h-20 justify-center min-w-fit">
                <div className="flex space-x-6">
                   <div className="flex flex-col items-center cursor-pointer relative group">
                      <label htmlFor="colorPicker" className="p-2 rounded hover:bg-gray-100 cursor-pointer flex flex-col items-center transition-colors">
                         <Palette size={26} className="text-[#0078d4]" />
                         <div className="h-1.5 w-10 mt-2 border border-gray-300 rounded-full shadow-sm" style={{ backgroundColor: pageColor }}></div>
                         <span className="text-[11px] font-medium text-gray-600 mt-1">Page Color</span>
                      </label>
                      <input 
                        id="colorPicker"
                        type="color" 
                        value={pageColor} 
                        onChange={(e) => { setPageColor(e.target.value); setSaveStatus('Saving...'); }}
                        className="absolute opacity-0 w-full h-full cursor-pointer"
                      />
                   </div>
                   
                   <div className="p-2 rounded hover:bg-gray-100 flex flex-col items-center group relative cursor-pointer transition-colors">
                       <Stamp size={26} className="text-[#0078d4]" />
                       <span className="text-[11px] font-medium text-gray-600 mt-2">Watermark</span>
                       
                       <div className="absolute top-full left-0 mt-2 w-40 bg-white shadow-xl border border-gray-200 rounded-md hidden group-hover:block z-50 overflow-hidden">
                          <button onClick={() => toggleWatermark('CONFIDENTIAL')} className="w-full text-left px-4 py-2 text-xs hover:bg-gray-50 text-gray-600 border-b border-gray-50">Confidential</button>
                          <button onClick={() => toggleWatermark('DRAFT')} className="w-full text-left px-4 py-2 text-xs hover:bg-gray-50 text-gray-600 border-b border-gray-50">Draft</button>
                          <button onClick={() => toggleWatermark('URGENT')} className="w-full text-left px-4 py-2 text-xs hover:bg-gray-50 text-gray-600 border-b border-gray-50">Urgent</button>
                          <button onClick={() => setWatermark(null)} className="w-full text-left px-4 py-2 text-xs hover:bg-red-50 text-red-600 font-medium">Remove Watermark</button>
                       </div>
                   </div>
                </div>
                <span className="text-[10px] text-gray-400 font-medium text-center uppercase tracking-wider w-full">Background</span>
             </div>
          </div>
        );
      case 'Layout':
        return (
          <div className="flex items-center space-x-6 px-6 h-24 animate-fadeIn overflow-x-auto">
             <div className="flex flex-col space-y-2 border-r border-gray-200 pr-6 h-20 justify-center min-w-fit">
                <div className="flex space-x-3">
                    <button onClick={() => setMargins('narrow')} className={`text-xs px-4 py-2 rounded hover:bg-blue-50 transition-colors border ${margins === 'narrow' ? 'bg-blue-50 border-blue-300 text-blue-700 shadow-sm' : 'border-gray-200 text-gray-600'}`}>Narrow</button>
                    <button onClick={() => setMargins('normal')} className={`text-xs px-4 py-2 rounded hover:bg-blue-50 transition-colors border ${margins === 'normal' ? 'bg-blue-50 border-blue-300 text-blue-700 shadow-sm' : 'border-gray-200 text-gray-600'}`}>Normal</button>
                    <button onClick={() => setMargins('wide')} className={`text-xs px-4 py-2 rounded hover:bg-blue-50 transition-colors border ${margins === 'wide' ? 'bg-blue-50 border-blue-300 text-blue-700 shadow-sm' : 'border-gray-200 text-gray-600'}`}>Wide</button>
                </div>
                <span className="text-[10px] text-gray-400 font-medium text-center uppercase tracking-wider">Page Margins</span>
             </div>
             <div className="flex flex-col space-y-2 h-20 justify-center min-w-fit">
                <div className="flex space-x-3">
                    <button onClick={() => setOrientation('portrait')} className={`text-xs px-4 py-2 rounded border flex items-center gap-2 hover:bg-blue-50 transition-colors ${orientation === 'portrait' ? 'bg-blue-50 border-blue-300 text-blue-700 shadow-sm' : 'border-gray-200 text-gray-600'}`}><div className="w-3 h-4 border border-current"></div> Portrait</button>
                    <button onClick={() => setOrientation('landscape')} className={`text-xs px-4 py-2 rounded border flex items-center gap-2 hover:bg-blue-50 transition-colors ${orientation === 'landscape' ? 'bg-blue-50 border-blue-300 text-blue-700 shadow-sm' : 'border-gray-200 text-gray-600'}`}><div className="w-4 h-3 border border-current"></div> Landscape</button>
                </div>
                <span className="text-[10px] text-gray-400 font-medium text-center uppercase tracking-wider">Orientation</span>
             </div>
          </div>
        );
      case 'References':
        return (
          <div className="flex items-center space-x-6 px-6 h-24 animate-fadeIn overflow-x-auto">
             <div className="flex flex-col space-y-2 border-r border-gray-200 pr-6 h-20 justify-center min-w-fit">
                <ToolbarAction icon={Book} label="Table of Contents" onClick={() => document.execCommand('insertHTML', false, '<div class="toc" style="background:#f9f9f9; padding:10px; margin:10px 0; border:1px solid #ddd;"><b>Table of Contents</b><br/><i>Add headings to populate</i></div>')} />
             </div>
             <div className="flex space-x-4 h-16 items-center min-w-fit">
                <ToolbarAction icon={Quote} label="Insert Citation" onClick={() => document.execCommand('insertText', false, ' (Author, 2024) ')} />
                <ToolbarAction icon={FileText} label="Insert Footnote" onClick={() => { 
                    document.execCommand('superscript'); 
                    document.execCommand('insertText', false, '1'); 
                    document.execCommand('superscript'); 
                }} />
             </div>
          </div>
        );
      case 'Mailings':
        return (
          <div className="flex items-center space-x-6 px-6 h-24 animate-fadeIn overflow-x-auto">
             <div className="flex flex-col space-y-2 border-r border-gray-200 pr-6 h-20 justify-center min-w-fit">
                <div className="flex space-x-2">
                    <ToolbarAction icon={Mail} label="Start Merge" onClick={() => alert("Mail Merge Wizard started...")} />
                    <ToolbarAction icon={Users} label="Select Recipients" onClick={() => alert("Recipient list loaded (Simulation)")} />
                </div>
                <span className="text-[10px] text-gray-400 font-medium text-center uppercase tracking-wider">Start Mail Merge</span>
             </div>
             <div className="flex flex-col space-y-2 border-r border-gray-200 pr-6 h-20 justify-center min-w-fit">
                <div className="flex space-x-3 items-center pt-1">
                    <button onClick={() => insertMergeField('Name')} className="px-3 py-1.5 text-xs border border-gray-300 rounded hover:bg-blue-50 hover:border-blue-300 hover:text-blue-700 transition-colors bg-white font-medium">{'{{Name}}'}</button>
                    <button onClick={() => insertMergeField('Address')} className="px-3 py-1.5 text-xs border border-gray-300 rounded hover:bg-blue-50 hover:border-blue-300 hover:text-blue-700 transition-colors bg-white font-medium">{'{{Address}}'}</button>
                    <button onClick={() => insertMergeField('Date')} className="px-3 py-1.5 text-xs border border-gray-300 rounded hover:bg-blue-50 hover:border-blue-300 hover:text-blue-700 transition-colors bg-white font-medium">{'{{Date}}'}</button>
                </div>
                <span className="text-[10px] text-gray-400 font-medium text-center uppercase tracking-wider">Write & Insert Fields</span>
             </div>
             <ToolbarAction icon={Play} label="Preview Results" onClick={() => alert("Previewing results for Recipient #1")} />
          </div>
        );
      case 'Review':
        return (
          <div className="flex items-center space-x-6 px-6 h-24 animate-fadeIn overflow-x-auto">
             <div className="flex space-x-4 h-16 items-center border-r border-gray-200 pr-6 min-w-fit">
               <ToolbarAction icon={CheckSquare} label="Check Grammar" onClick={() => { 
                   setAiPrompt("Check grammar and spelling"); 
                   setShowAIPanel(true); 
                   callGemini("Check grammar and spelling for this text", editorRef.current?.innerText); 
               }} />
               <ToolbarAction icon={FileText} label="Word Count" onClick={() => alert(`Word Count: ${totalWords}\nCharacters: ${editorRef.current?.innerText.length || 0}`)} />
             </div>
          </div>
        );
      case 'View':
        return (
          <div className="flex items-center space-x-6 px-6 h-24 animate-fadeIn overflow-x-auto">
             <div className="flex flex-col space-y-2 border-r border-gray-200 pr-6 h-20 justify-center min-w-fit">
                <div className="flex space-x-3">
                    <button onClick={() => setViewMode('print')} className={`text-xs px-4 py-2 rounded border flex items-center gap-2 transition-colors ${viewMode === 'print' ? 'bg-blue-50 border-blue-300 text-blue-700 shadow-sm' : 'border-gray-200 text-gray-600 hover:bg-gray-50'}`}><Layout size={16}/> Print Layout</button>
                    <button onClick={() => setViewMode('web')} className={`text-xs px-4 py-2 rounded border flex items-center gap-2 transition-colors ${viewMode === 'web' ? 'bg-blue-50 border-blue-300 text-blue-700 shadow-sm' : 'border-gray-200 text-gray-600 hover:bg-gray-50'}`}><Grid size={16}/> Web Layout</button>
                </div>
                <span className="text-[10px] text-gray-400 font-medium text-center uppercase tracking-wider">Views</span>
             </div>
             <div className="flex items-center space-x-3">
                <button onClick={() => setZoomLevel(Math.max(50, zoomLevel - 10))} className="p-2 hover:bg-gray-100 rounded-full transition-colors"><Minus size={20}/></button>
                <span className="w-16 text-center font-medium text-sm bg-gray-100 py-1 rounded">{zoomLevel}%</span>
                <button onClick={() => setZoomLevel(Math.min(200, zoomLevel + 10))} className="p-2 hover:bg-gray-100 rounded-full transition-colors"><Plus size={20}/></button>
             </div>
          </div>
        );
      case 'Home':
      default:
        return (
        <div className="flex items-center h-24 px-6 space-x-6 overflow-x-auto scrollbar-thin scrollbar-thumb-gray-300 pb-2 sm:pb-0 animate-fadeIn">
          {/* Clipboard */}
          <div className="flex flex-col items-center justify-center px-2 space-y-2 text-gray-600 border-r border-gray-200 pr-6 shrink-0 h-20">
             <button onClick={() => { navigator.clipboard.readText().then(text => { document.execCommand('insertText', false, text); updateWordCount(); setSaveStatus('Saving...'); }).catch(err => alert("Use Ctrl+V")); }} className="p-2 hover:bg-gray-100 rounded transition-colors group flex flex-col items-center">
               <FileText size={26} className="text-[#0078d4]" />
               <span className="text-[11px] font-medium mt-1 group-hover:text-black">Paste</span>
             </button>
          </div>

          {/* Font */}
          <div className="flex flex-col space-y-2 border-r border-gray-200 pr-6 shrink-0 h-20 justify-center min-w-fit">
            <div className="flex space-x-3">
              <div className="relative">
                <select onChange={(e) => formatDoc('fontName', e.target.value)} className="appearance-none border border-gray-300 rounded-sm text-xs h-8 w-40 px-2 focus:border-[#0078d4] focus:outline-none bg-white hover:border-gray-400 transition-colors">
                    <option value="Arial">Arial</option>
                    <option value="Arial Black">Arial Black</option>
                    <option value="Book Antiqua">Book Antiqua</option>
                    <option value="Brush Script MT">Brush Script MT</option>
                    <option value="Calibri">Calibri</option>
                    <option value="Cambria">Cambria</option>
                    <option value="Comic Sans MS">Comic Sans MS</option>
                    <option value="Consolas">Consolas</option>
                    <option value="Courier New">Courier New</option>
                    <option value="Garamond">Garamond</option>
                    <option value="Geneva">Geneva</option>
                    <option value="Georgia">Georgia</option>
                    <option value="Helvetica">Helvetica</option>
                    <option value="Impact">Impact</option>
                    <option value="Lucida Console">Lucida Console</option>
                    <option value="Lucida Sans Unicode">Lucida Sans</option>
                    <option value="Monaco">Monaco</option>
                    <option value="Palatino Linotype">Palatino Linotype</option>
                    <option value="Roboto">Roboto</option>
                    <option value="Segoe UI">Segoe UI</option>
                    <option value="Tahoma">Tahoma</option>
                    <option value="Times New Roman">Times New Roman</option>
                    <option value="Trebuchet MS">Trebuchet MS</option>
                    <option value="Verdana">Verdana</option>
                </select>
                <ChevronDown size={12} className="absolute right-2 top-2.5 pointer-events-none text-gray-500" />
              </div>
              <div className="relative">
                  <select onChange={(e) => formatDoc('fontSize', e.target.value)} className="appearance-none border border-gray-300 rounded-sm text-xs h-8 w-16 px-2 focus:border-[#0078d4] focus:outline-none bg-white hover:border-gray-400 transition-colors" defaultValue="3">
                    <option value="1">8</option><option value="2">10</option><option value="3">12</option><option value="4">14</option><option value="5">18</option><option value="6">24</option><option value="7">36</option>
                  </select>
                  <ChevronDown size={12} className="absolute right-2 top-2.5 pointer-events-none text-gray-500" />
              </div>
            </div>
            <div className="flex space-x-1.5">
              <ToolbarBtn icon={Bold} onClick={() => formatDoc('bold')} />
              <ToolbarBtn icon={Italic} onClick={() => formatDoc('italic')} />
              <ToolbarBtn icon={Underline} onClick={() => formatDoc('underline')} />
              <div className="w-px h-5 bg-gray-300 mx-2 self-center"></div>
              <ToolbarBtn icon={Highlighter} onClick={() => formatDoc('backColor', 'yellow')} />
              <ToolbarBtn icon={Type} onClick={() => formatDoc('foreColor', 'red')} />
            </div>
          </div>

          {/* Paragraph */}
          <div className="flex flex-col space-y-2 border-r border-gray-200 pr-6 shrink-0 h-20 justify-center min-w-fit">
             <div className="flex space-x-1.5">
               <ToolbarBtn icon={List} onClick={() => formatDoc('insertUnorderedList')} />
               <ToolbarBtn icon={ListOrdered} onClick={() => formatDoc('insertOrderedList')} />
             </div>
             <div className="flex space-x-1.5">
               <ToolbarBtn icon={AlignLeft} onClick={() => formatDoc('justifyLeft')} />
               <ToolbarBtn icon={AlignCenter} onClick={() => formatDoc('justifyCenter')} />
               <ToolbarBtn icon={AlignRight} onClick={() => formatDoc('justifyRight')} />
             </div>
          </div>

          {/* Styles */}
          <div className="flex items-center space-x-3 shrink-0 h-20">
             {['Normal', 'Heading 1', 'Heading 2', 'Title'].map((style, i) => (
               <button key={style} onClick={() => formatDoc('formatBlock', style === 'Normal' ? 'P' : style === 'Title' ? 'H1' : style === 'Heading 1' ? 'H2' : 'H3')} className="h-16 w-24 border border-gray-200 rounded bg-white hover:bg-[#f0f8ff] hover:border-[#0078d4] flex flex-col items-start p-2.5 transition-colors group text-left">
                  <span className={`text-xs mt-0.5 w-full truncate ${i===0?'font-normal':i===1?'font-semibold text-[#2b579a] text-sm':i===2?'font-medium text-[#2b579a]': 'font-bold text-lg'}`}>AaBbCc</span>
                  <span className="text-[10px] text-gray-500 mt-auto w-full truncate group-hover:text-[#0078d4]">{style}</span>
               </button>
             ))}
          </div>
        </div>
        );
    }
  };


  // ==========================
  // RENDER
  // ==========================

  if (view === 'dashboard') {
      return <Dashboard documents={documents} onCreate={createNewDocument} onOpen={openDocument} onDelete={deleteDocument} />;
  }

  return (
    <div className="flex flex-col h-screen bg-[#f0f0f0] font-sans overflow-hidden text-[#323130]">
      <input type="file" ref={fileInputRef} className="hidden" />
      <input type="file" ref={imageInputRef} className="hidden" accept="image/*" onChange={handleImageUpload} />

      {/* Editor Top Bar */}
      <div className="h-14 bg-white text-gray-800 flex items-center justify-between px-4 shadow-sm border-b border-gray-200 z-50 shrink-0">
        <div className="flex items-center space-x-4">
          <div className="p-2 hover:bg-[#e6f2ff] rounded cursor-pointer transition-colors" onClick={goToDashboard} title="Back to Dashboard">
             <FileText size={26} className="text-[#0078d4]" />
          </div>
          
          <div>
             <input 
              type="text" 
              value={fileName}
              onChange={(e) => { setFileName(e.target.value); setSaveStatus('Saving...'); }}
              className="bg-transparent text-lg font-normal text-gray-800 focus:outline-none focus:border focus:border-[#0078d4] focus:bg-white px-2 rounded transition-colors w-64 -ml-2"
            />
            <div className="flex items-center gap-4 text-[12px] text-gray-500 -mt-1">
                 <div className="flex gap-2">
                    {['File', 'Home', 'Insert', 'Design', 'Layout', 'References', 'Mailings', 'Review', 'View'].map(t => (
                        <button key={t} onClick={() => setActiveTab(t)} className={`hover:bg-gray-100 px-2 py-0.5 rounded ${activeTab === t ? 'font-semibold text-black bg-gray-100' : ''}`}>{t}</button>
                    ))}
                 </div>
                 <div className="flex items-center gap-1 pl-4 border-l border-gray-300">
                     {saveStatus === 'Saving...' ? <Cloud size={14} className="animate-pulse" /> : <Check size={14} />}
                     <span>{saveStatus}</span>
                 </div>
            </div>
          </div>
        </div>
        
        <div className="flex items-center gap-3">
           <button className="p-2 rounded-full hover:bg-gray-100"><MessageSquarePlus size={20} className="text-gray-600" /></button>
           <button onClick={() => setShowAIPanel(!showAIPanel)} className="flex items-center gap-2 bg-gradient-to-r from-[#0078d4] to-[#005a9e] text-white px-4 py-2 rounded-full shadow-sm hover:shadow hover:from-[#006bd0] hover:to-[#004c8c] transition-all">
               <Sparkles size={16} />
               <span className="text-sm font-medium">AI Assistant</span>
           </button>
           <div className="w-9 h-9 rounded-full bg-purple-600 text-white flex items-center justify-center text-sm font-bold ml-2">JD</div>
        </div>
      </div>

      {/* Ribbon/Toolbar (Removed rounding, increased height) */}
      <div className={`bg-[#f3f6fc] border-b border-gray-300 shadow-sm shrink-0 z-40 select-none flex items-center w-full px-4 ${isRibbonOpen ? 'h-auto py-1' : 'h-0 overflow-hidden'}`}>
         {renderToolbar()}
      </div>

      {/* Main Content */}
      <div className="flex-1 flex overflow-hidden relative bg-[#f9fbfd]">
        <div className="flex-1 overflow-y-auto overflow-x-hidden p-8 flex flex-col items-center relative scroll-smooth">
            <div className="relative group transition-all duration-300 ease-in-out mt-6" style={{ transform: `scale(${zoomLevel / 100})`, transformOrigin: 'top center' }}>
                {watermark && (
                    <div className="absolute inset-0 flex items-center justify-center pointer-events-none z-0 overflow-hidden select-none">
                        <span className="text-gray-300 font-bold transform -rotate-45 whitespace-nowrap opacity-30" style={{ fontSize: '10rem', mixBlendMode: 'multiply' }}>{watermark}</span>
                    </div>
                )}
                <div 
                  className="bg-white outline-none print:shadow-none print:m-0 print:w-full selection:bg-[#b3d7ff]"
                  contentEditable
                  ref={editorRef}
                  onPaste={handlePaste}
                  onInput={handleEditorChange}
                  onKeyUp={updateWordCount}
                  onClick={updateWordCount}
                  suppressContentEditableWarning
                  style={{...getPageStyle(), transform: 'none', marginBottom: 0}} 
                >
                </div>
            </div>
        </div>

        {/* AI Sidebar */}
        <div className={`bg-white shadow-xl border-l border-gray-200 flex flex-col transition-all duration-300 absolute right-0 top-0 bottom-0 z-30 ${showAIPanel ? 'w-80 translate-x-0' : 'w-80 translate-x-full'}`}>
             <div className="p-4 border-b border-gray-100 flex justify-between items-center bg-white">
                 <h2 className="font-bold text-gray-800 text-sm flex items-center gap-2"><Sparkles size={16} className="text-purple-600"/> AI Assistant</h2>
                 <button onClick={() => setShowAIPanel(false)}><X size={18} className="text-gray-400"/></button>
             </div>
             <div className="flex-1 overflow-y-auto p-4 space-y-4 bg-gray-50">
                 {aiResponse ? (
                     <div className="bg-white p-3 rounded shadow-sm border text-sm prose prose-sm">{aiResponse}</div>
                 ) : <p className="text-center text-gray-400 text-sm mt-10">How can I help you edit this document?</p>}
                 
                 {aiResponse && <button onClick={insertAIContent} className="w-full bg-[#0078d4] text-white py-2 rounded text-sm">Insert</button>}
             </div>
             <div className="p-3 bg-white border-t">
                 <div className="relative">
                    <input value={aiPrompt} onChange={(e) => setAiPrompt(e.target.value)} placeholder="Ask AI..." className="w-full border rounded-full px-4 py-2 text-sm pr-10 focus:outline-none focus:border-[#0078d4]" onKeyDown={(e) => e.key === 'Enter' && !e.shiftKey && callGemini(aiPrompt, selectedText)} />
                    <button onClick={() => callGemini(aiPrompt, selectedText)} className="absolute right-2 top-1.5 text-[#0078d4]"><Sparkles size={16}/></button>
                 </div>
             </div>
        </div>
      </div>
    </div>
  );
};

// Components
const ToolbarBtn = ({ icon: Icon, onClick, active }) => (
  <button onClick={onClick} className={`p-2 rounded-sm hover:bg-gray-200 text-gray-700 transition-all ${active ? 'bg-gray-300' : ''}`}><Icon size={18} strokeWidth={2} /></button>
);

const ToolbarAction = ({ icon: Icon, label, onClick }) => (
  <button onClick={onClick} className="flex flex-col items-center justify-center px-4 py-2 rounded hover:bg-gray-100 text-gray-600 min-w-[70px] transition-colors group">
    <Icon size={24} strokeWidth={1.5} className="text-gray-500 group-hover:text-[#0078d4]" />
    <span className="text-[11px] font-medium group-hover:text-[#0078d4] mt-1">{label}</span>
  </button>
);

export default App;
