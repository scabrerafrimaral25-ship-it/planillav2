import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.js?url';
import { Upload, FileJson, Search, Plus, Trash2, Printer, Download, Save, Home, AlertCircle, FileText, FileSpreadsheet } from 'lucide-react';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

type Pallet = {
  containerId: string;
  palletId: string;
  quantity: number;
  boxes: number;
  weight: number;
  description: string;
};

type Session = {
  version: number;
  savedAt: string;
  sessionName: string;
  masterData: Pallet[];
  searchIds: string[];
  originalData: any[][];
  colLote: number;
  colContenedor: number;
};

function parseNumber(val: any): number {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  let str = String(val).trim();
  if (str.includes('.') && str.includes(',')) {
    if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      str = str.replace(/,/g, '');
    }
  } else if (str.includes(',')) {
    if (str.match(/,\d{1,2}$/)) {
      str = str.replace(',', '.');
    } else {
      str = str.replace(/,/g, '');
    }
  }
  return parseFloat(str) || 0;
}

export default function App() {
  const [screen, setScreen] = useState<'home' | 'upload' | 'workspace'>('home');
  const [session, setSession] = useState<Session>({
    version: 1,
    savedAt: new Date().toISOString(),
    sessionName: 'Nueva Sesión',
    masterData: [],
    searchIds: [],
    originalData: [],
    colLote: -1,
    colContenedor: -1
  });
  const [toast, setToast] = useState<string | null>(null);
  const [dragActive, setDragActive] = useState(false);
  const [searchInput, setSearchInput] = useState('');
  const [isAlertOpen, setIsAlertOpen] = useState(false);
  const [alertMessage, setAlertMessage] = useState('');

  const fileInputRef = useRef<HTMLInputElement>(null);
  const sessionInputRef = useRef<HTMLInputElement>(null);
  const extractInputRef = useRef<HTMLInputElement>(null);

  const showToast = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(null), 3000);
  };

  const showAlert = (msg: string) => {
    setAlertMessage(msg);
    setIsAlertOpen(true);
  };

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processExcel(e.dataTransfer.files[0]);
    }
  };

  const processExcel = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        if (!e.target?.result) throw new Error("No se pudo leer el archivo");
        const data = new Uint8Array(e.target.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        let headerRowIdx = -1;
        let colContenedor = 2;
        let colCantidad = 3;
        let colCajas = 4;
        let colKilos = 5;
        let colDesc = 6;
        let colLote = 7;

        for (let i = 0; i < Math.min(50, json.length); i++) {
          const row = json[i] as any[];
          if (!row) continue;
          const rowStr = row.map(c => String(c || '').toLowerCase());
          
          const hasContenedor = rowStr.some(c => c.includes('contenedor') || c.includes('cntr') || c.includes('equipo'));
          const hasLote = rowStr.some(c => c.includes('lote') || c.includes('pallet') || c.includes('id'));
          
          if (hasContenedor || hasLote) {
            headerRowIdx = i;
            
            const findCol = (primaryKeywords: string[], secondaryKeywords: string[], fallback: number) => {
              let idx = rowStr.findIndex(c => primaryKeywords.some(k => c.includes(k)));
              if (idx === -1 && secondaryKeywords.length > 0) {
                idx = rowStr.findIndex(c => secondaryKeywords.some(k => c.includes(k)));
              }
              return idx !== -1 ? idx : fallback;
            };
            
            colContenedor = findCol(['contenedor', 'cntr', 'equipo'], [], 2);
            colLote = findCol(['lote', 'pallet', 'sscc'], ['código', 'codigo', 'id'], 7);
            colCantidad = findCol(['cantidad', 'cant'], [], 3);
            colCajas = findCol(['cajas', 'bultos', 'bx'], [], 4);
            colKilos = findCol(['kilos', 'peso', 'kg', 'neto'], [], 5);
            colDesc = findCol(['descripci', 'producto', 'item'], [], 6);
            break;
          }
        }

        const startRow = headerRowIdx !== -1 ? headerRowIdx + 1 : 0;
        const pallets: Pallet[] = [];
        
        for (let i = startRow; i < json.length; i++) {
          const row = json[i] as any[];
          if (!row || row.length === 0) continue;
          
          const containerId = String(row[colContenedor] || '').trim();
          const palletId = String(row[colLote] || '').trim();
          
          if (!containerId && !palletId) continue;
          if (containerId.toLowerCase().includes('total') || palletId.toLowerCase().includes('total')) continue;
          
          pallets.push({
            containerId,
            palletId,
            quantity: parseNumber(row[colCantidad]),
            boxes: parseNumber(row[colCajas]),
            weight: parseNumber(row[colKilos]),
            description: String(row[colDesc] || '').trim(),
          });
        }
        
        if (pallets.length === 0) {
          showAlert('No se encontraron datos válidos. Verifica que el Excel tenga columnas como "Contenedor" y "Lote".');
          return;
        }
        
        setSession(prev => ({ 
          ...prev, 
          masterData: pallets, 
          searchIds: [],
          originalData: json,
          colLote: colLote,
          colContenedor: colContenedor
        }));
        setScreen('workspace');
        showToast(`Cargados ${pallets.length} pallets exitosamente`);
      } catch (err: any) {
        console.error("Error procesando Excel:", err);
        showAlert(`Error al procesar el archivo: ${err.message || 'Formato irreconocible'}`);
      }
    };
    reader.onerror = () => showAlert('Error de lectura del archivo en el navegador.');
    reader.readAsArrayBuffer(file);
  };

  const loadSession = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target?.result as string);
        if (data.version && data.masterData) {
          setSession(data);
          setScreen('workspace');
          showToast('Sesión restaurada');
        } else {
          showAlert('El archivo no tiene el formato de sesión válido.');
        }
      } catch (err) {
        showAlert('Error al leer el archivo de sesión.');
      }
    };
    reader.readAsText(file);
  };

  const saveSession = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify({
      ...session,
      savedAt: new Date().toISOString()
    }));
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href",     dataStr);
    downloadAnchorNode.setAttribute("download", `${session.sessionName || 'sesion'}.json`);
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
    showToast('Sesión guardada');
  };

  const handleAddSearchIds = () => {
    if (!searchInput.trim()) return;
    const newIds = searchInput.split(/[\s,\n]+/).filter(id => id.trim() !== '');
    
    // Check for non-existent IDs (ignoring leading zeros for comparison)
    const normalize = (id: string) => id.replace(/^0+/, '').trim();
    
    const notFound: string[] = [];
    const idsToAdd: string[] = [];
    
    newIds.forEach(id => {
      const normId = normalize(id);
      const matchedPallet = session.masterData.find(p => normalize(p.palletId) === normId);
      
      if (matchedPallet) {
        idsToAdd.push(matchedPallet.palletId); // Use the exact ID from master data
      } else {
        notFound.push(id);
        idsToAdd.push(id); // Add it anyway so they can see it in the list
      }
    });
    
    if (notFound.length > 0) {
      showAlert(`Los siguientes IDs no existen en la planilla maestra:\n${notFound.join(', ')}`);
    }

    setSession(prev => {
      const uniqueIds = Array.from(new Set([...prev.searchIds, ...idsToAdd]));
      return { ...prev, searchIds: uniqueIds };
    });
    setSearchInput('');
  };

  const handleExtractFromFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    try {
      let text = '';
      if (file.name.toLowerCase().endsWith('.pdf')) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const textContent = await page.getTextContent();
          text += textContent.items.map((item: any) => item.str).join(' ') + ' ';
        }
      } else {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          json.forEach((row: any) => {
            text += row.join(' ') + ' ';
          });
        });
      }
      
      const matches = text.match(/\b\d{6,7}\b/g) || [];
      const extractedIds = Array.from(new Set(matches)).filter(id => !/^0+$/.test(id));
      
      if (extractedIds.length > 0) {
        setSearchInput(prev => prev + (prev ? ' ' : '') + extractedIds.join(' '));
        showToast(`Se extrajeron ${extractedIds.length} IDs del archivo`);
      } else {
        showAlert('No se encontraron números de 6 o 7 dígitos en el archivo.');
      }
    } catch (err: any) {
      console.error("Error extrayendo datos:", err);
      showAlert(`Error al extraer datos del archivo: ${err.message || 'Error desconocido'}`);
    }
    
    // Reset input
    if (extractInputRef.current) extractInputRef.current.value = '';
  };

  const exportExcel = () => {
    // Get containers that have searched pallets
    const containersWithSearched = Array.from(new Set(
      session.masterData
        .filter(p => session.searchIds.includes(p.palletId))
        .map(p => p.containerId)
    ));

    const wsData: any[][] = [];
    
    // Header rows
    wsData.push([]); // Row 1: Empty
    wsData.push(['PLANILLA DE CARGA', '', '', '', '', '', '']); // Row 2
    wsData.push(['Contenedor', 'Cant.', 'Bultos', 'Peso', 'Descripción', '', 'Pallet ID']); // Row 3
    
    // Data rows
    containersWithSearched.forEach(containerId => {
      const containerPallets = session.masterData.filter(p => p.containerId === containerId);
      
      containerPallets.forEach(p => {
        wsData.push([
          p.containerId,
          p.quantity,
          p.boxes,
          p.weight,
          p.description,
          '', // Empty column F
          p.palletId
        ]);
      });
      
      wsData.push([]); // Empty row between containers
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // Apply styles
    // Merge A2:G2
    ws['!merges'] = [
      { s: { r: 1, c: 0 }, e: { r: 1, c: 6 } }
    ];
    
    // Column widths
    ws['!cols'] = [
      { wch: 18 }, // Contenedor
      { wch: 6 },  // Cant
      { wch: 8 },  // Bultos
      { wch: 8 },  // Peso
      { wch: 60 }, // Descripción
      { wch: 4 },  // Empty
      { wch: 12 }, // Pallet ID
    ];

    // Style Header Row 2
    if (ws['A2']) {
      ws['A2'].s = {
        font: { bold: true, sz: 14 },
        alignment: { horizontal: 'center', vertical: 'center' }
      };
    }
    
    // Style Header Row 3
    const headers = ['A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3'];
    headers.forEach(ref => {
      if (ws[ref]) {
        ws[ref].s = {
          font: { bold: true },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
          }
        };
      }
    });

    // Style Data Rows
    const normalize = (id: string) => id.replace(/^0+/, '').trim();
    const searchIdsNorm = session.searchIds.map(normalize);

    for (let R = 3; R < wsData.length; R++) {
      const row = wsData[R];
      if (!row || row.length === 0) continue; // Empty row
      
      const palletId = String(row[6] || '').trim();
      if (!palletId) continue;
      
      const normPalletId = normalize(palletId);
      const isSearched = searchIdsNorm.includes(normPalletId);
      
      // Apply borders to all cells in the row
      for (let C = 0; C <= 6; C++) {
        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cellRef]) {
          ws[cellRef] = { t: 's', v: '' }; // Create empty cell if missing
        }
        
        ws[cellRef].s = {
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
          }
        };
        
        // Highlight Pallet ID cell if searched
        if (C === 6 && isSearched) {
          ws[cellRef].s.fill = {
            patternType: 'solid',
            fgColor: { rgb: 'FFFFFF00' } // Yellow
          };
          ws[cellRef].s.font = { bold: true };
        }
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Plan de Carga");
    XLSX.writeFile(wb, `${session.sessionName || 'plan_carga'}.xlsx`);
  };

  // Render logic
  const renderHome = () => (
    <div className="min-h-screen flex items-center justify-center p-4">
      <div className="max-w-4xl w-full grid grid-cols-1 md:grid-cols-2 gap-8">
        <button 
          onClick={() => setScreen('upload')}
          className="bg-white p-10 rounded-2xl shadow-sm hover:shadow-md transition-shadow flex flex-col items-center justify-center gap-4 text-accent border border-gray-100 group"
        >
          <div className="p-4 bg-accent/10 rounded-full group-hover:bg-accent/20 transition-colors">
            <Plus size={48} className="text-accent" />
          </div>
          <h2 className="text-2xl font-bold text-gray-800">Nueva Carga</h2>
          <p className="text-gray-500 text-center">Sube un archivo Excel maestro para comenzar una nueva planificación.</p>
        </button>
        
        <button 
          onClick={() => sessionInputRef.current?.click()}
          className="bg-white p-10 rounded-2xl shadow-sm hover:shadow-md transition-shadow flex flex-col items-center justify-center gap-4 text-green border border-gray-100 group"
        >
          <div className="p-4 bg-green/10 rounded-full group-hover:bg-green/20 transition-colors">
            <FileJson size={48} className="text-green" />
          </div>
          <h2 className="text-2xl font-bold text-gray-800">Abrir Sesión</h2>
          <p className="text-gray-500 text-center">Restaura un archivo .json guardado previamente.</p>
          <input 
            type="file" 
            accept=".json" 
            className="hidden" 
            ref={sessionInputRef} 
            onChange={loadSession} 
          />
        </button>
      </div>
    </div>
  );

  const renderUpload = () => (
    <div className="min-h-screen flex flex-col items-center justify-center p-4">
      <button 
        onClick={() => setScreen('home')}
        className="absolute top-6 left-6 flex items-center gap-2 text-gray-500 hover:text-gray-800"
      >
        <Home size={20} /> Volver al inicio
      </button>
      
      <div className="max-w-2xl w-full bg-white rounded-2xl shadow-sm p-8 border border-gray-100">
        <h2 className="text-2xl font-bold text-gray-800 mb-6 text-center">Subir Excel Maestro</h2>
        
        <div 
          className={`border-2 border-dashed rounded-xl p-12 text-center cursor-pointer transition-colors ${dragActive ? 'border-accent bg-accent/5' : 'border-gray-300 hover:border-accent hover:bg-gray-50'}`}
          onDragEnter={handleDrag}
          onDragLeave={handleDrag}
          onDragOver={handleDrag}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
        >
          <Upload size={48} className="mx-auto text-gray-400 mb-4" />
          <p className="text-lg text-gray-700 mb-2">Arrastra y suelta tu archivo Excel aquí</p>
          <p className="text-sm text-gray-500">o haz clic para seleccionar</p>
          <input 
            type="file" 
            accept=".xlsx, .xls" 
            className="hidden" 
            ref={fileInputRef} 
            onChange={(e) => e.target.files?.[0] && processExcel(e.target.files[0])} 
          />
        </div>
        
        <div className="mt-8 bg-gray-50 p-4 rounded-lg text-sm text-gray-600">
          <h3 className="font-semibold text-gray-800 mb-2">Formato esperado:</h3>
          <ul className="list-disc pl-5 space-y-1">
            <li>Detección automática de columnas si encuentra encabezados "Contenedor" y "Lote" en las primeras 50 filas.</li>
            <li>Fallback por posición: Col C (Contenedor), D (Cantidad), E (Cajas), F (Kilos), G (Descripción), H (Lote).</li>
            <li>Soporta números con formato europeo (1.234,56) y americano (1,234.56).</li>
          </ul>
        </div>
      </div>
    </div>
  );

  const renderWorkspace = () => {
    // Calculate derived data
    const containersWithSearched = Array.from(new Set(
      session.masterData
        .filter(p => session.searchIds.includes(p.palletId))
        .map(p => p.containerId)
    ));

    const searchedPallets = session.masterData.filter(p => session.searchIds.includes(p.palletId));
    const totalBultos = searchedPallets.reduce((sum, p) => sum + p.boxes, 0);
    const totalKilos = searchedPallets.reduce((sum, p) => sum + p.weight, 0);
    const cotes = Array.from(new Set(
      searchedPallets
        .map(p => p.description.match(/COTE P\d+/i)?.[0])
        .filter(Boolean)
    ));

    return (
      <div className="min-h-screen flex flex-col">
        {/* Header */}
        <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between sticky top-0 z-10 no-print">
          <div className="flex items-center gap-4">
            <div className="bg-accent text-white p-2 rounded-lg">
              <Upload size={24} />
            </div>
            <h1 className="text-xl font-bold text-gray-800 hidden sm:block">Gestor de Carga</h1>
            <div className="bg-gray-100 text-gray-600 px-3 py-1 rounded-full text-sm font-medium">
              {session.masterData.length} pallets
            </div>
          </div>
          
          <div className="flex items-center gap-4">
            <input 
              type="text" 
              value={session.sessionName}
              onChange={(e) => setSession({...session, sessionName: e.target.value})}
              className="border border-gray-300 rounded-md px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-accent/50 w-48"
              placeholder="Nombre de sesión"
            />
            <button onClick={() => window.print()} className="p-2 text-gray-500 hover:text-gray-800 hover:bg-gray-100 rounded-md" title="Imprimir">
              <Printer size={20} />
            </button>
            <button onClick={exportExcel} className="p-2 text-green hover:text-green-700 hover:bg-green/10 rounded-md" title="Exportar Excel">
              <Download size={20} />
            </button>
            <button onClick={saveSession} className="p-2 text-accent hover:text-accent-700 hover:bg-accent/10 rounded-md" title="Guardar Sesión">
              <Save size={20} />
            </button>
            <button onClick={() => setScreen('home')} className="p-2 text-gray-500 hover:text-gray-800 hover:bg-gray-100 rounded-md" title="Inicio">
              <Home size={20} />
            </button>
          </div>
        </header>

        <main className="flex-1 flex flex-col md:flex-row overflow-hidden">
          {/* Sidebar */}
          <aside className="w-full md:w-80 bg-white border-r border-gray-200 flex flex-col no-print h-[calc(100vh-73px)] overflow-y-auto">
            <div className="p-4 border-b border-gray-200">
              <h3 className="font-semibold text-gray-800 mb-3 flex items-center gap-2">
                <Search size={18} /> Buscar Pallets
              </h3>
              <textarea 
                value={searchInput}
                onChange={(e) => setSearchInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    handleAddSearchIds();
                  }
                }}
                placeholder="Ingresa IDs separados por espacio, coma o salto de línea..."
                className="w-full h-24 border border-gray-300 rounded-md p-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-accent/50 resize-none mb-2"
              />
              <div className="flex gap-2">
                <button 
                  onClick={handleAddSearchIds}
                  className="flex-1 bg-accent text-white py-2 rounded-md text-sm font-medium hover:bg-accent/90 transition-colors"
                >
                  Agregar
                </button>
                <button 
                  onClick={() => extractInputRef.current?.click()}
                  className="flex-1 bg-gray-100 text-gray-700 py-2 rounded-md text-sm font-medium hover:bg-gray-200 transition-colors flex items-center justify-center gap-1"
                  title="Extraer de Excel o PDF"
                >
                  <FileText size={16} /> Importar
                </button>
                <input 
                  type="file" 
                  accept=".pdf, .xlsx, .xls" 
                  className="hidden" 
                  ref={extractInputRef} 
                  onChange={handleExtractFromFile} 
                />
              </div>
            </div>
            
            <div className="flex-1 overflow-y-auto p-4">
              <div className="flex items-center justify-between mb-3">
                <h4 className="font-medium text-gray-700 text-sm">IDs Ingresados ({session.searchIds.length})</h4>
                {session.searchIds.length > 0 && (
                  <div className="flex gap-1">
                    <button 
                      onClick={() => {
                        const validIds = session.searchIds.filter(id => session.masterData.some(p => p.palletId === id));
                        setSession({...session, searchIds: validIds});
                      }}
                      className="text-xs text-red hover:underline"
                      title="Borrar no encontrados"
                    >
                      Limpiar
                    </button>
                    <span className="text-gray-300">|</span>
                    <button 
                      onClick={() => setSession({...session, searchIds: []})}
                      className="text-xs text-gray-500 hover:underline"
                    >
                      Borrar todo
                    </button>
                  </div>
                )}
              </div>
              
              <ul className="space-y-1">
                {session.searchIds.map(id => {
                  const exists = session.masterData.some(p => p.palletId === id);
                  return (
                    <li key={id} className={`flex items-center justify-between px-2 py-1.5 rounded text-sm font-mono ${exists ? 'bg-gray-50 text-gray-800' : 'bg-red/10 text-red'}`}>
                      <span>{id}</span>
                      <button 
                        onClick={() => setSession({...session, searchIds: session.searchIds.filter(i => i !== id)})}
                        className="text-gray-400 hover:text-red"
                      >
                        <Trash2 size={14} />
                      </button>
                    </li>
                  );
                })}
                {session.searchIds.length === 0 && (
                  <li className="text-sm text-gray-400 text-center py-4">No hay IDs ingresados</li>
                )}
              </ul>
            </div>
          </aside>

          {/* Content */}
          <div className="flex-1 overflow-y-auto p-6 bg-bg">
            {containersWithSearched.length === 0 ? (
              <div className="h-full flex flex-col items-center justify-center text-gray-400">
                <Search size={64} className="mb-4 opacity-20" />
                <p className="text-lg">Ingresa IDs de pallets para ver el plan de carga</p>
              </div>
            ) : (
              <div className="max-w-5xl mx-auto space-y-6">
                {/* Summary Card */}
                <div className="bg-[#1e3a8a] text-white rounded-xl shadow-md p-6">
                  <h2 className="text-xl font-bold mb-4 opacity-90">Resumen de Carga</h2>
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-6">
                    <div>
                      <p className="text-blue-200 text-sm mb-1">Pallets Buscados</p>
                      <p className="text-3xl font-mono font-bold">{searchedPallets.length}</p>
                    </div>
                    <div>
                      <p className="text-blue-200 text-sm mb-1">Total Bultos</p>
                      <p className="text-3xl font-mono font-bold">{totalBultos}</p>
                    </div>
                    <div>
                      <p className="text-blue-200 text-sm mb-1">Total Peso (kg)</p>
                      <p className="text-3xl font-mono font-bold">{totalKilos.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</p>
                    </div>
                    <div>
                      <p className="text-blue-200 text-sm mb-1">Cotes de Ingreso</p>
                      <p className="text-lg font-mono font-medium leading-tight">
                        {cotes.length > 0 ? cotes.join(', ') : '-'}
                      </p>
                    </div>
                  </div>
                </div>

                {/* Container Cards */}
                {containersWithSearched.map(containerId => {
                  const containerPallets = session.masterData.filter(p => p.containerId === containerId);
                  const contTotalBultos = containerPallets.reduce((sum, p) => sum + p.boxes, 0);
                  const contTotalKilos = containerPallets.reduce((sum, p) => sum + p.weight, 0);
                  const contTotalCant = containerPallets.reduce((sum, p) => sum + p.quantity, 0);
                  
                  return (
                    <div key={containerId} className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden break-inside-avoid">
                      <div className="bg-gray-50 px-6 py-4 border-b border-gray-200 flex justify-between items-center">
                        <h3 className="text-lg font-bold text-gray-800 flex items-center gap-2">
                          <span className="w-3 h-3 rounded-full bg-accent inline-block"></span>
                          Contenedor: <span className="font-mono">{containerId}</span>
                        </h3>
                        <span className="text-sm text-gray-500 font-medium">{containerPallets.length} pallets en total</span>
                      </div>
                      <div className="overflow-x-auto">
                        <table className="w-full text-left text-sm">
                          <thead className="bg-white text-gray-500 border-b border-gray-200">
                            <tr>
                              <th className="px-6 py-3 font-medium">Lote / ID</th>
                              <th className="px-6 py-3 font-medium">Descripción</th>
                              <th className="px-6 py-3 font-medium text-right">Bultos</th>
                              <th className="px-6 py-3 font-medium text-right">Cantidad</th>
                              <th className="px-6 py-3 font-medium text-right">Peso (kg)</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-100">
                            {containerPallets.map(p => {
                              const isSearched = session.searchIds.includes(p.palletId);
                              return (
                                <tr key={p.palletId} className={`${isSearched ? 'bg-yellow-50' : 'opacity-40 grayscale'}`}>
                                  <td className="px-6 py-3 font-mono font-medium text-gray-900">{p.palletId}</td>
                                  <td className="px-6 py-3 text-gray-600 max-w-xs truncate" title={p.description}>{p.description}</td>
                                  <td className="px-6 py-3 text-right font-mono text-gray-700">{p.boxes}</td>
                                  <td className="px-6 py-3 text-right font-mono text-gray-700">{p.quantity}</td>
                                  <td className="px-6 py-3 text-right font-mono text-gray-700">{p.weight.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                                </tr>
                              );
                            })}
                          </tbody>
                          <tfoot className="bg-gray-50 font-bold text-gray-800 border-t-2 border-gray-200">
                            <tr>
                              <td colSpan={2} className="px-6 py-3 text-right">TOTALES CONTENEDOR:</td>
                              <td className="px-6 py-3 text-right font-mono">{contTotalBultos}</td>
                              <td className="px-6 py-3 text-right font-mono">{contTotalCant}</td>
                              <td className="px-6 py-3 text-right font-mono">{contTotalKilos.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                            </tr>
                          </tfoot>
                        </table>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </main>
      </div>
    );
  };

  return (
    <>
      {screen === 'home' && renderHome()}
      {screen === 'upload' && renderUpload()}
      {screen === 'workspace' && renderWorkspace()}

      {/* Toast */}
      {toast && (
        <div className="fixed bottom-4 right-4 bg-gray-800 text-white px-4 py-3 rounded-lg shadow-lg flex items-center gap-3 z-50 animate-in fade-in slide-in-from-bottom-4 no-print">
          <div className="w-2 h-2 bg-green rounded-full"></div>
          {toast}
        </div>
      )}

      {/* Alert Modal */}
      {isAlertOpen && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4 no-print">
          <div className="bg-white rounded-xl shadow-xl max-w-md w-full p-6 animate-in zoom-in-95">
            <div className="flex items-start gap-4">
              <div className="bg-red/10 p-3 rounded-full flex-shrink-0">
                <AlertCircle className="text-red" size={24} />
              </div>
              <div>
                <h3 className="text-lg font-bold text-gray-900 mb-2">Atención</h3>
                <p className="text-gray-600 whitespace-pre-wrap text-sm">{alertMessage}</p>
              </div>
            </div>
            <div className="mt-6 flex justify-end">
              <button 
                onClick={() => setIsAlertOpen(false)}
                className="bg-gray-100 text-gray-800 px-4 py-2 rounded-lg font-medium hover:bg-gray-200 transition-colors"
              >
                Entendido
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
