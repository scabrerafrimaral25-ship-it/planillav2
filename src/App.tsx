import React, { useState, useRef, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx-js-style';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf';
import pdfWorker from 'pdfjs-dist/legacy/build/pdf.worker.min.js?url';
import { Upload, FileJson, Search, Plus, Trash2, Printer, Download, Save, Home, AlertCircle, FileText, FileSpreadsheet, Database, LayoutList } from 'lucide-react';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

type Pallet = {
  containerId: string;
  palletId: string;
  quantity: number;
  boxes: number;
  weight: number;
  description: string;
  rawRow?: string;
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
  const [activeTab, setActiveTab] = useState<'plan' | 'database'>('plan');
  const [globalSearch, setGlobalSearch] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
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

  const fixedMasterData = useMemo(() => {
    let currentContainer = '';
    return session.masterData.map(p => {
      if (p.containerId) {
        currentContainer = p.containerId;
        return p;
      } else {
        return { ...p, containerId: currentContainer };
      }
    });
  }, [session.masterData]);

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

  const processExcel = async (file: File) => {
    setIsProcessing(true);
    
    try {
      const arrayBuffer = await file.arrayBuffer();
      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      
      const pallets: Pallet[] = [];
      let allJson: any[] = [];
      let lastColLote = 7;
      let lastColContenedor = 2;
      let lastColCantidad = 3;
      let lastColCajas = 4;
      let lastColKilos = 5;
      let lastColDesc = 6;

      workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        allJson = allJson.concat(json);
        
        let headerRowIdx = -1;
        let colContenedor = lastColContenedor;
        let colCantidad = lastColCantidad;
        let colCajas = lastColCajas;
        let colKilos = lastColKilos;
        let colDesc = lastColDesc;
        let colLote = lastColLote;

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
            
            colContenedor = findCol(['contenedor', 'cntr', 'equipo'], [], lastColContenedor);
            colLote = findCol(['lote', 'pallet', 'sscc'], ['código', 'codigo', 'id'], lastColLote);
            colCantidad = findCol(['cantidad', 'cant'], [], lastColCantidad);
            colCajas = findCol(['cajas', 'bultos', 'bx'], [], lastColCajas);
            colKilos = findCol(['kilos', 'peso', 'kg', 'neto'], [], lastColKilos);
            colDesc = findCol(['descripci', 'producto', 'item'], [], lastColDesc);
            
            lastColLote = colLote;
            lastColContenedor = colContenedor;
            lastColCantidad = colCantidad;
            lastColCajas = colCajas;
            lastColKilos = colKilos;
            lastColDesc = colDesc;
            break;
          }
        }

        const startRow = headerRowIdx !== -1 ? headerRowIdx + 1 : 0;
        let currentContainerId = '';
        
        for (let i = startRow; i < json.length; i++) {
          const row = json[i] as any[];
          if (!row || row.length === 0) continue;
          
          let containerId = String(row[colContenedor] || '').trim();
          if (containerId && !containerId.toLowerCase().includes('total')) {
            currentContainerId = containerId;
          } else if (!containerId) {
            containerId = currentContainerId;
          }
          
          const palletId = String(row[colLote] || '').trim();
          const description = String(row[colDesc] || '').trim();
          const quantity = parseNumber(row[colCantidad]);
          
          if (!palletId && !description && !quantity) continue;
          if (containerId.toLowerCase().includes('total') || palletId.toLowerCase().includes('total')) continue;
          
          pallets.push({
            containerId,
            palletId,
            quantity,
            boxes: parseNumber(row[colCajas]),
            weight: parseNumber(row[colKilos]),
            description,
            rawRow: row.map(c => String(c || '')).join(' '),
          });
        }
      });
      
      if (pallets.length === 0) {
        showAlert('No se encontraron datos válidos. Verifica que el Excel tenga columnas como "Contenedor" y "Lote".');
        return;
      }
      
      setSession(prev => ({ 
        ...prev, 
        masterData: pallets, 
        searchIds: [],
        originalData: allJson,
        colLote: lastColLote,
        colContenedor: lastColContenedor
      }));
      setScreen('workspace');
      showToast(`Cargados ${pallets.length} pallets exitosamente`);
    } catch (err: any) {
      console.error("Error procesando Excel:", err);
      showAlert(`Error al procesar el archivo: ${err.message || 'Formato irreconocible'}`);
    } finally {
      setIsProcessing(false);
    }
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
    
    // Check for non-existent IDs (ignoring leading zeros and special chars for comparison)
    const normalize = (id: string) => String(id).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, '').toLowerCase();
    
    const notFound: string[] = [];
    const idsToAdd: string[] = [];
    
    newIds.forEach(id => {
      const normId = normalize(id);
      if (!normId) return;
      
      const matchedPallet = fixedMasterData.find(p => {
        const matchId = normalize(p.palletId) === normId;
        // Also check if the ID is somewhere in the description or raw row as a fallback
        const matchDesc = normalize(p.description).includes(normId);
        const matchRaw = p.rawRow ? normalize(p.rawRow).includes(normId) : false;
        return matchId || matchDesc || matchRaw;
      });
      
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
    const normalize = (val: string) => String(val).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, '').toLowerCase();
    const normalizedSearchIds = session.searchIds.map(id => normalize(id));
    
    const searchedPallets = fixedMasterData.filter(p => {
      const matchId = normalizedSearchIds.includes(normalize(p.palletId));
      const matchDesc = normalizedSearchIds.some(id => normalize(p.description).includes(id));
      const matchRaw = p.rawRow ? normalizedSearchIds.some(id => normalize(p.rawRow!).includes(id)) : false;
      return matchId || matchDesc || matchRaw;
    });

    const containersWithSearched = Array.from(new Set(
      searchedPallets.map(p => p.containerId)
    ));

    if (containersWithSearched.length === 0) {
      showAlert('No hay datos para exportar. Por favor, ingresa IDs de pallets primero.');
      return;
    }

    const wsData: any[][] = [];
    
    // Header rows
    wsData.push([]); // Row 1: Empty
    wsData.push(['PLANILLA DE CARGA', '', '', '', '', '', '']); // Row 2
    wsData.push(['Contenedor', 'Cant.', 'Bultos', 'Peso', 'Descripción', '', 'Pallet ID']); // Row 3
    
    // Data rows
    containersWithSearched.forEach(containerId => {
      const containerPallets = fixedMasterData.filter(p => p.containerId === containerId);
      
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

    // Summary section
    const totalBultos = searchedPallets.reduce((sum, p) => sum + p.boxes, 0);
    const totalKilos = searchedPallets.reduce((sum, p) => sum + p.weight, 0);
    const cotes = Array.from(new Set(
      searchedPallets
        .map(p => p.description.match(/COTE P\d+/i)?.[0])
        .filter(Boolean)
    ));

    wsData.push([]); // Empty row before summary
    wsData.push([]); 
    
    const summaryStartRow = wsData.length;
    
    wsData.push(['', '', '', '', 'RESUMEN TOTAL (SOLO BUSCADOS)', '', '']);
    wsData.push(['', '', '', '', 'TOTAL PALLETS', 'CAJAS', 'KG']);
    wsData.push(['', '', '', '', searchedPallets.length, totalBultos, totalKilos]);
    wsData.push([]);
    wsData.push(['', '', '', '', 'COTES DE INGRESO (UNICOS)', '', '']);
    cotes.forEach(cote => {
      wsData.push(['', '', '', '', cote, '', '']);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // Apply styles
    // Merge A2:G2
    ws['!merges'] = [
      { s: { r: 1, c: 0 }, e: { r: 1, c: 6 } }
    ];
    
    // Merge summary headers
    ws['!merges'].push({ s: { r: summaryStartRow, c: 4 }, e: { r: summaryStartRow, c: 6 } });
    ws['!merges'].push({ s: { r: summaryStartRow + 4, c: 4 }, e: { r: summaryStartRow + 4, c: 6 } });
    for (let i = 0; i < cotes.length; i++) {
      ws['!merges'].push({ s: { r: summaryStartRow + 5 + i, c: 4 }, e: { r: summaryStartRow + 5 + i, c: 6 } });
    }
    
    // Column widths
    ws['!cols'] = [
      { wch: 18 }, // Contenedor
      { wch: 6 },  // Cant
      { wch: 8 },  // Bultos
      { wch: 8 },  // Peso
      { wch: 60 }, // Descripción
      { wch: 8 },  // CAJAS
      { wch: 12 }, // Pallet ID / KG
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
    for (let R = 3; R < wsData.length; R++) {
      const row = wsData[R];
      if (!row || row.length === 0) continue; // Empty row
      
      // Check if it's a summary row
      if (R >= summaryStartRow) {
        if (row[4] !== undefined && row[4] !== '') {
          for (let C = 4; C <= 6; C++) {
            const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };
            
            ws[cellRef].s = {
              border: {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                left: { style: 'thin' },
                right: { style: 'thin' }
              }
            };
            
            if (R === summaryStartRow || R === summaryStartRow + 4) {
              ws[cellRef].s.font = { bold: true };
              ws[cellRef].s.fill = { patternType: 'solid', fgColor: { rgb: 'E5E7EB' } }; // Light gray
            }
          }
        }
        continue;
      }
      
      const palletId = String(row[6] || '').trim();
      const description = String(row[4] || '').trim();
      if (!palletId && !description) continue;
      
      const normPalletId = normalize(palletId);
      const normDesc = normalize(description);
      
      const isSearched = normalizedSearchIds.some(id => 
        normPalletId === id || normDesc.includes(id)
      );
      
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
      
      <div className="max-w-2xl w-full bg-white rounded-2xl shadow-sm p-8 border border-gray-100 relative">
        {isProcessing && (
          <div className="absolute inset-0 bg-white/80 backdrop-blur-sm rounded-2xl flex flex-col items-center justify-center z-10">
            <div className="w-12 h-12 border-4 border-accent border-t-transparent rounded-full animate-spin mb-4"></div>
            <p className="text-gray-700 font-medium">Procesando archivo...</p>
          </div>
        )}
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
    const normalize = (val: string) => String(val).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, '').toLowerCase();
    const normalizedSearchIds = session.searchIds.map(id => normalize(id));
    
    const searchedPallets = session.masterData.filter(p => {
      const matchId = normalizedSearchIds.includes(normalize(p.palletId));
      const matchDesc = normalizedSearchIds.some(id => normalize(p.description).includes(id));
      const matchRaw = p.rawRow ? normalizedSearchIds.some(id => normalize(p.rawRow!).includes(id)) : false;
      return matchId || matchDesc || matchRaw;
    });
    
    const containersWithSearched = Array.from(new Set(
      searchedPallets.map(p => p.containerId)
    ));

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
          
          <div className="flex bg-gray-100 p-1 rounded-lg">
            <button
              onClick={() => setActiveTab('plan')}
              className={`flex items-center gap-2 px-4 py-1.5 rounded-md text-sm font-medium transition-colors ${activeTab === 'plan' ? 'bg-white text-accent shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
            >
              <LayoutList size={16} /> Plan de Carga
            </button>
            <button
              onClick={() => setActiveTab('database')}
              className={`flex items-center gap-2 px-4 py-1.5 rounded-md text-sm font-medium transition-colors ${activeTab === 'database' ? 'bg-white text-accent shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
            >
              <Database size={16} /> Base de Datos
            </button>
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
          {activeTab === 'plan' ? (
            <>
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
                            const normalize = (val: string) => String(val).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, '').toLowerCase();
                            const validIds = session.searchIds.filter(id => {
                              const normId = normalize(id);
                              return fixedMasterData.some(p => 
                                normalize(p.palletId) === normId || 
                                normalize(p.description).includes(normId) || 
                                (p.rawRow ? normalize(p.rawRow).includes(normId) : false)
                              );
                            });
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
                      const normalize = (val: string) => String(val).replace(/[^a-zA-Z0-9]/g, '').replace(/^0+/, '').toLowerCase();
                      const normId = normalize(id);
                      const exists = fixedMasterData.some(p => 
                        normalize(p.palletId) === normId || 
                        normalize(p.description).includes(normId) || 
                        (p.rawRow ? normalize(p.rawRow).includes(normId) : false)
                      );
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

                    <div className="flex justify-between items-center mb-6">
                      <h2 className="text-2xl font-bold text-gray-900">Planilla de Carga Generada</h2>
                      <div className="flex gap-3">
                        <button onClick={() => window.print()} className="flex items-center gap-2 px-4 py-2 border border-gray-300 rounded-lg text-gray-700 hover:bg-gray-50 font-medium text-sm transition-colors">
                          <Printer size={16} /> Imprimir
                        </button>
                        <button onClick={exportExcel} className="flex items-center gap-2 px-4 py-2 bg-accent text-white rounded-lg hover:bg-accent/90 font-medium text-sm transition-colors shadow-sm">
                          <Download size={16} /> Exportar Excel
                        </button>
                      </div>
                    </div>

                    {/* Container Cards */}
                    {containersWithSearched.map(containerId => {
                      const containerPallets = fixedMasterData.filter(p => p.containerId === containerId);
                      const contTotalBultos = containerPallets.reduce((sum, p) => sum + p.boxes, 0);
                      const contTotalKilos = containerPallets.reduce((sum, p) => sum + p.weight, 0);
                      const contTotalCant = containerPallets.reduce((sum, p) => sum + p.quantity, 0);
                      
                      return (
                        <div key={containerId} className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden break-inside-avoid">
                          <div className="px-6 py-4 flex justify-between items-center bg-white">
                            <div className="flex items-center gap-4">
                              <div className="w-10 h-10 rounded-lg bg-green-100 text-green-600 flex items-center justify-center">
                                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M10 17h4V5H2v12h3"/><path d="M20 17h2v-9h-5V5h-7"/><path d="M15 17h2"/><circle cx="8.5" cy="17.5" r="1.5"/><circle cx="18.5" cy="17.5" r="1.5"/></svg>
                              </div>
                              <div>
                                <h3 className="text-lg font-bold text-gray-900 flex items-center gap-2">
                                  {containerId}
                                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-gray-400"><path d="m18 15-6-6-6 6"/></svg>
                                </h3>
                                <p className="text-sm text-gray-500">Contenedor</p>
                              </div>
                            </div>
                            <div className="flex items-center gap-8 text-right">
                              <div>
                                <p className="text-xs text-gray-500 uppercase tracking-wider mb-1">Pallets</p>
                                <p className="text-lg font-bold text-gray-900">{containerPallets.length}</p>
                              </div>
                              <div>
                                <p className="text-xs text-gray-500 uppercase tracking-wider mb-1">Peso Total</p>
                                <p className="text-lg font-bold text-gray-900">{contTotalKilos.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 3})} kg</p>
                              </div>
                            </div>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left text-sm">
                              <thead className="text-gray-500 border-b border-gray-100 bg-white">
                                <tr>
                                  <th className="px-6 py-3 font-medium">Pallet ID</th>
                                  <th className="px-6 py-3 font-medium">Descripción</th>
                                  <th className="px-6 py-3 font-medium text-right">Bultos</th>
                                  <th className="px-6 py-3 font-medium text-right">Cant.</th>
                                  <th className="px-6 py-3 font-medium text-right">Peso</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-gray-50">
                                {containerPallets.map((p, idx) => {
                                  const isSearched = normalizedSearchIds.some(id => 
                                    normalize(p.palletId) === id || 
                                    normalize(p.description).includes(id) || 
                                    (p.rawRow ? normalize(p.rawRow).includes(id) : false)
                                  );
                                  return (
                                    <tr key={`${p.palletId}-${idx}`} className={`${isSearched ? 'bg-[#fffaeb]' : 'bg-white hover:bg-gray-50'} relative transition-colors`}>
                                      {isSearched && (
                                        <td className="absolute left-0 top-0 bottom-0 w-1 bg-indigo-500"></td>
                                      )}
                                      <td className={`px-6 py-3 font-mono ${isSearched ? 'font-bold text-indigo-600' : 'text-gray-500'}`}>{p.palletId}</td>
                                      <td className={`px-6 py-3 max-w-xs truncate ${isSearched ? 'text-gray-900 font-medium' : 'text-gray-500'}`} title={p.description}>{p.description}</td>
                                      <td className={`px-6 py-3 text-right ${isSearched ? 'text-gray-900 font-medium' : 'text-gray-500'}`}>{p.boxes}</td>
                                      <td className={`px-6 py-3 text-right ${isSearched ? 'text-gray-900 font-medium' : 'text-gray-500'}`}>{p.quantity}</td>
                                      <td className={`px-6 py-3 text-right ${isSearched ? 'text-gray-900 font-medium' : 'text-gray-500'}`}>{p.weight}</td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                              <tfoot className="bg-white font-bold text-gray-800 border-t border-gray-100">
                                <tr>
                                  <td colSpan={2} className="px-6 py-4 text-right text-sm">Total Contenedor</td>
                                  <td className="px-6 py-4 text-right">{contTotalBultos}</td>
                                  <td className="px-6 py-4 text-right">{contTotalCant}</td>
                                  <td className="px-6 py-4 text-right">{contTotalKilos.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 3})}</td>
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
            </>
          ) : (
            <div className="flex-1 flex flex-col bg-white overflow-hidden">
              <div className="p-4 border-b border-gray-200 flex items-center justify-between bg-gray-50">
                <h2 className="text-lg font-bold text-gray-800">Base de Datos de Pallets</h2>
                <div className="relative w-72">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
                  <input 
                    type="text" 
                    placeholder="Buscar por ID, contenedor o descripción..."
                    value={globalSearch}
                    onChange={(e) => setGlobalSearch(e.target.value)}
                    className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-accent/50"
                  />
                </div>
              </div>
              <div className="flex-1 overflow-auto">
                <table className="w-full text-left text-sm">
                  <thead className="bg-white text-gray-500 border-b border-gray-200 sticky top-0 z-10 shadow-sm">
                    <tr>
                      <th className="px-6 py-3 font-medium">Contenedor</th>
                      <th className="px-6 py-3 font-medium">Lote / ID</th>
                      <th className="px-6 py-3 font-medium">Descripción</th>
                      <th className="px-6 py-3 font-medium text-right">Bultos</th>
                      <th className="px-6 py-3 font-medium text-right">Cantidad</th>
                      <th className="px-6 py-3 font-medium text-right">Peso (kg)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {session.masterData
                      .filter(p => 
                        !globalSearch || 
                        p.palletId.toLowerCase().includes(globalSearch.toLowerCase()) ||
                        p.containerId.toLowerCase().includes(globalSearch.toLowerCase()) ||
                        p.description.toLowerCase().includes(globalSearch.toLowerCase())
                      )
                      .slice(0, 500) // Limit to 500 for performance
                      .map((p, idx) => (
                      <tr key={`${p.palletId}-${idx}`} className="hover:bg-gray-50">
                        <td className="px-6 py-3 font-mono text-gray-600">{p.containerId}</td>
                        <td className="px-6 py-3 font-mono font-medium text-gray-900">{p.palletId}</td>
                        <td className="px-6 py-3 text-gray-600 max-w-md truncate" title={p.description}>{p.description}</td>
                        <td className="px-6 py-3 text-right font-mono text-gray-700">{p.boxes}</td>
                        <td className="px-6 py-3 text-right font-mono text-gray-700">{p.quantity}</td>
                        <td className="px-6 py-3 text-right font-mono text-gray-700">{p.weight.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                      </tr>
                    ))}
                    {session.masterData.length > 500 && !globalSearch && (
                      <tr>
                        <td colSpan={6} className="px-6 py-4 text-center text-gray-500 text-sm italic">
                          Mostrando los primeros 500 pallets. Usa el buscador para encontrar pallets específicos.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}
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
