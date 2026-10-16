import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  FileSpreadsheet,
  Upload,
  Layers,
  CheckCircle2,
  AlertCircle,
  X,
  Coins,
  Check,
  TrendingDown,
  ArrowRight,
  Tag,
  ToggleLeft,
  ToggleRight,
  Percent,
  SlidersHorizontal,
  Sparkles,
  PlusCircle,
  Edit3,
  Copy,
  FolderTree
} from 'lucide-react';
import { BarcodePriceItem, SheetBarcodeSummary } from '../types';

interface PriceTemplateModalProps {
  isOpen: boolean;
  onClose: () => void;
  items: BarcodePriceItem[];
  initialFile?: File | null;
  initialBase64?: string;
  initialSheetSummaries?: SheetBarcodeSummary[];
}

export const PriceTemplateModal: React.FC<PriceTemplateModalProps> = ({
  isOpen,
  onClose,
  items,
  initialFile,
  initialBase64,
  initialSheetSummaries,
}) => {
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [templateBase64, setTemplateBase64] = useState<string>('');
  const [templateSheets, setTemplateSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');

  // Multi-sheet control (User requested: multiple sheets with different brands/categories)
  const [applyToAllSheets, setApplyToAllSheets] = useState<boolean>(true);
  const [autoDetectBarcodeCol, setAutoDetectBarcodeCol] = useState<boolean>(true);

  // Column Target Mode: 'new_columns' (新建列追加，不覆盖原有内容) | 'existing_columns' (覆盖写入现有指定列)
  const [columnTargetMode, setColumnTargetMode] = useState<'new_columns' | 'existing_columns'>('new_columns');
  const [newOriginalColHeader, setNewOriginalColHeader] = useState<string>('提取-最低原价 (€)');
  const [newDiscountColHeader, setNewDiscountColHeader] = useState<string>('提取-折后特价 (€)');
  const [insertPlatformCol, setInsertPlatformCol] = useState<boolean>(true);
  const [newPlatformColHeader, setNewPlatformColHeader] = useState<string>('最低价渠道');

  const [templateColumns, setTemplateColumns] = useState<{
    index: number;
    letter: string;
    name: string;
    sampleText?: string;
  }[]>([]);
  const [templatePreviewRows, setTemplatePreviewRows] = useState<any[][]>([]);
  const [templateRawData, setTemplateRawData] = useState<any[][]>([]);
  const [templateHeaderRowIndex, setTemplateHeaderRowIndex] = useState<number>(1);
  const [templateStartDataRowIndex, setTemplateStartDataRowIndex] = useState<number>(2);

  // Mappings for Existing Columns Mode
  const [templateBarcodeCol, setTemplateBarcodeCol] = useState<number>(1);
  const [fillMode, setFillMode] = useState<'original_and_discount' | 'single' | 'multi'>('original_and_discount');

  // Mode 1: Original + Optional Discount Column
  const [originalPriceType, setOriginalPriceType] = useState<string>('lowest_original');
  const [originalPriceCol, setOriginalPriceCol] = useState<number>(0);
  const [insertDiscountCol, setInsertDiscountCol] = useState<boolean>(false);
  const [discountPriceType, setDiscountPriceType] = useState<string>('lowest_discount');
  const [discountPriceCol, setDiscountPriceCol] = useState<number>(0);
  const [discountNonSaleHandling, setDiscountNonSaleHandling] = useState<'empty' | 'original' | 'dash'>('empty');

  // Mode 2: Single Column
  const [singlePriceType, setSinglePriceType] = useState<string>('lowest_original');
  const [singlePriceCol, setSinglePriceCol] = useState<number>(0);

  // Mode 3: Multi-column mapping
  const [lowestOriginalPriceCol, setLowestOriginalPriceCol] = useState<number>(0);
  const [lowestDiscountPriceCol, setLowestDiscountPriceCol] = useState<number>(0);
  const [lowestEffectivePriceCol, setLowestEffectivePriceCol] = useState<number>(0);
  const [averagePriceCol, setAveragePriceCol] = useState<number>(0);
  const [meloniOriginalPriceCol, setMeloniOriginalPriceCol] = useState<number>(0);
  const [meloniDiscountPriceCol, setMeloniDiscountPriceCol] = useState<number>(0);
  const [megaOriginalPriceCol, setMegaOriginalPriceCol] = useState<number>(0);
  const [megaDiscountPriceCol, setMegaDiscountPriceCol] = useState<number>(0);
  const [maurysOriginalPriceCol, setMaurysOriginalPriceCol] = useState<number>(0);
  const [maurysDiscountPriceCol, setMaurysDiscountPriceCol] = useState<number>(0);
  const [risparmioOriginalPriceCol, setRisparmioOriginalPriceCol] = useState<number>(0);
  const [risparmioDiscountPriceCol, setRisparmioDiscountPriceCol] = useState<number>(0);
  const [platformNameCol, setPlatformNameCol] = useState<number>(0);
  const [titleCol, setTitleCol] = useState<number>(0);

  // Format settings
  const [valueFormat, setValueFormat] = useState<'number' | 'currency' | 'text'>('number');
  const [unmatchedFill, setUnmatchedFill] = useState<'empty' | 'zero' | 'dash'>('empty');

  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [statusText, setStatusText] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Initialize with passed file if present
  useEffect(() => {
    if (initialFile && initialBase64) {
      setTemplateFile(initialFile);
      setTemplateBase64(initialBase64);
      try {
        const binary = atob(initialBase64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        const wb = XLSX.read(bytes.buffer, { type: 'array' });
        setTemplateSheets(wb.SheetNames);
        const first = wb.SheetNames[0] || '';
        setSelectedSheet(first);
        const ws = wb.Sheets[first];
        const rawData = ws ? (XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][]) : [];
        parseSheetStructure(rawData);
      } catch {}
    }
  }, [initialFile, initialBase64]);

  if (!isOpen) return null;

  const parseSheetStructure = (rawData: any[][], preferredHeaderRow?: number) => {
    if (!rawData || rawData.length === 0) return;
    setTemplateRawData(rawData);

    let headerRowIdx = preferredHeaderRow !== undefined ? preferredHeaderRow - 1 : -1;

    if (headerRowIdx < 0) {
      let bestScore = -1;
      let bestRow = 0;
      const maxScan = Math.min(rawData.length, 6);

      for (let r = 0; r < maxScan; r++) {
        const row = rawData[r] || [];
        let score = 0;
        let nonNullCount = 0;

        for (let c = 0; c < row.length; c++) {
          const val = String(row[c] || '').trim();
          if (val) {
            nonNullCount++;
            const l = val.toLowerCase();
            if (l.includes('barcode') || l.includes('ean') || l.includes('c.barre') || l.includes('条码')) score += 50;
            if (l.includes('prezzo') || l.includes('price') || l.includes('costo') || l.includes('价格') || l.includes('单价') || l.includes('原价')) score += 40;
            if (l.includes('descrizione') || l.includes('prodotto') || l.includes('品名') || l.includes('名称')) score += 20;
          }
        }
        score += nonNullCount * 2;
        if (score > bestScore && nonNullCount >= 2) {
          bestScore = score;
          bestRow = r;
        }
      }
      headerRowIdx = bestRow >= 0 ? bestRow : 0;
    }

    const headerRow = rawData[headerRowIdx] || [];
    const maxCols = Math.max(...rawData.slice(0, 15).map((r) => (Array.isArray(r) ? r.length : 0)), headerRow.length);
    const dataRows = rawData.slice(headerRowIdx + 1, headerRowIdx + 30);
    const cols = [];

    for (let c = 0; c < maxCols; c++) {
      const val = String(headerRow[c] || '').trim();
      const letter = c < 26 ? String.fromCharCode(65 + c) : `${String.fromCharCode(64 + Math.floor(c / 26))}${String.fromCharCode(65 + (c % 26))}`;

      const samples: string[] = [];
      for (let r = 0; r < dataRows.length; r++) {
        const cellVal = String(dataRows[r]?.[c] ?? '').trim();
        if (cellVal && !samples.includes(cellVal)) {
          samples.push(cellVal.length > 25 ? cellVal.slice(0, 25) + '…' : cellVal);
          if (samples.length >= 2) break;
        }
      }

      cols.push({
        index: c + 1,
        letter,
        name: val || `第 ${c + 1} 列 (${letter}列)`,
        sampleText: samples.length > 0 ? samples.join(', ') : '暂无数据',
      });
    }

    setTemplateColumns(cols);
    setTemplateHeaderRowIndex(headerRowIdx + 1);
    setTemplateStartDataRowIndex(headerRowIdx + 2);
    setTemplatePreviewRows(rawData.slice(headerRowIdx, headerRowIdx + 10));

    // Intelligent auto-detection of Barcode, Original Price, and Discount Price columns
    let bestBarcodeCol = 1;
    let maxBarcodeScore = -1;
    let bestOrigPriceCol = 0;
    let maxOrigScore = -1;
    let bestDiscPriceCol = 0;
    let maxDiscScore = -1;

    cols.forEach((col) => {
      const headerText = col.name.toLowerCase();
      let bScore = 0;
      let oScore = 0;
      let dScore = 0;

      if (headerText.includes('barcode') || headerText.includes('ean') || headerText.includes('条形码') || headerText.includes('条码')) {
        bScore += 120;
      }
      if (headerText.includes('原价') || headerText.includes('正价') || headerText.includes('进价') || headerText.includes('prezzo listino') || headerText.includes('prezzo base') || headerText.includes('prezzo')) {
        oScore += 100;
      }
      if (headerText.includes('打折') || headerText.includes('折后') || headerText.includes('特价') || headerText.includes('促销') || headerText.includes('sconto') || headerText.includes('promo')) {
        dScore += 120;
      }

      const cIdx = col.index - 1;
      let numericBarcodeCount = 0;
      let numericPriceCount = 0;
      let blankCount = 0;

      dataRows.forEach((r) => {
        const val = String(r[cIdx] ?? '').trim();
        if (!val) {
          blankCount++;
          return;
        }
        const digitsOnly = val.replace(/[^0-9]/g, '');
        if (digitsOnly.length >= 8 && digitsOnly.length <= 14 && /^\d+$/.test(val)) {
          numericBarcodeCount++;
        }
        if (/^[0-9]+([.,][0-9]{1,2})?$/.test(val) && parseFloat(val.replace(',', '.')) < 1000) {
          numericPriceCount++;
        }
      });

      if (numericBarcodeCount > 0) bScore += numericBarcodeCount * 15;
      if (numericPriceCount > 0 && numericBarcodeCount === 0) {
        oScore += numericPriceCount * 10;
        dScore += numericPriceCount * 5;
      }
      if (blankCount >= dataRows.length * 0.7) {
        oScore += 20;
        dScore += 25;
      }

      if (bScore > maxBarcodeScore) {
        maxBarcodeScore = bScore;
        bestBarcodeCol = col.index;
      }
      if (oScore > maxOrigScore && col.index !== bestBarcodeCol) {
        maxOrigScore = oScore;
        bestOrigPriceCol = col.index;
      }
      if (dScore > maxDiscScore && col.index !== bestBarcodeCol && col.index !== bestOrigPriceCol) {
        maxDiscScore = dScore;
        bestDiscPriceCol = col.index;
      }
    });

    if (bestBarcodeCol) setTemplateBarcodeCol(bestBarcodeCol);

    if (bestOrigPriceCol && bestOrigPriceCol !== bestBarcodeCol) {
      setOriginalPriceCol(bestOrigPriceCol);
      setSinglePriceCol(bestOrigPriceCol);
      setLowestOriginalPriceCol(bestOrigPriceCol);
    } else {
      const fallback = cols.find((c) => c.index !== bestBarcodeCol);
      if (fallback) {
        setOriginalPriceCol(fallback.index);
        setSinglePriceCol(fallback.index);
        setLowestOriginalPriceCol(fallback.index);
      }
    }

    if (bestDiscPriceCol && bestDiscPriceCol !== bestBarcodeCol && bestDiscPriceCol !== bestOrigPriceCol) {
      setDiscountPriceCol(bestDiscPriceCol);
      setLowestDiscountPriceCol(bestDiscPriceCol);
    } else {
      const fallbackDisc = cols.find((c) => c.index !== bestBarcodeCol && c.index !== bestOrigPriceCol);
      if (fallbackDisc) {
        setDiscountPriceCol(fallbackDisc.index);
        setLowestDiscountPriceCol(fallbackDisc.index);
      }
    }
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setTemplateFile(file);

    const reader = new FileReader();
    reader.onload = (event) => {
      const arrayBuffer = event.target?.result as ArrayBuffer;
      const bytes = new Uint8Array(arrayBuffer);
      let binary = '';
      for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      const b64 = btoa(binary);
      setTemplateBase64(b64);

      const wb = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetNames = wb.SheetNames;
      setTemplateSheets(sheetNames);
      const first = sheetNames[0] || '';
      setSelectedSheet(first);
      const ws = wb.Sheets[first];
      const rawData = ws ? (XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][]) : [];
      parseSheetStructure(rawData);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleSheetChange = (sheetName: string) => {
    setSelectedSheet(sheetName);
    if (!templateBase64) return;
    try {
      const binary = atob(templateBase64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      const wb = XLSX.read(bytes.buffer, { type: 'array' });
      const ws = wb.Sheets[sheetName];
      const rawData = ws ? (XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][]) : [];
      parseSheetStructure(rawData);
    } catch {}
  };

  const handleExecute = async () => {
    if (!templateBase64) {
      alert('请先上传或确认您的 Excel 模板文件！');
      return;
    }

    if (columnTargetMode === 'existing_columns') {
      if (fillMode === 'original_and_discount') {
        if (!originalPriceCol) {
          alert('请选择填入原价的目标列！');
          return;
        }
        if (insertDiscountCol && !discountPriceCol) {
          alert('您已开启插入打折价列，请选择填入打折价的目标列！');
          return;
        }
        if (insertDiscountCol && originalPriceCol === discountPriceCol) {
          alert('原价列和打折价列不能选择同一个列！请分别指定不同的目标列。');
          return;
        }
      } else if (fillMode === 'single') {
        if (!singlePriceCol) {
          alert('请选择要填入价格的目标列！');
          return;
        }
      }
    }

    setIsProcessing(true);
    setStatusText('正在按配置将价格精准注入您的 Excel 表格...');

    try {
      const res = await fetch('/api/prices/fill-template', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          templateBase64,
          filename: templateFile?.name || 'custom_prices.xlsx',
          sheetName: selectedSheet,
          applyToAllSheets,
          autoDetectBarcodeCol,

          // Column Target Mode
          columnTargetMode,
          newOriginalColHeader,
          newDiscountColHeader,
          insertPlatformCol,
          newPlatformColHeader,

          barcodeColIndex: templateBarcodeCol,
          headerRowIndex: templateHeaderRowIndex,
          startRow: templateStartDataRowIndex,
          fillMode,

          // Mode 1: Original + Optional Discount
          originalPriceType,
          originalPriceColIndex: originalPriceCol,
          insertDiscountCol,
          discountPriceType,
          discountPriceColIndex: discountPriceCol,
          discountNonSaleHandling,

          // Mode 2: Single Column
          singlePriceType,
          singlePriceColIndex: singlePriceCol,

          // Mode 3: Multi Column
          columnMappings: {
            lowestOriginalPriceCol: lowestOriginalPriceCol > 0 ? lowestOriginalPriceCol : undefined,
            lowestDiscountPriceCol: lowestDiscountPriceCol > 0 ? lowestDiscountPriceCol : undefined,
            lowestEffectivePriceCol: lowestEffectivePriceCol > 0 ? lowestEffectivePriceCol : undefined,
            averagePriceCol: averagePriceCol > 0 ? averagePriceCol : undefined,
            meloniOriginalPriceCol: meloniOriginalPriceCol > 0 ? meloniOriginalPriceCol : undefined,
            meloniDiscountPriceCol: meloniDiscountPriceCol > 0 ? meloniDiscountPriceCol : undefined,
            megaOriginalPriceCol: megaOriginalPriceCol > 0 ? megaOriginalPriceCol : undefined,
            megaDiscountPriceCol: megaDiscountPriceCol > 0 ? megaDiscountPriceCol : undefined,
            maurysOriginalPriceCol: maurysOriginalPriceCol > 0 ? maurysOriginalPriceCol : undefined,
            maurysDiscountPriceCol: maurysDiscountPriceCol > 0 ? maurysDiscountPriceCol : undefined,
            risparmioOriginalPriceCol: risparmioOriginalPriceCol > 0 ? risparmioOriginalPriceCol : undefined,
            risparmioDiscountPriceCol: risparmioDiscountPriceCol > 0 ? risparmioDiscountPriceCol : undefined,
            platformNameCol: platformNameCol > 0 ? platformNameCol : undefined,
            titleCol: titleCol > 0 ? titleCol : undefined,
          },
          valueFormat,
          unmatchedFill,
          items,
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || '回填失败');
      }

      const filledCount = res.headers.get('X-Filled-Count') || '0';
      const processedSheets = res.headers.get('X-Processed-Sheets') || '1';
      const blob = await res.blob();
      const blobUrl = window.URL.createObjectURL(blob);
      const safeName = templateFile?.name.replace(/\.xlsx$/i, '_已插入价格.xlsx') || 'prices_filled.xlsx';

      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = safeName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(blobUrl);

      const msg = `🎉 成功生成文件：【${safeName}】！\n已处理 ${processedSheets} 个工作表，共精准插入 ${filledCount} 条商品价格！\n` +
        (columnTargetMode === 'new_columns'
          ? '✓ 采用新建追加列模式，100% 完整保留了原表格全部列与历史内容！'
          : '✓ 已按指定列回填原有表格。');

      setStatusText(msg);
      alert(msg);
      onClose();
    } catch (err: any) {
      alert(`回填失败: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-slate-900 border border-slate-700/80 rounded-2xl w-full max-w-3xl shadow-2xl overflow-hidden flex flex-col max-h-[92vh] animate-in fade-in zoom-in-95 duration-200">
        {/* Header */}
        <div className="bg-slate-950 px-5 py-4 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-amber-500/20 border border-amber-500/30 flex items-center justify-center text-amber-400">
              <Coins className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-extrabold text-white flex items-center gap-2">
                <span>套用 Excel 模板插入价格并导出</span>
                <span className="text-[11px] font-mono px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                  支持新建列 · 支持多工作表
                </span>
              </h3>
              <p className="text-xs text-slate-400 mt-0.5">
                支持在原表格末尾新建列（避免覆盖任何原有内容），同时支持一次性为多个不同品牌/分类的工作表自动插入价格。
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Body */}
        <div className="p-5 space-y-4 overflow-y-auto font-sans text-xs">
          {/* Step 1: Upload / Verify Template */}
          <div className="bg-slate-950/80 border border-slate-800 rounded-xl p-4 space-y-3">
            <div className="flex items-center justify-between">
              <span className="font-bold text-xs text-slate-200 flex items-center gap-1.5">
                <Upload className="w-3.5 h-3.5 text-amber-400" />
                <span>步骤一：确认目标 Excel 模板文件 (.xlsx)</span>
              </span>
              {templateFile && (
                <span className="text-[11px] text-emerald-400 bg-emerald-950/60 border border-emerald-800/40 px-2 py-0.5 rounded-full font-mono">
                  ✓ 已载入: {templateFile.name} ({templateSheets.length} 个工作表)
                </span>
              )}
            </div>

            <div
              onClick={() => fileInputRef.current?.click()}
              className="border-2 border-dashed border-slate-700 hover:border-amber-500 bg-slate-900/60 hover:bg-slate-900/90 rounded-xl p-4 text-center cursor-pointer transition flex flex-col items-center justify-center gap-1.5"
            >
              <Upload className="w-6 h-6 text-slate-400 hover:text-amber-400" />
              <div className="text-xs font-semibold text-slate-300">
                {templateFile ? `当前文件：${templateFile.name}（点击可更换）` : '点击选择或拖拽您的 Excel 模板文件至此处 (.xlsx)'}
              </div>
              <div className="text-[11px] text-slate-500">
                支持包含多个工作表（不同品牌/分类）的表格，系统将保留全部排版与格式
              </div>
              <input
                type="file"
                ref={fileInputRef}
                onChange={handleFileSelect}
                accept=".xlsx"
                className="hidden"
              />
            </div>

            {/* Multi-Sheet Detection & Range Selector */}
            {templateSheets.length > 0 && (
              <div className="bg-slate-900/90 border border-slate-800 p-3 rounded-lg space-y-2">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <FolderTree className="w-4 h-4 text-cyan-400" />
                    <span className="font-bold text-slate-200 text-xs">检测到表格包含 {templateSheets.length} 个工作表：</span>
                  </div>

                  {/* Multi-Sheet Toggle */}
                  <label className="flex items-center gap-1.5 cursor-pointer text-xs select-none">
                    <input
                      type="checkbox"
                      checked={applyToAllSheets}
                      onChange={(e) => setApplyToAllSheets(e.target.checked)}
                      className="rounded border-slate-700 bg-slate-950 text-amber-500 focus:ring-0 w-4 h-4 cursor-pointer"
                    />
                    <span className={`font-bold ${applyToAllSheets ? 'text-amber-300' : 'text-slate-400'}`}>
                      🌟 同时为所有工作表 ({templateSheets.length}个) 插入价格 (强烈推荐)
                    </span>
                  </label>
                </div>

                {/* Sheets pills */}
                <div className="flex flex-wrap gap-1.5 pt-1">
                  {templateSheets.map((s) => {
                    const isCurrent = s === selectedSheet;
                    return (
                      <button
                        key={s}
                        type="button"
                        onClick={() => handleSheetChange(s)}
                        className={`text-[11px] px-2.5 py-1 rounded-lg border font-mono transition cursor-pointer flex items-center gap-1.5 ${
                          isCurrent
                            ? 'bg-amber-500/20 text-amber-300 border-amber-500/50 font-bold'
                            : 'bg-slate-950 text-slate-400 border-slate-800 hover:border-slate-700 hover:text-slate-200'
                        }`}
                      >
                        <span>{s}</span>
                        {applyToAllSheets && (
                          <span className="w-1.5 h-1.5 rounded-full bg-emerald-400" title="该表将被插入价格" />
                        )}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}
          </div>

          {/* Step 2: Column Insertion Mode (NEW vs EXISTING) */}
          {templateFile && (
            <div className="bg-slate-950/80 border border-slate-800 rounded-xl p-4 space-y-3.5 animate-in fade-in">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2.5">
                <span className="font-bold text-xs text-slate-200 flex items-center gap-1.5">
                  <Tag className="w-3.5 h-3.5 text-amber-400" />
                  <span>步骤二：选择价格写入方式（新建列 vs 覆盖现有列）</span>
                </span>
                <span className="text-[11px] text-slate-400">
                  原需求：可以新建列，避免覆盖原来的内容
                </span>
              </div>

              {/* Column Target Mode Switcher */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div
                  onClick={() => setColumnTargetMode('new_columns')}
                  className={`p-3.5 rounded-xl border transition-all cursor-pointer space-y-1.5 ${
                    columnTargetMode === 'new_columns'
                      ? 'bg-amber-950/30 border-amber-500/70 shadow-lg text-white'
                      : 'bg-slate-900/60 border-slate-800 text-slate-400 hover:border-slate-700'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-xs text-amber-300 flex items-center gap-1.5">
                      <PlusCircle className="w-4 h-4 text-amber-400" />
                      <span>新建列模式 (自动追加在末尾)</span>
                    </span>
                    <span className="text-[10px] bg-emerald-500/20 text-emerald-300 px-2 py-0.5 rounded-full border border-emerald-500/40 font-bold">
                      推荐 · 0风险
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-300 leading-relaxed">
                    在原表格最后一列后自动创建全新的价格列，<strong>100% 绝不覆盖原表格的任何列与任何内容</strong>，自动继承表头样式。
                  </p>
                </div>

                <div
                  onClick={() => setColumnTargetMode('existing_columns')}
                  className={`p-3.5 rounded-xl border transition-all cursor-pointer space-y-1.5 ${
                    columnTargetMode === 'existing_columns'
                      ? 'bg-amber-950/30 border-amber-500/70 shadow-lg text-white'
                      : 'bg-slate-900/60 border-slate-800 text-slate-400 hover:border-slate-700'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-xs text-slate-200 flex items-center gap-1.5">
                      <Edit3 className="w-4 h-4 text-cyan-400" />
                      <span>覆盖写入已有指定列</span>
                    </span>
                    <span className="text-[10px] bg-slate-800 text-slate-400 px-2 py-0.5 rounded-full">
                      已有空白列时用
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-400 leading-relaxed">
                    将提取的价格写入原表格中您已提前留出的某一空白列（例如已有的 C列 单价列）。
                  </p>
                </div>
              </div>

              {/* Mode A: New Columns Settings */}
              {columnTargetMode === 'new_columns' && (
                <div className="bg-slate-900/90 border border-slate-800 p-4 rounded-xl space-y-3.5">
                  <div className="font-bold text-xs text-amber-300 flex items-center gap-1.5">
                    <SlidersHorizontal className="w-3.5 h-3.5" />
                    <span>新建追加列具体配置：</span>
                  </div>

                  {/* 1. New Original Price Column */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 bg-slate-950/80 p-3 rounded-lg border border-slate-800">
                    <div className="space-y-1">
                      <label className="text-[11px] text-slate-300 font-semibold block">原价数据来源：</label>
                      <select
                        value={originalPriceType}
                        onChange={(e) => setOriginalPriceType(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 cursor-pointer"
                      >
                        <option value="lowest_original">🌟 全网综合最低原价 (推荐，不含促销)</option>
                        <option value="lowest_effective">🎯 全网最低有效到手价 (含打折)</option>
                        <option value="meloni">🟣 Meloni 原价</option>
                        <option value="megacedi">🟢 MegaCedi 批发底价</option>
                        <option value="maurys">🔵 Maury's 原价</option>
                        <option value="risparmiocasa">🟡 Risparmio Casa 原价</option>
                        <option value="piume">🌸 PiùMe 原价</option>
                      </select>
                    </div>

                    <div className="space-y-1">
                      <label className="text-[11px] text-blue-300 font-bold block">新建原价列标题名称：</label>
                      <input
                        type="text"
                        value={newOriginalColHeader}
                        onChange={(e) => setNewOriginalColHeader(e.target.value)}
                        placeholder="例如: 提取-最低原价 (€)"
                        className="w-full bg-slate-900 border border-blue-500/50 rounded-lg px-2.5 py-1.5 text-xs text-blue-200 font-medium focus:outline-none focus:border-blue-400"
                      />
                    </div>
                  </div>

                  {/* 2. Optional Discount Price Column */}
                  <div
                    className={`border rounded-lg p-3 space-y-2.5 transition-all ${
                      insertDiscountCol
                        ? 'bg-rose-950/20 border-rose-500/50'
                        : 'bg-slate-950/50 border-slate-800 hover:border-slate-700'
                    }`}
                  >
                    <div className="flex items-center justify-between">
                      <label
                        onClick={() => setInsertDiscountCol(!insertDiscountCol)}
                        className="flex items-center gap-2 cursor-pointer select-none"
                      >
                        <input
                          type="checkbox"
                          checked={insertDiscountCol}
                          onChange={() => {}}
                          className="rounded border-slate-700 bg-slate-900 text-rose-500 focus:ring-0 w-4 h-4 cursor-pointer"
                        />
                        <span className={`text-xs font-bold ${insertDiscountCol ? 'text-rose-300' : 'text-slate-300'}`}>
                          需要时才选择额外新建一列打折价 (活动促销特价列)
                        </span>
                      </label>

                      <button
                        type="button"
                        onClick={() => setInsertDiscountCol(!insertDiscountCol)}
                        className={`text-[11px] px-2.5 py-0.5 rounded-full border transition cursor-pointer flex items-center gap-1 ${
                          insertDiscountCol
                            ? 'bg-rose-500/20 text-rose-300 border-rose-500/40 font-bold'
                            : 'bg-slate-900 text-slate-400 border-slate-800'
                        }`}
                      >
                        <Percent className="w-3 h-3" />
                        <span>{insertDiscountCol ? '已启用新建打折价列' : '未开启 (默认只建原价列)'}</span>
                      </button>
                    </div>

                    {insertDiscountCol && (
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1 border-t border-rose-900/40 animate-in fade-in">
                        <div className="space-y-1">
                          <label className="text-[11px] text-slate-300">选用哪种打折价：</label>
                          <select
                            value={discountPriceType}
                            onChange={(e) => setDiscountPriceType(e.target.value)}
                            className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 cursor-pointer"
                          >
                            <option value="lowest_discount">🏷️ 全网最低折后特价 (推荐)</option>
                            <option value="meloni_discount">🟣 Meloni 折后特价</option>
                            <option value="megacedi_discount">🟢 MegaCedi 折后特价</option>
                            <option value="maurys_discount">🔵 Maury's 折后特价</option>
                            <option value="risparmiocasa_discount">🟡 Risparmio Casa 折后特价</option>
                            <option value="piume_discount">🌸 PiùMe 折后特价</option>
                          </select>
                        </div>

                        <div className="space-y-1">
                          <label className="text-[11px] text-rose-300 font-bold">新建打折价列标题：</label>
                          <input
                            type="text"
                            value={newDiscountColHeader}
                            onChange={(e) => setNewDiscountColHeader(e.target.value)}
                            placeholder="例如: 提取-折后特价 (€)"
                            className="w-full bg-slate-900 border border-rose-500/50 rounded-lg px-2.5 py-1.5 text-xs text-rose-200 font-medium focus:outline-none focus:border-rose-400"
                          />
                        </div>

                        <div className="space-y-1">
                          <label className="text-[11px] text-slate-300">无促销活动时：</label>
                          <select
                            value={discountNonSaleHandling}
                            onChange={(e) => setDiscountNonSaleHandling(e.target.value as any)}
                            className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 cursor-pointer"
                          >
                            <option value="empty">留空 (保持空白，推荐)</option>
                            <option value="original">填入原价 (无折扣时填原价)</option>
                            <option value="dash">填入破折号 (-)</option>
                          </select>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* 3. Optional Platform Column */}
                  <div className="flex items-center justify-between bg-slate-950/60 p-2.5 rounded-lg border border-slate-800">
                    <label
                      onClick={() => setInsertPlatformCol(!insertPlatformCol)}
                      className="flex items-center gap-2 cursor-pointer select-none text-xs text-slate-300"
                    >
                      <input
                        type="checkbox"
                        checked={insertPlatformCol}
                        onChange={() => {}}
                        className="rounded border-slate-700 bg-slate-900 text-teal-500 focus:ring-0 w-4 h-4 cursor-pointer"
                      />
                      <span>同时新建一列「最低价渠道/平台来源」</span>
                    </label>

                    {insertPlatformCol && (
                      <input
                        type="text"
                        value={newPlatformColHeader}
                        onChange={(e) => setNewPlatformColHeader(e.target.value)}
                        className="bg-slate-900 border border-teal-500/50 rounded px-2 py-0.5 text-xs text-teal-300 w-36 font-medium"
                      />
                    )}
                  </div>
                </div>
              )}

              {/* Mode B: Existing Columns Mode */}
              {columnTargetMode === 'existing_columns' && (
                <div className="bg-slate-900/90 border border-slate-800 p-4 rounded-xl space-y-3">
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-200">
                      选择原价填入的目标列 (已存在于当前表格)：
                    </label>
                    <select
                      value={originalPriceCol}
                      onChange={(e) => setOriginalPriceCol(Number(e.target.value))}
                      className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 cursor-pointer font-mono"
                    >
                      <option value={0}>[请选择目标列]</option>
                      {templateColumns.map((col) => (
                        <option key={col.index} value={col.index}>
                          {col.letter}列 - {col.name} 【样例: {col.sampleText}】
                        </option>
                      ))}
                    </select>
                  </div>

                  <div className="flex items-center justify-between pt-1">
                    <label
                      onClick={() => setInsertDiscountCol(!insertDiscountCol)}
                      className="flex items-center gap-2 cursor-pointer select-none text-xs text-slate-300"
                    >
                      <input
                        type="checkbox"
                        checked={insertDiscountCol}
                        onChange={() => {}}
                        className="rounded border-slate-700 bg-slate-900 text-rose-500 focus:ring-0 w-4 h-4 cursor-pointer"
                      />
                      <span>需要时才选择插入一列打折价 (已有列)</span>
                    </label>
                  </div>

                  {insertDiscountCol && (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                      <div className="space-y-1">
                        <label className="text-[11px] text-rose-300 font-bold">打折价填入目标列：</label>
                        <select
                          value={discountPriceCol}
                          onChange={(e) => setDiscountPriceCol(Number(e.target.value))}
                          className="w-full bg-slate-950 border border-rose-600 rounded-lg px-2.5 py-1.5 text-xs text-rose-200 font-mono cursor-pointer"
                        >
                          <option value={0}>[请选择目标列]</option>
                          {templateColumns.map((col) => (
                            <option key={col.index} value={col.index}>
                              {col.letter}列 - {col.name} 【样例: {col.sampleText}】
                            </option>
                          ))}
                        </select>
                      </div>

                      <div className="space-y-1">
                        <label className="text-[11px] text-slate-300">无促销活动时：</label>
                        <select
                          value={discountNonSaleHandling}
                          onChange={(e) => setDiscountNonSaleHandling(e.target.value as any)}
                          className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 cursor-pointer"
                        >
                          <option value="empty">留空 (保持空白)</option>
                          <option value="original">填入原价</option>
                          <option value="dash">填入破折号 (-)</option>
                        </select>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* Barcode Column Detection Config */}
              <div className="bg-amber-950/20 border border-amber-600/40 p-3 rounded-lg space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-xs text-amber-300 flex items-center gap-1.5">
                    <CheckCircle2 className="w-3.5 h-3.5 text-amber-400" />
                    <span>条形码匹配列规则：</span>
                  </span>
                  <label className="flex items-center gap-1.5 cursor-pointer text-[11px] text-amber-200 select-none">
                    <input
                      type="checkbox"
                      checked={autoDetectBarcodeCol}
                      onChange={(e) => setAutoDetectBarcodeCol(e.target.checked)}
                      className="rounded border-amber-700 bg-slate-950 text-amber-500 focus:ring-0 w-3.5 h-3.5 cursor-pointer"
                    />
                    <span>自动识别每个工作表的条码列 (适应多表不同列位置，推荐)</span>
                  </label>
                </div>

                {!autoDetectBarcodeCol && (
                  <div className="space-y-1 pt-1 animate-in fade-in">
                    <label className="text-[11px] text-slate-300">强制指定当前工作表条码列：</label>
                    <select
                      value={templateBarcodeCol}
                      onChange={(e) => setTemplateBarcodeCol(Number(e.target.value))}
                      className="w-full bg-slate-950 border border-amber-700/60 rounded-lg px-2.5 py-1.5 text-xs text-amber-200 focus:outline-none font-mono cursor-pointer"
                    >
                      {templateColumns.map((col) => (
                        <option key={col.index} value={col.index}>
                          {col.letter}列 - {col.name} 【样例: {col.sampleText}】
                        </option>
                      ))}
                    </select>
                  </div>
                )}
              </div>

              {/* Format & Fallbacks */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 bg-slate-900/60 p-3 rounded-lg border border-slate-800">
                <div className="space-y-1">
                  <label className="text-[11px] font-bold text-slate-300 block">写入数值格式：</label>
                  <select
                    value={valueFormat}
                    onChange={(e) => setValueFormat(e.target.value as any)}
                    className="w-full bg-slate-950 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 cursor-pointer"
                  >
                    <option value="number">纯数字保留两位小数 (12.50，方便后续求和/公式)</option>
                    <option value="currency">欧元货币格式 (€12.50)</option>
                    <option value="text">纯文本 (€12.50 文本)</option>
                  </select>
                </div>

                <div className="space-y-1">
                  <label className="text-[11px] font-bold text-slate-300 block">未匹配到条码时的单元格：</label>
                  <select
                    value={unmatchedFill}
                    onChange={(e) => setUnmatchedFill(e.target.value as any)}
                    className="w-full bg-slate-950 border border-slate-700 rounded px-2 py-1 text-xs text-slate-200 cursor-pointer"
                  >
                    <option value="empty">保持留空 (不覆盖原内容)</option>
                    <option value="zero">填入 0.00</option>
                    <option value="dash">填入横线 (-)</option>
                  </select>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="bg-slate-950 px-5 py-3.5 border-t border-slate-800 flex items-center justify-between">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-400 hover:text-white hover:bg-slate-800 transition cursor-pointer"
          >
            取消
          </button>

          <button
            type="button"
            onClick={handleExecute}
            disabled={!templateFile || isProcessing}
            className="bg-gradient-to-r from-amber-500 via-amber-600 to-amber-700 hover:from-amber-400 hover:to-amber-600 disabled:opacity-40 disabled:cursor-not-allowed text-slate-950 font-black text-xs px-6 py-2.5 rounded-xl shadow-xl shadow-amber-500/20 transition flex items-center gap-2 cursor-pointer"
          >
            {isProcessing ? (
              <>
                <div className="w-4 h-4 border-2 border-slate-950 border-t-transparent rounded-full animate-spin" />
                <span>正在执行插入并导出...</span>
              </>
            ) : (
              <>
                <FileSpreadsheet className="w-4 h-4 fill-slate-950" />
                <span>
                  {columnTargetMode === 'new_columns'
                    ? `新建列导出 (${applyToAllSheets ? `全部${templateSheets.length}个工作表` : '当前工作表'})`
                    : `按指定列覆盖回填导出`}
                </span>
              </>
            )}
          </button>
        </div>
      </div>
    </div>
  );
};
