import React, { useState, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  Coins,
  Store,
  Play,
  Square,
  FileSpreadsheet,
  Upload,
  RefreshCw,
  Search,
  Filter,
  CheckCircle2,
  AlertCircle,
  TrendingDown,
  TrendingUp,
  ExternalLink,
  Trash2,
  Download,
  Layers,
  Sparkles,
  Info,
  Copy,
  Check,
  Percent,
  CheckSquare,
  SquareDashed,
  Tag,
  Eye,
  SlidersHorizontal,
  Flame,
  FolderTree,
  KeyRound,
  ShieldCheck,
  HelpCircle
} from 'lucide-react';
import { BarcodePriceItem, PriceExtractorStats, PriceSourceSite, SheetBarcodeSummary } from '../types';
import { SAMPLE_BARCODES_FROM_USER } from './BarcodeMatcher';
import { PriceTemplateModal } from './PriceTemplateModal';
import { PlatformAuthModal } from './PlatformAuthModal';

interface PlatformOption {
  id: PriceSourceSite;
  name: string;
  domain: string;
  description: string;
  loginBadge: string;
  isFreeLogin: boolean;
  colorClass: string;
  badgeClass: string;
}

const AVAILABLE_PLATFORMS: PlatformOption[] = [
  {
    id: 'maurys',
    name: "Maury's Online",
    domain: 'maurysonline.it',
    description: '意大利连锁百货大卖场，标价与促销折扣公开',
    loginBadge: '🟢 100% 免登录看价格',
    isFreeLogin: true,
    colorClass: 'border-cyan-500/50 bg-cyan-950/20 text-cyan-300',
    badgeClass: 'bg-cyan-500/20 text-cyan-300 border-cyan-500/40',
  },
  {
    id: 'risparmiocasa',
    name: 'Risparmio Casa',
    domain: 'shop.risparmiocasa.com',
    description: '意大利国民家清个护连锁，官方直营价与特价公开',
    loginBadge: '🟢 100% 免登录看价格',
    isFreeLogin: true,
    colorClass: 'border-amber-500/50 bg-amber-950/20 text-amber-300',
    badgeClass: 'bg-amber-500/20 text-amber-300 border-amber-500/40',
  },
  {
    id: 'piume',
    name: 'PiùMe',
    domain: 'piume.it',
    description: '意大利知名美妆日化清洁大型连锁，数万商品官方售价公开',
    loginBadge: '🟢 100% 免登录看价格',
    isFreeLogin: true,
    colorClass: 'border-rose-500/50 bg-rose-950/20 text-rose-300',
    badgeClass: 'bg-rose-500/20 text-rose-300 border-rose-500/40',
  },
  {
    id: 'meloni',
    name: 'Meloni Store',
    domain: 'melonistore.com',
    description: '日化百货B2B供货商，常规活动品公开，专供批发底价支持配置账号/Cookie',
    loginBadge: '🟡 B2B底价 (可配账号/Cookie)',
    isFreeLogin: false,
    colorClass: 'border-purple-500/50 bg-purple-950/20 text-purple-300',
    badgeClass: 'bg-purple-500/20 text-purple-300 border-purple-500/40',
  },
  {
    id: 'megacedi',
    name: 'MegaCedi',
    domain: 'megacedi.com',
    description: 'B2B 批发仓储体系，支持箱规装箱数，批发底价可配Cookie',
    loginBadge: '🟡 B2B底价 (可配Cookie)',
    isFreeLogin: false,
    colorClass: 'border-emerald-500/50 bg-emerald-950/20 text-emerald-300',
    badgeClass: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40',
  },
];

export const PriceExtractor: React.FC = () => {
  const [inputText, setInputText] = useState<string>('');
  const [selectedSites, setSelectedSites] = useState<PriceSourceSite[]>([
    'meloni',
    'megacedi',
    'maurys',
    'risparmiocasa',
    'piume',
  ]);
  const [results, setResults] = useState<BarcodePriceItem[]>([]);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [statusText, setStatusText] = useState<string>('');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [filterMode, setFilterMode] = useState<'all' | 'matched' | 'discounted' | 'regular' | 'spread' | 'unmatched'>('all');
  const [priceViewMode, setPriceViewMode] = useState<'both' | 'original_only' | 'discount_only'>('both');
  const [isTemplateModalOpen, setIsTemplateModalOpen] = useState<boolean>(false);
  const [isLoginModalOpen, setIsLoginModalOpen] = useState<boolean>(false);
  const [authModalPlatform, setAuthModalPlatform] = useState<PriceSourceSite>('meloni');
  const [copiedKey, setCopiedKey] = useState<string | null>(null);

  const openAuthModal = (plat: PriceSourceSite) => {
    setAuthModalPlatform(plat);
    setIsLoginModalOpen(true);
  };

  // Multi-sheet workbook persistence for template export
  const [uploadedWorkbookFile, setUploadedWorkbookFile] = useState<File | null>(null);
  const [uploadedWorkbookBase64, setUploadedWorkbookBase64] = useState<string>('');
  const [sheetSummaries, setSheetSummaries] = useState<SheetBarcodeSummary[]>([]);

  const [stats, setStats] = useState<PriceExtractorStats>({
    total: 0,
    processed: 0,
    matchedCount: 0,
    unmatchedCount: 0,
    meloniFound: 0,
    megaFound: 0,
    maurysFound: 0,
    risparmioFound: 0,
    piumeFound: 0,
  });

  const abortControllerRef = useRef<AbortController | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Restore cache from localStorage on mount
  useEffect(() => {
    try {
      const cached = localStorage.getItem('price_extractor_results_cache');
      const cachedStats = localStorage.getItem('price_extractor_stats_cache');
      if (cached) {
        const parsed = JSON.parse(cached);
        if (Array.isArray(parsed) && parsed.length > 0) {
          setResults(parsed);
          setStatusText(`✓ 已从本地浏览器记忆中恢复上次比价的 ${parsed.length} 条商品价格数据`);
        }
      }
      if (cachedStats) {
        const parsedStats = JSON.parse(cachedStats);
        if (parsedStats && typeof parsedStats === 'object') {
          setStats(parsedStats);
        }
      }
    } catch {}
  }, []);

  // Save to localStorage when updated
  useEffect(() => {
    if (results.length > 0) {
      try {
        localStorage.setItem('price_extractor_results_cache', JSON.stringify(results));
        localStorage.setItem('price_extractor_stats_cache', JSON.stringify(stats));
      } catch {}
    }
  }, [results, stats]);

  // Clean parsed barcodes from input
  const parsedBarcodes = useMemo(() => {
    const raw = inputText
      .split(/[\r\n,; \t]+/)
      .map((b) => b.replace(/[^0-9A-Za-z]/g, '').trim())
      .filter((b) => b.length >= 6);
    return Array.from(new Set(raw));
  }, [inputText]);

  // Toggle single platform
  const toggleSite = (site: PriceSourceSite) => {
    setSelectedSites((prev) => {
      if (prev.includes(site)) {
        if (prev.length === 1) {
          alert('请至少保留一个需要比价的网站！');
          return prev;
        }
        return prev.filter((s) => s !== site);
      } else {
        return [...prev, site];
      }
    });
  };

  const selectAllSites = () => {
    setSelectedSites(['meloni', 'megacedi', 'maurys', 'risparmiocasa', 'piume']);
  };

  const selectRetailOnly = () => {
    setSelectedSites(['maurys', 'risparmiocasa', 'piume']);
  };

  // Quick Seed Barcodes
  const handleLoadSample = () => {
    setInputText(SAMPLE_BARCODES_FROM_USER.join('\n'));
    setStatusText(`已一键载入 ${SAMPLE_BARCODES_FROM_USER.length} 个待提取商品条码！`);
  };

  // Handle Excel file upload (Supports multi-sheet workbooks for different brands/categories)
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const arrayBuffer = event.target?.result as ArrayBuffer;
        const bytes = new Uint8Array(arrayBuffer);
        let binary = '';
        for (let i = 0; i < bytes.byteLength; i++) {
          binary += String.fromCharCode(bytes[i]);
        }
        const b64 = btoa(binary);

        setUploadedWorkbookFile(file);
        setUploadedWorkbookBase64(b64);

        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetNames = workbook.SheetNames || [];
        const summaries: SheetBarcodeSummary[] = [];
        const allBarcodes: string[] = [];

        sheetNames.forEach((sName) => {
          const sheet = workbook.Sheets[sName];
          if (!sheet) return;
          const json = XLSX.utils.sheet_to_json<any>(sheet, { header: 1 });
          if (!json || json.length === 0) return;

          let barcodeCol = -1;
          const header = json[0];
          if (Array.isArray(header)) {
            for (let i = 0; i < header.length; i++) {
              const h = String(header[i] || '').toLowerCase().trim();
              if (h.includes('barcode') || h.includes('ean') || h.includes('条形码') || h.includes('条码') || h.includes('c.barre') || h.includes('codice')) {
                barcodeCol = i;
                break;
              }
            }
          }

          const startRow = barcodeCol !== -1 ? 1 : 0;
          const currentSheetBarcodes: string[] = [];

          for (let r = startRow; r < json.length; r++) {
            const row = json[r];
            if (!Array.isArray(row)) continue;
            if (barcodeCol !== -1) {
              const val = String(row[barcodeCol] || '').replace(/[^0-9A-Za-z]/g, '').trim();
              if (val.length >= 6) currentSheetBarcodes.push(val);
            } else {
              for (const cell of row) {
                const val = String(cell || '').replace(/[^0-9A-Za-z]/g, '').trim();
                if (val.length >= 8 && val.length <= 14 && /^\d+$/.test(val)) {
                  currentSheetBarcodes.push(val);
                  break;
                }
              }
            }
          }

          const uniqueThisSheet = Array.from(new Set(currentSheetBarcodes));
          if (uniqueThisSheet.length > 0) {
            summaries.push({
              sheetName: sName,
              count: uniqueThisSheet.length,
              barcodes: uniqueThisSheet,
            });
            allBarcodes.push(...uniqueThisSheet);
          }
        });

        setSheetSummaries(summaries);
        const uniqueAll = Array.from(new Set(allBarcodes));

        if (uniqueAll.length > 0) {
          setInputText(uniqueAll.join('\n'));
          const sheetInfo = summaries.length > 1
            ? `共 ${summaries.length} 个工作表（${summaries.map((s) => `${s.sheetName}: ${s.count}个`).slice(0, 4).join(', ')}${summaries.length > 4 ? '…' : ''}）`
            : `工作表「${sheetNames[0]}」`;
          setStatusText(`✓ 成功从「${file.name}」中读取到 ${sheetInfo}，汇总去重得到 ${uniqueAll.length} 个待提取商品条码！`);
        } else {
          setStatusText(`未能从「${file.name}」的任何工作表中识别到条码，请确保包含 EAN 或 8-14 位纯数字。`);
        }
      } catch (err: any) {
        setStatusText(`读取文件失败: ${err.message}`);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Start Streaming Price Extraction
  const startExtraction = async () => {
    if (parsedBarcodes.length === 0 || isProcessing) return;
    if (selectedSites.length === 0) {
      alert('请选择至少一个目标网站！');
      return;
    }

    setIsProcessing(true);
    setStatusText(`正在初始化多源价格提取引擎，目标 ${selectedSites.length} 个网站，共 ${parsedBarcodes.length} 个条码...`);
    setResults([]);
    setStats({
      total: parsedBarcodes.length,
      processed: 0,
      matchedCount: 0,
      unmatchedCount: 0,
      meloniFound: 0,
      megaFound: 0,
      maurysFound: 0,
      risparmioFound: 0,
      piumeFound: 0,
    });

    const controller = new AbortController();
    abortControllerRef.current = controller;

    try {
      const response = await fetch('/api/prices/stream', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          barcodes: parsedBarcodes,
          sites: selectedSites,
          delayMs: 150,
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const reader = response.body?.getReader();
      const decoder = new TextDecoder('utf-8');
      if (!reader) return;

      let buffer = '';
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const events = buffer.split('\n\n');
        buffer = events.pop() || '';

        for (const ev of events) {
          if (!ev.trim()) continue;
          const lines = ev.split('\n');
          let eventName = 'message';
          let dataStr = '';

          for (const line of lines) {
            if (line.startsWith('event:')) eventName = line.replace('event:', '').trim();
            else if (line.startsWith('data:')) dataStr += line.replace('data:', '').trim();
          }

          if (!dataStr) continue;
          try {
            const data = JSON.parse(dataStr);
            if (eventName === 'status') {
              setStatusText(data.message || '');
            } else if (eventName === 'item') {
              setResults((prev) => {
                const next = [...prev];
                const idx = next.findIndex((item) => item.barcode === data.barcode);
                if (idx !== -1) next[idx] = data;
                else next.push(data);
                return next;
              });
            } else if (eventName === 'progress') {
              setStats({
                total: data.total,
                processed: data.processed,
                matchedCount: data.matchedCount,
                unmatchedCount: data.unmatchedCount,
                meloniFound: data.meloniFound,
                megaFound: data.megaFound,
                maurysFound: data.maurysFound,
                risparmioFound: data.risparmioFound,
                piumeFound: data.piumeFound || 0,
              });
              setStatusText(`已提取 (${data.processed}/${data.total})，当前条码: ${data.currentBarcode}`);
            } else if (eventName === 'complete') {
              setStatusText(data.message || '全网比价提取完成！');
            }
          } catch {}
        }
      }
    } catch (err: any) {
      if (err.name !== 'AbortError') {
        setStatusText(`提取出错: ${err.message}`);
      } else {
        setStatusText('已停止本次价格提取任务');
      }
    } finally {
      setIsProcessing(false);
      abortControllerRef.current = null;
    }
  };

  const stopExtraction = () => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
      abortControllerRef.current = null;
    }
    setIsProcessing(false);
  };

  // Copy helper
  const copyText = (txt: string, key: string) => {
    navigator.clipboard.writeText(txt);
    setCopiedKey(key);
    setTimeout(() => setCopiedKey(null), 2000);
  };

  // Filtered Results
  const filteredResults = useMemo(() => {
    return results.filter((item) => {
      if (filterMode === 'matched' && !item.matched) return false;
      if (filterMode === 'discounted' && (!item.hasDiscount || !item.lowestDiscountPrice)) return false;
      if (filterMode === 'regular' && (!item.matched || item.hasDiscount)) return false;
      if (filterMode === 'spread' && (!item.priceDiff || item.priceDiff <= 0)) return false;
      if (filterMode === 'unmatched' && item.matched) return false;

      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase();
        return item.barcode.includes(q) || item.title.toLowerCase().includes(q);
      }
      return true;
    });
  }, [results, filterMode, searchTerm]);

  // Discount Count
  const discountedCount = useMemo(() => {
    return results.filter((r) => r.hasDiscount && r.lowestDiscountPrice !== null).length;
  }, [results]);

  // Export 1: Standalone Comparison Excel (.xlsx) with Original & Discount Columns
  const handleExportXlsx = async () => {
    if (results.length === 0) return;
    setStatusText('正在生成全平台价格对比 Excel 工作簿...');

    try {
      const res = await fetch('/api/prices/export-xlsx', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items: results }),
      });

      if (!res.ok) throw new Error('导出 Excel 失败');
      const blob = await res.blob();
      const blobUrl = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = `multisite_price_comparison_${results.length}items.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(blobUrl);
      setStatusText(`🎉 成功导出全平台比价 Excel 表格 (${results.length} 项)！包含最低原价与打折特价明细。`);
    } catch (err: any) {
      alert(`导出表格失败: ${err.message}`);
    }
  };

  // Export 2: Standard CSV with Original & Discount Columns
  const handleExportCsv = () => {
    if (results.length === 0) return;

    const headers = [
      '条形码 (Barcode)',
      '商品名称 (Title)',
      '综合最低原价 (€)',
      '综合折后特价 (€)',
      '促销状态',
      '最低价渠道',
      '综合最高价 (€)',
      '平台价差 (€)',
      '平均参考价 (€)',
      'Meloni 原价',
      'Meloni 特价',
      'MegaCedi 原价/底价',
      'MegaCedi 特价',
      "Maury's 原价",
      "Maury's 特价",
      'Risparmio Casa 原价',
      'Risparmio Casa 特价',
      'PiùMe 原价',
      'PiùMe 特价',
      '提取时间',
    ];

    const escapeCsv = (val: any) => {
      if (val === null || val === undefined) return '""';
      const str = String(val).replace(/"/g, '""');
      return `"${str}"`;
    };

    const rows = results.map((r) => [
      escapeCsv(r.barcode),
      escapeCsv(r.title),
      escapeCsv(r.lowestOriginalPrice != null ? r.lowestOriginalPrice.toFixed(2) : (r.lowestPrice != null ? r.lowestPrice.toFixed(2) : '')),
      escapeCsv(r.lowestDiscountPrice != null ? r.lowestDiscountPrice.toFixed(2) : ''),
      escapeCsv(r.hasDiscount ? '打折促销中' : '正价'),
      escapeCsv(r.lowestPlatform || '-'),
      escapeCsv(r.highestPrice != null ? r.highestPrice.toFixed(2) : ''),
      escapeCsv(r.priceDiff != null ? r.priceDiff.toFixed(2) : ''),
      escapeCsv(r.averagePrice != null ? r.averagePrice.toFixed(2) : ''),
      escapeCsv(r.prices.meloni?.formattedOriginalPrice || r.prices.meloni?.formattedPrice || '未收录'),
      escapeCsv(r.prices.meloni?.formattedDiscountPrice || '-'),
      escapeCsv(r.prices.megacedi?.formattedOriginalPrice || r.prices.megacedi?.formattedPrice || '未收录'),
      escapeCsv(r.prices.megacedi?.formattedDiscountPrice || '-'),
      escapeCsv(r.prices.maurys?.formattedOriginalPrice || r.prices.maurys?.formattedPrice || '未收录'),
      escapeCsv(r.prices.maurys?.formattedDiscountPrice || '-'),
      escapeCsv(r.prices.risparmiocasa?.formattedOriginalPrice || r.prices.risparmiocasa?.formattedPrice || '未收录'),
      escapeCsv(r.prices.risparmiocasa?.formattedDiscountPrice || '-'),
      escapeCsv(r.prices.piume?.formattedOriginalPrice || r.prices.piume?.formattedPrice || '未收录'),
      escapeCsv(r.prices.piume?.formattedDiscountPrice || '-'),
      escapeCsv(r.extractedAt),
    ]);

    const csvContent = '\uFEFF' + [headers.join(','), ...rows.map((r) => r.join(','))].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `price_comparison_with_discount_${results.length}items.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const progressPercent = stats.total > 0 ? Math.min(100, Math.round((stats.processed / stats.total) * 100)) : 0;

  return (
    <div className="space-y-6">
      {/* Hero Banner for Price Extractor */}
      <div className="relative overflow-hidden rounded-2xl bg-gradient-to-br from-amber-950/40 via-slate-900 to-slate-950 border border-amber-500/30 p-6 shadow-2xl">
        <div className="absolute top-0 right-0 -mr-16 -mt-16 w-64 h-64 bg-amber-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-black uppercase tracking-wider bg-amber-500 text-slate-950">
                专属功能专区
              </span>
              <span className="text-xs font-mono text-amber-300 font-semibold flex items-center gap-1">
                <Tag className="w-3.5 h-3.5" />
                <span>区分原价与打折特价 · 严格套用自定义 Excel 模板</span>
              </span>
            </div>
            <h2 className="text-xl md:text-2xl font-black text-white tracking-tight flex items-center gap-2">
              <span>全网多平台价格提取与自定义模板回填</span>
              <Coins className="w-6 h-6 text-amber-400 animate-bounce" />
            </h2>
            <p className="text-xs text-slate-300 max-w-2xl leading-relaxed">
              专门开辟的价格提取核心模块：可一键并发提取 <strong>Meloni Store、MegaCedi、Maury's Online、Risparmio Casa</strong> 四大平台的商品售价；严格区分 <strong>正价原价</strong> 与 <strong>促销打折特价</strong>；支持平时只填原价，需要时才按需选择插入一列打折价，完美套用您原本的 Excel 模板。
            </p>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() => setIsTemplateModalOpen(true)}
              className="bg-gradient-to-r from-amber-500 via-amber-600 to-amber-700 hover:from-amber-400 hover:to-amber-600 text-slate-950 font-black text-xs px-4 py-2.5 rounded-xl shadow-lg shadow-amber-500/20 transition flex items-center gap-2 cursor-pointer group"
            >
              <FileSpreadsheet className="w-4 h-4 fill-slate-950 group-hover:scale-110 transition-transform" />
              <span>套用我的自定义 Excel 模板导出</span>
            </button>
          </div>
        </div>
      </div>

      {/* Main Work Area: Left Config (Sites) + Right Input (Barcodes) */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
        {/* Left: Platform Selection Area (5 Cols) */}
        <div className="lg:col-span-5 bg-slate-900/90 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
          <div className="flex items-center justify-between border-b border-slate-800 pb-3">
            <h3 className="font-extrabold text-sm text-white flex items-center gap-2">
              <Store className="w-4 h-4 text-amber-400" />
              <span>选定要提取价格的网站 ({selectedSites.length}/4)</span>
            </h3>
            <div className="flex items-center gap-1.5 text-[11px]">
              <button
                type="button"
                onClick={selectAllSites}
                className="text-cyan-400 hover:underline cursor-pointer"
              >
                全选
              </button>
              <span className="text-slate-600">·</span>
              <button
                type="button"
                onClick={selectRetailOnly}
                className="text-amber-400 hover:underline cursor-pointer"
              >
                仅零售
              </button>
            </div>
          </div>

          <div className="space-y-2.5">
            {AVAILABLE_PLATFORMS.map((site) => {
              const isSelected = selectedSites.includes(site.id);
              return (
                <div
                  key={site.id}
                  onClick={() => toggleSite(site.id)}
                  className={`p-3 rounded-xl border transition-all cursor-pointer flex items-center justify-between gap-3 ${
                    isSelected ? site.colorClass : 'border-slate-800 bg-slate-950/60 text-slate-400 hover:border-slate-700'
                  }`}
                >
                  <div className="flex items-center gap-2.5">
                    <input
                      type="checkbox"
                      checked={isSelected}
                      onChange={() => {}}
                      className="rounded border-slate-700 bg-slate-900 text-amber-500 focus:ring-0 cursor-pointer w-4 h-4"
                    />
                    <div>
                      <div className="flex flex-wrap items-center gap-1.5">
                        <span className="font-bold text-xs text-white">{site.name}</span>
                        <code className="text-[10px] text-slate-400 font-mono">{site.domain}</code>
                        <span className={`text-[9px] px-1.5 py-0.2 rounded font-sans ${
                          site.isFreeLogin ? 'bg-emerald-950/80 text-emerald-300 border border-emerald-800/50' : 'bg-slate-900 text-amber-300/80 border border-slate-700'
                        }`}>
                          {site.loginBadge}
                        </span>
                        {/* Login / Cookie button for Meloni and MegaCedi */}
                        {site.id === 'meloni' && (
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              openAuthModal('meloni');
                            }}
                            className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-purple-500/20 hover:bg-purple-500/30 text-purple-300 border border-purple-500/40 text-[10px] font-semibold transition cursor-pointer"
                            title="点击配置 Meloni Store 账号密码或 Cookie，解锁批发专享底价"
                          >
                            <KeyRound className="w-3 h-3 text-purple-400" />
                            <span>在此填写账号/Cookie</span>
                          </button>
                        )}
                        {site.id === 'megacedi' && (
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              openAuthModal('megacedi');
                            }}
                            className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-emerald-500/20 hover:bg-emerald-500/30 text-emerald-300 border border-emerald-500/40 text-[10px] font-semibold transition cursor-pointer"
                            title="点击配置 MegaCedi 账号密码或 Cookie，解锁商超采购底价"
                          >
                            <KeyRound className="w-3 h-3 text-emerald-400" />
                            <span>在此填写账号/Cookie</span>
                          </button>
                        )}
                      </div>
                      <p className="text-[10px] text-slate-400 mt-0.5">{site.description}</p>
                    </div>
                  </div>

                  <span
                    className={`text-[10px] font-mono px-2 py-0.5 rounded-full border shrink-0 ${
                      isSelected ? site.badgeClass : 'bg-slate-900 text-slate-500 border-slate-800'
                    }`}
                  >
                    {isSelected ? '已选' : '跳过'}
                  </span>
                </div>
              );
            })}
          </div>

          {/* Login transparency explanation */}
          <div className="bg-slate-950/80 p-3.5 rounded-xl border border-slate-800 text-[11px] text-slate-300 space-y-2">
            <div className="flex items-center justify-between text-amber-300 font-bold">
              <span className="flex items-center gap-1.5">
                <HelpCircle className="w-3.5 h-3.5 text-amber-400" />
                <span>价格提取免登录规则说明：</span>
              </span>
            </div>
            <div className="space-y-1.5 text-slate-400 leading-relaxed text-[10px]">
              <div className="flex items-start gap-1.5">
                <span className="text-emerald-400 font-bold shrink-0">✓ 100% 免登录：</span>
                <span><strong>Maury's Online</strong> 与 <strong>Risparmio Casa</strong> 是公开零售大卖场，全部商品原价与促销折扣<strong>完全公开透明</strong>，无需任何账号即可直接抓取！</span>
              </div>
              <div className="flex items-start gap-1.5">
                <span className="text-amber-400 font-bold shrink-0">⚠ 会员制底价：</span>
                <div className="space-y-1">
                  <span><strong>Meloni</strong> 与 <strong>MegaCedi</strong> 是意大利专业 B2B 批发仓储商，部分商品实行批发保护（网页提示 Accedi 需商户登录查看批发底价）。</span>
                  <div className="flex flex-wrap items-center gap-2 pt-0.5">
                    <button
                      type="button"
                      onClick={() => openAuthModal('meloni')}
                      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-lg bg-purple-950/80 border border-purple-500/50 text-purple-300 hover:text-white hover:bg-purple-900/80 text-[10px] font-bold cursor-pointer transition"
                    >
                      <KeyRound className="w-3 h-3 text-purple-400" />
                      <span>👉 点此填写 Meloni 账号/Cookie</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => openAuthModal('megacedi')}
                      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-lg bg-emerald-950/80 border border-emerald-500/50 text-emerald-300 hover:text-white hover:bg-emerald-900/80 text-[10px] font-bold cursor-pointer transition"
                    >
                      <KeyRound className="w-3 h-3 text-emerald-400" />
                      <span>👉 点此填写 MegaCedi 账号/Cookie</span>
                    </button>
                  </div>
                </div>
              </div>
              <div className="pt-1 text-slate-500">
                💡 <strong>最省心用法：</strong>平时若不想登录，点击上方【仅零售】即可一键只比对两家完全免登录的商城，实时秒出真实售价！
              </div>
            </div>
          </div>
        </div>

        {/* Right: Barcode Input Area (7 Cols) */}
        <div className="lg:col-span-7 bg-slate-900/90 border border-slate-800 rounded-2xl p-5 shadow-lg flex flex-col justify-between space-y-4">
          <div className="space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <h3 className="font-extrabold text-sm text-white">
                  待提取商品条形码 (EAN)
                </h3>
                <span className="text-xs font-mono font-bold px-2 py-0.5 rounded-full bg-slate-950 text-cyan-300 border border-slate-800">
                  {parsedBarcodes.length} 个条码
                </span>
              </div>

              {/* Action Buttons */}
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  className="text-xs bg-slate-800 hover:bg-slate-700 text-slate-200 px-2.5 py-1.5 rounded-lg border border-slate-700 transition flex items-center gap-1.5 cursor-pointer"
                  title="上传已有的商品 Excel 导入条码"
                >
                  <Upload className="w-3.5 h-3.5 text-emerald-400" />
                  <span>Excel 表格导入</span>
                </button>
                <input
                  type="file"
                  ref={fileInputRef}
                  onChange={handleFileUpload}
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                />

                <button
                  type="button"
                  onClick={handleLoadSample}
                  className="text-xs bg-slate-800 hover:bg-slate-700 text-amber-300 px-2.5 py-1.5 rounded-lg border border-slate-700 transition flex items-center gap-1.5 cursor-pointer"
                  title="载入系统预设的 300+ 待比价条码"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  <span>载入 300+ 测试条码</span>
                </button>

                {inputText && (
                  <button
                    type="button"
                    onClick={() => setInputText('')}
                    className="text-xs text-slate-500 hover:text-rose-400 p-1 cursor-pointer"
                    title="清空"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                )}
              </div>
            </div>

            {/* Multi-sheet summary banner if Excel was uploaded with multiple sheets */}
            {sheetSummaries.length > 0 && (
              <div className="bg-slate-950/80 border border-slate-800 rounded-xl p-3 space-y-2 animate-in fade-in">
                <div className="flex flex-wrap items-center justify-between gap-1 text-[11px]">
                  <span className="text-slate-200 font-bold flex items-center gap-1.5">
                    <FolderTree className="w-3.5 h-3.5 text-cyan-400" />
                    <span>已读取 {sheetSummaries.length} 个工作表（不同品牌/分类产品）：</span>
                  </span>
                  <span className="text-[10px] text-emerald-400 font-mono bg-emerald-950/60 border border-emerald-800/40 px-2 py-0.5 rounded-full">
                    全部表的条码已提取汇总于下方
                  </span>
                </div>
                <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto scrollbar-thin">
                  {sheetSummaries.map((s) => (
                    <span
                      key={s.sheetName}
                      className="text-[10px] font-mono bg-slate-900 border border-slate-700/80 text-cyan-300 px-2 py-0.5 rounded-lg flex items-center gap-1"
                    >
                      <span className="text-slate-400 font-sans">{s.sheetName}:</span>
                      <span className="font-bold text-amber-300">{s.count} 个条码</span>
                    </span>
                  ))}
                </div>
                <p className="text-[10px] text-slate-400 flex items-center gap-1">
                  <Sparkles className="w-3 h-3 text-amber-400 shrink-0" />
                  <span>提示：提取完价格后，可直接点击「套用我的自定义 Excel 模板导出」，系统将自动在各工作表中新建列写入价格！</span>
                </p>
              </div>
            )}

            {/* Input Textarea */}
            <div className="relative">
              <textarea
                value={inputText}
                onChange={(e) => setInputText(e.target.value)}
                placeholder="请在此粘贴或输入商品条形码（每行一个，或用逗号、空格分隔，支持 8-14 位标准 EAN 条码）..."
                className="w-full h-36 bg-slate-950 border border-slate-800 focus:border-amber-500 rounded-xl p-3 text-xs text-slate-200 font-mono placeholder:text-slate-600 focus:outline-none transition resize-none scrollbar-thin scrollbar-thumb-slate-700"
              />
              {parsedBarcodes.length > 0 && (
                <div className="absolute right-3 bottom-3 text-[10px] text-slate-400 bg-slate-900/90 px-2 py-0.5 rounded border border-slate-800 font-mono">
                  已解析 {parsedBarcodes.length} 个不重复条码
                </div>
              )}
            </div>
          </div>

          {/* Trigger Button & Status Message */}
          <div className="space-y-3 pt-2">
            <div className="flex items-center gap-3">
              {!isProcessing ? (
                <button
                  type="button"
                  onClick={startExtraction}
                  disabled={parsedBarcodes.length === 0}
                  className="flex-1 bg-gradient-to-r from-amber-500 via-amber-600 to-amber-700 hover:from-amber-400 hover:to-amber-600 disabled:opacity-40 disabled:cursor-not-allowed text-slate-950 font-black text-sm py-3 px-6 rounded-xl shadow-xl shadow-amber-500/20 transition flex items-center justify-center gap-2 group cursor-pointer"
                >
                  <Play className="w-4 h-4 fill-slate-950 group-hover:scale-110 transition-transform" />
                  <span>启动全网多平台价格提取 ({selectedSites.length} 站并发)</span>
                </button>
              ) : (
                <button
                  type="button"
                  onClick={stopExtraction}
                  className="flex-1 bg-rose-600 hover:bg-rose-500 text-white font-black text-sm py-3 px-6 rounded-xl shadow-xl shadow-rose-950 transition flex items-center justify-center gap-2 cursor-pointer"
                >
                  <Square className="w-4 h-4 fill-white" />
                  <span>停止提取任务</span>
                </button>
              )}
            </div>

            {/* Live Status & Progress */}
            {isProcessing && (
              <div className="space-y-1.5 animate-in fade-in">
                <div className="flex items-center justify-between text-xs text-slate-400">
                  <span className="flex items-center gap-2 text-amber-300 font-semibold truncate max-w-[80%]">
                    <RefreshCw className="w-3.5 h-3.5 animate-spin text-amber-400 shrink-0" />
                    <span className="truncate">{statusText}</span>
                  </span>
                  <span className="font-mono text-cyan-300 font-bold">{progressPercent}%</span>
                </div>
                <div className="w-full h-2 bg-slate-950 rounded-full overflow-hidden border border-slate-800">
                  <div
                    className="h-full bg-gradient-to-r from-amber-500 via-emerald-400 to-cyan-400 transition-all duration-200"
                    style={{ width: `${progressPercent}%` }}
                  />
                </div>
              </div>
            )}

            {!isProcessing && statusText && (
              <p className="text-xs text-slate-400 flex items-center gap-1.5">
                <Info className="w-3.5 h-3.5 text-cyan-400 shrink-0" />
                <span>{statusText}</span>
              </p>
            )}
          </div>
        </div>
      </div>

      {/* KPI Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-9 gap-2.5">
        <div className="bg-slate-900 border border-slate-800 p-3 rounded-xl">
          <span className="text-[11px] text-slate-400 block">待提取条码</span>
          <span className="text-lg font-extrabold text-white font-mono">{parsedBarcodes.length}</span>
        </div>

        <div className="bg-slate-900 border border-slate-800 p-3 rounded-xl">
          <span className="text-[11px] text-slate-400 block">已完成检索</span>
          <span className="text-lg font-extrabold text-cyan-400 font-mono">{stats.processed}</span>
        </div>

        <div className="bg-slate-900 border border-slate-800 p-3 rounded-xl">
          <span className="text-[11px] text-slate-400 block">成功匹配商品</span>
          <span className="text-lg font-extrabold text-emerald-400 font-mono">{stats.matchedCount}</span>
        </div>

        {/* Promo / Discount Stat */}
        <div className="bg-slate-900 border border-rose-900/40 p-3 rounded-xl bg-gradient-to-b from-rose-950/20 to-transparent">
          <span className="text-[11px] text-rose-300 font-semibold flex items-center gap-1">
            <Flame className="w-3 h-3 text-rose-400" />
            <span>促销打折</span>
          </span>
          <div className="flex items-baseline gap-1 mt-0.5">
            <span className="text-lg font-extrabold text-rose-400 font-mono">{discountedCount}</span>
            <span className="text-[10px] text-slate-500">件</span>
          </div>
        </div>

        <div className="bg-slate-900 border border-slate-800 p-3 rounded-xl">
          <span className="text-[11px] text-purple-300 block">Meloni</span>
          <span className="text-lg font-extrabold text-purple-400 font-mono">{stats.meloniFound}</span>
        </div>

        <div className="bg-slate-900 border border-slate-800 p-3 rounded-xl">
          <span className="text-[11px] text-emerald-300 block">MegaCedi</span>
          <span className="text-lg font-extrabold text-emerald-400 font-mono">{stats.megaFound}</span>
        </div>

        <div className="bg-slate-900 border border-slate-800 p-3 rounded-xl">
          <span className="text-[11px] text-cyan-300 block">Maury's</span>
          <span className="text-lg font-extrabold text-cyan-400 font-mono">{stats.maurysFound}</span>
        </div>

        <div className="bg-slate-900 border border-slate-800 p-3 rounded-xl">
          <span className="text-[11px] text-amber-300 block">Risparmio</span>
          <span className="text-lg font-extrabold text-amber-400 font-mono">{stats.risparmioFound}</span>
        </div>

        <div className="bg-slate-900 border border-rose-900/40 p-3 rounded-xl bg-gradient-to-b from-rose-950/20 to-transparent">
          <span className="text-[11px] text-rose-300 block">PiùMe</span>
          <span className="text-lg font-extrabold text-rose-400 font-mono">{stats.piumeFound}</span>
        </div>
      </div>

      {/* Table & Results View */}
      <div className="bg-slate-900/90 border border-slate-800 rounded-2xl p-5 shadow-xl space-y-4">
        {/* Table Toolbar */}
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2">
            {/* Filter Mode Tabs */}
            <div className="flex flex-wrap items-center bg-slate-950 p-1 rounded-xl border border-slate-800 text-xs">
              <button
                type="button"
                onClick={() => setFilterMode('all')}
                className={`px-3 py-1.5 rounded-lg font-medium transition cursor-pointer ${
                  filterMode === 'all' ? 'bg-amber-500 text-slate-950 font-bold' : 'text-slate-400 hover:text-white'
                }`}
              >
                全部条码 ({results.length})
              </button>
              <button
                type="button"
                onClick={() => setFilterMode('matched')}
                className={`px-3 py-1.5 rounded-lg font-medium transition cursor-pointer ${
                  filterMode === 'matched' ? 'bg-emerald-500 text-slate-950 font-bold' : 'text-slate-400 hover:text-white'
                }`}
              >
                已匹配比价 ({results.filter((r) => r.matched).length})
              </button>
              <button
                type="button"
                onClick={() => setFilterMode('discounted')}
                className={`px-3 py-1.5 rounded-lg font-medium transition cursor-pointer flex items-center gap-1 ${
                  filterMode === 'discounted' ? 'bg-rose-500 text-white font-bold' : 'text-rose-400 hover:text-rose-300'
                }`}
                title="只查看具有促销打折价的商品"
              >
                <Flame className="w-3 h-3" />
                <span>仅打折促销 ({discountedCount})</span>
              </button>
              <button
                type="button"
                onClick={() => setFilterMode('regular')}
                className={`px-3 py-1.5 rounded-lg font-medium transition cursor-pointer ${
                  filterMode === 'regular' ? 'bg-blue-600 text-white font-bold' : 'text-blue-400 hover:text-blue-300'
                }`}
                title="只查看正价（无活动打折）的商品"
              >
                仅正价商品 ({results.filter((r) => r.matched && !r.hasDiscount).length})
              </button>
              <button
                type="button"
                onClick={() => setFilterMode('spread')}
                className={`px-3 py-1.5 rounded-lg font-medium transition cursor-pointer ${
                  filterMode === 'spread' ? 'bg-cyan-500 text-slate-950 font-bold' : 'text-slate-400 hover:text-white'
                }`}
              >
                有价差空间 ({results.filter((r) => r.priceDiff && r.priceDiff > 0).length})
              </button>
              <button
                type="button"
                onClick={() => setFilterMode('unmatched')}
                className={`px-3 py-1.5 rounded-lg font-medium transition cursor-pointer ${
                  filterMode === 'unmatched' ? 'bg-slate-700 text-white font-bold' : 'text-slate-400 hover:text-white'
                }`}
              >
                未匹配 ({results.filter((r) => !r.matched).length})
              </button>
            </div>

            {/* Keyword Search */}
            <div className="relative min-w-[200px]">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                placeholder="搜索条形码或商品品名..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 focus:border-amber-500 rounded-xl pl-9 pr-3 py-1.5 text-xs text-slate-200 placeholder:text-slate-500 focus:outline-none transition"
              />
            </div>
          </div>

          {/* Action Bar & Display Toggle */}
          <div className="flex flex-wrap items-center gap-2">
            {/* Display Switcher: All vs Original only vs Discount only */}
            <div className="flex items-center bg-slate-950 px-2 py-1 rounded-xl border border-slate-800 text-[11px] gap-1 text-slate-400">
              <span className="text-[10px] text-slate-500 mr-1 flex items-center gap-0.5">
                <Eye className="w-3 h-3" />
                <span>显示:</span>
              </span>
              <button
                type="button"
                onClick={() => setPriceViewMode('both')}
                className={`px-2 py-0.5 rounded cursor-pointer ${
                  priceViewMode === 'both' ? 'bg-amber-500 text-slate-950 font-bold' : 'hover:text-white'
                }`}
              >
                区分原价+特价
              </button>
              <button
                type="button"
                onClick={() => setPriceViewMode('original_only')}
                className={`px-2 py-0.5 rounded cursor-pointer ${
                  priceViewMode === 'original_only' ? 'bg-blue-600 text-white font-bold' : 'hover:text-white'
                }`}
                title="隐藏短期打折特价，纯看标准原价"
              >
                只看原价
              </button>
              <button
                type="button"
                onClick={() => setPriceViewMode('discount_only')}
                className={`px-2 py-0.5 rounded cursor-pointer ${
                  priceViewMode === 'discount_only' ? 'bg-rose-600 text-white font-bold' : 'hover:text-white'
                }`}
                title="仅看折后特价"
              >
                只看特价
              </button>
            </div>

            {/* Custom Excel Template Fill Button (Core user request) */}
            <button
              type="button"
              onClick={() => setIsTemplateModalOpen(true)}
              disabled={results.length === 0}
              className="bg-gradient-to-r from-amber-500 via-amber-600 to-amber-700 hover:from-amber-400 hover:to-amber-600 text-slate-950 text-xs font-black px-3.5 py-1.5 rounded-xl shadow-lg shadow-amber-500/20 transition flex items-center gap-1.5 disabled:opacity-40 disabled:cursor-not-allowed group cursor-pointer"
              title="上传您自己的 Excel 模板，按条码精准将提取的价格填入指定列"
            >
              <FileSpreadsheet className="w-4 h-4 fill-slate-950 group-hover:rotate-6 transition-transform" />
              <span>套用我的自定义 Excel 模板导出</span>
            </button>

            {/* Direct XLSX */}
            <button
              type="button"
              onClick={handleExportXlsx}
              disabled={results.length === 0}
              className="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold px-3 py-1.5 rounded-xl shadow transition flex items-center gap-1.5 disabled:opacity-40 cursor-pointer"
              title="下载包含颜色标记与对比格式的 Excel 文件"
            >
              <Download className="w-3.5 h-3.5" />
              <span>多源比价 Excel (.xlsx)</span>
            </button>

            {/* Direct CSV */}
            <button
              type="button"
              onClick={handleExportCsv}
              disabled={results.length === 0}
              className="bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold px-3 py-1.5 rounded-xl border border-slate-700 transition flex items-center gap-1.5 disabled:opacity-40 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>标准 CSV</span>
            </button>

            {/* Clear */}
            {results.length > 0 && (
              <button
                type="button"
                onClick={() => {
                  setResults([]);
                  setStats({
                    total: 0,
                    processed: 0,
                    matchedCount: 0,
                    unmatchedCount: 0,
                    meloniFound: 0,
                    megaFound: 0,
                    maurysFound: 0,
                    risparmioFound: 0,
                    piumeFound: 0,
                  });
                }}
                disabled={isProcessing}
                className="p-1.5 text-slate-500 hover:text-rose-400 rounded-lg hover:bg-rose-500/10 transition cursor-pointer"
                title="清空当前结果"
              >
                <Trash2 className="w-4 h-4" />
              </button>
            )}
          </div>
        </div>

        {/* Results Table */}
        {filteredResults.length > 0 ? (
          <div className="overflow-x-auto border border-slate-800 rounded-xl bg-slate-950 scrollbar-thin scrollbar-thumb-slate-700">
            <table className="w-full text-left text-xs border-collapse font-sans">
              <thead>
                <tr className="bg-slate-900 border-b border-slate-800 text-slate-300 text-[11px]">
                  <th className="py-3 px-3 font-semibold">条形码 (Barcode)</th>
                  <th className="py-3 px-3 font-semibold min-w-[200px]">商品品名 (Title)</th>
                  <th className="py-3 px-3 font-bold text-amber-300 bg-amber-950/20 min-w-[170px]">
                    {priceViewMode === 'original_only'
                      ? '综合最低原价 (正价)'
                      : priceViewMode === 'discount_only'
                      ? '综合最低折后特价'
                      : '综合最优价 (原价 / 打折特价)'}
                  </th>
                  <th className="py-3 px-3 font-semibold text-purple-300 min-w-[130px]">Meloni Store</th>
                  <th className="py-3 px-3 font-semibold text-emerald-300 min-w-[130px]">MegaCedi 批价</th>
                  <th className="py-3 px-3 font-semibold text-cyan-300 min-w-[130px]">Maury's Online</th>
                  <th className="py-3 px-3 font-semibold text-amber-300 min-w-[130px]">Risparmio Casa</th>
                  <th className="py-3 px-3 font-semibold text-rose-300 min-w-[130px]">PiùMe</th>
                  <th className="py-3 px-3 font-semibold text-right">平台价差</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/80 font-mono text-xs">
                {filteredResults.map((r) => {
                  const hasLowest = r.lowestPrice !== null;
                  const hasDiscount = r.hasDiscount && r.lowestDiscountPrice !== null;
                  const origPrice = r.lowestOriginalPrice ?? r.lowestPrice;

                  return (
                    <tr key={r.barcode} className="hover:bg-slate-900/40 transition">
                      {/* Barcode */}
                      <td className="py-2.5 px-3 whitespace-nowrap">
                        <div className="flex items-center gap-1.5">
                          <span className="font-bold text-slate-200">{r.barcode}</span>
                          <button
                            type="button"
                            onClick={() => copyText(r.barcode, `b-${r.barcode}`)}
                            className="text-slate-500 hover:text-amber-400 p-0.5 cursor-pointer"
                            title="复制条码"
                          >
                            {copiedKey === `b-${r.barcode}` ? (
                              <Check className="w-3 h-3 text-emerald-400" />
                            ) : (
                              <Copy className="w-3 h-3" />
                            )}
                          </button>
                        </div>
                      </td>

                      {/* Title */}
                      <td className="py-2.5 px-3 font-sans">
                        <div className="space-y-0.5">
                          <span className="line-clamp-1 text-slate-300 font-medium" title={r.title}>
                            {r.title}
                          </span>
                          {hasDiscount && (
                            <span className="inline-flex items-center gap-1 text-[10px] text-rose-300 bg-rose-950/60 border border-rose-800/40 px-1.5 py-0.2 rounded font-sans">
                              <Flame className="w-2.5 h-2.5 text-rose-400" />
                              <span>促销打折活动中</span>
                            </span>
                          )}
                        </div>
                      </td>

                      {/* Composite Lowest / Distinguish Original vs Discount Price */}
                      <td className="py-2.5 px-3 bg-amber-950/10 whitespace-nowrap">
                        {hasLowest ? (
                          <div className="space-y-1">
                            {/* Mode A: View Both */}
                            {priceViewMode === 'both' && (
                              <>
                                {hasDiscount ? (
                                  <div className="space-y-1">
                                    <div className="flex items-center gap-1.5">
                                      <span className="text-rose-400 font-black text-sm bg-rose-950/60 px-2 py-0.5 rounded border border-rose-500/40">
                                        特价 €{r.lowestDiscountPrice!.toFixed(2)}
                                      </span>
                                      <span className="text-[10px] font-bold text-rose-300 bg-rose-500/20 px-1 rounded">
                                        促
                                      </span>
                                    </div>
                                    <div className="text-[11px] text-slate-400 flex items-center gap-1 font-sans">
                                      <span>原价:</span>
                                      <span className="line-through text-slate-400 font-mono">
                                        €{origPrice ? origPrice.toFixed(2) : '-'}
                                      </span>
                                      {r.lowestPlatform && (
                                        <span className="text-slate-500 text-[10px]">({r.lowestPlatform})</span>
                                      )}
                                    </div>
                                  </div>
                                ) : (
                                  <div className="space-y-0.5">
                                    <div className="flex items-center gap-1.5">
                                      <span className="text-blue-400 font-black text-sm bg-blue-950/60 px-2 py-0.5 rounded border border-blue-500/40">
                                        原价 €{origPrice ? origPrice.toFixed(2) : '-'}
                                      </span>
                                      <span className="text-[10px] text-blue-300 font-sans">正价</span>
                                    </div>
                                    {r.lowestPlatform && (
                                      <div className="text-[10px] text-slate-400 font-sans">
                                        来自 {r.lowestPlatform}
                                      </div>
                                    )}
                                  </div>
                                )}
                              </>
                            )}

                            {/* Mode B: Original Only */}
                            {priceViewMode === 'original_only' && (
                              <div className="flex items-center gap-1.5">
                                <span className="text-blue-400 font-black text-sm bg-blue-950/60 px-2 py-0.5 rounded border border-blue-500/40">
                                  原价 €{origPrice ? origPrice.toFixed(2) : '-'}
                                </span>
                                {r.lowestOriginalPlatform && (
                                  <span className="text-[10px] text-slate-400 font-sans">
                                    ({r.lowestOriginalPlatform})
                                  </span>
                                )}
                              </div>
                            )}

                            {/* Mode C: Discount Only */}
                            {priceViewMode === 'discount_only' && (
                              <div>
                                {hasDiscount ? (
                                  <span className="text-rose-400 font-black text-sm bg-rose-950/60 px-2 py-0.5 rounded border border-rose-500/40">
                                    特价 €{r.lowestDiscountPrice!.toFixed(2)}
                                  </span>
                                ) : (
                                  <span className="text-slate-500 text-xs italic font-sans">无特价 (正价)</span>
                                )}
                              </div>
                            )}
                          </div>
                        ) : (
                          <span className="text-slate-500 text-xs italic font-sans">未取得金额</span>
                        )}
                      </td>

                      {/* Meloni Store */}
                      <td className="py-2.5 px-3 whitespace-nowrap">
                        {r.prices.meloni?.found ? (
                          <div className="space-y-0.5">
                            {r.prices.meloni.hasDiscount ? (
                              <div>
                                <div className="flex items-center gap-1">
                                  <span className="text-rose-400 font-bold">
                                    {r.prices.meloni.formattedDiscountPrice}
                                  </span>
                                  {r.prices.meloni.discountPercent && (
                                    <span className="text-[9px] text-rose-300 bg-rose-950/60 px-1 rounded">
                                      {r.prices.meloni.discountPercent}
                                    </span>
                                  )}
                                </div>
                                <div className="text-[10px] text-slate-400 font-sans">
                                  原价: <span className="line-through">{r.prices.meloni.formattedOriginalPrice}</span>
                                </div>
                              </div>
                            ) : (
                              <span className="text-purple-300 font-bold block">
                                {r.prices.meloni.formattedPrice}
                              </span>
                            )}
                            {r.prices.meloni.url && (
                              <a
                                href={r.prices.meloni.url}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-[10px] text-purple-400/80 hover:underline flex items-center gap-0.5"
                              >
                                <span>原网</span>
                                <ExternalLink className="w-2.5 h-2.5" />
                              </a>
                            )}
                          </div>
                        ) : (
                          <span className="text-slate-600 text-xs">-</span>
                        )}
                      </td>

                      {/* MegaCedi */}
                      <td className="py-2.5 px-3 whitespace-nowrap">
                        {r.prices.megacedi?.found ? (
                          <div className="space-y-0.5">
                            <span className="text-emerald-300 font-bold block">
                              {r.prices.megacedi.formattedPrice}
                            </span>
                            {r.prices.megacedi.url && (
                              <a
                                href={r.prices.megacedi.url}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-[10px] text-emerald-400/80 hover:underline flex items-center gap-0.5"
                              >
                                <span>原网</span>
                                <ExternalLink className="w-2.5 h-2.5" />
                              </a>
                            )}
                          </div>
                        ) : (
                          <span className="text-slate-600 text-xs">-</span>
                        )}
                      </td>

                      {/* Maurys */}
                      <td className="py-2.5 px-3 whitespace-nowrap">
                        {r.prices.maurys?.found ? (
                          <div className="space-y-0.5">
                            {r.prices.maurys.hasDiscount ? (
                              <div>
                                <div className="flex items-center gap-1">
                                  <span className="text-rose-400 font-bold">
                                    {r.prices.maurys.formattedDiscountPrice}
                                  </span>
                                  {r.prices.maurys.discountPercent && (
                                    <span className="text-[9px] text-rose-300 bg-rose-950/60 px-1 rounded">
                                      {r.prices.maurys.discountPercent}
                                    </span>
                                  )}
                                </div>
                                <div className="text-[10px] text-slate-400 font-sans">
                                  原价: <span className="line-through">{r.prices.maurys.formattedOriginalPrice}</span>
                                </div>
                              </div>
                            ) : (
                              <span className="text-cyan-300 font-bold block">
                                {r.prices.maurys.formattedPrice}
                              </span>
                            )}
                            {r.prices.maurys.url && (
                              <a
                                href={r.prices.maurys.url}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-[10px] text-cyan-400/80 hover:underline flex items-center gap-0.5"
                              >
                                <span>原网</span>
                                <ExternalLink className="w-2.5 h-2.5" />
                              </a>
                            )}
                          </div>
                        ) : (
                          <span className="text-slate-600 text-xs">-</span>
                        )}
                      </td>

                      {/* Risparmio Casa */}
                      <td className="py-2.5 px-3 whitespace-nowrap">
                        {r.prices.risparmiocasa?.found ? (
                          <div className="space-y-0.5">
                            {r.prices.risparmiocasa.hasDiscount ? (
                              <div>
                                <div className="flex items-center gap-1">
                                  <span className="text-rose-400 font-bold">
                                    {r.prices.risparmiocasa.formattedDiscountPrice}
                                  </span>
                                  {r.prices.risparmiocasa.discountPercent && (
                                    <span className="text-[9px] text-rose-300 bg-rose-950/60 px-1 rounded">
                                      {r.prices.risparmiocasa.discountPercent}
                                    </span>
                                  )}
                                </div>
                                <div className="text-[10px] text-slate-400 font-sans">
                                  原价: <span className="line-through">{r.prices.risparmiocasa.formattedOriginalPrice}</span>
                                </div>
                              </div>
                            ) : (
                              <span className="text-amber-300 font-bold block">
                                {r.prices.risparmiocasa.formattedPrice}
                              </span>
                            )}
                            {r.prices.risparmiocasa.url && (
                              <a
                                href={r.prices.risparmiocasa.url}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-[10px] text-amber-400/80 hover:underline flex items-center gap-0.5"
                              >
                                <span>原网</span>
                                <ExternalLink className="w-2.5 h-2.5" />
                              </a>
                            )}
                          </div>
                        ) : (
                          <span className="text-slate-600 text-xs">-</span>
                        )}
                      </td>

                      {/* PiùMe */}
                      <td className="py-2.5 px-3 whitespace-nowrap">
                        {r.prices.piume?.found ? (
                          <div className="space-y-0.5">
                            {r.prices.piume.hasDiscount ? (
                              <div>
                                <div className="flex items-center gap-1">
                                  <span className="text-rose-400 font-bold">
                                    {r.prices.piume.formattedDiscountPrice}
                                  </span>
                                  {r.prices.piume.discountPercent && (
                                    <span className="text-[9px] text-rose-300 bg-rose-950/60 px-1 rounded">
                                      {r.prices.piume.discountPercent}
                                    </span>
                                  )}
                                </div>
                                <div className="text-[10px] text-slate-400 font-sans">
                                  原价: <span className="line-through">{r.prices.piume.formattedOriginalPrice}</span>
                                </div>
                              </div>
                            ) : (
                              <span className="text-rose-300 font-bold block">
                                {r.prices.piume.formattedPrice}
                              </span>
                            )}
                            {r.prices.piume.url && (
                              <a
                                href={r.prices.piume.url}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-[10px] text-rose-400/80 hover:underline flex items-center gap-0.5"
                              >
                                <span>原网</span>
                                <ExternalLink className="w-2.5 h-2.5" />
                              </a>
                            )}
                          </div>
                        ) : (
                          <span className="text-slate-600 text-xs">-</span>
                        )}
                      </td>

                      {/* Price Diff */}
                      <td className="py-2.5 px-3 text-right whitespace-nowrap">
                        {r.priceDiff && r.priceDiff > 0 ? (
                          <span className="text-amber-400 font-bold bg-amber-950/40 px-2 py-0.5 rounded border border-amber-800/40">
                            Δ €{r.priceDiff.toFixed(2)}
                          </span>
                        ) : (
                          <span className="text-slate-600">-</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        ) : (
          /* Empty placeholder */
          <div className="border border-dashed border-slate-800 rounded-xl p-12 text-center flex flex-col items-center justify-center space-y-3">
            <Coins className="w-12 h-12 text-slate-700" />
            <div className="space-y-1">
              <h4 className="text-sm font-bold text-slate-300">暂无价格比对数据</h4>
              <p className="text-xs text-slate-500 max-w-md">
                在上方粘贴条码或点击「载入 300+ 测试条码」，然后点击「启动全网多平台价格提取」即可开始并发检索。
              </p>
            </div>
            <button
              type="button"
              onClick={handleLoadSample}
              className="text-xs bg-slate-800 hover:bg-slate-700 text-amber-300 px-3 py-1.5 rounded-lg border border-slate-700 transition flex items-center gap-1.5 cursor-pointer"
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>载入测试条码体验</span>
            </button>
          </div>
        )}
      </div>

      {/* Custom Excel Template Injection Modal */}
      <PriceTemplateModal
        isOpen={isTemplateModalOpen}
        onClose={() => setIsTemplateModalOpen(false)}
        items={results}
        initialFile={uploadedWorkbookFile}
        initialBase64={uploadedWorkbookBase64}
        initialSheetSummaries={sheetSummaries}
      />

      {/* Multi-Platform Auth & Cookie Modal */}
      <PlatformAuthModal
        isOpen={isLoginModalOpen}
        onClose={() => setIsLoginModalOpen(false)}
        initialPlatform={authModalPlatform}
      />
    </div>
  );
};
