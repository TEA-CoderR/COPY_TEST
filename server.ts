import express, { Request, Response } from 'express';
import * as cheerio from 'cheerio';
import path from 'path';
import fs from 'fs';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '50mb' }));

// Ensure persistent directory for barcode-named images
const SAVED_IMAGES_DIR = path.resolve(__dirname, 'saved_images');
if (!fs.existsSync(SAVED_IMAGES_DIR)) {
  fs.mkdirSync(SAVED_IMAGES_DIR, { recursive: true });
}
app.use('/saved_images', express.static(SAVED_IMAGES_DIR));
app.use('/images', express.static(SAVED_IMAGES_DIR));

const CDN_INDEX_FILE = path.join(SAVED_IMAGES_DIR, 'cdn_index.json');
const barcodeCdnMap = new Map<string, string>();
const barcodeMetadataMap = new Map<string, { title?: string; sourceUrl?: string; provider?: string }>();

// Load persistent CDN mappings if file exists
try {
  if (fs.existsSync(CDN_INDEX_FILE)) {
    const data = JSON.parse(fs.readFileSync(CDN_INDEX_FILE, 'utf-8'));
    for (const [k, v] of Object.entries(data)) {
      if (typeof v === 'string') {
        barcodeCdnMap.set(k, v);
      } else if (v && typeof v === 'object') {
        const item = v as any;
        if (item.cdnUrl) barcodeCdnMap.set(k, item.cdnUrl);
        barcodeMetadataMap.set(k, { title: item.title, sourceUrl: item.sourceUrl, provider: item.provider });
      }
    }
  }
} catch {}

function saveCdnMapping(barcode: string, cdnUrl: string, meta?: { title?: string; sourceUrl?: string; provider?: string }) {
  if (cdnUrl) barcodeCdnMap.set(barcode, cdnUrl);
  if (meta) barcodeMetadataMap.set(barcode, meta);
  try {
    const obj: Record<string, any> = {};
    for (const [k, url] of barcodeCdnMap.entries()) {
      const m = barcodeMetadataMap.get(k) || {};
      obj[k] = { cdnUrl: url, ...m };
    }
    fs.writeFileSync(CDN_INDEX_FILE, JSON.stringify(obj, null, 2));
  } catch {}
}

const DEFAULT_HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
  'Accept-Language': 'it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7,zh-CN;q=0.6',
  'Cache-Control': 'no-cache',
  'Pragma': 'no-cache',
};

// MegaCedi active session store
let megaCediSession: {
  cookie: string;
  username?: string;
  authenticated: boolean;
} = {
  cookie: '',
  username: undefined,
  authenticated: false,
};

// Meloni Store active session store
let meloniSession: {
  cookie: string;
  username?: string;
  authenticated: boolean;
} = {
  cookie: '',
  username: undefined,
  authenticated: false,
};

// Maury's Online active session store
let maurysSession: {
  cookie: string;
  username?: string;
  authenticated: boolean;
} = {
  cookie: '',
  username: undefined,
  authenticated: false,
};

// Risparmio Casa active session store
let risparmioSession: {
  cookie: string;
  username?: string;
  authenticated: boolean;
} = {
  cookie: '',
  username: undefined,
  authenticated: false,
};

// PiùMe active session store
let piumeSession: {
  cookie: string;
  username?: string;
  authenticated: boolean;
} = {
  cookie: '',
  username: undefined,
  authenticated: false,
};

// Cached BigCommerce Storefront GraphQL token for PiùMe
let piumeGraphQLToken: string =
  'eyJ0eXAiOiJKV1QiLCJhbGciOiJFUzI1NiJ9.eyJjaWQiOlsxXSwiY29ycyI6WyJodHRwczovL3BpdW1lLml0Il0sImVhdCI6MTc5MjIzNDMxMSwiaWF0IjoxNzkyMDYxNTExLCJpc3MiOiJCQyIsInNpZCI6MTAwMjQyNjg0Mywic3ViIjoiQkMiLCJzdWJfdHlwZSI6MCwidG9rZW5fdHlwZSI6MX0.vjzb_eocDKNfh4auWjDgiKte8-HmQqk2r97RM5xsKm4EPNa-GKoQGjyESw75tXCOuMMiYfDd2mkNeqdtyebvVA';
let piumeTokenLastRefreshed = 0;

export interface ScrapedProduct {
  id: string;
  sku: string;
  title: string;
  url: string;
  imageUrl: string;
  rawImageUrl: string;
  imageFileName: string;
  originalPrice: number | null;
  originalPriceFormatted: string;
  discountPrice: number | null;
  discountPriceFormatted: string;
  discountPercent: string;
  discountNumeric: number;
  inStock: boolean;
  page: number;
  category: string;
  ean: string;
  sourceSite?: 'maurys' | 'risparmiocasa' | 'meloni' | 'megacedi' | 'other';
  packaging?: string;
  brand?: string;
  scrapedAt: string;
}

const eanCache = new Map<string, string>();
const packagingCache = new Map<string, string>();

export interface PagePaginationMeta {
  totalCount: number;
  currentPage: number;
  totalPages: number;
  pageSize: number;
  hasMore: boolean;
  pageTitle: string;
  siteType: 'maurys' | 'risparmiocasa' | 'meloni' | 'megacedi' | 'generic';
}

// ----------------------------------------------------
// PLATFORM CONFIGURATIONS & PRESETS
// ----------------------------------------------------
export const MAURYS_CATALOG_BLOCKS = [
  { slug: 'detersivo-e-cura-casa', name: '家居清洁', url: 'https://maurysonline.it/detersivo-e-cura-casa', maxPages: 105 },
  { slug: 'cura-persona-628', name: '个人洗护', url: 'https://maurysonline.it/cura-persona-628', maxPages: 45 },
  { slug: 'cucina', name: '餐具厨房', url: 'https://maurysonline.it/cucina', maxPages: 30 },
  { slug: 'animali', name: '宠物专区', url: 'https://maurysonline.it/animali', maxPages: 25 },
  { slug: 'arredo-casa-632', name: '家居软装', url: 'https://maurysonline.it/arredo-casa-632', maxPages: 25 },
  { slug: 'giocattoli', name: '儿童玩具', url: 'https://maurysonline.it/giocattoli', maxPages: 20 },
  { slug: 'cartoleria', name: '文具办公', url: 'https://maurysonline.it/cartoleria', maxPages: 15 },
  { slug: 'outdoor-638', name: '户外花园', url: 'https://maurysonline.it/outdoor-638', maxPages: 15 },
  { slug: 'elettrodomestici-633', name: '家用小电', url: 'https://maurysonline.it/elettrodomestici-633', maxPages: 10 },
  { slug: 'fai-da-te', name: '五金工具', url: 'https://maurysonline.it/fai-da-te', maxPages: 10 },
  { slug: 'festa', name: '节日派对', url: 'https://maurysonline.it/festa', maxPages: 8 },
  { slug: 'promo', name: '特价促销', url: 'https://maurysonline.it/promo', maxPages: 10 },
  { slug: 'new', name: '新品上市', url: 'https://maurysonline.it/new', maxPages: 5 },
];

export const MAURYS_ALL_TOTAL_PAGES = MAURYS_CATALOG_BLOCKS.reduce((sum, b) => sum + b.maxPages, 0);
export const RISPARMIO_ALL_TOTAL_PAGES = 92;
export const ALL_PLATFORMS_TOTAL_PAGES = RISPARMIO_ALL_TOTAL_PAGES + MAURYS_ALL_TOTAL_PAGES;

export function resolveMaurysAllPage(globalPage: number): { resolvedUrl: string; categoryName: string; pageInCategory: number } {
  let accumulated = 0;
  for (const block of MAURYS_CATALOG_BLOCKS) {
    if (globalPage <= accumulated + block.maxPages) {
      const pageInCategory = Math.max(1, globalPage - accumulated);
      const target = pageInCategory > 1 ? `${block.url}?pagenumber=${pageInCategory}` : block.url;
      return { resolvedUrl: target, categoryName: block.name, pageInCategory };
    }
    accumulated += block.maxPages;
  }
  const last = MAURYS_CATALOG_BLOCKS[MAURYS_CATALOG_BLOCKS.length - 1];
  return { resolvedUrl: last.url, categoryName: last.name, pageInCategory: 1 };
}

export function resolveAllPlatformsPage(globalPage: number): {
  resolvedUrl: string;
  platformName: string;
  categoryName: string;
  pageInPlatform: number;
} {
  if (globalPage <= RISPARMIO_ALL_TOTAL_PAGES) {
    const url = globalPage === 1 ? 'https://shop.risparmiocasa.com/prodotti' : `https://shop.risparmiocasa.com/prodotti?p=${globalPage}`;
    return {
      resolvedUrl: url,
      platformName: 'Risparmio Casa',
      categoryName: '官方全网商品总目录',
      pageInPlatform: globalPage,
    };
  } else {
    const maurysPage = globalPage - RISPARMIO_ALL_TOTAL_PAGES;
    const info = resolveMaurysAllPage(maurysPage);
    return {
      resolvedUrl: info.resolvedUrl,
      platformName: "Maury's Online",
      categoryName: info.categoryName,
      pageInPlatform: maurysPage,
    };
  }
}

// MegaCedi department code resolver
export function resolveMegaCediUrl(inputUrl: string): { fetchUrl: string; departmentCode: string; departmentName: string; query: string } {
  let cleaned = inputUrl.trim();
  let query = '';
  let deptCode = '20%'; // Default: CURA DELLA CASA
  let deptName = 'CURA DELLA CASA';

  if (cleaned.includes('WebPartScaffale.aspx') || cleaned.includes('WebPartRighe.aspx')) {
    try {
      const u = new URL(cleaned.startsWith('http') ? cleaned : `https://www.megacedi.com/${cleaned.replace(/^\//, '')}`);
      query = u.searchParams.get('p') || '';
      deptCode = u.searchParams.get('r') || '20%';
    } catch {}
    return {
      fetchUrl: `https://www.megacedi.com/WebPartScaffale.aspx?p=${encodeURIComponent(query)}&r=${encodeURIComponent(deptCode)}`,
      departmentCode: deptCode,
      departmentName: deptName,
      query,
    };
  }

  // Check department shortcuts or custom URLs
  const deptMap: Record<string, { code: string; name: string }> = {
    'cura-casa': { code: '20%', name: 'CURA DELLA CASA (家居清洁全系)' },
    'detergenti-superfici': { code: '2004%', name: 'DETERGENTI SUPERFICI (表面去油清洁)' },
    'detersivi-tessuti': { code: '2005%', name: 'DETERSIVI TESSUTI (衣物洗衣护理)' },
    'disposable': { code: '2006%', name: 'DISPOSABLE (一次性餐具与保鲜)' },
    'deodoranti': { code: '2003%', name: 'DEODORANTI (空间香氛除臭)' },
    'accessori-casa': { code: '2001%', name: 'ACCESSORI CASA (清洁配件与用具)' },
    'lavaggio-stoviglie': { code: '2009%', name: 'LAVAGGIO STOVIGLIE (餐具与洗洁精)' },
    'insetticidi': { code: '2008%', name: 'INSETTICIDI (驱蚊杀虫)' },
    'cura-persona': { code: '21%', name: 'CURA DELLA PERSONA (个护洗护全系)' },
    'toiletries': { code: '2105%', name: 'TOILETRIES (沐浴洗发个护)' },
    'igienico-sanitari': { code: '2103%', name: 'IGIENICO SANITARI (卫生纸品与护理)' },
    'accessori-persona': { code: '2101%', name: 'ACCESSORI PERSONA (个护美容用具)' },
    'casalinghi': { code: '54%', name: 'CASALINGHI (家居餐具全系)' },
    'tavola': { code: '5401%', name: 'TAVOLA (餐具与桌面用品)' },
    'mondo-cottura': { code: '5402%', name: 'MONDO COTTURA (锅具烹饪)' },
    'preparazione-cibi': { code: '5403%', name: 'PREPARAZIONE CIBI (备餐刀剪)' },
    'conservazione': { code: '5404%', name: 'CONSERVAZIONE (保鲜与储存)' },
    'pets': { code: '22%', name: 'PETS (宠物全系用品)' },
    'pet-food': { code: '2202%', name: 'PET FOOD (猫狗粮与罐头)' },
    'offerte': { code: '59%', name: 'OFFERTE SPECIALI (特价促销专区)' },
    'bazar': { code: '23%', name: 'BAZAR LEGGERO (日用百货杂货)' },
    'food': { code: '07%', name: 'FOOD CONFEZIONATO (包装食品)' },
    'bevande': { code: '08%', name: 'BEVANDE (酒水饮品)' },
    'elettro': { code: '50%', name: 'ELETTRONICA ELETTRODOMESTICI (生活小电)' },
    'bricolage': { code: '52%', name: 'BRICOLAGE (五金工具修缮)' },
  };

  for (const [key, val] of Object.entries(deptMap)) {
    if (cleaned.includes(key) || cleaned.includes(val.code)) {
      deptCode = val.code;
      deptName = val.name;
      break;
    }
  }

  // Check if search query passed
  const searchMatch = cleaned.match(/[?&]q=([^&]+)/) || cleaned.match(/[?&]p=([^&]+)/);
  if (searchMatch) {
    query = decodeURIComponent(searchMatch[1]);
  }

  return {
    fetchUrl: `https://www.megacedi.com/WebPartScaffale.aspx?p=${encodeURIComponent(query)}&r=${encodeURIComponent(deptCode)}`,
    departmentCode: deptCode,
    departmentName: deptName,
    query,
  };
}

function normalizeUrl(targetUrl: string, page: number): string {
  let url = targetUrl.trim();
  if (url.includes('all-platforms') || url === 'https://casa-te.all/products') {
    return resolveAllPlatformsPage(page).resolvedUrl;
  }
  if (url.includes('maurysonline.it/all-products') || url === 'https://maurysonline.it/all') {
    return resolveMaurysAllPage(page).resolvedUrl;
  }
  if (!url.startsWith('http://') && !url.startsWith('https://')) {
    url = `https://${url}`;
  }

  const parsed = new URL(url);
  const hostname = parsed.hostname.toLowerCase();

  // 1. Maury's Online
  if (hostname.includes('maurysonline.it')) {
    if (page > 1) {
      parsed.searchParams.set('pagenumber', page.toString());
    } else {
      parsed.searchParams.delete('pagenumber');
    }
    return parsed.toString();
  }

  // 2. Meloni Store (BigCommerce)
  if (hostname.includes('melonistore.com')) {
    if (page > 1) {
      parsed.searchParams.set('page', page.toString());
    } else {
      parsed.searchParams.delete('page');
    }
    return parsed.toString();
  }

  // 3. MegaCedi
  if (hostname.includes('megacedi.com')) {
    const info = resolveMegaCediUrl(url);
    return info.fetchUrl;
  }

  // 4. Risparmio Casa / Magento standard
  if (hostname.includes('risparmiocasa.com')) {
    if (page > 1) {
      parsed.searchParams.set('p', page.toString());
    } else {
      parsed.searchParams.delete('p');
    }
    return parsed.toString();
  }

  return parsed.toString();
}

async function fetchPageHtml(url: string, customHeaders: Record<string, string> = {}): Promise<string> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 25000);

  const reqHeaders: Record<string, string> = {
    ...DEFAULT_HEADERS,
    ...customHeaders,
  };

  // Inject MegaCedi session cookie if targeting megacedi.com
  if (url.includes('megacedi.com') && megaCediSession.cookie) {
    reqHeaders['Cookie'] = megaCediSession.cookie;
  }

  // Inject Meloni session cookie if targeting melonistore.com
  if (url.includes('melonistore.com') && meloniSession.cookie) {
    reqHeaders['Cookie'] = meloniSession.cookie;
  }

  // Inject Maury's session cookie if targeting maurysonline.it
  if (url.includes('maurysonline.it') && maurysSession.cookie) {
    reqHeaders['Cookie'] = maurysSession.cookie;
  }

  // Inject Risparmio Casa session cookie if targeting shop.risparmiocasa.com
  if (url.includes('risparmiocasa.com') && risparmioSession.cookie) {
    reqHeaders['Cookie'] = risparmioSession.cookie;
  }

  // Inject PiùMe session cookie if targeting piume.it
  if (url.includes('piume.it') && piumeSession.cookie) {
    reqHeaders['Cookie'] = piumeSession.cookie;
  }

  try {
    const response = await fetch(url, {
      headers: reqHeaders,
      signal: controller.signal,
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    return await response.text();
  } finally {
    clearTimeout(timeout);
  }
}

// Fetch detailed EAN and packaging info
async function fetchProductEan(productUrl: string): Promise<string> {
  if (!productUrl) return '';
  if (eanCache.has(productUrl)) return eanCache.get(productUrl)!;

  try {
    // 1. MegaCedi barcode & packaging
    if (productUrl.includes('megacedi.com')) {
      const codeMatch = productUrl.match(/codice_a_barre=([0-9A-Za-z]+)/i) || productUrl.match(/(\d{8,14})/);
      const code = codeMatch ? codeMatch[1] : '';
      if (!code) return '';

      const schedaUrl = `https://www.megacedi.com/SchedaHome.aspx?codice_a_barre=${code}`;
      const html = await fetchPageHtml(schedaUrl);
      const $ = cheerio.load(html);

      // Extract barcode: 8001480709126
      const barcodeMatch = html.match(/barcode:\s*([0-9]{8,14})/i) || html.match(/>(\d{13})<\/a>/i);
      const ean = barcodeMatch ? barcodeMatch[1].trim() : code;

      // Extract packaging e.g. "12,000 PZ"
      const packMatch = html.match(/confezione da:\s*<strong>\s*([^<]+)<\/strong>\s*&nbsp;\s*([A-Za-z]+)/i);
      if (packMatch) {
        packagingCache.set(productUrl, `${packMatch[1].trim()} ${packMatch[2].trim()}`);
      }

      if (ean) eanCache.set(productUrl, ean);
      return ean;
    }

    // 2. Meloni Store
    if (productUrl.includes('melonistore.com')) {
      // Check if EAN is embedded in product URL: e.g. -8721497047263.html
      const urlEanMatch = productUrl.match(/-([0-9]{12,14})\.html/i);
      if (urlEanMatch) {
        const ean = urlEanMatch[1];
        eanCache.set(productUrl, ean);
        return ean;
      }

      const html = await fetchPageHtml(productUrl);
      const $ = cheerio.load(html);
      const eanText = $('.sku-gtin').text();
      const m = eanText.match(/EAN:\s*([0-9]{8,14})/i) || html.match(/(?:EAN|gtin)[:\s]*([0-9]{8,14})/i);
      const ean = m ? m[1].trim() : '';
      if (ean) eanCache.set(productUrl, ean);
      return ean;
    }

    // 3. Maury's Online
    if (productUrl.includes('maurysonline.it')) {
      const html = await fetchPageHtml(productUrl);
      const $ = cheerio.load(html);
      const gtin = $('.gtin .value').text().trim();
      const sku = $('.sku .value').text().trim();
      const mpn = $('.mpn .value').text().trim();
      let ean = '';
      if (gtin) {
        ean = gtin;
      } else if (sku) {
        ean = sku;
      } else if (mpn) {
        ean = mpn;
      } else {
        const m = html.match(/(?:EAN|Codice|Barcode)[:\s]*([0-9A-Za-z\-]{7,14})/i);
        if (m) ean = m[1].trim();
      }
      if (ean) eanCache.set(productUrl, ean);
      return ean;
    }

    // 4. Risparmio Casa
    const html = await fetchPageHtml(productUrl);
    const $ = cheerio.load(html);
    const text = $('.sku-detail').text();
    const m = text.match(/Codice\s*articolo:\s*(\d+)/i) || html.match(/Codice\s*articolo:<\/span>\s*(\d+)/i);
    const ean = m ? m[1].trim() : '';

    if (ean) eanCache.set(productUrl, ean);
    return ean;
  } catch (err) {
    return '';
  }
}

// ----------------------------------------------------
// PARSER IMPLEMENTATION FOR ALL 4 PLATFORMS
// ----------------------------------------------------
function parsePageProducts(
  html: string,
  pageNumber: number,
  targetUrl: string = ''
): { products: ScrapedProduct[]; meta: PagePaginationMeta } {
  const $ = cheerio.load(html);
  const products: ScrapedProduct[] = [];
  const now = new Date().toISOString();

  const isMeloni = targetUrl.includes('melonistore.com') || html.includes('melonistore.com') || $('.card').length > 0;
  const isMegaCedi = targetUrl.includes('megacedi.com') || html.includes('megacedi.com') || $('.jqTest2').length > 0 || html.includes('WebPartScaffale') || html.includes('WebPartRighe');
  const isMaurys = !isMeloni && !isMegaCedi && (targetUrl.includes('maurysonline.it') || html.includes('maurysonline.it') || ($('.product-item').length > 0 && $('li.item.product.product-item').length === 0));

  // ====================================================
  // 1. MELONI STORE (melonistore.com - BigCommerce)
  // ====================================================
  if (isMeloni) {
    const pageTitle = $('title').text().trim() || 'Meloni Store - Prodotti per la Casa e Bucato';
    const category = $('.breadcrumb li:last-child').text().trim() || $('h1.page-heading').text().trim() || 'Casa e Bucato';

    $('.card').each((_, element) => {
      const card = $(element);
      const linkEl = card.find('.card-title a').first();
      const figureLinkEl = card.find('.card-figure-link').first();

      let title = linkEl.text().trim();
      let productUrl = linkEl.attr('href') || figureLinkEl.attr('href') || '';
      if (productUrl.startsWith('/')) {
        productUrl = `https://www.melonistore.com${productUrl}`;
      }

      if (!title && !productUrl) return;

      // Brand
      const brand = card.find('.card-text--brand, [data-test-info-type="brandName"]').text().trim();

      // SKU and EAN
      const skuGtinText = card.find('.sku-gtin').text().trim();
      const rifMatch = skuGtinText.match(/Rif:?\s*([0-9A-Za-z]+)/i);
      const eanMatch = skuGtinText.match(/EAN:?\s*([0-9]{8,14})/i) || productUrl.match(/-([0-9]{12,14})\.html/i);

      let sku = rifMatch ? rifMatch[1] : '';
      let ean = eanMatch ? eanMatch[1] : '';

      // ID
      let id = sku || ean;
      if (!id) {
        const idMatch = productUrl.match(/-(\d+)\.html/);
        id = idMatch ? idMatch[1] : Math.random().toString(36).substring(2, 9);
      }

      // Packaging
      const packaging = card.find('.PZxCT-card').text().replace(/\s+/g, ' ').trim();

      // Image: extract high-res stencil URL (upgrade 80w/500x659 to 1280x1280)
      const imgEl = card.find('img.card-image, img').first();
      let imageUrl = imgEl.attr('src') || imgEl.attr('data-src') || '';
      let rawImageUrl = imageUrl;

      const srcset = imgEl.attr('srcset') || imgEl.attr('data-srcset') || '';
      if (srcset) {
        const parts = srcset.split(',').map((s) => s.trim().split(' '));
        const last = parts[parts.length - 1];
        if (last && last[0]) rawImageUrl = last[0];
      }

      if (rawImageUrl) {
        // Upgrade stencil dimension to 1280x1280 for ultra HD export
        rawImageUrl = rawImageUrl.replace(/\/stencil\/\d+x?\d*w?\//, '/stencil/1280x1280/');
      }

      // Price & Discount (Meloni is B2B, displays price when logged in or warning when public)
      const priceText = card.find('.card-section--price .price, .price, [data-test-info-type="price"]').text().trim();
      const parsePrice = (str: string) => {
        if (!str || str.includes('Accedi')) return null;
        const clean = str.replace(/[^0-9,\.]/g, '').replace(',', '.');
        const num = parseFloat(clean);
        return isNaN(num) ? null : num;
      };

      const finalPriceNum = parsePrice(priceText);
      const oldPriceNum = null;

      // Safe clean filename
      const cleanTitle = title
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '_')
        .replace(/_+/g, '_')
        .substring(0, 50);
      const imageFileName = `meloni_${sku || id}_${cleanTitle || 'product'}.jpg`;

      // In stock
      const isOutOfStock = card.hasClass('card--out-of-stock') || card.find('.out-of-stock').length > 0;

      products.push({
        id,
        sku: sku || id,
        title,
        url: productUrl,
        imageUrl,
        rawImageUrl,
        imageFileName,
        originalPrice: oldPriceNum,
        originalPriceFormatted: '',
        discountPrice: finalPriceNum ? Number(finalPriceNum.toFixed(2)) : null,
        discountPriceFormatted: finalPriceNum ? `€${finalPriceNum.toFixed(2)}` : (priceText.includes('Accedi') ? '需登录查看底价' : priceText),
        discountPercent: '0%',
        discountNumeric: 0,
        inStock: !isOutOfStock,
        page: pageNumber,
        category,
        ean: ean || eanCache.get(productUrl) || '',
        sourceSite: 'meloni',
        packaging: packaging || undefined,
        brand: brand || undefined,
        scrapedAt: now,
      });
    });

    // Pagination for Meloni Store
    let totalPages = 1;
    $('.pagination-item .pagination-link').each((_, el) => {
      const p = parseInt($(el).text().trim(), 10);
      if (!isNaN(p) && p > totalPages) totalPages = p;
    });

    const pageSize = products.length > 0 ? products.length : 36;
    const totalCount = totalPages * pageSize;

    return {
      products,
      meta: {
        totalCount,
        currentPage: pageNumber,
        totalPages: Math.max(totalPages, pageNumber),
        pageSize,
        hasMore: pageNumber < totalPages,
        pageTitle,
        siteType: 'meloni',
      },
    };
  }

  // ====================================================
  // 2. MEGACEDI (megacedi.com - B2B Wholesale Portal)
  // ====================================================
  if (isMegaCedi) {
    const pageTitle = $('title').text().trim() || 'MegaCedi Catalogo Prodotti B2B';
    const deptInfo = resolveMegaCediUrl(targetUrl);
    const category = deptInfo.departmentName;

    // Collect all raw items in the returned department shelf / table
    const rawItems: {
      id: string;
      title: string;
      imgUrl: string;
      packaging?: string;
    }[] = [];

    // Parse from WebPartScaffale.aspx (grid of jqTest2 images)
    $('img.jqTest2').each((_, el) => {
      const img = $(el);
      const id = img.attr('id') || '';
      const alt = img.attr('alt') || '';
      let imgSrc = img.attr('src') || '';
      if (imgSrc.startsWith('/')) imgSrc = `https://www.megacedi.com${imgSrc}`;
      if (imgSrc.startsWith('http:')) imgSrc = imgSrc.replace('http:', 'https:');

      if (id && alt) {
        if (!rawItems.some((item) => item.id === id)) {
          rawItems.push({
            id,
            title: alt.trim(),
            imgUrl: imgSrc,
          });
        }
      }
    });

    // Also parse from WebPartRighe.aspx if table format
    $('tr').each((_, el) => {
      const row = $(el);
      const link = row.find('a[href*="codice_a_barre="]').first();
      const codeMatch = link.attr('href')?.match(/codice_a_barre=([0-9A-Za-z]+)/i);
      const id = codeMatch ? codeMatch[1] : link.text().trim();
      const titleFont = row.find('font[style*="font-size:medium"]').text().trim();
      const confTd = row.find('td').eq(2).text().replace(/\s+/g, ' ').trim();

      if (id && titleFont && !rawItems.some((item) => item.id === id)) {
        rawItems.push({
          id,
          title: titleFont,
          imgUrl: `https://www.megacedi.com/immagini/medie/${id}.jpg`,
          packaging: confTd || undefined,
        });
      }
    });

    // Virtual pagination: Slice 48 items per page for responsive streaming
    const pageSize = 48;
    const totalCount = rawItems.length;
    const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));

    const startIndex = (pageNumber - 1) * pageSize;
    const pagedItems = rawItems.slice(startIndex, startIndex + pageSize);

    pagedItems.forEach((raw) => {
      const detailUrl = `https://www.megacedi.com/SchedaArticoloP.aspx?codice_a_barre=${raw.id}`;
      const medImgUrl = `https://www.megacedi.com/immagini/medie/${raw.id}.jpg`;
      const cleanTitle = raw.title
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '_')
        .replace(/_+/g, '_')
        .substring(0, 50);
      const imageFileName = `megacedi_${raw.id}_${cleanTitle || 'product'}.jpg`;

      // Brand extraction from title (e.g. "ACE CANDEG.DENSO PIU'" -> ACE)
      const brandMatch = raw.title.match(/^([A-Za-z0-9'&.]+)\s+/);
      const brand = brandMatch ? brandMatch[1].toUpperCase() : undefined;

      const cachedEan = eanCache.get(detailUrl) || eanCache.get(raw.id) || '';
      const cachedPack = packagingCache.get(detailUrl) || raw.packaging || undefined;

      products.push({
        id: raw.id,
        sku: raw.id,
        title: raw.title,
        url: detailUrl,
        imageUrl: medImgUrl,
        rawImageUrl: medImgUrl,
        imageFileName,
        originalPrice: null,
        originalPriceFormatted: '',
        discountPrice: null,
        discountPriceFormatted: megaCediSession.authenticated ? '会员专属底价' : '需登录查看底价',
        discountPercent: '0%',
        discountNumeric: 0,
        inStock: true,
        page: pageNumber,
        category,
        ean: cachedEan,
        packaging: cachedPack,
        brand,
        sourceSite: 'megacedi',
        scrapedAt: now,
      });
    });

    return {
      products,
      meta: {
        totalCount: totalCount || products.length,
        currentPage: pageNumber,
        totalPages: Math.max(totalPages, pageNumber),
        pageSize,
        hasMore: pageNumber < totalPages,
        pageTitle: `${pageTitle} (${category})`,
        siteType: 'megacedi',
      },
    };
  }

  // ====================================================
  // 3. MAURY'S ONLINE (maurysonline.it)
  // ====================================================
  if (isMaurys) {
    const defaultCategory = $('.breadcrumb li:last-child').text().trim() ||
      $('h1').text().trim() ||
      'Detersivo e Cura Casa';

    $('.product-item').each((_, element) => {
      const item = $(element);

      // Title & Link
      const linkEl = item.find('.product-title a').first();
      const pictureLinkEl = item.find('.picture a').first();
      let title = linkEl.text().trim() || pictureLinkEl.attr('title')?.trim() || '';
      let productUrl = linkEl.attr('href') || pictureLinkEl.attr('href') || '';
      if (productUrl.startsWith('/')) {
        productUrl = `https://maurysonline.it${productUrl}`;
      }

      if (!title && !productUrl) return;

      // Product ID
      let id = item.attr('data-productid') || '';
      if (!id) {
        const btn = item.find('button[onclick*="catalog/"]');
        const m = btn.attr('onclick')?.match(/catalog\/(\d+)/);
        if (m) id = m[1];
      }
      if (!id) {
        id = Math.random().toString(36).substring(2, 9);
      }

      // Images
      const imgEl = item.find('img.picture-img, .picture img').first();
      let imageUrl = imgEl.attr('data-src') || imgEl.attr('src') || '';

      // High-res master raw image (replace _360.jpeg or _510.jpeg to get original master image)
      let rawImageUrl = imageUrl;
      if (imageUrl) {
        rawImageUrl = imageUrl.replace(/_(?:360|510|105)\.jpe?g/i, '.jpeg');
      }

      // SKU extraction
      let sku = id;
      const skuMatch = imageUrl.match(/_(kit\d+)/i) || title.match(/\b(kit\d+)\b/i);
      if (skuMatch) {
        sku = skuMatch[1].toUpperCase();
      }

      // Prices & Discount
      const oldPriceEl = item.find('.price.old-price');
      const actualPriceEl = item.find('.price.actual-price-red, .price.actual-price, .price');

      const oldPriceText = oldPriceEl.text().trim();
      const actualPriceText = actualPriceEl.first().text().trim();

      const parsePrice = (str: string) => {
        if (!str) return null;
        const clean = str.replace(/[^0-9,\.]/g, '').replace(',', '.');
        const num = parseFloat(clean);
        return isNaN(num) ? null : num;
      };

      const oldPriceNum = parsePrice(oldPriceText);
      let finalPriceNum = parsePrice(actualPriceText);

      let discountNumeric = 0;
      let discountPercent = '0%';
      if (oldPriceNum && finalPriceNum && oldPriceNum > finalPriceNum) {
        discountNumeric = Math.round(((oldPriceNum - finalPriceNum) / oldPriceNum) * 100);
        discountPercent = `-${discountNumeric}%`;
      }

      // Out of stock
      const stockEl = item.find('.stock.unavailable, .out-of-stock');
      const inStock = stockEl.length === 0;

      // Safe clean filename
      const cleanTitle = title
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '_')
        .replace(/_+/g, '_')
        .substring(0, 50);
      const imageFileName = `maurys_${id}_${cleanTitle || 'product'}.jpg`;

      products.push({
        id,
        sku,
        title,
        url: productUrl,
        imageUrl,
        rawImageUrl,
        imageFileName,
        originalPrice: oldPriceNum ? Number(oldPriceNum.toFixed(2)) : null,
        originalPriceFormatted: oldPriceText || (oldPriceNum ? `€${oldPriceNum.toFixed(2)}` : ''),
        discountPrice: finalPriceNum ? Number(finalPriceNum.toFixed(2)) : null,
        discountPriceFormatted: actualPriceText || (finalPriceNum ? `€${finalPriceNum.toFixed(2)}` : ''),
        discountPercent,
        discountNumeric,
        inStock,
        page: pageNumber,
        category: defaultCategory,
        ean: eanCache.get(productUrl) || (sku.startsWith('KIT') ? sku : ''),
        sourceSite: 'maurys',
        scrapedAt: now,
      });
    });

    // Pagination for Maury's
    let totalPages = 1;
    const lastPageEl = $('.pager .last-page a, .pagination .last-page a');
    if (lastPageEl.length) {
      const pageAttr = lastPageEl.attr('data-page');
      if (pageAttr) {
        totalPages = parseInt(pageAttr, 10);
      }
    }
    if (totalPages === 1) {
      $('.pager .individual-page a, .pagination .individual-page a').each((_, el) => {
        const p = parseInt($(el).attr('data-page') || $(el).text().trim(), 10);
        if (!isNaN(p) && p > totalPages) totalPages = p;
      });
    }

    const pageSize = products.length > 0 ? products.length : 30;
    const totalCount = totalPages * pageSize;
    const pageTitle = $('title').text().trim() || "Maury's Online - Detersivo e Cura Casa";

    return {
      products,
      meta: {
        totalCount,
        currentPage: pageNumber,
        totalPages: Math.max(totalPages, pageNumber),
        pageSize,
        hasMore: pageNumber < totalPages,
        pageTitle,
        siteType: 'maurys',
      },
    };
  }

  // ====================================================
  // 4. RISPARMIO CASA (shop.risparmiocasa.com)
  // ====================================================
  $('li.item.product.product-item').each((_, element) => {
    const item = $(element);

    // Product Link & Title
    const linkEl = item.find('a.product-item-link');
    const title = linkEl.text().trim();
    let productUrl = linkEl.attr('href') || '';
    if (productUrl.startsWith('/')) {
      productUrl = `https://shop.risparmiocasa.com${productUrl}`;
    }

    if (!title && !productUrl) return;

    // ID
    const priceBox = item.find('.price-box');
    const id = priceBox.attr('data-product-id') || item.find('input[name="product"]').val()?.toString() || Math.random().toString(36).substring(2, 9);

    // Images
    const imgEl = item.find('img.product-image-photo').first();
    let imageUrl = imgEl.attr('data-src') || imgEl.attr('src') || '';
    if (!imageUrl) {
      const hoverImg = item.find('img.hover_image');
      imageUrl = hoverImg.attr('data-src') || hoverImg.attr('src') || '';
    }

    // Prices
    const oldPriceWrapper = item.find('[data-price-type="oldPrice"]');
    const finalPriceWrapper = item.find('[data-price-type="finalPrice"]');

    const oldPriceVal = oldPriceWrapper.attr('data-price-amount');
    const finalPriceVal = finalPriceWrapper.attr('data-price-amount');

    let oldPriceNum = oldPriceVal ? parseFloat(oldPriceVal) : null;
    let finalPriceNum = finalPriceVal ? parseFloat(finalPriceVal) : null;

    let oldPriceFormatted = oldPriceWrapper.find('.price').text().trim();
    let finalPriceFormatted = finalPriceWrapper.find('.price').text().trim();

    if (!finalPriceFormatted) {
      const genericPrice = item.find('.price').first().text().trim();
      finalPriceFormatted = genericPrice;
      if (!finalPriceNum && genericPrice) {
        const cleaned = genericPrice.replace(/[^0-9,\.]/g, '').replace(',', '.');
        finalPriceNum = parseFloat(cleaned) || null;
      }
    }

    // Discount percentage
    let discountPercent = item.find('.product-label.sale-label').text().trim();
    let discountNumeric = 0;

    if (discountPercent) {
      const numMatch = discountPercent.match(/(\d+)/);
      if (numMatch) {
        discountNumeric = parseInt(numMatch[1], 10);
      }
    } else if (oldPriceNum && finalPriceNum && oldPriceNum > finalPriceNum) {
      discountNumeric = Math.round(((oldPriceNum - finalPriceNum) / oldPriceNum) * 100);
      discountPercent = `-${discountNumeric}%`;
    }

    // Out of stock check
    const stockEl = item.find('.stock.unavailable, .out-of-stock');
    const inStock = stockEl.length === 0;

    // High-resolution original master image: strip dynamic resize query params
    let rawImageUrl = imageUrl;
    if (imageUrl) {
      try {
        const u = new URL(imageUrl);
        u.search = '';
        rawImageUrl = u.toString();
      } catch {}
    }

    // Clean filename
    const cleanTitle = title
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '_')
      .replace(/_+/g, '_')
      .substring(0, 50);
    const imageFileName = `rc_${id}_${cleanTitle || 'product'}.jpg`;

    products.push({
      id,
      sku: id,
      title,
      url: productUrl,
      imageUrl,
      rawImageUrl,
      imageFileName,
      originalPrice: oldPriceNum ? Number(oldPriceNum.toFixed(2)) : null,
      originalPriceFormatted: oldPriceFormatted || (oldPriceNum ? `€${oldPriceNum.toFixed(2)}` : ''),
      discountPrice: finalPriceNum ? Number(finalPriceNum.toFixed(2)) : null,
      discountPriceFormatted: finalPriceFormatted || (finalPriceNum ? `€${finalPriceNum.toFixed(2)}` : ''),
      discountPercent: discountPercent || '0%',
      discountNumeric,
      inStock,
      page: pageNumber,
      category: 'Offerte',
      ean: eanCache.get(productUrl) || '',
      sourceSite: 'risparmiocasa',
      scrapedAt: now,
    });
  });

  // Extract pagination info for Risparmio Casa
  let totalCount = 0;
  const toolbarNumbers = $('.toolbar-number');
  if (toolbarNumbers.length > 0) {
    const lastNumText = toolbarNumbers.last().text().trim();
    const count = parseInt(lastNumText, 10);
    if (!isNaN(count) && count > 0) {
      totalCount = count;
    }
  }

  if (totalCount === 0) {
    const match = html.match(/class="toolbar-number"[^>]*>(\d+)<\/span>/gi);
    if (match && match.length >= 3) {
      const parsed = parseInt(match[match.length - 1].replace(/<[^>]+>/g, ''), 10);
      if (!isNaN(parsed)) totalCount = parsed;
    }
  }

  const pageSize = products.length > 0 ? products.length : 48;
  const totalPages = totalCount > 0 ? Math.ceil(totalCount / 48) : 1;
  const pageTitle = $('title').text().trim() || 'Risparmio Casa Offerte';

  return {
    products,
    meta: {
      totalCount: totalCount || products.length,
      currentPage: pageNumber,
      totalPages: Math.max(totalPages, pageNumber),
      pageSize,
      hasMore: pageNumber < totalPages,
      pageTitle,
      siteType: 'risparmiocasa',
    },
  };
}

// ----------------------------------------------------
// API ROUTES & AUTHENTICATION ENGINES
// ----------------------------------------------------

async function loginMeloniStore(username: string, password: string): Promise<{ success: boolean; cookie?: string; message?: string; error?: string }> {
  try {
    const loginPageUrl = 'https://www.melonistore.com/login.php';
    const initRes = await fetch(loginPageUrl, { headers: DEFAULT_HEADERS });
    const initHtml = await initRes.text();
    const $ = cheerio.load(initHtml);
    const token = $('input[name="authenticity_token"]').val() || '';

    const setCookies = initRes.headers.get('set-cookie') || '';
    const initialCookies = setCookies
      .split(',')
      .map((c) => c.split(';')[0].trim())
      .filter((c) => c.includes('='))
      .join('; ');

    const formParams = new URLSearchParams();
    formParams.set('login_email', username.trim());
    formParams.set('login_pass', password.trim());
    if (token) formParams.set('authenticity_token', token.toString());

    const postRes = await fetch('https://www.melonistore.com/login.php?action=check_login', {
      method: 'POST',
      headers: {
        ...DEFAULT_HEADERS,
        'Content-Type': 'application/x-www-form-urlencoded',
        'Cookie': initialCookies,
        'Referer': loginPageUrl,
      },
      body: formParams.toString(),
      redirect: 'manual',
    });

    const authCookies = postRes.headers.get('set-cookie') || '';
    const merged = [initialCookies, authCookies].filter(Boolean).join('; ');

    return {
      success: true,
      cookie: merged || initialCookies,
      message: `Meloni Store 账号 ${username} 认证请求已提交并保存会话凭据！`,
    };
  } catch (err: any) {
    return {
      success: false,
      error: 'Meloni Store 登录失败: ' + err.message,
    };
  }
}

// Unified Platform Auth Status
app.get('/api/auth/status', (_req: Request, res: Response) => {
  res.json({
    success: true,
    statuses: {
      megacedi: {
        authenticated: megaCediSession.authenticated,
        username: megaCediSession.username,
        hasCustomCookie: !!megaCediSession.cookie,
        siteName: 'MegaCedi',
        domain: 'megacedi.com',
        requiresLogin: true,
        message: megaCediSession.authenticated
          ? `已成功认证 MegaCedi 账号 (${megaCediSession.username || 'Cookie Session'})`
          : '当前为公开目录浏览模式（部分专供商超的 B2B 批发底价受限）',
      },
      meloni: {
        authenticated: meloniSession.authenticated,
        username: meloniSession.username,
        hasCustomCookie: !!meloniSession.cookie,
        siteName: 'Meloni Store',
        domain: 'melonistore.com',
        requiresLogin: true,
        message: meloniSession.authenticated
          ? `已成功认证 Meloni Store 会员 (${meloniSession.username || 'Cookie Session'})`
          : '当前为公开零售价格模式（专供批发品提示 Accedi）',
      },
      maurys: {
        authenticated: maurysSession.authenticated,
        username: maurysSession.username,
        hasCustomCookie: !!maurysSession.cookie,
        siteName: "Maury's Online",
        domain: 'maurysonline.it',
        requiresLogin: false,
        message: maurysSession.authenticated
          ? `已认证会员会话 (${maurysSession.username || 'Cookie'})`
          : '公开零售商城（100% 免登录，全网原价与特价公开）',
      },
      risparmiocasa: {
        authenticated: risparmioSession.authenticated,
        username: risparmioSession.username,
        hasCustomCookie: !!risparmioSession.cookie,
        siteName: 'Risparmio Casa',
        domain: 'shop.risparmiocasa.com',
        requiresLogin: false,
        message: risparmioSession.authenticated
          ? `已认证会员会话 (${risparmioSession.username || 'Cookie'})`
          : '公开连锁大超市（100% 免登录，全网标价与促销公开）',
      },
      piume: {
        authenticated: piumeSession.authenticated,
        username: piumeSession.username,
        hasCustomCookie: !!piumeSession.cookie,
        siteName: 'PiùMe',
        domain: 'piume.it',
        requiresLogin: false,
        message: piumeSession.authenticated
          ? `已认证会员会话 (${piumeSession.username || 'Cookie'})`
          : '公开美妆日化连锁（100% 免登录，全网标价与促销公开）',
      },
    },
  });
});

// Unified Login / Cookie Save Endpoint
app.post('/api/auth/login', async (req: Request, res: Response) => {
  try {
    const { platform = 'megacedi', username, password, temporaryCode, cookie } = req.body;

    // A. Meloni Store
    if (platform === 'meloni') {
      if (cookie && typeof cookie === 'string' && cookie.trim()) {
        const trimmed = cookie.trim();
        meloniSession = {
          cookie: trimmed,
          username: username || 'Meloni商户Cookie',
          authenticated: true,
        };
        return res.json({
          success: true,
          platform: 'meloni',
          message: 'Meloni Store Cookie 保存成功！后续抓取将自动携带该会话。',
          status: {
            authenticated: true,
            username: meloniSession.username,
            hasCustomCookie: true,
          },
        });
      }

      if (!username || !password) {
        return res.status(400).json({ success: false, error: '请提供 Meloni 账号邮箱和密码，或直接填入有效 Cookie' });
      }

      const loginResult = await loginMeloniStore(username, password);
      if (loginResult.success && loginResult.cookie) {
        meloniSession = {
          cookie: loginResult.cookie,
          username: username.trim(),
          authenticated: true,
        };
        return res.json({
          success: true,
          platform: 'meloni',
          message: loginResult.message || `Meloni Store 用户 ${username} 登录成功！`,
          status: {
            authenticated: true,
            username: meloniSession.username,
            hasCustomCookie: true,
          },
        });
      } else {
        return res.status(400).json({ success: false, error: loginResult.error || '登录失败，建议改用 Cookie 模式' });
      }
    }

    // B. MegaCedi
    if (platform === 'megacedi') {
      if (cookie && typeof cookie === 'string' && cookie.trim()) {
        const trimmedCookie = cookie.trim();
        megaCediSession = {
          cookie: trimmedCookie,
          username: username || 'Cookie授权用户',
          authenticated: true,
        };
        return res.json({
          success: true,
          platform: 'megacedi',
          message: 'MegaCedi Cookie 凭证保存成功！后续采集将携带该认证凭据。',
          status: {
            authenticated: true,
            username: megaCediSession.username,
            hasCustomCookie: true,
          },
        });
      }

      if (!username || !password) {
        return res.status(400).json({ success: false, error: '请提供用户名和密码或有效 Cookie' });
      }

      const loginPageUrl = 'https://www.megacedi.com/GestioneUtentiPubblici/home.aspx';
      const initRes = await fetch(loginPageUrl, { headers: DEFAULT_HEADERS });
      const initHtml = await initRes.text();
      const $ = cheerio.load(initHtml);

      const viewState = $('input[name="__VIEWSTATE"]').val() || '';
      const viewStateGen = $('input[name="__VIEWSTATEGENERATOR"]').val() || '';
      const eventValidation = $('input[name="__EVENTVALIDATION"]').val() || '';

      const setCookies = initRes.headers.get('set-cookie') || '';
      const initialSessionCookies = setCookies
        .split(',')
        .map((c) => c.split(';')[0].trim())
        .filter((c) => c.includes('='))
        .join('; ');

      const postUrl = 'https://www.megacedi.com/GestioneUtentiPubblici/Loginpubb.aspx?ReturnUrl=%2fGestioneUtentiPubblici%2fhome.aspx';
      const formParams = new URLSearchParams();
      formParams.set('__EVENTTARGET', '');
      formParams.set('__EVENTARGUMENT', '');
      formParams.set('__VIEWSTATE', viewState.toString());
      if (viewStateGen) formParams.set('__VIEWSTATEGENERATOR', viewStateGen.toString());
      if (eventValidation) formParams.set('__EVENTVALIDATION', eventValidation.toString());
      formParams.set('ctl00$ContentPlaceHolder1$Login1$UserName', username.trim());
      formParams.set('ctl00$ContentPlaceHolder1$Password1', password.trim());
      if (temporaryCode) {
        formParams.set('ctl00$ContentPlaceHolder1$CodAttivazione', temporaryCode.trim());
      }
      formParams.set('ctl00$ContentPlaceHolder1$Login1$LoginButton', 'Accedi');

      const authRes = await fetch(postUrl, {
        method: 'POST',
        headers: {
          ...DEFAULT_HEADERS,
          'Content-Type': 'application/x-www-form-urlencoded',
          'Cookie': initialSessionCookies,
          'Referer': loginPageUrl,
        },
        body: formParams.toString(),
        redirect: 'manual',
      });

      const authCookiesHeader = authRes.headers.get('set-cookie') || '';
      const mergedCookies = [initialSessionCookies, authCookiesHeader].filter(Boolean).join('; ');

      megaCediSession = {
        cookie: mergedCookies || initialSessionCookies,
        username: username.trim(),
        authenticated: true,
      };

      return res.json({
        success: true,
        platform: 'megacedi',
        message: `MegaCedi 用户 ${username} 登录验证已成功提交并保存会话！`,
        status: {
          authenticated: true,
          username,
          hasCustomCookie: true,
        },
      });
    }

    // C. Maury's Online (Optional custom cookie)
    if (platform === 'maurys') {
      if (cookie && typeof cookie === 'string' && cookie.trim()) {
        maurysSession = {
          cookie: cookie.trim(),
          username: username || "Maury's会员Cookie",
          authenticated: true,
        };
      } else if (username) {
        maurysSession = {
          cookie: '',
          username: username.trim(),
          authenticated: true,
        };
      }
      return res.json({
        success: true,
        platform: 'maurys',
        message: "Maury's Online 授权配置已更新！",
        status: {
          authenticated: maurysSession.authenticated,
          username: maurysSession.username,
          hasCustomCookie: !!maurysSession.cookie,
        },
      });
    }

    // D. Risparmio Casa (Optional custom cookie)
    if (platform === 'risparmiocasa') {
      if (cookie && typeof cookie === 'string' && cookie.trim()) {
        risparmioSession = {
          cookie: cookie.trim(),
          username: username || 'RisparmioCasa会员Cookie',
          authenticated: true,
        };
      } else if (username) {
        risparmioSession = {
          cookie: '',
          username: username.trim(),
          authenticated: true,
        };
      }
      return res.json({
        success: true,
        platform: 'risparmiocasa',
        message: 'Risparmio Casa 授权配置已更新！',
        status: {
          authenticated: risparmioSession.authenticated,
          username: risparmioSession.username,
          hasCustomCookie: !!risparmioSession.cookie,
        },
      });
    }

    // E. PiùMe (Optional custom cookie / fidelity login)
    if (platform === 'piume') {
      if (cookie && typeof cookie === 'string' && cookie.trim()) {
        piumeSession = {
          cookie: cookie.trim(),
          username: username || 'PiùMe会员Cookie',
          authenticated: true,
        };
      } else if (username) {
        piumeSession = {
          cookie: '',
          username: username.trim(),
          authenticated: true,
        };
      }
      return res.json({
        success: true,
        platform: 'piume',
        message: 'PiùMe 授权配置已更新！',
        status: {
          authenticated: piumeSession.authenticated,
          username: piumeSession.username,
          hasCustomCookie: !!piumeSession.cookie,
        },
      });
    }

    res.status(400).json({ success: false, error: '未知的平台标识: ' + platform });
  } catch (error: any) {
    res.status(500).json({
      success: false,
      error: '平台登录失败: ' + error.message,
    });
  }
});

// Unified Logout Endpoint
app.post('/api/auth/logout', (req: Request, res: Response) => {
  const { platform } = req.body;
  if (!platform || platform === 'all') {
    megaCediSession = { cookie: '', username: undefined, authenticated: false };
    meloniSession = { cookie: '', username: undefined, authenticated: false };
    maurysSession = { cookie: '', username: undefined, authenticated: false };
    risparmioSession = { cookie: '', username: undefined, authenticated: false };
    piumeSession = { cookie: '', username: undefined, authenticated: false };
  } else if (platform === 'meloni') {
    meloniSession = { cookie: '', username: undefined, authenticated: false };
  } else if (platform === 'megacedi') {
    megaCediSession = { cookie: '', username: undefined, authenticated: false };
  } else if (platform === 'maurys') {
    maurysSession = { cookie: '', username: undefined, authenticated: false };
  } else if (platform === 'risparmiocasa') {
    risparmioSession = { cookie: '', username: undefined, authenticated: false };
  } else if (platform === 'piume') {
    piumeSession = { cookie: '', username: undefined, authenticated: false };
  }
  res.json({ success: true, message: `已清除 ${platform || '所有平台'} 登录会话凭据` });
});

// Meloni Session Status
app.get('/api/meloni/status', (_req: Request, res: Response) => {
  res.json({
    success: true,
    status: {
      authenticated: meloniSession.authenticated,
      username: meloniSession.username,
      hasCustomCookie: !!meloniSession.cookie,
      message: meloniSession.authenticated
        ? `已成功认证 Meloni Store 会员 (${meloniSession.username || 'Cookie Session'})`
        : '当前为公开零售模式（专供 B2B 商品提示需登录）',
    },
  });
});

// Meloni Login
app.post('/api/meloni/login', async (req: Request, res: Response) => {
  req.body.platform = 'meloni';
  return app._router.handle(req, res, () => {});
});

// Meloni Logout
app.post('/api/meloni/logout', (_req: Request, res: Response) => {
  meloniSession = { cookie: '', username: undefined, authenticated: false };
  res.json({ success: true, message: 'Meloni Store 会话已清除' });
});

// MegaCedi Session Status (Backward Compatibility)
app.get('/api/megacedi/status', (_req: Request, res: Response) => {
  res.json({
    success: true,
    status: {
      authenticated: megaCediSession.authenticated,
      username: megaCediSession.username,
      hasCustomCookie: !!megaCediSession.cookie,
      message: megaCediSession.authenticated
        ? `已成功认证 MegaCedi 账号 (${megaCediSession.username || 'Cookie Session'})`
        : '当前为公开目录浏览模式（支持采集全品类商品库、条形码与包装规格）',
    },
  });
});

// MegaCedi Login (Backward Compatibility)
app.post('/api/megacedi/login', async (req: Request, res: Response) => {
  req.body.platform = 'megacedi';
  return app._router.handle(req, res, () => {});
});

// MegaCedi Logout (Backward Compatibility)
app.post('/api/megacedi/logout', (_req: Request, res: Response) => {
  megaCediSession = { cookie: '', username: undefined, authenticated: false };
  res.json({ success: true, message: 'MegaCedi 会话已清除' });
});

// 1. API: Get page info & initial count
app.get('/api/info', async (req: Request, res: Response) => {
  try {
    const targetUrl = (req.query.url as string) || 'https://maurysonline.it/detersivo-e-cura-casa';
    const isAllPlatforms = targetUrl.includes('all-platforms') || targetUrl === 'https://casa-te.all/products';
    const isMaurysAll = !isAllPlatforms && (targetUrl.includes('maurysonline.it/all-products') || targetUrl === 'https://maurysonline.it/all');

    if (isAllPlatforms) {
      const [rcHtml, maurysHtml] = await Promise.all([
        fetchPageHtml('https://shop.risparmiocasa.com/prodotti'),
        fetchPageHtml('https://maurysonline.it/detersivo-e-cura-casa'),
      ]);
      const rcParsed = parsePageProducts(rcHtml, 1, 'https://shop.risparmiocasa.com/prodotti');
      const maurysParsed = parsePageProducts(maurysHtml, 1, 'https://maurysonline.it/detersivo-e-cura-casa');
      const sampleProducts = [...rcParsed.products.slice(0, 3), ...maurysParsed.products.slice(0, 3)];
      await Promise.all(
        sampleProducts.map(async (prod) => {
          prod.ean = await fetchProductEan(prod.url);
        })
      );
      return res.json({
        success: true,
        url: targetUrl,
        title: "CASA & TE - 全部平台所有产品数据 (Maury's Online + Risparmio Casa + Meloni Store + MegaCedi)",
        totalCount: 4383 + MAURYS_ALL_TOTAL_PAGES * 30 + 10000,
        totalPages: ALL_PLATFORMS_TOTAL_PAGES,
        siteType: 'generic',
        sampleProducts,
        pageSize: 48,
      });
    }

    if (isMaurysAll) {
      const html = await fetchPageHtml('https://maurysonline.it/detersivo-e-cura-casa');
      const { products } = parsePageProducts(html, 1, 'https://maurysonline.it/detersivo-e-cura-casa');
      const sampleProducts = products.slice(0, 6);
      await Promise.all(
        sampleProducts.map(async (prod) => {
          prod.ean = await fetchProductEan(prod.url);
        })
      );
      return res.json({
        success: true,
        url: targetUrl,
        title: "Maury's Online - 全站全品类全量商品 (All Categories)",
        totalCount: MAURYS_ALL_TOTAL_PAGES * 30,
        totalPages: MAURYS_ALL_TOTAL_PAGES,
        siteType: 'maurys',
        sampleProducts,
        pageSize: 30,
      });
    }

    const normalized = normalizeUrl(targetUrl, 1);
    const html = await fetchPageHtml(normalized);
    const { products, meta } = parsePageProducts(html, 1, targetUrl);

    const sampleProducts = products.slice(0, 6);
    // Enrich sample products with real EANs/SKUs
    await Promise.all(
      sampleProducts.map(async (prod) => {
        if (!prod.ean) {
          prod.ean = await fetchProductEan(prod.url);
        }
      })
    );

    res.json({
      success: true,
      url: targetUrl,
      title: meta.pageTitle,
      totalCount: meta.totalCount,
      totalPages: meta.totalPages,
      siteType: meta.siteType,
      sampleProducts,
      pageSize: meta.pageSize,
    });
  } catch (error: any) {
    res.status(500).json({
      success: false,
      error: error.message || '获取网页信息失败',
    });
  }
});

// 2. API: Server-Sent Events (SSE) streaming for real-time batch scraping
app.get('/api/scrape/stream', async (req: Request, res: Response) => {
  const targetUrl = (req.query.url as string) || 'https://maurysonline.it/detersivo-e-cura-casa';
  const startPage = Math.max(1, parseInt((req.query.startPage as string) || '1', 10));
  const maxPagesRequested = parseInt((req.query.endPage as string) || '3', 10);
  const delayMs = Math.min(2000, Math.max(200, parseInt((req.query.delay as string) || '400', 10)));

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders?.();

  const sendEvent = (event: string, data: any) => {
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  };

  let isAborted = false;
  req.on('close', () => {
    isAborted = true;
  });

  try {
    const isAllPlatforms = targetUrl.includes('all-platforms') || targetUrl === 'https://casa-te.all/products';
    const isMaurysAll = !isAllPlatforms && (targetUrl.includes('maurysonline.it/all-products') || targetUrl === 'https://maurysonline.it/all');
    const isMeloni = targetUrl.includes('melonistore.com');
    const isMegaCedi = targetUrl.includes('megacedi.com');

    const siteLabel = isAllPlatforms
      ? '全部平台聚合'
      : isMaurysAll
      ? "Maury's Online (全站全量)"
      : isMeloni
      ? 'Meloni Store'
      : isMegaCedi
      ? 'MegaCedi'
      : targetUrl.includes('maurysonline.it')
      ? "Maury's Online"
      : 'Risparmio Casa';

    sendEvent('status', { message: `开始检测 [${siteLabel}] 目标页面元数据: ${targetUrl}` });

    let availableTotalPages = 1;
    let totalCount = 0;
    let siteType: 'maurys' | 'risparmiocasa' | 'meloni' | 'megacedi' | 'generic' = isAllPlatforms
      ? 'generic'
      : isMaurysAll
      ? 'maurys'
      : isMeloni
      ? 'meloni'
      : isMegaCedi
      ? 'megacedi'
      : 'generic';
    let pageTitle = '';
    let firstPageHtml = '';

    if (isAllPlatforms) {
      availableTotalPages = ALL_PLATFORMS_TOTAL_PAGES;
      totalCount = 4383 + MAURYS_ALL_TOTAL_PAGES * 30;
      siteType = 'generic';
      pageTitle = "全部平台所有产品数据 (Maury's Online + Risparmio Casa + Meloni Store + MegaCedi)";
    } else if (isMaurysAll) {
      availableTotalPages = MAURYS_ALL_TOTAL_PAGES;
      totalCount = MAURYS_ALL_TOTAL_PAGES * 30;
      siteType = 'maurys';
      pageTitle = "Maury's Online - 全站全品类全量商品";
    } else {
      const firstTarget = normalizeUrl(targetUrl, 1);
      firstPageHtml = await fetchPageHtml(firstTarget);
      const firstParse = parsePageProducts(firstPageHtml, 1, targetUrl);
      availableTotalPages = firstParse.meta.totalPages;
      totalCount = firstParse.meta.totalCount;
      siteType = firstParse.meta.siteType;
      pageTitle = firstParse.meta.pageTitle;
    }

    const actualEndPage = Math.min(maxPagesRequested, availableTotalPages);

    sendEvent('meta', {
      totalCount,
      availableTotalPages,
      startPage,
      endPage: actualEndPage,
      siteType,
      pageTitle,
    });

    let allScrapedCount = 0;
    const seenProductIds = new Set<string>();

    for (let p = startPage; p <= actualEndPage; p++) {
      if (isAborted) break;

      const platInfo = isAllPlatforms ? resolveAllPlatformsPage(p) : null;
      const catInfo = isMaurysAll ? resolveMaurysAllPage(p) : null;
      const currentUrl = isAllPlatforms ? platInfo!.resolvedUrl : normalizeUrl(targetUrl, p);

      sendEvent('page_start', {
        page: p,
        totalPages: actualEndPage,
        url: currentUrl,
      });

      if (isAllPlatforms && platInfo) {
        sendEvent('status', {
          message: `[${platInfo.platformName} · ${platInfo.categoryName}] 正在抓取第 ${platInfo.pageInPlatform} 页 (全部平台总进度 P.${p}/${actualEndPage})...`,
        });
      } else if (isMaurysAll && catInfo) {
        sendEvent('status', {
          message: `正在抓取全站分类 [${catInfo.categoryName}] (第 ${catInfo.pageInCategory} 页 · 总第 ${p}/${actualEndPage} 页)...`,
        });
      } else if (isMegaCedi) {
        sendEvent('status', {
          message: `[MegaCedi] 正在抓取第 ${p} 页 (${(p - 1) * 48 + 1} - ${p * 48} 件)...`,
        });
      }

      let pageHtml = '';
      if (p === 1 && !isAllPlatforms && !isMaurysAll && firstPageHtml) {
        pageHtml = firstPageHtml;
      } else {
        if (p > startPage) {
          await new Promise((r) => setTimeout(r, delayMs));
        }
        pageHtml = await fetchPageHtml(currentUrl);
      }

      const { products } = parsePageProducts(pageHtml, p, currentUrl);

      // Deduplicate products
      const uniqueProducts: ScrapedProduct[] = [];
      for (const prod of products) {
        if (!seenProductIds.has(prod.id)) {
          seenProductIds.add(prod.id);
          uniqueProducts.push(prod);
        }
      }

      const shouldFetchEan = req.query.fetchEan !== 'false';
      if (shouldFetchEan && uniqueProducts.length > 0) {
        // Only fetch EAN for products that don't already have one
        const needEan = uniqueProducts.filter((pr) => !pr.ean);
        if (needEan.length > 0) {
          sendEvent('status', { message: `正在抓取第 ${p} 页商品的 Codice / 官方 13 位 EAN 条形码...` });
          for (let i = 0; i < needEan.length; i += 6) {
            const chunk = needEan.slice(i, i + 6);
            await Promise.all(
              chunk.map(async (prod) => {
                prod.ean = await fetchProductEan(prod.url);
              })
            );
          }
        }
      }

      allScrapedCount += uniqueProducts.length;

      sendEvent('page_data', {
        page: p,
        totalPages: actualEndPage,
        itemsCount: uniqueProducts.length,
        accumulatedCount: allScrapedCount,
        products: uniqueProducts,
      });

      if (products.length === 0 && !isAllPlatforms && !isMaurysAll && !isMegaCedi) {
        break;
      }
    }

    sendEvent('complete', {
      totalScraped: allScrapedCount,
      message: `🎉 成功采集 ${allScrapedCount} 件商品数据！`,
    });
    res.end();
  } catch (error: any) {
    sendEvent('error', {
      error: error.message || '采集过程发生错误',
    });
    res.end();
  }
});

// 3. API: Batch Scrape (standard POST endpoint)
app.post('/api/scrape/batch', async (req: Request, res: Response) => {
  try {
    const { url = 'https://maurysonline.it/detersivo-e-cura-casa', startPage = 1, endPage = 1, delayMs = 300 } = req.body;
    const start = Math.max(1, Number(startPage));
    const end = Math.max(start, Number(endPage));

    const allProducts: ScrapedProduct[] = [];
    const seenIds = new Set<string>();

    for (let p = start; p <= end; p++) {
      const pageUrl = normalizeUrl(url, p);
      const html = await fetchPageHtml(pageUrl);
      const { products } = parsePageProducts(html, p, url);

      for (const prod of products) {
        if (!seenIds.has(prod.id)) {
          seenIds.add(prod.id);
          allProducts.push(prod);
        }
      }

      if (products.length === 0) break;
      if (p < end) {
        await new Promise((r) => setTimeout(r, delayMs));
      }
    }

    res.json({
      success: true,
      count: allProducts.length,
      products: allProducts,
    });
  } catch (error: any) {
    res.status(500).json({
      success: false,
      error: error.message || '采集失败',
    });
  }
});

// 4. API: Enrich products with EAN
app.post('/api/scrape/enrich-eans', async (req: Request, res: Response) => {
  try {
    const { products = [] } = req.body;
    for (let i = 0; i < products.length; i += 6) {
      const chunk = products.slice(i, i + 6);
      await Promise.all(
        chunk.map(async (prod: ScrapedProduct) => {
          if (!prod.ean && prod.url) {
            prod.ean = await fetchProductEan(prod.url);
          }
        })
      );
    }
    res.json({ success: true, count: products.length, products });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 5. API: Image Proxy endpoint to avoid CORS issues when bundling ZIP on client
app.get('/api/proxy-image', async (req: Request, res: Response) => {
  const imageUrl = req.query.url as string;
  if (!imageUrl) {
    return res.status(400).send('Missing url param');
  }

  try {
    const isMaurys = imageUrl.includes('maurysonline.it');
    const isMeloni = imageUrl.includes('melonistore.com') || imageUrl.includes('bigcommerce.com');
    const isMegaCedi = imageUrl.includes('megacedi.com');

    const referer = isMaurys
      ? 'https://maurysonline.it/'
      : isMeloni
      ? 'https://www.melonistore.com/'
      : isMegaCedi
      ? 'https://www.megacedi.com/'
      : 'https://shop.risparmiocasa.com/';

    const fetchRes = await fetch(imageUrl, {
      headers: {
        'User-Agent': DEFAULT_HEADERS['User-Agent'],
        'Accept': 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
        'Referer': referer,
      },
    });

    if (!fetchRes.ok) {
      return res.status(fetchRes.status).send(`Failed to fetch image: ${fetchRes.statusText}`);
    }

    const contentType = fetchRes.headers.get('content-type') || 'image/jpeg';
    res.setHeader('Content-Type', contentType);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Cache-Control', 'public, max-age=86400');

    const buffer = await fetchRes.arrayBuffer();
    res.send(Buffer.from(buffer));
  } catch (err: any) {
    res.status(500).send('Proxy error: ' + err.message);
  }
});

// ----------------------------------------------------
// BARCODE IMAGE MATCHER & LOCAL SERVER STORAGE ENGINE
// ----------------------------------------------------
let meloniEanIndex: Map<string, string> | null = null;
let isIndexingMeloni = false;

async function getOrBuildMeloniIndex(): Promise<Map<string, string>> {
  if (meloniEanIndex && meloniEanIndex.size > 0) return meloniEanIndex;

  if (isIndexingMeloni) {
    while (isIndexingMeloni) {
      await new Promise((r) => setTimeout(r, 100));
    }
    if (meloniEanIndex && meloniEanIndex.size > 0) return meloniEanIndex;
  }

  isIndexingMeloni = true;
  try {
    const [res1, res2] = await Promise.all([
      fetch('https://www.melonistore.com/xmlsitemap.php?type=products&page=1', { headers: DEFAULT_HEADERS }),
      fetch('https://www.melonistore.com/xmlsitemap.php?type=products&page=2', { headers: DEFAULT_HEADERS }),
    ]);

    const [xml1, xml2] = await Promise.all([res1.text(), res2.text()]);
    const allXml = xml1 + xml2;

    const map = new Map<string, string>();
    const matches = allXml.matchAll(/<loc>(https:\/\/www\.melonistore\.com\/[^<]+-([0-9]{8,14})\.html)<\/loc>/g);
    for (const m of matches) {
      map.set(m[2], m[1]);
    }

    meloniEanIndex = map;
    return map;
  } catch (err) {
    console.error('Failed to build Meloni EAN index:', err);
    return meloniEanIndex || new Map();
  } finally {
    isIndexingMeloni = false;
  }
}

async function matchAndSaveBarcode(
  barcode: string,
  forceRefetch = false
): Promise<{
  barcode: string;
  matched: boolean;
  provider: 'meloni' | 'megacedi' | 'cached' | 'none';
  title?: string;
  imageUrl?: string;
  rawImageUrl?: string;
  sourceUrl?: string;
  filename?: string;
  fileSize?: number;
  message?: string;
  savedAt?: string;
}> {
  const cleanBarcode = barcode.trim();
  if (!cleanBarcode) {
    return { barcode, matched: false, provider: 'none', message: '条形码为空' };
  }

  const localFilePath = path.join(SAVED_IMAGES_DIR, `${cleanBarcode}.jpg`);

  // Step 0: Check if file already exists in local server cache
  if (!forceRefetch && fs.existsSync(localFilePath)) {
    try {
      const stat = fs.statSync(localFilePath);
      if (stat.size > 500) {
        const meta = barcodeMetadataMap.get(cleanBarcode);
        return {
          barcode: cleanBarcode,
          matched: true,
          provider: (meta?.provider as any) || 'cached',
          title: meta?.title,
          filename: `${cleanBarcode}.jpg`,
          imageUrl: `/saved_images/${cleanBarcode}.jpg`,
          rawImageUrl: barcodeCdnMap.get(cleanBarcode) || '',
          sourceUrl: meta?.sourceUrl,
          fileSize: stat.size,
          savedAt: stat.mtime.toISOString(),
          message: '已存在于服务器图库',
        };
      }
    } catch {}
  }

  // Step 1: Priority 1 - Search Meloni Store (meloni Map)
  try {
    const meloniMap = await getOrBuildMeloniIndex();
    let prodUrl = meloniMap.get(cleanBarcode);

    // If not in sitemap index, try direct search on Meloni
    if (!prodUrl) {
      try {
        const searchUrl = `https://www.melonistore.com/search.php?search_query=${encodeURIComponent(cleanBarcode)}`;
        const searchHtml = await fetchPageHtml(searchUrl);
        const $s = cheerio.load(searchHtml);
        const firstCardLink = $s('.card-title a').first().attr('href');
        if (firstCardLink) {
          prodUrl = firstCardLink;
          meloniMap.set(cleanBarcode, firstCardLink);
        }
      } catch {}
    }

    if (prodUrl) {
      const html = await fetchPageHtml(prodUrl);
      const $ = cheerio.load(html);

      const title =
        $('h1.productView-title').text().trim() ||
        $('h1').first().text().trim() ||
        $('title').text().trim();

      // Extract high resolution product picture
      let imgUrl =
        $('meta[property="og:image"]').attr('content') ||
        $('img.productView-image--default, img[src*="stencil"]').attr('src') ||
        $('img.card-image').attr('src') ||
        '';

      if (imgUrl) {
        // Upgrade BigCommerce stencil resolution to 1280x1280
        imgUrl = imgUrl.replace(/\/stencil\/\d+x?\d*w?\//, '/stencil/1280x1280/');
        imgUrl = imgUrl.replace(/\.\d+\.\d+\.(png|jpg|jpeg)/i, '.1280.1280.$1');
      }

      if (imgUrl) {
        saveCdnMapping(cleanBarcode, imgUrl, { title, sourceUrl: prodUrl, provider: 'meloni' });
        const imgRes = await fetch(imgUrl, {
          headers: {
            ...DEFAULT_HEADERS,
            Referer: 'https://www.melonistore.com/',
          },
        });

        if (imgRes.ok) {
          const buf = await imgRes.arrayBuffer();
          fs.writeFileSync(localFilePath, Buffer.from(buf));
          return {
            barcode: cleanBarcode,
            matched: true,
            provider: 'meloni',
            title,
            imageUrl: `/saved_images/${cleanBarcode}.jpg`,
            rawImageUrl: imgUrl,
            sourceUrl: prodUrl,
            filename: `${cleanBarcode}.jpg`,
            fileSize: buf.byteLength,
            savedAt: new Date().toISOString(),
            message: '从 Meloni Store 成功匹配并保存',
          };
        }
      }
    }
  } catch (err: any) {
    console.error(`Meloni lookup error for ${cleanBarcode}:`, err.message);
  }

  // Step 2: Priority 2 - Search MegaCedi
  try {
    const megaSearchUrl = `https://www.megacedi.com/WebPartScaffale.aspx?p=${encodeURIComponent(cleanBarcode)}&r=%`;
    const html = await fetchPageHtml(megaSearchUrl);
    const $ = cheerio.load(html);

    const img = $('img.jqTest2').first();
    if (img.length > 0) {
      const internalId = img.attr('id') || '';
      const title = img.attr('alt')?.trim() || '';

      if (internalId) {
        let imgUrl = `https://www.megacedi.com/immagini/medie/${internalId}.jpg`;
        let imgRes = await fetch(imgUrl, {
          headers: {
            ...DEFAULT_HEADERS,
            Referer: 'https://www.megacedi.com/',
          },
        });

        // Fallback to piccole if medie 404
        if (!imgRes.ok) {
          imgUrl = `https://www.megacedi.com/immagini/piccole/${internalId}.jpg`;
          imgRes = await fetch(imgUrl, {
            headers: {
              ...DEFAULT_HEADERS,
              Referer: 'https://www.megacedi.com/',
            },
          });
        }

        if (imgRes.ok) {
          const megaDetailUrl = `https://www.megacedi.com/SchedaArticoloP.aspx?codice_a_barre=${internalId}`;
          saveCdnMapping(cleanBarcode, imgUrl, { title, sourceUrl: megaDetailUrl, provider: 'megacedi' });
          const buf = await imgRes.arrayBuffer();
          fs.writeFileSync(localFilePath, Buffer.from(buf));
          return {
            barcode: cleanBarcode,
            matched: true,
            provider: 'megacedi',
            title,
            imageUrl: `/saved_images/${cleanBarcode}.jpg`,
            rawImageUrl: imgUrl,
            sourceUrl: megaDetailUrl,
            filename: `${cleanBarcode}.jpg`,
            fileSize: buf.byteLength,
            savedAt: new Date().toISOString(),
            message: '从 MegaCedi 成功匹配并保存',
          };
        }
      }
    }
  } catch (err: any) {
    console.error(`MegaCedi lookup error for ${cleanBarcode}:`, err.message);
  }

  // Step 3: Neither supplier has it
  return {
    barcode: cleanBarcode,
    matched: false,
    provider: 'none',
    filename: `${cleanBarcode}.jpg`,
    message: '在 Meloni 与 MegaCedi 均未检索到此条形码',
  };
}

// API: Stream Barcode Matcher (SSE)
app.post('/api/matcher/stream', async (req: Request, res: Response) => {
  const { barcodes = [], forceRefetch = false } = req.body;
  const rawList: string[] = Array.isArray(barcodes) ? barcodes : [];

  // Deduplicate and clean barcodes
  const uniqueBarcodes = Array.from(
    new Set(rawList.map((b) => String(b).replace(/[^0-9A-Za-z]/g, '').trim()).filter((b) => b.length >= 6))
  );

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders?.();

  const sendEvent = (event: string, data: any) => {
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  };

  let isAborted = false;
  res.on('close', () => {
    isAborted = true;
  });

  const total = uniqueBarcodes.length;
  sendEvent('start', {
    total,
    message: `准备开始批量匹配与存储 ${total} 个商品条码...`,
  });

  // Pre-load Meloni Store index once before loop
  sendEvent('status', { message: '正在加载/更新 Meloni Store 全网商品索引 (约1.3万件)...' });
  await getOrBuildMeloniIndex();
  sendEvent('status', { message: '索引就绪，开始检索商品图片...' });

  let processed = 0;
  let matchedMeloni = 0;
  let matchedMega = 0;
  let matchedCached = 0;
  let unmatched = 0;

  // Process in small concurrency batches of 3 to balance speed and gentle load
  const concurrency = 3;
  for (let i = 0; i < uniqueBarcodes.length; i += concurrency) {
    if (isAborted) break;

    const chunk = uniqueBarcodes.slice(i, i + concurrency);
    const results = await Promise.all(
      chunk.map(async (code) => {
        return await matchAndSaveBarcode(code, forceRefetch);
      })
    );

    for (const r of results) {
      processed++;
      if (r.matched) {
        if (r.provider === 'meloni') matchedMeloni++;
        else if (r.provider === 'megacedi') matchedMega++;
        else if (r.provider === 'cached') matchedCached++;
      } else {
        unmatched++;
      }

      sendEvent('item', r);
    }

    sendEvent('progress', {
      processed,
      total,
      matchedMeloni,
      matchedMega,
      matchedCached,
      unmatched,
      percent: Math.round((processed / total) * 100),
      currentBarcode: chunk[chunk.length - 1],
    });

    // Small delay between chunks
    await new Promise((r) => setTimeout(r, 150));
  }

  sendEvent('complete', {
    total,
    processed,
    matchedTotal: matchedMeloni + matchedMega + matchedCached,
    matchedMeloni,
    matchedMega,
    matchedCached,
    unmatched,
    message: `🎉 条码批量图库检索完成！已成功匹配并入库 ${matchedMeloni + matchedMega + matchedCached} 张高清商品图片。`,
  });

  res.end();
});

// API: Get list of all saved images in local server storage
app.get('/api/matcher/saved-list', (_req: Request, res: Response) => {
  try {
    if (!fs.existsSync(SAVED_IMAGES_DIR)) {
      return res.json({ success: true, count: 0, images: [] });
    }

    const files = fs.readdirSync(SAVED_IMAGES_DIR);
    const images = files
      .filter((f) => f.endsWith('.jpg') || f.endsWith('.png') || f.endsWith('.jpeg'))
      .map((f) => {
        const fullPath = path.join(SAVED_IMAGES_DIR, f);
        const stat = fs.statSync(fullPath);
        const barcode = f.replace(/\.[^.]+$/, '');
        return {
          barcode,
          filename: f,
          size: stat.size,
          url: `/saved_images/${f}`,
          modifiedAt: stat.mtime.toISOString(),
        };
      });

    res.json({
      success: true,
      count: images.length,
      images,
    });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// API: Delete single saved image
app.delete('/api/matcher/image/:barcode', (req: Request, res: Response) => {
  try {
    const { barcode } = req.params;
    const filePath = path.join(SAVED_IMAGES_DIR, `${barcode}.jpg`);
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
      return res.json({ success: true, message: `已从服务器删除条码 ${barcode}.jpg 图片` });
    }
    res.status(404).json({ success: false, error: '文件不存在' });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// API: Clear all saved images
app.post('/api/matcher/clear-all', (_req: Request, res: Response) => {
  try {
    if (fs.existsSync(SAVED_IMAGES_DIR)) {
      const files = fs.readdirSync(SAVED_IMAGES_DIR);
      for (const f of files) {
        fs.unlinkSync(path.join(SAVED_IMAGES_DIR, f));
      }
    }
    res.json({ success: true, message: '服务器条码图片库已全部清空' });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// API: Download all saved images as ZIP (using STORE for instant, uncompressed packaging)
app.all('/api/matcher/download-zip', async (req: Request, res: Response) => {
  try {
    if (!fs.existsSync(SAVED_IMAGES_DIR)) {
      return res.status(400).json({ success: false, error: '服务器图库暂无已保存的图片。请先在上方输入条码并点击「开始批量检索与入库」！' });
    }

    let files = fs.readdirSync(SAVED_IMAGES_DIR).filter((f) => f.endsWith('.jpg') || f.endsWith('.png'));
    
    // Optional filter by requested barcodes
    const reqBarcodes = req.body?.barcodes || (typeof req.query?.barcodes === 'string' ? req.query.barcodes.split(',') : null);
    if (Array.isArray(reqBarcodes) && reqBarcodes.length > 0) {
      const barcodeSet = new Set(reqBarcodes.map((b: string) => String(b).replace(/[^0-9A-Za-z]/g, '').trim()));
      files = files.filter((f) => barcodeSet.has(f.replace(/\.[^.]+$/, '')));
    }

    // Optional pagination / chunking to allow downloading in smaller parts (e.g. 500 items per part)
    const offset = Math.max(0, parseInt(String(req.query?.offset || req.body?.offset || '0'), 10) || 0);
    const limit = Math.max(0, parseInt(String(req.query?.limit || req.body?.limit || '0'), 10) || 0);
    const totalFilesCount = files.length;

    if (limit > 0) {
      files = files.slice(offset, offset + limit);
    }

    if (files.length === 0) {
      return res.status(400).json({ success: false, error: '指定分卷范围内暂无图片！' });
    }

    const zip = new JSZip();
    const mappingRows = ['条形码(Barcode),图片文件名(Filename),服务器相对URL(ServerURL)'];

    for (const f of files) {
      const fullPath = path.join(SAVED_IMAGES_DIR, f);
      try {
        const buffer = fs.readFileSync(fullPath);
        zip.file(f, buffer);
        const barcode = f.replace(/\.[^.]+$/, '');
        mappingRows.push(`"${barcode}","${f}","/saved_images/${f}"`);
      } catch {}
    }

    zip.file('mapping_index.csv', '\uFEFF' + mappingRows.join('\r\n'));

    // Use STORE (no compression) because JPGs are already compressed; this is instant & uses low CPU
    const zipBuffer = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'STORE',
    });

    const filename = limit > 0
      ? `barcode_images_part_${offset + 1}_to_${offset + files.length}_of_${totalFilesCount}.zip`
      : `barcode_images_all_${files.length}items.zip`;

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', String(zipBuffer.length));
    res.setHeader('Access-Control-Expose-Headers', 'Content-Length, Content-Disposition');
    res.send(zipBuffer);
  } catch (err: any) {
    console.error('Error generating ZIP:', err);
    res.status(500).json({ success: false, error: '打包ZIP失败: ' + err.message });
  }
});

// API: Export Excel with ACTUAL EMBEDDED IMAGES (Zero network / zero login needed)
app.post('/api/matcher/export-embedded-xlsx', async (req: Request, res: Response) => {
  try {
    const { items = [] } = req.body;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'CASA & TE';
    const ws = wb.addWorksheet('条码图片库');

    ws.columns = [
      { header: '商品图片 (已内嵌)', key: 'imgPlaceholder', width: 14 },
      { header: '条形码 (Barcode)', key: 'barcode', width: 18 },
      { header: '匹配状态', key: 'status', width: 15 },
      { header: '供应商来源', key: 'provider', width: 16 },
      { header: '商品名称 (Title)', key: 'title', width: 45 },
      { header: '本地图片文件名', key: 'filename', width: 22 },
      { header: '原厂免登录公开CDN图源', key: 'cdnUrl', width: 50 },
      { header: 'Excel免登录显图公式', key: 'formula', width: 55 },
      { header: '供应商详情链接', key: 'sourceUrl', width: 35 },
    ];

    // Style header row
    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E293B' },
    };
    headerRow.height = 26;

    let rowIndex = 2;
    for (const item of items) {
      const barcode = String(item.barcode || '').trim();
      const cdnUrl = item.rawImageUrl || barcodeCdnMap.get(barcode) || '';
      const formula = cdnUrl ? `=IMAGE("${cdnUrl}")` : '';

      const row = ws.addRow({
        imgPlaceholder: '',
        barcode: barcode,
        status: item.matched ? '匹配成功 (已入库)' : '未找到',
        provider:
          item.provider === 'meloni'
            ? 'Meloni Store'
            : item.provider === 'megacedi'
            ? 'MegaCedi'
            : item.provider === 'cached'
            ? '本地已存'
            : '无',
        title: item.title || '',
        filename: `${barcode}.jpg`,
        cdnUrl: cdnUrl,
        formula: formula,
        sourceUrl: item.sourceUrl || '',
      });

      const localImgPath = path.join(SAVED_IMAGES_DIR, `${barcode}.jpg`);
      let imgBuf: Buffer | null = null;
      if (fs.existsSync(localImgPath)) {
        try {
          imgBuf = fs.readFileSync(localImgPath);
        } catch {}
      } else if (cdnUrl) {
        try {
          const resp = await fetch(cdnUrl, { signal: AbortSignal.timeout(6000) });
          if (resp.ok) {
            imgBuf = Buffer.from(await resp.arrayBuffer());
            try { fs.writeFileSync(localImgPath, imgBuf); } catch {}
          }
        } catch {}
      }

      if (imgBuf) {
        try {
          const imgId = wb.addImage({
            buffer: imgBuf as any,
            extension: 'jpeg',
          });
          row.height = 60;
          ws.addImage(imgId, {
            tl: { col: 0.1, row: rowIndex - 0.9 },
            ext: { width: 55, height: 55 },
            editAs: 'oneCell',
          });
        } catch {}
      }

      rowIndex++;
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="barcode_embedded_images_${items.length}items.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err: any) {
    console.error('Error generating embedded Excel:', err);
    res.status(500).send('导出嵌入图片 Excel 失败: ' + err.message);
  }
});

// API: Fill custom user-uploaded Excel template with matched images and titles
app.post('/api/matcher/fill-template', async (req: Request, res: Response) => {
  try {
    const {
      templateBase64,
      filename = 'custom_template.xlsx',
      sheetName,
      sheetIndex = 1,
      barcodeColIndex = 1, // 1-based column index
      titleColIndex, // optional 1-based column index
      imageColIndex, // 1-based column index
      fillType = 'embedded', // 'embedded' | 'formula' | 'cdnUrl'
      startRow = 2,
    } = req.body;

    if (!templateBase64) {
      return res.status(400).json({ success: false, error: '请上传有效的 Excel 模板文件 (.xlsx)' });
    }

    const templateBuffer = Buffer.from(templateBase64, 'base64');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(templateBuffer as any);

    const ws = sheetName ? wb.getWorksheet(sheetName) : wb.getWorksheet(sheetIndex);
    if (!ws) {
      return res.status(400).json({ success: false, error: `未找到指定工作表: ${sheetName || sheetIndex}` });
    }

    let filledImagesCount = 0;
    let filledTitlesCount = 0;
    const totalRows = ws.rowCount;

    // Scan each row in the user's template
    for (let rowNumber = startRow; rowNumber <= totalRows; rowNumber++) {
      const row = ws.getRow(rowNumber);
      const rawBarcode = row.getCell(barcodeColIndex).value;
      if (!rawBarcode) continue;

      // Extract numeric/alphanumeric barcode
      const cleanBarcode = String(rawBarcode).replace(/[^0-9A-Za-z]/g, '').trim();
      if (!cleanBarcode || cleanBarcode.length < 6) continue;

      // Look up cached metadata
      const meta = barcodeMetadataMap.get(cleanBarcode);
      const cdnUrl = barcodeCdnMap.get(cleanBarcode) || '';
      const localFilePath = path.join(SAVED_IMAGES_DIR, `${cleanBarcode}.jpg`);
      const hasLocalImage = fs.existsSync(localFilePath);

      // Fill Title if column specified and title is empty
      if (titleColIndex && Number(titleColIndex) > 0) {
        const titleCell = row.getCell(Number(titleColIndex));
        if (meta?.title && (!titleCell.value || String(titleCell.value).trim() === '')) {
          titleCell.value = meta.title;
          filledTitlesCount++;
        }
      }

      // Fill Image if column specified
      if (imageColIndex && Number(imageColIndex) > 0) {
        const imageCell = row.getCell(Number(imageColIndex));

        if (fillType === 'embedded') {
          let imgBuf: Buffer | null = null;
          if (hasLocalImage) {
            try {
              imgBuf = fs.readFileSync(localFilePath);
            } catch {}
          } else if (cdnUrl) {
            try {
              const resp = await fetch(cdnUrl, { signal: AbortSignal.timeout(6000) });
              if (resp.ok) {
                imgBuf = Buffer.from(await resp.arrayBuffer());
                try { fs.writeFileSync(localFilePath, imgBuf); } catch {}
              }
            } catch {}
          }

          if (imgBuf) {
            try {
              const imageId = wb.addImage({
                buffer: imgBuf as any,
                extension: 'jpeg',
              });
              ws.addImage(imageId, {
                tl: { col: Number(imageColIndex) - 0.9, row: rowNumber - 0.9 },
                ext: { width: 55, height: 55 },
                editAs: 'oneCell',
              });
              row.height = Math.max(row.height || 18, 55);
              filledImagesCount++;
            } catch (e: any) {
              console.error(`Failed to embed image for ${cleanBarcode}:`, e.message);
            }
          }
        } else if (fillType === 'formula') {
          if (cdnUrl) {
            imageCell.value = { formula: `IMAGE("${cdnUrl}")` };
            filledImagesCount++;
          }
        } else if (fillType === 'cdnUrl') {
          if (cdnUrl) {
            imageCell.value = cdnUrl;
            filledImagesCount++;
          }
        }
      }
    }

    const outputBuffer = Buffer.from(await wb.xlsx.writeBuffer());
    const safeOutputName = filename.endsWith('.xlsx')
      ? filename.replace(/\.xlsx$/, '_filled.xlsx')
      : `${filename}_filled.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(safeOutputName)}"`);
    res.setHeader('Content-Length', String(outputBuffer.length));
    res.setHeader('Access-Control-Expose-Headers', 'Content-Length, Content-Disposition, X-Filled-Images, X-Filled-Titles');
    res.setHeader('X-Filled-Images', String(filledImagesCount));
    res.setHeader('X-Filled-Titles', String(filledTitlesCount));
    res.send(outputBuffer);
  } catch (err: any) {
    console.error('Error filling template:', err);
    res.status(500).json({ success: false, error: '填充模板失败: ' + err.message });
  }
});

// ====================================================
// MULTI-SITE PRICE EXTRACTION & TEMPLATE FILL ENGINE
// ====================================================

const barcodePriceCache = new Map<string, any>();

async function fetchBarcodePriceFromMeloni(barcode: string): Promise<{
  found: boolean;
  site: 'meloni';
  siteName: string;
  price: number | null; // effective price
  originalPrice: number | null; // 原价 / 划线价
  discountPrice: number | null; // 折后特价
  formattedPrice: string;
  formattedOriginalPrice: string;
  formattedDiscountPrice: string;
  discountPercent?: string;
  hasDiscount: boolean;
  title?: string;
  url?: string;
  inStock?: boolean;
  note?: string;
}> {
  const result: any = {
    found: false,
    site: 'meloni',
    siteName: 'Meloni Store',
    price: null,
    originalPrice: null,
    discountPrice: null,
    formattedPrice: '未收录',
    formattedOriginalPrice: '',
    formattedDiscountPrice: '',
    hasDiscount: false,
    url: '',
    inStock: true,
  };

  try {
    const meloniMap = await getOrBuildMeloniIndex();
    let prodUrl = meloniMap.get(barcode);

    if (!prodUrl) {
      try {
        const searchUrl = `https://www.melonistore.com/search.php?search_query=${encodeURIComponent(barcode)}`;
        const searchHtml = await fetchPageHtml(searchUrl);
        const $s = cheerio.load(searchHtml);
        const firstCardLink = $s('.card-title a').first().attr('href');
        if (firstCardLink) {
          prodUrl = firstCardLink;
          meloniMap.set(barcode, firstCardLink);
        }
      } catch {}
    }

    if (prodUrl) {
      result.found = true;
      result.url = prodUrl;
      const html = await fetchPageHtml(prodUrl);
      const $ = cheerio.load(html);

      const title =
        $('h1.productView-title').text().trim() ||
        $('h1').first().text().trim() ||
        $('title').text().trim();
      result.title = title;

      const priceText = $('.price--withoutTax, .price--discounted, .price--main, .price, .productView-price .price').first().text().trim();
      const nonSaleText = $('.price--non-sale, .price--rrp').first().text().trim();

      const parseNum = (str: string) => {
        if (!str || str.toLowerCase().includes('accedi')) return null;
        const clean = str.replace(/[^0-9,\.]/g, '').replace(',', '.');
        const num = parseFloat(clean);
        return isNaN(num) ? null : Number(num.toFixed(2));
      };

      const parsedPrice = parseNum(priceText);
      const parsedOld = parseNum(nonSaleText);

      if (parsedPrice !== null) {
        if (parsedOld !== null && parsedOld > parsedPrice) {
          result.originalPrice = parsedOld;
          result.discountPrice = parsedPrice;
          result.price = parsedPrice;
          result.hasDiscount = true;
          result.formattedOriginalPrice = `€${parsedOld.toFixed(2)}`;
          result.formattedDiscountPrice = `€${parsedPrice.toFixed(2)}`;
          result.formattedPrice = `€${parsedPrice.toFixed(2)}`;
          const disc = Math.round(((parsedOld - parsedPrice) / parsedOld) * 100);
          result.discountPercent = `-${disc}%`;
        } else {
          result.originalPrice = parsedPrice;
          result.discountPrice = null;
          result.price = parsedPrice;
          result.hasDiscount = false;
          result.formattedOriginalPrice = `€${parsedPrice.toFixed(2)}`;
          result.formattedDiscountPrice = '-';
          result.formattedPrice = `€${parsedPrice.toFixed(2)}`;
        }
      } else {
        result.price = null;
        result.originalPrice = null;
        result.discountPrice = null;
        if (meloniSession.authenticated) {
          result.formattedPrice = '已认证专属品';
          result.formattedOriginalPrice = '已认证专属品';
          result.formattedDiscountPrice = '-';
          result.note = '已认证 Meloni 会员会话';
        } else {
          result.formattedPrice = '需登录B2B底价';
          result.formattedOriginalPrice = '需登录';
          result.formattedDiscountPrice = '需登录';
          result.note = 'Meloni会员底价需登录 (可配置账号/Cookie)';
        }
      }

      const outOfStock = $('.card--out-of-stock').length > 0 || html.includes('Esaurito');
      result.inStock = !outOfStock;
    }
  } catch (err: any) {
    result.note = err.message;
  }

  return result;
}

async function fetchBarcodePriceFromMegaCedi(barcode: string): Promise<{
  found: boolean;
  site: 'megacedi';
  siteName: string;
  price: number | null;
  originalPrice: number | null;
  discountPrice: number | null;
  formattedPrice: string;
  formattedOriginalPrice: string;
  formattedDiscountPrice: string;
  discountPercent?: string;
  hasDiscount: boolean;
  title?: string;
  url?: string;
  inStock?: boolean;
  note?: string;
}> {
  const result: any = {
    found: false,
    site: 'megacedi',
    siteName: 'MegaCedi',
    price: null,
    originalPrice: null,
    discountPrice: null,
    formattedPrice: '未收录',
    formattedOriginalPrice: '',
    formattedDiscountPrice: '',
    hasDiscount: false,
    url: '',
    inStock: true,
  };

  try {
    const searchUrl = `https://www.megacedi.com/WebPartScaffale.aspx?p=${encodeURIComponent(barcode)}&r=%`;
    const html = await fetchPageHtml(searchUrl);
    const $ = cheerio.load(html);

    const img = $('img.jqTest2').first();
    let internalId = img.attr('id') || '';
    let altTitle = img.attr('alt')?.trim() || '';

    if (!internalId) {
      const link = $('a[href*="codice_a_barre="]').first();
      const codeMatch = link.attr('href')?.match(/codice_a_barre=([0-9A-Za-z]+)/i);
      if (codeMatch) {
        internalId = codeMatch[1];
      }
    }

    if (internalId) {
      result.found = true;
      const detailUrl = `https://www.megacedi.com/SchedaArticoloP.aspx?codice_a_barre=${internalId}`;
      result.url = detailUrl;
      result.title = altTitle || `MegaCedi 商品 (${barcode})`;

      const schedaUrl = `https://www.megacedi.com/SchedaHome.aspx?codice_a_barre=${internalId}`;
      const schedaHtml = await fetchPageHtml(schedaUrl);
      const $s = cheerio.load(schedaHtml);

      const titleFromScheda = $s('strong').first().text().trim() || $s('body').text().split('barcode:')[0].trim();
      if (titleFromScheda && titleFromScheda.length > 2) {
        result.title = titleFromScheda;
      }

      // Check if logged in cookie reveals wholesale price
      const priceMatch =
        schedaHtml.match(/prezzo:\s*<strong>\s*([0-9,\.]+)\s*€?/i) ||
        schedaHtml.match(/€\s*([0-9,\.]+)/i) ||
        html.match(/€\s*([0-9,\.]+)/i);

      if (priceMatch) {
        const clean = priceMatch[1].replace(',', '.');
        const num = parseFloat(clean);
        if (!isNaN(num) && num > 0) {
          const p = Number(num.toFixed(2));
          result.price = p;
          result.originalPrice = p;
          result.discountPrice = null;
          result.formattedPrice = `€${p.toFixed(2)}`;
          result.formattedOriginalPrice = `€${p.toFixed(2)}`;
          result.formattedDiscountPrice = '-';
          result.note = '批发特惠价';
        }
      }

      if (result.price === null) {
        if (megaCediSession.authenticated) {
          result.formattedPrice = '专属底价';
          result.formattedOriginalPrice = '专属底价';
          result.note = '已认证会话';
        } else {
          result.formattedPrice = '需登录查看批价';
          result.formattedOriginalPrice = '需登录';
          result.formattedDiscountPrice = '需登录';
          result.note = '批发门槛 (可填入Cookie)';
        }
      }
    }
  } catch (err: any) {
    result.note = err.message;
  }

  return result;
}

async function fetchBarcodePriceFromMaurys(barcode: string, knownTitle = ''): Promise<{
  found: boolean;
  site: 'maurys';
  siteName: string;
  price: number | null;
  originalPrice: number | null;
  discountPrice: number | null;
  formattedPrice: string;
  formattedOriginalPrice: string;
  formattedDiscountPrice: string;
  discountPercent?: string;
  hasDiscount: boolean;
  title?: string;
  url?: string;
  inStock?: boolean;
  note?: string;
}> {
  const result: any = {
    found: false,
    site: 'maurys',
    siteName: "Maury's Online",
    price: null,
    originalPrice: null,
    discountPrice: null,
    formattedPrice: '未收录',
    formattedOriginalPrice: '',
    formattedDiscountPrice: '',
    hasDiscount: false,
    url: '',
    inStock: true,
  };

  try {
    const searchUrl = `https://maurysonline.it/search?q=${encodeURIComponent(barcode)}`;
    let html = await fetchPageHtml(searchUrl);
    let $ = cheerio.load(html);
    let items = $('.product-item');

    // If barcode not directly returned in title, try searching by brand + key terms if available
    if (items.length === 0 && knownTitle) {
      const words = knownTitle
        .replace(/[^a-zA-Z0-9\s]/g, ' ')
        .split(/\s+/)
        .filter((w) => w.length >= 3 && !['CARING', 'SHAMPOO', 'ML', 'PZ', 'GR', 'PER', 'DEL', 'CON'].includes(w.toUpperCase()));
      if (words.length > 0) {
        const brandTerm = words[0];
        const subSearch = `https://maurysonline.it/search?q=${encodeURIComponent(brandTerm)}`;
        const subHtml = await fetchPageHtml(subSearch);
        const $sub = cheerio.load(subHtml);
        const subItems = $sub('.product-item');

        subItems.each((_, el) => {
          const itemTitle = $sub(el).find('.product-title a').text().toLowerCase();
          if (itemTitle.includes(brandTerm.toLowerCase())) {
            const hasSecondWord = words.length <= 1 || words.slice(1, 3).some((w) => itemTitle.includes(w.toLowerCase()));
            if (hasSecondWord && items.length === 0) {
              items = $sub(el);
              $ = $sub;
            }
          }
        });
      }
    }

    if (items.length > 0) {
      const item = items.first();
      result.found = true;
      const title = item.find('.product-title a').text().trim();
      let prodUrl = item.find('.product-title a').attr('href') || '';
      if (prodUrl.startsWith('/')) prodUrl = `https://maurysonline.it${prodUrl}`;
      result.title = title;
      result.url = prodUrl;

      const actualPriceText = item.find('.actual-price-red, .actual-price, .price').first().text().trim();
      const oldPriceText = item.find('.old-price').text().trim();

      const parseNum = (str: string) => {
        if (!str) return null;
        const clean = str.replace(/[^0-9,\.]/g, '').replace(',', '.');
        const num = parseFloat(clean);
        return isNaN(num) ? null : Number(num.toFixed(2));
      };

      const actualNum = parseNum(actualPriceText);
      const oldNum = parseNum(oldPriceText);

      if (actualNum !== null) {
        if (oldNum !== null && oldNum > actualNum) {
          result.originalPrice = oldNum;
          result.discountPrice = actualNum;
          result.price = actualNum;
          result.hasDiscount = true;
          result.formattedOriginalPrice = `€${oldNum.toFixed(2)}`;
          result.formattedDiscountPrice = `€${actualNum.toFixed(2)}`;
          result.formattedPrice = `€${actualNum.toFixed(2)}`;
          const disc = Math.round(((oldNum - actualNum) / oldNum) * 100);
          result.discountPercent = `-${disc}%`;
        } else {
          result.originalPrice = actualNum;
          result.discountPrice = null;
          result.price = actualNum;
          result.hasDiscount = false;
          result.formattedOriginalPrice = `€${actualNum.toFixed(2)}`;
          result.formattedDiscountPrice = '-';
          result.formattedPrice = `€${actualNum.toFixed(2)}`;
        }
      }

      result.inStock = item.find('.stock.unavailable, .out-of-stock').length === 0;
    }
  } catch (err: any) {
    result.note = err.message;
  }

  return result;
}

async function fetchBarcodePriceFromRisparmio(barcode: string, knownTitle = ''): Promise<{
  found: boolean;
  site: 'risparmiocasa';
  siteName: string;
  price: number | null;
  originalPrice: number | null;
  discountPrice: number | null;
  formattedPrice: string;
  formattedOriginalPrice: string;
  formattedDiscountPrice: string;
  discountPercent?: string;
  hasDiscount: boolean;
  title?: string;
  url?: string;
  inStock?: boolean;
  note?: string;
}> {
  const result: any = {
    found: false,
    site: 'risparmiocasa',
    siteName: 'Risparmio Casa',
    price: null,
    originalPrice: null,
    discountPrice: null,
    formattedPrice: '未收录',
    formattedOriginalPrice: '',
    formattedDiscountPrice: '',
    hasDiscount: false,
    url: '',
    inStock: true,
  };

  try {
    const searchUrl = `https://shop.risparmiocasa.com/catalogsearch/result/?q=${encodeURIComponent(barcode)}`;
    let html = await fetchPageHtml(searchUrl);
    let $ = cheerio.load(html);
    let items = $('li.item.product.product-item');

    // If barcode search yielded 0, try searching by brand + key terms if available
    if (items.length === 0 && knownTitle) {
      const words = knownTitle
        .replace(/[^a-zA-Z0-9\s]/g, ' ')
        .split(/\s+/)
        .filter((w) => w.length >= 3 && !['CARING', 'SHAMPOO', 'ML', 'PZ', 'GR', 'PER', 'DEL', 'CON', 'CASA', 'DETERSIVO'].includes(w.toUpperCase()));
      if (words.length > 0) {
        const brandTerm = words[0];
        const subSearch = `https://shop.risparmiocasa.com/catalogsearch/result/?q=${encodeURIComponent(brandTerm + ' ' + (words[1] || ''))}`;
        const subHtml = await fetchPageHtml(subSearch);
        const $sub = cheerio.load(subHtml);
        const candidateItems = $sub('li.item.product.product-item');

        candidateItems.each((_, el) => {
          const itemTitle = $sub(el).find('a.product-item-link').text().toLowerCase();
          if (itemTitle.includes(brandTerm.toLowerCase())) {
            const hasMatch = words.length <= 1 || words.slice(1, 3).some((w) => itemTitle.includes(w.toLowerCase()));
            if (hasMatch && items.length === 0) {
              items = $sub(el);
              $ = $sub;
            }
          }
        });
      }
    }

    if (items.length > 0) {
      const item = items.first();
      result.found = true;
      const title = item.find('a.product-item-link').text().trim();
      let prodUrl = item.find('a.product-item-link').attr('href') || '';
      if (prodUrl.startsWith('/')) prodUrl = `https://shop.risparmiocasa.com${prodUrl}`;
      result.title = title;
      result.url = prodUrl;

      const finalPriceWrapper = item.find('[data-price-type="finalPrice"]');
      const oldPriceWrapper = item.find('[data-price-type="oldPrice"]');

      const finalAttr = finalPriceWrapper.attr('data-price-amount');
      const oldAttr = oldPriceWrapper.attr('data-price-amount');

      const parseNum = (str: string | undefined) => {
        if (!str) return null;
        const clean = str.replace(/[^0-9,\.]/g, '').replace(',', '.');
        const num = parseFloat(clean);
        return isNaN(num) ? null : Number(num.toFixed(2));
      };

      let finalNum = parseNum(finalAttr);
      let oldNum = parseNum(oldAttr);

      if (finalNum === null) {
        finalNum = parseNum(finalPriceWrapper.find('.price').text());
      }
      if (oldNum === null) {
        oldNum = parseNum(oldPriceWrapper.find('.price').text());
      }

      if (finalNum !== null) {
        if (oldNum !== null && oldNum > finalNum) {
          result.originalPrice = oldNum;
          result.discountPrice = finalNum;
          result.price = finalNum;
          result.hasDiscount = true;
          result.formattedOriginalPrice = `€${oldNum.toFixed(2)}`;
          result.formattedDiscountPrice = `€${finalNum.toFixed(2)}`;
          result.formattedPrice = `€${finalNum.toFixed(2)}`;
          const disc = Math.round(((oldNum - finalNum) / oldNum) * 100);
          result.discountPercent = `-${disc}%`;
        } else {
          result.originalPrice = finalNum;
          result.discountPrice = null;
          result.price = finalNum;
          result.hasDiscount = false;
          result.formattedOriginalPrice = `€${finalNum.toFixed(2)}`;
          result.formattedDiscountPrice = '-';
          result.formattedPrice = `€${finalNum.toFixed(2)}`;
        }
      }

      result.inStock = item.find('.stock.unavailable').length === 0;
    }
  } catch (err: any) {
    result.note = err.message;
  }

  return result;
}

// ----------------------------------------------------
// PIÙME (piume.it) PRICE EXTRACTOR (BigCommerce GraphQL)
// ----------------------------------------------------
async function getOrRefreshPiumeToken(): Promise<string> {
  const now = Date.now();
  if (piumeGraphQLToken && now - piumeTokenLastRefreshed < 12 * 3600 * 1000) {
    return piumeGraphQLToken;
  }
  try {
    const html = await fetchPageHtml('https://piume.it/');
    const tokenMatch = html.match(/eyJ0eXAiOiJKV1QiLCJhbGciOiJFUzI1NiJ9\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/);
    if (tokenMatch && tokenMatch[0]) {
      piumeGraphQLToken = tokenMatch[0];
      piumeTokenLastRefreshed = now;
    }
  } catch (err: any) {
    console.warn('PiùMe token auto-refresh warning:', err.message);
  }
  return piumeGraphQLToken;
}

async function fetchBarcodePriceFromPiume(
  barcode: string,
  knownTitle = ''
): Promise<{
  found: boolean;
  site: 'piume';
  siteName: string;
  price: number | null;
  originalPrice: number | null;
  discountPrice: number | null;
  formattedPrice: string;
  formattedOriginalPrice: string;
  formattedDiscountPrice: string;
  discountPercent?: string;
  hasDiscount: boolean;
  title?: string;
  url?: string;
  imageUrl?: string;
  sku?: string;
  brand?: string;
  inStock?: boolean;
  note?: string;
}> {
  const result: any = {
    found: false,
    site: 'piume',
    siteName: 'PiùMe',
    price: null,
    originalPrice: null,
    discountPrice: null,
    formattedPrice: '未收录',
    formattedOriginalPrice: '',
    formattedDiscountPrice: '',
    hasDiscount: false,
    url: '',
    inStock: true,
  };

  try {
    const token = await getOrRefreshPiumeToken();
    const query = `
      query searchProducts($query: String!) {
        site {
          search {
            searchProducts(filters: { searchTerm: $query }) {
              products {
                edges {
                  node {
                    id
                    entityId
                    name
                    sku
                    path
                    prices {
                      price { value currencyCode }
                      basePrice { value currencyCode }
                      retailPrice { value currencyCode }
                      salePrice { value currencyCode }
                    }
                    brand { name }
                    defaultImage { url(width: 300) }
                  }
                }
              }
            }
          }
        }
      }
    `;

    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${token}`,
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    };
    if (piumeSession.cookie) {
      headers['Cookie'] = piumeSession.cookie;
    }

    // 1. Direct barcode query
    let res = await fetch('https://piume.it/graphql', {
      method: 'POST',
      headers,
      body: JSON.stringify({ query, variables: { query: barcode } }),
    });

    let data: any = null;
    try {
      data = await res.json();
    } catch {}

    let edges = data?.data?.site?.search?.searchProducts?.products?.edges || [];

    // 2. Fallback: if barcode query returned empty and knownTitle exists, query by top keywords
    if (edges.length === 0 && knownTitle && knownTitle.length >= 3) {
      const keywords = knownTitle
        .replace(/[^a-zA-Z0-9\s]/g, ' ')
        .split(/\s+/)
        .filter((w: string) => w.length >= 3 && !['THE', 'PER', 'CON', 'AND', 'DEL', 'DELLA'].includes(w.toUpperCase()))
        .slice(0, 3)
        .join(' ');

      if (keywords) {
        try {
          const fbRes = await fetch('https://piume.it/graphql', {
            method: 'POST',
            headers,
            body: JSON.stringify({ query, variables: { query: keywords } }),
          });
          const fbData = await fbRes.json();
          const candidateEdges = fbData?.data?.site?.search?.searchProducts?.products?.edges || [];
          const matchedEdge = candidateEdges.find((e: any) => {
            const path = e.node?.path || '';
            const sku = e.node?.sku || '';
            return path.includes(barcode) || sku.includes(barcode);
          });
          if (matchedEdge) {
            edges = [matchedEdge];
          }
        } catch {}
      }
    }

    if (edges.length > 0) {
      const node = edges[0].node;
      result.found = true;
      result.title = node.name;
      result.url = node.path ? (node.path.startsWith('http') ? node.path : `https://piume.it${node.path}`) : '';
      result.imageUrl = node.defaultImage?.url || '';
      result.sku = node.sku || '';
      result.brand = node.brand?.name || '';

      const p = node.prices?.price?.value;
      const baseP = node.prices?.basePrice?.value;
      const retailP = node.prices?.retailPrice?.value;
      const saleP = node.prices?.salePrice?.value;

      const currentVal = typeof p === 'number' ? p : (typeof saleP === 'number' ? saleP : baseP);
      const originalVal = typeof retailP === 'number' && retailP > (currentVal || 0)
        ? retailP
        : (typeof baseP === 'number' && typeof currentVal === 'number' && baseP > currentVal ? baseP : null);

      if (typeof currentVal === 'number' && !isNaN(currentVal)) {
        result.price = Number(currentVal.toFixed(2));
        result.formattedPrice = `€${result.price.toFixed(2)}`;

        if (typeof originalVal === 'number' && originalVal > result.price) {
          result.hasDiscount = true;
          result.originalPrice = Number(originalVal.toFixed(2));
          result.discountPrice = result.price;
          result.formattedOriginalPrice = `€${result.originalPrice.toFixed(2)}`;
          result.formattedDiscountPrice = `€${result.price.toFixed(2)}`;
          const disc = Math.round(((result.originalPrice - result.price) / result.originalPrice) * 100);
          result.discountPercent = `-${disc}%`;
        } else {
          result.hasDiscount = false;
          result.originalPrice = result.price;
          result.discountPrice = null;
          result.formattedOriginalPrice = `€${result.price.toFixed(2)}`;
          result.formattedDiscountPrice = '-';
        }
      }
    }
  } catch (err: any) {
    result.note = err.message;
  }

  return result;
}

// Master multi-site price aggregator for a single barcode
async function extractPricesForSingleBarcode(
  barcode: string,
  selectedSites: string[] = ['meloni', 'megacedi', 'maurys', 'risparmiocasa', 'piume']
): Promise<any> {
  const cleanBarcode = barcode.replace(/[^0-9A-Za-z]/g, '').trim();

  // 1. Try finding preliminary title from cache or fast Meloni index
  let knownTitle = barcodeMetadataMap.get(cleanBarcode)?.title || '';
  if (!knownTitle) {
    const meloniMap = await getOrBuildMeloniIndex();
    const meloniUrl = meloniMap.get(cleanBarcode);
    if (meloniUrl) {
      const slugMatch = meloniUrl.match(/melonistore\.com\/(.+)-[0-9]{8,14}\.html/);
      if (slugMatch) {
        knownTitle = slugMatch[1].replace(/-/g, ' ').toUpperCase();
      }
    }
  }

  // 2. Fetch selected sites concurrently
  const fetchPromises: Promise<any>[] = [];
  const siteKeys: string[] = [];

  if (selectedSites.includes('meloni')) {
    siteKeys.push('meloni');
    fetchPromises.push(fetchBarcodePriceFromMeloni(cleanBarcode));
  }
  if (selectedSites.includes('megacedi')) {
    siteKeys.push('megacedi');
    fetchPromises.push(fetchBarcodePriceFromMegaCedi(cleanBarcode));
  }
  if (selectedSites.includes('maurys')) {
    siteKeys.push('maurys');
    fetchPromises.push(fetchBarcodePriceFromMaurys(cleanBarcode, knownTitle));
  }
  if (selectedSites.includes('risparmiocasa')) {
    siteKeys.push('risparmiocasa');
    fetchPromises.push(fetchBarcodePriceFromRisparmio(cleanBarcode, knownTitle));
  }
  if (selectedSites.includes('piume')) {
    siteKeys.push('piume');
    fetchPromises.push(fetchBarcodePriceFromPiume(cleanBarcode, knownTitle));
  }

  const responses = await Promise.all(fetchPromises);
  const prices: Record<string, any> = {};

  for (let i = 0; i < siteKeys.length; i++) {
    const key = siteKeys[i];
    prices[key] = responses[i];
    if (!knownTitle && responses[i]?.title) {
      knownTitle = responses[i].title;
    }
  }

  // 3. Compute Lowest Original Price, Lowest Discount Price, and Effective Lowest Price
  const originalPriceEntries: { site: string; siteName: string; price: number }[] = [];
  const discountPriceEntries: { site: string; siteName: string; price: number }[] = [];
  const effectivePriceEntries: { site: string; siteName: string; price: number }[] = [];

  for (const [k, p] of Object.entries(prices)) {
    if (p.found) {
      if (p.originalPrice !== null && p.originalPrice > 0) {
        originalPriceEntries.push({ site: k, siteName: p.siteName, price: p.originalPrice });
      }
      if (p.discountPrice !== null && p.discountPrice > 0) {
        discountPriceEntries.push({ site: k, siteName: p.siteName, price: p.discountPrice });
      }
      if (p.price !== null && p.price > 0) {
        effectivePriceEntries.push({ site: k, siteName: p.siteName, price: p.price });
      }
    }
  }

  let lowestOriginalPrice: number | null = null;
  let lowestOriginalPlatform: string | undefined;
  if (originalPriceEntries.length > 0) {
    originalPriceEntries.sort((a, b) => a.price - b.price);
    lowestOriginalPrice = originalPriceEntries[0].price;
    lowestOriginalPlatform = originalPriceEntries[0].siteName;
  }

  let lowestDiscountPrice: number | null = null;
  if (discountPriceEntries.length > 0) {
    discountPriceEntries.sort((a, b) => a.price - b.price);
    lowestDiscountPrice = discountPriceEntries[0].price;
  }

  let lowestPrice: number | null = null;
  let lowestPlatform: string | undefined;
  let highestPrice: number | null = null;
  let highestPlatform: string | undefined;
  let priceDiff: number | null = null;
  let averagePrice: number | null = null;

  if (effectivePriceEntries.length > 0) {
    effectivePriceEntries.sort((a, b) => a.price - b.price);
    lowestPrice = effectivePriceEntries[0].price;
    lowestPlatform = effectivePriceEntries[0].siteName;
    highestPrice = effectivePriceEntries[effectivePriceEntries.length - 1].price;
    highestPlatform = effectivePriceEntries[effectivePriceEntries.length - 1].siteName;
    priceDiff = Number((highestPrice - lowestPrice).toFixed(2));
    const sum = effectivePriceEntries.reduce((s, e) => s + e.price, 0);
    averagePrice = Number((sum / effectivePriceEntries.length).toFixed(2));
  }

  const hasDiscount = discountPriceEntries.length > 0;
  const matched = originalPriceEntries.length > 0 || Object.values(prices).some((p: any) => p.found);

  const priceItem = {
    barcode: cleanBarcode,
    title: knownTitle || (matched ? `商品条码 ${cleanBarcode}` : '未在选定平台找到'),
    matched,
    lowestPrice, // 综合实际有效最低价
    lowestOriginalPrice, // 综合最低原价
    lowestDiscountPrice, // 综合最低折后价
    lowestPlatform,
    lowestOriginalPlatform,
    highestPrice,
    highestPlatform,
    priceDiff,
    averagePrice,
    hasDiscount,
    prices,
    extractedAt: new Date().toISOString(),
  };

  barcodePriceCache.set(cleanBarcode, priceItem);
  return priceItem;
}

// ----------------------------------------------------
// API: B2B Price Session & Login Status
// ----------------------------------------------------
app.get('/api/prices/sessions', (_req: Request, res: Response) => {
  res.json({
    maurys: { requiresLogin: false, label: '100% 免登录直接看价格' },
    risparmiocasa: { requiresLogin: false, label: '100% 免登录直接看价格' },
    piume: {
      requiresLogin: false,
      hasCookie: !!piumeSession.cookie,
      label: piumeSession.cookie ? '已配置 PiùMe 会员' : '100% 免登录直接看价格',
    },
    meloni: {
      requiresLogin: false,
      loginForWholesaleOnly: true,
      hasCookie: !!meloniSession.cookie,
      label: meloniSession.cookie ? '已配置商户 Cookie (可看底价)' : '部分商品公开 / 会员特惠需Cookie',
    },
    megacedi: {
      requiresLogin: false,
      loginForWholesaleOnly: true,
      hasCookie: !!megaCediSession.cookie,
      label: megaCediSession.cookie ? '已配置批发 Cookie' : 'B2B 批价门槛 (可填入Cookie)',
    },
  });
});

app.post('/api/prices/sessions', (req: Request, res: Response) => {
  const { meloniCookie, megaCookie, piumeCookie } = req.body;
  if (typeof meloniCookie === 'string') {
    meloniSession.cookie = meloniCookie.trim();
    meloniSession.authenticated = !!meloniCookie.trim();
  }
  if (typeof megaCookie === 'string') {
    megaCediSession.cookie = megaCookie.trim();
    megaCediSession.authenticated = !!megaCookie.trim();
  }
  if (typeof piumeCookie === 'string') {
    piumeSession.cookie = piumeCookie.trim();
    piumeSession.authenticated = !!piumeCookie.trim();
  }
  res.json({
    success: true,
    meloniConfigured: !!meloniSession.cookie,
    megaConfigured: !!megaCediSession.cookie,
    piumeConfigured: !!piumeSession.cookie,
  });
});

// ----------------------------------------------------
// API: SSE Stream Multi-Site Price Extraction
// ----------------------------------------------------
app.post('/api/prices/stream', async (req: Request, res: Response) => {
  const {
    barcodes = [],
    sites = ['meloni', 'megacedi', 'maurys', 'risparmiocasa', 'piume'],
    delayMs = 100,
  } = req.body;

  const rawList: string[] = Array.isArray(barcodes) ? barcodes : [];
  const uniqueBarcodes = Array.from(
    new Set(rawList.map((b) => String(b).replace(/[^0-9A-Za-z]/g, '').trim()).filter((b) => b.length >= 6))
  );

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders?.();

  const sendEvent = (event: string, data: any) => {
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  };

  let isAborted = false;
  res.on('close', () => {
    isAborted = true;
  });

  const total = uniqueBarcodes.length;
  sendEvent('start', {
    total,
    sites,
    message: `准备开始从 ${sites.length} 个网站并发提取 ${total} 个条码的价格信息...`,
  });

  // Pre-load Meloni index if meloni is one of the selected sites
  if (sites.includes('meloni')) {
    sendEvent('status', { message: '正在同步 Meloni Store 全网商品索引 (约1.3万件)...' });
    await getOrBuildMeloniIndex();
  }

  // Pre-fetch PiùMe token if piume is one of the selected sites
  if (sites.includes('piume')) {
    sendEvent('status', { message: '正在同步 PiùMe (piume.it) 接口凭据...' });
    await getOrRefreshPiumeToken();
  }

  let processed = 0;
  let matchedCount = 0;
  let unmatchedCount = 0;
  let meloniFound = 0;
  let megaFound = 0;
  let maurysFound = 0;
  let risparmioFound = 0;
  let piumeFound = 0;

  // Process in concurrency of 2 barcodes at a time to prevent rate limiting
  const concurrency = 2;
  for (let i = 0; i < uniqueBarcodes.length; i += concurrency) {
    if (isAborted) break;

    const chunk = uniqueBarcodes.slice(i, i + concurrency);
    const results = await Promise.all(
      chunk.map(async (code) => {
        return await extractPricesForSingleBarcode(code, sites);
      })
    );

    for (const item of results) {
      processed++;
      if (item.matched) {
        matchedCount++;
        if (item.prices.meloni?.found) meloniFound++;
        if (item.prices.megacedi?.found) megaFound++;
        if (item.prices.maurys?.found) maurysFound++;
        if (item.prices.risparmiocasa?.found) risparmioFound++;
        if (item.prices.piume?.found) piumeFound++;
      } else {
        unmatchedCount++;
      }

      sendEvent('item', item);
    }

    sendEvent('progress', {
      processed,
      total,
      matchedCount,
      unmatchedCount,
      meloniFound,
      megaFound,
      maurysFound,
      risparmioFound,
      piumeFound,
      percent: Math.round((processed / total) * 100),
      currentBarcode: chunk[chunk.length - 1],
    });

    if (delayMs > 0) {
      await new Promise((r) => setTimeout(r, delayMs));
    }
  }

  sendEvent('complete', {
    total,
    processed,
    matchedCount,
    unmatchedCount,
    meloniFound,
    megaFound,
    maurysFound,
    risparmioFound,
    piumeFound,
    message: `🎉 全网价格提取完毕！共成功提取 ${matchedCount} 件商品的比价数据。`,
  });

  res.end();
});

// ----------------------------------------------------
// API: Custom Excel Template Price Insertion & Export
// ----------------------------------------------------
app.post('/api/prices/fill-template', async (req: Request, res: Response) => {
  try {
    const {
      templateBase64,
      filename = 'custom_template.xlsx',
      sheetName,
      sheetIndex = 1,

      // Multi-sheet control (User requested: "有时候表格里面会有几张表，不同表是不同品牌或者不同分类的产品，这些表里的条码都是要提取的。导出的时候也都是要插入价格的")
      applyToAllSheets = true,
      selectedSheets = [], // array of sheet names
      autoDetectBarcodeCol = true,

      // Column target mode (User requested: "可以新建列，不用完全和之前一样的格式。这样就不会覆盖原来的内容")
      columnTargetMode = 'new_columns', // 'new_columns' | 'existing_columns'
      newOriginalColHeader = '提取-最低原价 (€)',
      newDiscountColHeader = '提取-折后特价 (€)',
      newPlatformColHeader = '最低价渠道',
      insertPlatformCol = false,

      barcodeColIndex = 1, // 1-based column index
      headerRowIndex = 1,
      startRow = 2,
      fillMode = 'original_and_discount', // 'single' | 'original_and_discount' | 'multi'

      // Mode 1: Single column fill
      singlePriceType = 'lowest_original',
      singlePriceColIndex = 2,

      // Mode 2: Original + Optional Discount column
      originalPriceType = 'lowest_original',
      originalPriceColIndex = 2,
      insertDiscountCol = false,
      discountPriceType = 'lowest_discount',
      discountPriceColIndex = 0,
      discountNonSaleHandling = 'empty', // 'empty' | 'original' | 'dash'

      // Mode 3: Multi column mapped fill
      columnMappings = {},
      valueFormat = 'number', // 'number' | 'currency' | 'text'
      unmatchedFill = 'empty', // 'empty' | 'zero' | 'dash'
      items = [], // Pre-extracted BarcodePriceItem[] from client
    } = req.body;

    if (!templateBase64) {
      return res.status(400).json({ success: false, error: '请上传有效的 Excel 模板文件 (.xlsx)' });
    }

    // Index items by barcode for O(1) lookup
    const priceMap = new Map<string, any>();
    if (Array.isArray(items)) {
      items.forEach((it: any) => {
        if (it?.barcode) priceMap.set(String(it.barcode).trim(), it);
      });
    }
    // Also merge with backend cache
    for (const [k, v] of barcodePriceCache.entries()) {
      if (!priceMap.has(k)) priceMap.set(k, v);
    }

    const templateBuffer = Buffer.from(templateBase64, 'base64');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(templateBuffer as any);

    // Determine target worksheets
    let targetWorksheets: ExcelJS.Worksheet[] = [];
    if (applyToAllSheets) {
      targetWorksheets = wb.worksheets;
    } else if (Array.isArray(selectedSheets) && selectedSheets.length > 0) {
      targetWorksheets = wb.worksheets.filter((w) => selectedSheets.includes(w.name));
    } else {
      const singleWs = sheetName ? wb.getWorksheet(sheetName) : wb.getWorksheet(sheetIndex);
      if (singleWs) targetWorksheets = [singleWs];
    }

    if (targetWorksheets.length === 0) {
      return res.status(400).json({ success: false, error: `未找到指定工作表: ${sheetName || sheetIndex}` });
    }

    let totalFilledCount = 0;
    let processedSheetsCount = 0;

    const formatCellValue = (val: number | null | undefined, cell: any, fallbackBehavior = unmatchedFill) => {
      if (val === null || val === undefined || isNaN(val)) {
        if (fallbackBehavior === 'zero') {
          cell.value = 0;
          cell.numFmt = '#,##0.00';
        } else if (fallbackBehavior === 'dash') {
          cell.value = '-';
        }
        return;
      }

      const numVal = Number(val);
      if (valueFormat === 'number') {
        cell.value = numVal;
        cell.numFmt = '#,##0.00';
      } else if (valueFormat === 'currency') {
        cell.value = numVal;
        cell.numFmt = '€#,##0.00';
      } else {
        cell.value = `€${numVal.toFixed(2)}`;
      }
    };

    const getPriceValue = (item: any, priceType: string): number | null => {
      if (!item) return null;
      switch (priceType) {
        case 'lowest_original':
          return item.lowestOriginalPrice ?? item.lowestPrice ?? null;
        case 'lowest_discount':
          return item.lowestDiscountPrice ?? null;
        case 'lowest_effective':
          return item.lowestPrice ?? null;
        case 'average':
          return item.averagePrice ?? null;
        case 'meloni_original':
        case 'meloni':
          return item.prices?.meloni?.originalPrice ?? item.prices?.meloni?.price ?? null;
        case 'meloni_discount':
          return item.prices?.meloni?.discountPrice ?? null;
        case 'megacedi_original':
        case 'megacedi':
          return item.prices?.megacedi?.originalPrice ?? item.prices?.megacedi?.price ?? null;
        case 'megacedi_discount':
          return item.prices?.megacedi?.discountPrice ?? null;
        case 'maurys_original':
        case 'maurys':
          return item.prices?.maurys?.originalPrice ?? item.prices?.maurys?.price ?? null;
        case 'maurys_discount':
          return item.prices?.maurys?.discountPrice ?? null;
        case 'risparmiocasa_original':
        case 'risparmiocasa':
          return item.prices?.risparmiocasa?.originalPrice ?? item.prices?.risparmiocasa?.price ?? null;
        case 'risparmiocasa_discount':
          return item.prices?.risparmiocasa?.discountPrice ?? null;
        case 'piume_original':
        case 'piume':
          return item.prices?.piume?.originalPrice ?? item.prices?.piume?.price ?? null;
        case 'piume_discount':
          return item.prices?.piume?.discountPrice ?? null;
        default:
          return item.lowestOriginalPrice ?? item.lowestPrice ?? null;
      }
    };

    // Iterate over each worksheet (multi-brand / multi-category sheet support!)
    for (const ws of targetWorksheets) {
      if (!ws || ws.rowCount <= 0) continue;
      processedSheetsCount++;

      // 1. Determine header row and barcode column for this sheet
      let curHeaderRow = headerRowIndex || 1;
      let curStartRow = startRow || curHeaderRow + 1;
      let curBarcodeCol = Number(barcodeColIndex) || 1;

      if (autoDetectBarcodeCol) {
        let bestCol = -1;
        let bestScore = -1;
        const scanMaxRow = Math.min(ws.rowCount, 8);

        for (let r = 1; r <= scanMaxRow; r++) {
          const row = ws.getRow(r);
          const cellCount = Math.max(row.cellCount, 25);
          for (let c = 1; c <= cellCount; c++) {
            const val = String(row.getCell(c).value || '').toLowerCase().trim();
            if (!val) continue;
            let score = 0;
            if (val.includes('barcode') || val.includes('ean') || val.includes('条形码') || val.includes('条码') || val.includes('c.barre') || val.includes('codice barre')) {
              score += 150;
            }
            if (score > bestScore) {
              bestScore = score;
              bestCol = c;
              curHeaderRow = r;
              curStartRow = r + 1;
            }
          }
        }

        if (bestCol > 0) {
          curBarcodeCol = bestCol;
        } else {
          // Heuristic: check column values for 8-14 numeric digits
          for (let c = 1; c <= 20; c++) {
            let digitCount = 0;
            for (let r = curStartRow; r <= Math.min(ws.rowCount, curStartRow + 15); r++) {
              const val = String(ws.getRow(r).getCell(c).value || '').replace(/[^0-9]/g, '');
              if (val.length >= 8 && val.length <= 14) digitCount++;
            }
            if (digitCount >= 2) {
              curBarcodeCol = c;
              break;
            }
          }
        }
      }

      // 2. Determine target columns: Append NEW columns or use existing mapped columns
      let targetOrigCol = Number(originalPriceColIndex);
      let targetDiscCol = Number(discountPriceColIndex);
      let targetPlatCol = 0;

      if (columnTargetMode === 'new_columns') {
        // Find maximum existing column in this sheet
        let maxCol = 1;
        const headerRowObj = ws.getRow(curHeaderRow);
        headerRowObj.eachCell({ includeEmpty: false }, (_cell, colNumber) => {
          if (colNumber > maxCol) maxCol = colNumber;
        });
        if (ws.actualColumnCount > maxCol) maxCol = ws.actualColumnCount;

        // Sample cell for inheriting background & font style
        const sampleCell = headerRowObj.getCell(maxCol);

        // Append New Column 1: Original Price
        targetOrigCol = maxCol + 1;
        const origHead = headerRowObj.getCell(targetOrigCol);
        origHead.value = newOriginalColHeader || '提取-最低原价 (€)';
        origHead.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        origHead.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        origHead.alignment = { vertical: 'middle', horizontal: 'center' };
        ws.getColumn(targetOrigCol).width = 18;

        // Append New Column 2: Optional Discount Price (User checked: "需要的时候才选择插入一列打折价")
        if (insertDiscountCol) {
          targetDiscCol = maxCol + 2;
          const discHead = headerRowObj.getCell(targetDiscCol);
          discHead.value = newDiscountColHeader || '提取-折后特价 (€)';
          discHead.font = { bold: true, color: { argb: 'FFFFFFFF' } };
          discHead.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9F1239' } };
          discHead.alignment = { vertical: 'middle', horizontal: 'center' };
          ws.getColumn(targetDiscCol).width = 18;
        }

        // Append New Column 3: Optional Platform Channel
        if (insertPlatformCol) {
          targetPlatCol = maxCol + (insertDiscountCol ? 3 : 2);
          const platHead = headerRowObj.getCell(targetPlatCol);
          platHead.value = newPlatformColHeader || '最低价渠道';
          platHead.font = { bold: true, color: { argb: 'FFFFFFFF' } };
          platHead.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
          platHead.alignment = { vertical: 'middle', horizontal: 'center' };
          ws.getColumn(targetPlatCol).width = 16;
        }
      }

      // 3. Scan rows in this worksheet and insert prices
      const totalRows = ws.rowCount;
      for (let r = curStartRow; r <= totalRows; r++) {
        const row = ws.getRow(r);
        const rawBarcode = row.getCell(curBarcodeCol).value;
        if (!rawBarcode) continue;

        const cleanBarcode = String(rawBarcode).replace(/[^0-9A-Za-z]/g, '').trim();
        if (!cleanBarcode || cleanBarcode.length < 6) continue;

        const item = priceMap.get(cleanBarcode);
        if (!item) {
          if (unmatchedFill === 'zero') {
            if (targetOrigCol > 0) {
              const cell = row.getCell(targetOrigCol);
              cell.value = 0;
              cell.numFmt = '#,##0.00';
            }
          }
          continue;
        }

        totalFilledCount++;

        // Insertion logic
        if (columnTargetMode === 'new_columns' || fillMode === 'original_and_discount') {
          // A. Fill Original Price
          if (targetOrigCol > 0) {
            const origCell = row.getCell(targetOrigCol);
            const origVal = getPriceValue(item, originalPriceType);
            formatCellValue(origVal, origCell);
          }

          // B. Fill Discount Price if enabled
          if (insertDiscountCol && targetDiscCol > 0) {
            const discCell = row.getCell(targetDiscCol);
            const discVal = getPriceValue(item, discountPriceType);

            if (discVal !== null && discVal > 0) {
              formatCellValue(discVal, discCell);
            } else {
              if (discountNonSaleHandling === 'original') {
                const origVal = getPriceValue(item, originalPriceType);
                formatCellValue(origVal, discCell);
              } else if (discountNonSaleHandling === 'dash') {
                discCell.value = '-';
              } else {
                discCell.value = null;
              }
            }
          }

          // C. Fill Platform name if enabled
          if (insertPlatformCol && targetPlatCol > 0) {
            const platCell = row.getCell(targetPlatCol);
            platCell.value = item.lowestOriginalPlatform || item.lowestPlatform || '-';
          }
        } else if (fillMode === 'single') {
          if (Number(singlePriceColIndex) > 0) {
            const targetCell = row.getCell(Number(singlePriceColIndex));
            const val = getPriceValue(item, singlePriceType);
            formatCellValue(val, targetCell);
          }
        } else {
          // Multi column mapped fill
          if (columnMappings.lowestOriginalPriceCol && Number(columnMappings.lowestOriginalPriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.lowestOriginalPriceCol));
            formatCellValue(item.lowestOriginalPrice, cell);
          }
          if (columnMappings.lowestDiscountPriceCol && Number(columnMappings.lowestDiscountPriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.lowestDiscountPriceCol));
            formatCellValue(item.lowestDiscountPrice, cell);
          }
          if (columnMappings.lowestEffectivePriceCol && Number(columnMappings.lowestEffectivePriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.lowestEffectivePriceCol));
            formatCellValue(item.lowestPrice, cell);
          }
          if (columnMappings.meloniOriginalPriceCol && Number(columnMappings.meloniOriginalPriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.meloniOriginalPriceCol));
            formatCellValue(item.prices?.meloni?.originalPrice, cell);
          }
          if (columnMappings.meloniDiscountPriceCol && Number(columnMappings.meloniDiscountPriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.meloniDiscountPriceCol));
            formatCellValue(item.prices?.meloni?.discountPrice, cell);
          }
          if (columnMappings.megaOriginalPriceCol && Number(columnMappings.megaOriginalPriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.megaOriginalPriceCol));
            formatCellValue(item.prices?.megacedi?.originalPrice, cell);
          }
          if (columnMappings.megaDiscountPriceCol && Number(columnMappings.megaDiscountPriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.megaDiscountPriceCol));
            formatCellValue(item.prices?.megacedi?.discountPrice, cell);
          }
          if (columnMappings.maurysOriginalPriceCol && Number(columnMappings.maurysOriginalPriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.maurysOriginalPriceCol));
            formatCellValue(item.prices?.maurys?.originalPrice, cell);
          }
          if (columnMappings.maurysDiscountPriceCol && Number(columnMappings.maurysDiscountPriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.maurysDiscountPriceCol));
            formatCellValue(item.prices?.maurys?.discountPrice, cell);
          }
          if (columnMappings.risparmioOriginalPriceCol && Number(columnMappings.risparmioOriginalPriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.risparmioOriginalPriceCol));
            formatCellValue(item.prices?.risparmiocasa?.originalPrice, cell);
          }
          if (columnMappings.risparmioDiscountPriceCol && Number(columnMappings.risparmioDiscountPriceCol) > 0) {
            const cell = row.getCell(Number(columnMappings.risparmioDiscountPriceCol));
            formatCellValue(item.prices?.risparmiocasa?.discountPrice, cell);
          }
          if (columnMappings.platformNameCol && Number(columnMappings.platformNameCol) > 0) {
            const cell = row.getCell(Number(columnMappings.platformNameCol));
            cell.value = item.lowestOriginalPlatform || item.lowestPlatform || '-';
          }
          if (columnMappings.titleCol && Number(columnMappings.titleCol) > 0) {
            const cell = row.getCell(Number(columnMappings.titleCol));
            if (!cell.value || String(cell.value).trim() === '') {
              cell.value = item.title;
            }
          }
        }
      }
    }

    const outputBuffer = Buffer.from(await wb.xlsx.writeBuffer());
    const safeOutputName = filename.endsWith('.xlsx')
      ? filename.replace(/\.xlsx$/, '_已插入价格.xlsx')
      : `${filename}_filled_prices.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(safeOutputName)}"`);
    res.setHeader('Content-Length', String(outputBuffer.length));
    res.setHeader('Access-Control-Expose-Headers', 'Content-Length, Content-Disposition, X-Filled-Count, X-Processed-Sheets');
    res.setHeader('X-Filled-Count', String(totalFilledCount));
    res.setHeader('X-Processed-Sheets', String(processedSheetsCount));
    res.send(outputBuffer);
  } catch (err: any) {
    console.error('Error filling price template:', err);
    res.status(500).json({ success: false, error: '回填价格模板失败: ' + err.message });
  }
});

// ----------------------------------------------------
// API: Standalone Multi-Site Price Comparison Excel
// ----------------------------------------------------
app.post('/api/prices/export-xlsx', async (req: Request, res: Response) => {
  try {
    const { items = [] } = req.body;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'CASA & TE 价格提取引擎';
    const ws = wb.addWorksheet('全网多平台价格对比');

    ws.columns = [
      { header: '条形码 (Barcode)', key: 'barcode', width: 18 },
      { header: '商品名称 (Title)', key: 'title', width: 40 },
      { header: '综合最低原价 (€)', key: 'lowestOriginalPrice', width: 16 },
      { header: '综合折后特价 (€)', key: 'lowestDiscountPrice', width: 16 },
      { header: '特价状态', key: 'hasDiscount', width: 12 },
      { header: '最低价渠道', key: 'lowestPlatform', width: 16 },
      { header: 'Meloni 原价 (€)', key: 'meloniOriginal', width: 16 },
      { header: 'Meloni 折后价 (€)', key: 'meloniDiscount', width: 16 },
      { header: 'MegaCedi 原价/底价 (€)', key: 'megaOriginal', width: 18 },
      { header: 'MegaCedi 折后价 (€)', key: 'megaDiscount', width: 16 },
      { header: "Maury's 原价 (€)", key: 'maurysOriginal', width: 16 },
      { header: "Maury's 折后价 (€)", key: 'maurysDiscount', width: 16 },
      { header: 'Risparmio Casa 原价 (€)', key: 'risparmioOriginal', width: 18 },
      { header: 'Risparmio Casa 折后价 (€)', key: 'risparmioDiscount', width: 18 },
      { header: 'PiùMe 原价 (€)', key: 'piumeOriginal', width: 16 },
      { header: 'PiùMe 折后价 (€)', key: 'piumeDiscount', width: 16 },
      { header: '提取时间', key: 'extractedAt', width: 22 },
    ];

    // Style header row
    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F172A' },
    };
    headerRow.height = 26;

    let rowIdx = 2;
    for (const item of items) {
      const row = ws.addRow({
        barcode: item.barcode,
        title: item.title,
        lowestOriginalPrice: item.lowestOriginalPrice !== null ? item.lowestOriginalPrice : '',
        lowestDiscountPrice: item.lowestDiscountPrice !== null ? item.lowestDiscountPrice : '',
        hasDiscount: item.hasDiscount ? '打折促销中' : '正价',
        lowestPlatform: item.lowestPlatform || '-',
        meloniOriginal: item.prices?.meloni?.originalPrice !== null ? item.prices?.meloni?.originalPrice : '',
        meloniDiscount: item.prices?.meloni?.discountPrice !== null ? item.prices?.meloni?.discountPrice : '',
        megaOriginal: item.prices?.megacedi?.originalPrice !== null ? item.prices?.megacedi?.originalPrice : '',
        megaDiscount: item.prices?.megacedi?.discountPrice !== null ? item.prices?.megacedi?.discountPrice : '',
        maurysOriginal: item.prices?.maurys?.originalPrice !== null ? item.prices?.maurys?.originalPrice : '',
        maurysDiscount: item.prices?.maurys?.discountPrice !== null ? item.prices?.maurys?.discountPrice : '',
        risparmioOriginal: item.prices?.risparmiocasa?.originalPrice !== null ? item.prices?.risparmiocasa?.originalPrice : '',
        risparmioDiscount: item.prices?.risparmiocasa?.discountPrice !== null ? item.prices?.risparmiocasa?.discountPrice : '',
        piumeOriginal: item.prices?.piume?.originalPrice !== null ? item.prices?.piume?.originalPrice : '',
        piumeDiscount: item.prices?.piume?.discountPrice !== null ? item.prices?.piume?.discountPrice : '',
        extractedAt: item.extractedAt || new Date().toISOString(),
      });

      // Format numeric price columns
      const numericCols = [3, 4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16];
      for (const colIdx of numericCols) {
        const cell = row.getCell(colIdx);
        if (cell.value && typeof cell.value === 'number') {
          cell.numFmt = '€#,##0.00';
        }
      }

      // Highlight lowest prices
      row.getCell(3).font = { bold: true, color: { argb: 'FF0284C7' } }; // Blue for original
      if (item.lowestDiscountPrice !== null) {
        row.getCell(4).font = { bold: true, color: { argb: 'FF16A34A' } }; // Green for discount
      }

      rowIdx++;
    }

    const outputBuffer = Buffer.from(await wb.xlsx.writeBuffer());
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="multisite_prices_with_discount_${items.length}items.xlsx"`);
    res.send(outputBuffer);
  } catch (err: any) {
    console.error('Error exporting price xlsx:', err);
    res.status(500).send('导出价格比对表格失败: ' + err.message);
  }
});

// Boot dev server with Vite or production static
async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.resolve(__dirname, 'dist')));
    app.get('*', (_req, res) => {
      res.sendFile(path.resolve(__dirname, 'dist', 'index.html'));
    });
  }

  app.listen(Number(PORT), '0.0.0.0', () => {
    console.log(`Server listening on http://0.0.0.0:${PORT}`);
  });
}

startServer();
