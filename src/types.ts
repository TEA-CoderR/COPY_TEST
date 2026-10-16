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

export interface SiteMetaInfo {
  url: string;
  title: string;
  totalCount: number;
  totalPages: number;
  pageSize: number;
  siteType?: 'maurys' | 'risparmiocasa' | 'meloni' | 'megacedi' | 'generic';
  sampleProducts?: ScrapedProduct[];
  requiresAuth?: boolean;
}

export type ViewMode = 'cards' | 'table' | 'code' | 'tiktok';
export type ExportFormat = 'csv' | 'xlsx' | 'json' | 'images-txt' | 'tiktok-csv' | 'zip-images';

export interface CategoryPreset {
  id: string;
  name: string;
  url: string;
  tag: string;
  site: 'maurys' | 'risparmiocasa' | 'meloni' | 'megacedi' | 'all';
  description: string;
}

export interface MegaCediSessionStatus {
  authenticated: boolean;
  username?: string;
  hasCustomCookie: boolean;
  message?: string;
}

export interface PlatformAuthStatus {
  authenticated: boolean;
  username?: string;
  hasCustomCookie: boolean;
  siteName: string;
  domain: string;
  requiresLogin: boolean;
  message?: string;
}

export type AllPlatformAuthStatuses = Record<PriceSourceSite, PlatformAuthStatus>;

export interface BarcodeMatchResult {
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
}

export interface BarcodeMatcherStats {
  total: number;
  processed: number;
  matchedMeloni: number;
  matchedMega: number;
  matchedCached: number;
  unmatched: number;
}

export type PriceSourceSite = 'meloni' | 'megacedi' | 'maurys' | 'risparmiocasa' | 'piume';

export interface SitePriceDetail {
  found: boolean;
  site: PriceSourceSite;
  siteName: string;
  price: number | null; // effective price
  originalPrice: number | null; // 原价 / 划线价
  discountPrice: number | null; // 折后特价 (若无活动则为 null)
  formattedPrice: string;
  formattedOriginalPrice?: string;
  formattedDiscountPrice?: string;
  discountPercent?: string;
  hasDiscount: boolean;
  title?: string;
  url?: string;
  inStock?: boolean;
  note?: string;
}

export interface BarcodePriceItem {
  barcode: string;
  title: string;
  matched: boolean;
  lowestPrice: number | null; // 综合有效最低价
  lowestOriginalPrice: number | null; // 综合最低原价
  lowestDiscountPrice: number | null; // 综合最低打折价 (若有打折)
  lowestPlatform?: string;
  lowestOriginalPlatform?: string;
  highestPrice: number | null;
  highestPlatform?: string;
  priceDiff?: number | null;
  averagePrice?: number | null;
  hasDiscount: boolean;
  prices: {
    meloni?: SitePriceDetail;
    megacedi?: SitePriceDetail;
    maurys?: SitePriceDetail;
    risparmiocasa?: SitePriceDetail;
    piume?: SitePriceDetail;
    [key: string]: SitePriceDetail | undefined;
  };
  extractedAt: string;
}

export interface PriceExtractorStats {
  total: number;
  processed: number;
  matchedCount: number;
  unmatchedCount: number;
  meloniFound: number;
  megaFound: number;
  maurysFound: number;
  risparmioFound: number;
  piumeFound: number;
}

export interface SheetBarcodeSummary {
  sheetName: string;
  count: number;
  barcodes: string[];
}

