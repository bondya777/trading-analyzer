import * as XLSX from 'xlsx';

// ============================================================
// 类型定义
// ============================================================

export interface AccountStats {
  账户代码: string;
  账户名称: string;
  交易金额: number;
  交易笔数: number;
}

export interface AnalysisResult {
  totalHgRecords: number;
  totalZqRecords: number;
  totalAccounts: number;
  hgAccounts: number;
  zqAccounts: number;
  filteredHgRecords: number;
  hgTop10ByValue: AccountStats[];
  hgTop10ByCount: AccountStats[];
  xyhgTop10ByValue: AccountStats[];
  xyhgTop10ByCount: AccountStats[];
  xyhgAccounts: XyhgAccountDetail[];
  longTermTop10ByValue: AccountStats[];
  longTermTop10ByCount: AccountStats[];
  longTermAccounts: BondAccountDetail[];
  fundTop10ByValue: AccountStats[];
  fundTop10ByCount: AccountStats[];
  fundAccounts: BondAccountDetail[];
  zhaiquanTop10ByValue: AccountStats[];
  zhaiquanTop10ByCount: AccountStats[];
  zhaiquanAccounts: BondAccountDetail[];
  qhTop10ByValue: AccountStats[];
  qhTop10ByCount: AccountStats[];
  qhAccounts: BondAccountDetail[];
  lowActivityAccounts: LowActivityAccount[];
  zeroTradeAccounts: AccountStats[];
  absTop10ByValue: AccountStats[];
  absTop10ByCount: AccountStats[];
  reitsTop10ByValue: AccountStats[];
  reitsTop10ByCount: AccountStats[];
  specialTop10: SpecialAccount[];
  // === 新增 ===
  monthlyTrend: MonthlyTrend[];
  heatmapData: HeatmapData[];
  heatmapMonths: string[];
  radarAccounts: RadarAccountData[];
  termStructure: TermStructureData[];
  anomalies: AnomalyRecord[];
  accountList: { 账户代码: string; 账户名称: string; 总笔数: number; 总金额: number }[];
  allAccountsDetail: Record<string, BondAccountDetail>;
}

export interface MonthlyTrend {
  月份: string;
  回购业务: number;
  债券业务: number;
  基金: number;
  可转债: number;
  国债期货: number;
}

export interface HeatmapData {
  账户代码: string;
  账户名称: string;
  [月份: string]: number | string;
}

export interface RadarAccountData {
  账户代码: string;
  账户名称: string;
  回购活跃度: number;
  债券活跃度: number;
  基金活跃度: number;
  期货活跃度: number;
  长期债券偏好: number;
  交易频次: number;
}

export interface TermStructureData {
  期限段: string;
  买入金额: number;
  卖出金额: number;
  净买卖: number;
}

export interface AnomalyRecord {
  账户代码: string;
  账户名称: string;
  业务品种: string;
  交易方向: string;
  名称: string;
  金额: number;
  日期?: string;
  类型: '超大额' | '高频率';
}

export interface XyhgAccountDetail {
  账户代码: string;
  账户名称: string;
  协议回购金额: number;
  协议回购笔数: number;
  全部回购类型: Record<string, number>;
  交易方向分布: Record<string, number>;
}

export interface BondCategoryDetail {
  类别: string;
  买入金额: number;
  买入笔数: number;
  卖出金额: number;
  卖出笔数: number;
  净买卖: number;
  平均期限?: number;
  品种分布: Record<string, number>;
}

export interface BondAccountDetail {
  账户代码: string;
  账户名称: string;
  总笔数: number;
  债券类别: BondCategoryDetail[];
  基金分析?: FundAnalysis;
  转债分析?: ZqAnalysis;
  期货分析?: QhAnalysis;
  买卖方向: Record<string, number>;
  月度金额?: Record<string, number>;
}

export interface FundAnalysis {
  买入金额: number;
  买入笔数: number;
  卖出金额: number;
  卖出笔数: number;
  涉及基金: string[];
}

export interface ZqAnalysis {
  买入金额: number;
  买入笔数: number;
  卖出金额: number;
  卖出笔数: number;
  涉及转债: string[];
}

export interface QhAnalysis {
  笔数: number;
  总金额: number;
  涉及品种: string[];
}

export interface LowActivityAccount {
  账户代码: string;
  账户名称: string;
  总交易笔数: number;
  总交易金额: number;
  持有品种: string[];
  交易明细: TradeDetail[];
}

export interface TradeDetail {
  业务品种: string;
  交易方向: string;
  名称: string;
  金额: number;
}

export interface SpecialAccount {
  排名: number;
  账户代码: string;
  账户名称: string;
  业务品种: string;
  金额: number;
  笔数: number;
}

// ============================================================
// 辅助函数
// ============================================================

function toNumber(v: unknown): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const n = parseFloat(v.replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

function toStr(v: unknown): string {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function parseDate(v: unknown): Date | undefined {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === 'number') {
    // Excel serial date: days since 1899-12-30
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(epoch.getTime() + v * 86400000);
  }
  if (typeof v === 'string') {
    const d = new Date(v);
    if (!isNaN(d.getTime())) return d;
    // try YYYYMMDD
    if (/^\d{8}$/.test(v)) {
      const y = parseInt(v.slice(0, 4)), m = parseInt(v.slice(4, 6)) - 1, day = parseInt(v.slice(6, 8));
      const dd = new Date(y, m, day);
      if (!isNaN(dd.getTime())) return dd;
    }
  }
  return undefined;
}

function getMonthKey(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function findDateField(row: Record<string, unknown>): Date | undefined {
  const dateFields = ['交易日期', '日期', '成交日期', '交割日期', '结算日期', 'Date', '日期.1'];
  for (const f of dateFields) {
    const v = row[f];
    if (v !== undefined && v !== null && v !== '') {
      const d = parseDate(v);
      if (d) return d;
    }
  }
  // Try any key that contains '日期' or 'date' (case-insensitive)
  for (const [k, v] of Object.entries(row)) {
    if (/日期|date/i.test(k)) {
      const d = parseDate(v);
      if (d) return d;
    }
  }
  return undefined;
}

function parseTerm(termStr: string): number | undefined {
  if (!termStr) return undefined;
  const s = termStr.trim();

  let m = s.match(/^(\d+)年以上$/);
  if (m) return parseFloat(m[1]) + 5;

  m = s.match(/^(\d+)[\-~](\d+)年$/);
  if (m) return (parseFloat(m[1]) + parseFloat(m[2])) / 2;

  m = s.match(/^(\d+(?:\.\d+)?)年$/);
  if (m) return parseFloat(m[1]);

  m = s.match(/^(\d+(?:\.\d+)?)月$/);
  if (m) return parseFloat(m[1]) / 12;

  m = s.match(/^(\d+(?:\.\d+)?)日$/);
  if (m) return parseFloat(m[1]) / 365;

  m = s.match(/^(\d+(?:\.\d+)?)天$/);
  if (m) return parseFloat(m[1]) / 365;

  return undefined;
}

function getTermForRow(row: Record<string, unknown>): number | undefined {
  const t1 = parseTerm(toStr(row['期限']));
  if (t1 !== undefined) return t1;
  return parseTerm(toStr(row['期限.1']));
}

function getTermBucket(term: number | undefined): string {
  if (term === undefined) return '未知';
  if (term < 1) return '0-1年';
  if (term < 3) return '1-3年';
  if (term < 5) return '3-5年';
  if (term < 10) return '5-10年';
  if (term < 30) return '10-30年';
  return '30年+';
}

// ============================================================
// Excel 解析
// ============================================================

export async function parseExcel(file: File): Promise<{ hg: Record<string, unknown>[]; zq: Record<string, unknown>[] }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetNames = workbook.SheetNames;
        let hgSheet = workbook.Sheets[sheetNames[0]];
        let zqSheet = workbook.Sheets[sheetNames[1]];

        for (const name of sheetNames) {
          if (name.includes('回购')) hgSheet = workbook.Sheets[name];
          if (name.includes('债券') || name.includes('其他')) zqSheet = workbook.Sheets[name];
        }

        const hg = XLSX.utils.sheet_to_json<Record<string, unknown>>(hgSheet);
        const zq = XLSX.utils.sheet_to_json<Record<string, unknown>>(zqSheet);

        resolve({ hg, zq });
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// ============================================================
// 数据预处理
// ============================================================

function preprocessHg(records: Record<string, unknown>[]) {
  return records.map(r => {
    const date = findDateField(r);
    return {
      ...r,
      _账户代码: toStr(r['内部证券账户代码']),
      _账户名称: toStr(r['内部证券账户名称']),
      _业务品种: toStr(r['业务品种']),
      _交易品种: toStr(r['交易品种']),
      _交易方向: toStr(r['交易方向']),
      _名称: toStr(r['名称']),
      _债券类别: '',
      _债券品种: '',
      _金额元: toNumber(r['交易金额(元)']),
      _金额亿: toNumber(r['交易金额(元)']) / 1e8,
      _日期: date,
      _月份: date ? getMonthKey(date) : undefined,
    };
  });
}

function preprocessZq(records: Record<string, unknown>[]) {
  return records.map(r => {
    const date = findDateField(r);
    return {
      ...r,
      _账户代码: toStr(r['内部证券账户代码']),
      _账户名称: toStr(r['内部证券账户名称']),
      _业务品种: toStr(r['业务品种']),
      _交易品种: toStr(r['交易品种']),
      _交易方向: toStr(r['交易方向']),
      _债券类别: toStr(r['债券类别']),
      _债券品种: toStr(r['债券品种']),
      _名称: toStr(r['名称']),
      _期限: toStr(r['期限']),
      _期限1: toStr(r['期限.1']),
      _金额元: toNumber(r['交易金额(元)']),
      _金额亿: toNumber(r['交易金额(元)']) / 1e8,
      _日期: date,
      _月份: date ? getMonthKey(date) : undefined,
      _期限数值: getTermForRow(r),
      _期限段: getTermBucket(getTermForRow(r)),
    };
  });
}

// ============================================================
// Top10 计算
// ============================================================

function getTop10(
  records: Record<string, unknown>[],
  groupKey: string = '_账户代码',
  nameKey: string = '_账户名称',
  valueKey: string = '_金额亿'
): { byValue: AccountStats[]; byCount: AccountStats[] } {
  const grouped = new Map<string, { name: string; amount: number; count: number }>();

  for (const r of records) {
    const code = toStr(r[groupKey]);
    const name = toStr(r[nameKey]) || code;
    if (!code) continue;

    const existing = grouped.get(code);
    if (existing) {
      existing.amount += toNumber(r[valueKey]);
      existing.count += 1;
    } else {
      grouped.set(code, { name, amount: toNumber(r[valueKey]), count: 1 });
    }
  }

  const arr = Array.from(grouped.entries()).map(([code, v]) => ({
    账户代码: code,
    账户名称: v.name,
    交易金额: Math.round(v.amount * 10000) / 10000,
    交易笔数: v.count,
  }));

  const byValue = [...arr].sort((a, b) => b.交易金额 - a.交易金额).slice(0, 10);
  const byCount = [...arr].sort((a, b) => b.交易笔数 - a.交易笔数).slice(0, 10);

  return { byValue, byCount };
}

// ============================================================
// 债券账户深度分析
// ============================================================

function analyzeBondAccount(
  dfZq: Record<string, unknown>[],
  code: string,
  name: string
): BondAccountDetail {
  const acct = dfZq.filter(r => toStr(r['_账户代码']) === code);

  const bondCats: BondCategoryDetail[] = [];
  for (const cat of ['利率债', '信用债']) {
    const catDf = acct.filter(r => toStr(r['_债券类别']) === cat);
    if (catDf.length === 0) continue;

    const buyDf = catDf.filter(r => ['买入', '分销买入'].includes(toStr(r['_交易方向'])));
    const sellDf = catDf.filter(r => ['卖出', '分销卖出'].includes(toStr(r['_交易方向'])));

    const buyAmount = buyDf.reduce((s, r) => s + toNumber(r['_金额亿']), 0);
    const sellAmount = sellDf.reduce((s, r) => s + toNumber(r['_金额亿']), 0);

    const terms = catDf.map(r => getTermForRow(r)).filter(t => t !== undefined) as number[];
    const avgTerm = terms.length > 0 ? terms.reduce((a, b) => a + b, 0) / terms.length : undefined;

    const varieties: Record<string, number> = {};
    for (const r of catDf) {
      const v = toStr(r['_债券品种']);
      if (v && v !== 'nan') {
        varieties[v] = (varieties[v] || 0) + 1;
      }
    }

    bondCats.push({
      类别: cat,
      买入金额: Math.round(buyAmount * 10000) / 10000,
      买入笔数: buyDf.length,
      卖出金额: Math.round(sellAmount * 10000) / 10000,
      卖出笔数: sellDf.length,
      净买卖: Math.round((buyAmount - sellAmount) * 10000) / 10000,
      平均期限: avgTerm !== undefined ? Math.round(avgTerm * 100) / 100 : undefined,
      品种分布: varieties,
    });
  }

  let fundAnalysis: FundAnalysis | undefined;
  const fundDf = acct.filter(r => toStr(r['_业务品种']).includes('基金') || toStr(r['_交易品种']).includes('基金'));
  if (fundDf.length > 0) {
    const fundBuy = fundDf.filter(r => ['买入', '分销买入'].includes(toStr(r['_交易方向'])));
    const fundSell = fundDf.filter(r => ['卖出', '分销卖出'].includes(toStr(r['_交易方向'])));
    fundAnalysis = {
      买入金额: Math.round(fundBuy.reduce((s, r) => s + toNumber(r['_金额亿']), 0) * 10000) / 10000,
      买入笔数: fundBuy.length,
      卖出金额: Math.round(fundSell.reduce((s, r) => s + toNumber(r['_金额亿']), 0) * 10000) / 10000,
      卖出笔数: fundSell.length,
      涉及基金: [...new Set(fundDf.map(r => toStr(r['_名称'])).filter(n => n && n !== 'nan'))].slice(0, 10),
    };
  }

  let zqAnalysis: ZqAnalysis | undefined;
  const zqDf = acct.filter(r =>
    toStr(r['_业务品种']).includes('转债') ||
    toStr(r['_债券品种']).includes('转债') ||
    toStr(r['_债券类别']).includes('转债')
  );
  if (zqDf.length > 0) {
    const zqBuy = zqDf.filter(r => ['买入', '分销买入'].includes(toStr(r['_交易方向'])));
    const zqSell = zqDf.filter(r => ['卖出', '分销卖出'].includes(toStr(r['_交易方向'])));
    zqAnalysis = {
      买入金额: Math.round(zqBuy.reduce((s, r) => s + toNumber(r['_金额亿']), 0) * 10000) / 10000,
      买入笔数: zqBuy.length,
      卖出金额: Math.round(zqSell.reduce((s, r) => s + toNumber(r['_金额亿']), 0) * 10000) / 10000,
      卖出笔数: zqSell.length,
      涉及转债: [...new Set(zqDf.map(r => toStr(r['_名称'])).filter(n => n && n !== 'nan'))].slice(0, 10),
    };
  }

  let qhAnalysis: QhAnalysis | undefined;
  const qhDf = acct.filter(r => toStr(r['_业务品种']).includes('期货') || toStr(r['_交易品种']).includes('期货'));
  if (qhDf.length > 0) {
    qhAnalysis = {
      笔数: qhDf.length,
      总金额: Math.round(qhDf.reduce((s, r) => s + toNumber(r['_金额亿']), 0) * 10000) / 10000,
      涉及品种: [...new Set(qhDf.map(r => toStr(r['_名称'])).filter(n => n && n !== 'nan'))].slice(0, 10),
    };
  }

  const dirSummary: Record<string, number> = {};
  for (const r of acct) {
    const d = toStr(r['_交易方向']) || '未知';
    dirSummary[d] = (dirSummary[d] || 0) + toNumber(r['_金额亿']);
  }
  for (const k of Object.keys(dirSummary)) {
    dirSummary[k] = Math.round(dirSummary[k] * 10000) / 10000;
  }

  const monthlyAmount: Record<string, number> = {};
  for (const r of acct) {
    const m = toStr(r['_月份']);
    if (m) {
      monthlyAmount[m] = (monthlyAmount[m] || 0) + toNumber(r['_金额亿']);
    }
  }
  for (const k of Object.keys(monthlyAmount)) {
    monthlyAmount[k] = Math.round(monthlyAmount[k] * 10000) / 10000;
  }

  return {
    账户代码: code,
    账户名称: name,
    总笔数: acct.length,
    债券类别: bondCats,
    基金分析: fundAnalysis,
    转债分析: zqAnalysis,
    期货分析: qhAnalysis,
    买卖方向: dirSummary,
    月度金额: monthlyAmount,
  };
}

// ============================================================
// 低活跃度账户分析
// ============================================================

function getAccountHoldings(records: Record<string, unknown>[], code: string): string[] {
  const acct = records.filter(r => toStr(r['_账户代码']) === code);
  const holdings = new Set<string>();

  for (const r of acct) {
    const biz = toStr(r['_业务品种']);
    if (biz && biz !== 'nan') holdings.add(`回购-${biz}`);

    const name = toStr(r['_名称']);
    if (name && name !== 'nan') holdings.add(name);

    const variety = toStr(r['_债券品种']);
    if (variety && variety !== 'nan') holdings.add(`债券品种-${variety}`);

    const cat = toStr(r['_债券类别']);
    if (cat && cat !== 'nan') holdings.add(`债券类别-${cat}`);
  }

  return Array.from(holdings);
}

// ============================================================
// 月度趋势分析
// ============================================================

function buildMonthlyTrend(
  dfHg: Record<string, unknown>[],
  dfZq: Record<string, unknown>[]
): { trend: MonthlyTrend[]; months: string[] } {
  const monthsSet = new Set<string>();
  for (const r of [...dfHg, ...dfZq]) {
    const m = toStr(r['_月份']);
    if (m) monthsSet.add(m);
  }
  const months = Array.from(monthsSet).sort();

  const trend: MonthlyTrend[] = [];
  for (const m of months) {
    const hgAmount = dfHg
      .filter(r => toStr(r['_月份']) === m)
      .reduce((s, r) => s + toNumber(r['_金额亿']), 0);
    const zqAmount = dfZq
      .filter(r => toStr(r['_月份']) === m)
      .reduce((s, r) => s + toNumber(r['_金额亿']), 0);
    const fundAmount = dfZq
      .filter(r => toStr(r['_月份']) === m && (toStr(r['_业务品种']).includes('基金') || toStr(r['_交易品种']).includes('基金')))
      .reduce((s, r) => s + toNumber(r['_金额亿']), 0);
    const zqConvertible = dfZq
      .filter(r => toStr(r['_月份']) === m && (toStr(r['_业务品种']).includes('转债') || toStr(r['_债券品种']).includes('转债')))
      .reduce((s, r) => s + toNumber(r['_金额亿']), 0);
    const qhAmount = dfZq
      .filter(r => toStr(r['_月份']) === m && (toStr(r['_业务品种']).includes('期货') || toStr(r['_交易品种']).includes('期货')))
      .reduce((s, r) => s + toNumber(r['_金额亿']), 0);

    trend.push({
      月份: m,
      回购业务: Math.round(hgAmount * 10000) / 10000,
      债券业务: Math.round(zqAmount * 10000) / 10000,
      基金: Math.round(fundAmount * 10000) / 10000,
      可转债: Math.round(zqConvertible * 10000) / 10000,
      国债期货: Math.round(qhAmount * 10000) / 10000,
    });
  }

  return { trend, months };
}

// ============================================================
// 热力图数据
// ============================================================

function buildHeatmap(
  dfAll: Record<string, unknown>[],
  months: string[]
): HeatmapData[] {
  const acctMap = new Map<string, { name: string; monthly: Record<string, number> }>();

  for (const r of dfAll) {
    const code = toStr(r['_账户代码']);
    const name = toStr(r['_账户名称']) || code;
    const m = toStr(r['_月份']);
    if (!code || !m) continue;

    const ex = acctMap.get(code);
    if (ex) {
      ex.monthly[m] = (ex.monthly[m] || 0) + toNumber(r['_金额亿']);
    } else {
      acctMap.set(code, { name, monthly: { [m]: toNumber(r['_金额亿']) } });
    }
  }

  // 取交易金额Top 30的账户用于热力图
  const sorted = Array.from(acctMap.entries())
    .map(([code, v]) => ({
      code,
      name: v.name,
      total: Object.values(v.monthly).reduce((a, b) => a + b, 0),
      monthly: v.monthly,
    }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 30);

  return sorted.map(v => {
    const row: HeatmapData = {
      账户代码: v.code,
      账户名称: v.name,
    };
    for (const m of months) {
      row[m] = Math.round((v.monthly[m] || 0) * 10000) / 10000;
    }
    return row;
  });
}

// ============================================================
// 雷达图数据
// ============================================================

function buildRadarData(
  dfHg: Record<string, unknown>[],
  dfZq: Record<string, unknown>[]
): RadarAccountData[] {
  const dfAll = [...dfHg, ...dfZq];


  const acctMap = new Map<string, { name: string; hgAmount: number; zqAmount: number; fundAmount: number; zqConvAmount: number; qhAmount: number; longTermAmount: number; count: number }>();

  for (const r of dfAll) {
    const code = toStr(r['_账户代码']);
    const name = toStr(r['_账户名称']) || code;
    if (!code) continue;

    const isHg = dfHg.includes(r as any);
    const amount = toNumber(r['_金额亿']);
    const isFund = toStr(r['_业务品种']).includes('基金') || toStr(r['_交易品种']).includes('基金');
    const isConvertible = toStr(r['_业务品种']).includes('转债') || toStr(r['_债券品种']).includes('转债');
    const isQh = toStr(r['_业务品种']).includes('期货') || toStr(r['_交易品种']).includes('期货');
    const isLongTerm = ['10-30年', '30年以上'].includes(toStr(r['_期限1'])) || ['10年', '15年', '20年', '30年', '50年'].includes(toStr(r['_期限']));

    const ex = acctMap.get(code);
    if (ex) {
      ex.count += 1;
      if (isHg) ex.hgAmount += amount;
      else ex.zqAmount += amount;
      if (isFund) ex.fundAmount += amount;
      if (isConvertible) ex.zqConvAmount += amount;
      if (isQh) ex.qhAmount += amount;
      if (isLongTerm) ex.longTermAmount += amount;
    } else {
      acctMap.set(code, {
        name,
        hgAmount: isHg ? amount : 0,
        zqAmount: !isHg ? amount : 0,
        fundAmount: isFund ? amount : 0,
        zqConvAmount: isConvertible ? amount : 0,
        qhAmount: isQh ? amount : 0,
        longTermAmount: isLongTerm ? amount : 0,
        count: 1,
      });
    }
  }

  // 取Top 15账户
  const sorted = Array.from(acctMap.entries())
    .map(([code, v]) => ({ code, ...v, total: v.hgAmount + v.zqAmount }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 15);

  // 全局归一化基准
  const maxCount = Math.max(...sorted.map(v => v.count), 1);
  const maxHgA = Math.max(...sorted.map(v => v.hgAmount), 0.0001);
  const maxZqA = Math.max(...sorted.map(v => v.zqAmount), 0.0001);
  const maxFundA = Math.max(...sorted.map(v => v.fundAmount), 0.0001);
  const maxQhA = Math.max(...sorted.map(v => v.qhAmount), 0.0001);
  

  return sorted.map(v => ({
    账户代码: v.code,
    账户名称: v.name,
    回购活跃度: Math.round((v.hgAmount / maxHgA) * 100),
    债券活跃度: Math.round((v.zqAmount / maxZqA) * 100),
    基金活跃度: Math.round((v.fundAmount / maxFundA) * 100),
    期货活跃度: Math.round((v.qhAmount / maxQhA) * 100),
    长期债券偏好: Math.round((v.longTermAmount / Math.max(v.zqAmount, 0.0001)) * 100),
    交易频次: Math.round((v.count / maxCount) * 100),
  }));
}

// ============================================================
// 期限结构数据
// ============================================================

function buildTermStructure(dfZq: Record<string, unknown>[]): TermStructureData[] {
  const buckets = ['0-1年', '1-3年', '3-5年', '5-10年', '10-30年', '30年+', '未知'];
  const data = new Map<string, { buy: number; sell: number }>();

  for (const b of buckets) data.set(b, { buy: 0, sell: 0 });

  for (const r of dfZq) {
    const bucket = toStr(r['_期限段']) || '未知';
    const dir = toStr(r['_交易方向']);
    const amount = toNumber(r['_金额亿']);
    const ex = data.get(bucket);
    if (!ex) continue;
    if (['买入', '分销买入'].includes(dir)) ex.buy += amount;
    else if (['卖出', '分销卖出'].includes(dir)) ex.sell += amount;
  }

  return buckets.map(b => {
    const ex = data.get(b)!;
    return {
      期限段: b,
      买入金额: Math.round(ex.buy * 10000) / 10000,
      卖出金额: Math.round(ex.sell * 10000) / 10000,
      净买卖: Math.round((ex.buy - ex.sell) * 10000) / 10000,
    };
  });
}

// ============================================================
// 异常检测
// ============================================================

function detectAnomalies(dfAll: Record<string, unknown>[]): AnomalyRecord[] {
  const anomalies: AnomalyRecord[] = [];

  // 超大额：单笔超过 50 亿
  for (const r of dfAll) {
    const amount = toNumber(r['_金额亿']);
    if (amount >= 50) {
      anomalies.push({
        账户代码: toStr(r['_账户代码']),
        账户名称: toStr(r['_账户名称']),
        业务品种: toStr(r['_业务品种']) || toStr(r['_交易品种']) || '未知',
        交易方向: toStr(r['_交易方向']) || '未知',
        名称: toStr(r['_名称']) || '',
        金额: Math.round(amount * 10000) / 10000,
        日期: toStr(r['_月份']) || undefined,
        类型: '超大额',
      });
    }
  }

  // 高频：同一账户同一月份同一品种超过 50 笔
  const freqMap = new Map<string, { count: number; code: string; name: string; biz: string; month: string }>();
  for (const r of dfAll) {
    const code = toStr(r['_账户代码']);
    const biz = toStr(r['_业务品种']) || toStr(r['_交易品种']);
    const month = toStr(r['_月份']);
    if (!code || !biz || !month) continue;
    const key = `${code}|${biz}|${month}`;
    const ex = freqMap.get(key);
    if (ex) {
      ex.count += 1;
    } else {
      freqMap.set(key, { count: 1, code, name: toStr(r['_账户名称']) || code, biz, month });
    }
  }
  for (const [, v] of freqMap.entries()) {
    if (v.count >= 50) {
      anomalies.push({
        账户代码: v.code,
        账户名称: v.name,
        业务品种: v.biz,
        交易方向: '',
        名称: '',
        金额: 0,
        日期: v.month,
        类型: '高频率',
      });
    }
  }

  // 按金额降序
  anomalies.sort((a, b) => b.金额 - a.金额);
  return anomalies.slice(0, 50);
}

// ============================================================
// 主分析函数
// ============================================================

export function runAnalysis(hgRaw: Record<string, unknown>[], zqRaw: Record<string, unknown>[]): AnalysisResult {
  const dfHg = preprocessHg(hgRaw);
  const dfZq = preprocessZq(zqRaw);

  const dfAll = [...dfHg, ...dfZq];

  const allHgCodes = new Set(dfHg.map(r => toStr(r['_账户代码'])).filter(Boolean));
  const allZqCodes = new Set(dfZq.map(r => toStr(r['_账户代码'])).filter(Boolean));
  const allAccounts = new Set([...allHgCodes, ...allZqCodes]);

  // 1.1 回购业务（过滤单笔超100亿）
  const dfHgFiltered = dfHg.filter(r => toNumber(r['_金额元']) <= 100 * 1e8);
  const hgTop10 = getTop10(dfHgFiltered);

  // 1.2 协议回购
  const xyhg = dfHgFiltered.filter(r => toStr(r['_业务品种']).includes('协议回购'));
  const xyhgTop10 = getTop10(xyhg);

  const xyhgAccounts: XyhgAccountDetail[] = [];
  for (const stat of xyhgTop10.byValue.slice(0, 5)) {
    const acctAll = dfHgFiltered.filter(r => toStr(r['_账户代码']) === stat.账户代码);
    const bizTypes: Record<string, number> = {};
    const dirDist: Record<string, number> = {};
    for (const r of acctAll) {
      const bz = toStr(r['_业务品种']);
      bizTypes[bz] = (bizTypes[bz] || 0) + 1;
      const dir = toStr(r['_交易方向']);
      dirDist[dir] = (dirDist[dir] || 0) + 1;
    }
    xyhgAccounts.push({
      账户代码: stat.账户代码,
      账户名称: stat.账户名称,
      协议回购金额: stat.交易金额,
      协议回购笔数: stat.交易笔数,
      全部回购类型: bizTypes,
      交易方向分布: dirDist,
    });
  }

  // 2.1 长期债券
  const longTerm = dfZq.filter(r =>
    ['10-30年', '30年以上'].includes(toStr(r['_期限1'])) ||
    ['10年', '15年', '20年', '30年', '50年'].includes(toStr(r['_期限']))
  );
  const longTermTop10 = getTop10(longTerm);

  const longTermAccounts: BondAccountDetail[] = [];
  for (const stat of longTermTop10.byValue.slice(0, 5)) {
    longTermAccounts.push(analyzeBondAccount(dfZq, stat.账户代码, stat.账户名称));
  }

  // 2.2 基金
  const fundDf = dfZq.filter(r => toStr(r['_业务品种']).includes('基金') || toStr(r['_交易品种']).includes('基金'));
  const fundTop10 = getTop10(fundDf);
  const fundAccounts: BondAccountDetail[] = [];
  for (const stat of fundTop10.byValue.slice(0, 5)) {
    fundAccounts.push(analyzeBondAccount(dfZq, stat.账户代码, stat.账户名称));
  }

  // 2.3 可转债
  const zqDf = dfZq.filter(r =>
    toStr(r['_业务品种']).includes('转债') ||
    toStr(r['_债券品种']).includes('转债') ||
    toStr(r['_债券类别']).includes('转债')
  );
  const zhaiquanTop10 = getTop10(zqDf);
  const zhaiquanAccounts: BondAccountDetail[] = [];
  for (const stat of zhaiquanTop10.byValue.slice(0, 5)) {
    zhaiquanAccounts.push(analyzeBondAccount(dfZq, stat.账户代码, stat.账户名称));
  }

  // 2.4 国债期货
  const qhDf = dfZq.filter(r => toStr(r['_业务品种']).includes('期货') || toStr(r['_交易品种']).includes('期货'));
  const qhTop10 = getTop10(qhDf);
  const qhAccounts: BondAccountDetail[] = [];
  for (const stat of qhTop10.byValue.slice(0, 5)) {
    qhAccounts.push(analyzeBondAccount(dfZq, stat.账户代码, stat.账户名称));
  }

  // 3. 低活跃度账户
  const activityMap = new Map<string, { name: string; count: number; amount: number }>();
  for (const r of dfAll) {
    const code = toStr(r['_账户代码']);
    const name = toStr(r['_账户名称']) || code;
    if (!code) continue;
    const ex = activityMap.get(code);
    if (ex) {
      ex.count += 1;
      ex.amount += toNumber(r['_金额亿']);
    } else {
      activityMap.set(code, { name, count: 1, amount: toNumber(r['_金额亿']) });
    }
  }

  const lowActivity: LowActivityAccount[] = [];
  for (const [code, v] of activityMap.entries()) {
    if (v.count <= 2) {
      const records = dfAll.filter(r => toStr(r['_账户代码']) === code);
      const holdings = getAccountHoldings(dfAll, code);
      const details: TradeDetail[] = records.map(r => ({
        业务品种: toStr(r['_业务品种']) || toStr(r['_交易品种']) || '未知',
        交易方向: toStr(r['_交易方向']) || '未知',
        名称: toStr(r['_名称']) || '',
        金额: Math.round(toNumber(r['_金额亿']) * 10000) / 10000,
      }));
      lowActivity.push({
        账户代码: code,
        账户名称: v.name,
        总交易笔数: v.count,
        总交易金额: Math.round(v.amount * 10000) / 10000,
        持有品种: holdings.slice(0, 15),
        交易明细: details,
      });
    }
  }
  lowActivity.sort((a, b) => a.总交易笔数 - b.总交易笔数);

  const zeroTrade: AccountStats[] = [];
  for (const [code, v] of activityMap.entries()) {
    if (v.count === 0) {
      zeroTrade.push({ 账户代码: code, 账户名称: v.name, 交易金额: 0, 交易笔数: 0 });
    }
  }

  // 4. ABS / REITs / 特殊品种
  const absDf = dfZq.filter(r =>
    /资产支持证券|ABS/i.test(toStr(r['_债券品种'])) ||
    /ABS|资产支持/i.test(toStr(r['_名称']))
  );
  const absTop10 = getTop10(absDf);

  const reitsDf = dfZq.filter(r =>
    /REITs|REIT/i.test(toStr(r['_债券品种'])) ||
    /REITs|REIT/i.test(toStr(r['_名称'])) ||
    /REITs|REIT/i.test(toStr(r['_业务品种']))
  );
  const reitsTop10 = getTop10(reitsDf);

  const specialTypes = ['分销买卖', '远期交易', '转托管', '债券回售'];
  const specialDf = dfZq.filter(r =>
    specialTypes.includes(toStr(r['_业务品种']))
  );
  const specialGrouped = new Map<string, { code: string; name: string; biz: string; amount: number; count: number }>();
  for (const r of specialDf) {
    const code = toStr(r['_账户代码']);
    const name = toStr(r['_账户名称']) || code;
    const biz = toStr(r['_业务品种']);
    if (!code) continue;
    const key = `${code}|${biz}`;
    const ex = specialGrouped.get(key);
    if (ex) {
      ex.amount += toNumber(r['_金额亿']);
      ex.count += 1;
    } else {
      specialGrouped.set(key, { code, name, biz, amount: toNumber(r['_金额亿']), count: 1 });
    }
  }
  const specialArr = Array.from(specialGrouped.values())
    .map((v, i) => ({ ...v, 排名: i + 1 }))
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 10)
    .map((v, i) => ({
      排名: i + 1,
      账户代码: v.code,
      账户名称: v.name,
      业务品种: v.biz,
      金额: Math.round(v.amount * 10000) / 10000,
      笔数: v.count,
    }));

  // ===== 新增分析 =====
  const { trend: monthlyTrend, months: heatmapMonths } = buildMonthlyTrend(dfHgFiltered, dfZq);
  const heatmapData = buildHeatmap(dfAll, heatmapMonths);
  const radarAccounts = buildRadarData(dfHgFiltered, dfZq);
  const termStructure = buildTermStructure(dfZq);
  const anomalies = detectAnomalies(dfAll);

  // 账户列表（用于搜索和对比）
  const accountList = Array.from(activityMap.entries())
    .map(([code, v]) => ({
      账户代码: code,
      账户名称: v.name,
      总笔数: v.count,
      总金额: Math.round(v.amount * 10000) / 10000,
    }))
    .sort((a, b) => b.总金额 - a.总金额);

  // 所有账户的详细分析（用于对比和搜索）
  const allAccountsDetail: Record<string, BondAccountDetail> = {};
  for (const code of allAccounts) {
    const name = activityMap.get(code)?.name || code;
    allAccountsDetail[code] = analyzeBondAccount(dfZq, code, name);
  }

  return {
    totalHgRecords: dfHg.length,
    totalZqRecords: dfZq.length,
    totalAccounts: allAccounts.size,
    hgAccounts: allHgCodes.size,
    zqAccounts: allZqCodes.size,
    filteredHgRecords: dfHgFiltered.length,
    hgTop10ByValue: hgTop10.byValue,
    hgTop10ByCount: hgTop10.byCount,
    xyhgTop10ByValue: xyhgTop10.byValue,
    xyhgTop10ByCount: xyhgTop10.byCount,
    xyhgAccounts,
    longTermTop10ByValue: longTermTop10.byValue,
    longTermTop10ByCount: longTermTop10.byCount,
    longTermAccounts,
    fundTop10ByValue: fundTop10.byValue,
    fundTop10ByCount: fundTop10.byCount,
    fundAccounts,
    zhaiquanTop10ByValue: zhaiquanTop10.byValue,
    zhaiquanTop10ByCount: zhaiquanTop10.byCount,
    zhaiquanAccounts,
    qhTop10ByValue: qhTop10.byValue,
    qhTop10ByCount: qhTop10.byCount,
    qhAccounts,
    lowActivityAccounts: lowActivity,
    zeroTradeAccounts: zeroTrade,
    absTop10ByValue: absTop10.byValue,
    absTop10ByCount: absTop10.byCount,
    reitsTop10ByValue: reitsTop10.byValue,
    reitsTop10ByCount: reitsTop10.byCount,
    specialTop10: specialArr,
    monthlyTrend,
    heatmapData,
    heatmapMonths,
    radarAccounts,
    termStructure,
    anomalies,
    accountList,
    allAccountsDetail,
  };
}
