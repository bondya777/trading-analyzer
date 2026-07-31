import { useState, useCallback, useRef, useMemo } from 'react';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  LineChart, Line, AreaChart, Area, RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
  ComposedChart,
} from 'recharts';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';
import { ScrollArea } from '@/components/ui/scroll-area';
import { parseExcel, runAnalysis, type AnalysisResult, type BondAccountDetail, type PositionTimeSeries } from '@/lib/analyzer';
import {
  Upload, TrendingUp, Activity, BarChart3, AlertTriangle, Package, FileText,
  Search, X, GitCompare, Thermometer, Wallet, Scale,
} from 'lucide-react';

// ============================================================
// 通用子组件
// ============================================================

function StatCard({ title, value, subtitle, icon: Icon, color = 'blue' }: {
  title: string; value: string | number; subtitle?: string;
  icon: React.ElementType; color?: string;
}) {
  const colorMap: Record<string, string> = {
    blue: 'text-blue-600', green: 'text-green-600', red: 'text-red-600',
    amber: 'text-amber-600', purple: 'text-purple-600', cyan: 'text-cyan-600',
  };
  return (
    <Card>
      <CardContent className="p-4">
        <div className="flex items-center justify-between">
          <div>
            <p className="text-xs text-muted-foreground">{title}</p>
            <p className={`text-xl font-bold ${colorMap[color] || colorMap.blue}`}>{value}</p>
            {subtitle && <p className="text-xs text-muted-foreground mt-1">{subtitle}</p>}
          </div>
          <Icon className={`w-5 h-5 ${colorMap[color] || colorMap.blue}`} />
        </div>
      </CardContent>
    </Card>
  );
}

function Top10Section({ title, byValue, byCount, color = '#3b82f6' }: {
  title: string;
  byValue: { 账户代码: string; 账户名称: string; 交易金额: number; 交易笔数: number }[];
  byCount: { 账户代码: string; 账户名称: string; 交易金额: number; 交易笔数: number }[];
  color?: string;
}) {
  const chartData = byValue.map((item) => ({
    name: item.账户名称.length > 12 ? item.账户名称.slice(0, 12) + '...' : item.账户名称,
    金额: item.交易金额,
    笔数: item.交易笔数,
  }));
  return (
    <div className="space-y-4">
      <h3 className="text-sm font-semibold text-muted-foreground">{title}</h3>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Card>
          <CardHeader className="pb-2"><CardTitle className="text-sm">按金额 Top10</CardTitle></CardHeader>
          <CardContent>
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={chartData} layout="vertical" margin={{ left: 80, right: 20 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" />
                <YAxis dataKey="name" type="category" width={75} tick={{ fontSize: 11 }} />
                <Tooltip formatter={(v: number) => `${v} 亿`} />
                <Bar dataKey="金额" fill={color} radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
            <Table>
              <TableHeader><TableRow><TableHead>排名</TableHead><TableHead>账户代码</TableHead><TableHead>账户名称</TableHead><TableHead>金额(亿)</TableHead></TableRow></TableHeader>
              <TableBody>
                {byValue.map((item, i) => (
                  <TableRow key={item.账户代码}><TableCell>{i + 1}</TableCell><TableCell className="font-mono text-xs">{item.账户代码}</TableCell><TableCell>{item.账户名称}</TableCell><TableCell className="text-right font-mono">{item.交易金额}</TableCell></TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
        <Card>
          <CardHeader className="pb-2"><CardTitle className="text-sm">按笔数 Top10</CardTitle></CardHeader>
          <CardContent>
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={byCount.map(item => ({ name: item.账户名称.length > 12 ? item.账户名称.slice(0, 12) + '...' : item.账户名称, 笔数: item.交易笔数 }))} layout="vertical" margin={{ left: 80, right: 20 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" />
                <YAxis dataKey="name" type="category" width={75} tick={{ fontSize: 11 }} />
                <Tooltip />
                <Bar dataKey="笔数" fill="#10b981" radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
            <Table>
              <TableHeader><TableRow><TableHead>排名</TableHead><TableHead>账户代码</TableHead><TableHead>账户名称</TableHead><TableHead>笔数</TableHead></TableRow></TableHeader>
              <TableBody>
                {byCount.map((item, i) => (
                  <TableRow key={item.账户代码}><TableCell>{i + 1}</TableCell><TableCell className="font-mono text-xs">{item.账户代码}</TableCell><TableCell>{item.账户名称}</TableCell><TableCell className="text-right font-mono">{item.交易笔数}</TableCell></TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

function AccountDetailCard({ detail }: { detail: BondAccountDetail }) {
  return (
    <Card className="mb-4">
      <CardHeader><CardTitle className="text-sm">{detail.账户名称} <span className="text-xs text-muted-foreground font-mono">({detail.账户代码})</span></CardTitle></CardHeader>
      <CardContent className="space-y-4">
        <div className="grid grid-cols-2 gap-2 text-sm">
          <div><span className="text-muted-foreground">总笔数:</span> <strong>{detail.总笔数}</strong></div>
        </div>
        {detail.债券类别.length > 0 && (
          <div>
            <h4 className="text-xs font-semibold mb-2">债券类别分析</h4>
            <div className="grid grid-cols-1 gap-2">
              {detail.债券类别.map(cat => (
                <div key={cat.类别} className="bg-muted rounded p-2 text-sm">
                  <div className="font-semibold">{cat.类别}</div>
                  <div className="grid grid-cols-4 gap-2 text-xs mt-1">
                    <div>买入: {cat.买入金额}亿 ({cat.买入笔数}笔)</div>
                    <div>卖出: {cat.卖出金额}亿 ({cat.卖出笔数}笔)</div>
                    <div>净买卖: {cat.净买卖}亿</div>
                    {cat.平均期限 !== undefined && <div>平均期限: {cat.平均期限}年</div>}
                  </div>
                  {Object.keys(cat.品种分布).length > 0 && (
                    <div className="mt-1 text-xs text-muted-foreground">品种: {Object.entries(cat.品种分布).map(([k, v]) => `${k}(${v})`).join(', ')}</div>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}
        {detail.基金分析 && (
          <div className="bg-blue-50 rounded p-2 text-sm">
            <div className="font-semibold text-blue-700">基金分析</div>
            <div className="grid grid-cols-2 gap-2 text-xs mt-1">
              <div>买入: {detail.基金分析.买入金额}亿 ({detail.基金分析.买入笔数}笔)</div>
              <div>卖出: {detail.基金分析.卖出金额}亿 ({detail.基金分析.卖出笔数}笔)</div>
            </div>
            {detail.基金分析.涉及基金.length > 0 && <div className="mt-1 text-xs text-muted-foreground">涉及: {detail.基金分析.涉及基金.join(', ')}</div>}
          </div>
        )}
        {detail.转债分析 && (
          <div className="bg-amber-50 rounded p-2 text-sm">
            <div className="font-semibold text-amber-700">可转债分析</div>
            <div className="grid grid-cols-2 gap-2 text-xs mt-1">
              <div>买入: {detail.转债分析.买入金额}亿 ({detail.转债分析.买入笔数}笔)</div>
              <div>卖出: {detail.转债分析.卖出金额}亿 ({detail.转债分析.卖出笔数}笔)</div>
            </div>
          </div>
        )}
        {detail.期货分析 && (
          <div className="bg-purple-50 rounded p-2 text-sm">
            <div className="font-semibold text-purple-700">国债期货</div>
            <div className="text-xs mt-1">{detail.期货分析.笔数}笔, 总金额 {detail.期货分析.总金额}亿</div>
          </div>
        )}
        {Object.keys(detail.买卖方向).length > 0 && (
          <div className="text-xs text-muted-foreground">方向: {Object.entries(detail.买卖方向).map(([k, v]) => `${k}: ${v}亿`).join(', ')}</div>
        )}
      </CardContent>
    </Card>
  );
}

// ============================================================
// 持仓规模相关子组件
// ============================================================

function SizeLayerChart({ data }: { data: { 规模档位: string; 账户数: number; 回购占比: number; 债券占比: number; 基金占比: number; 转债占比: number; 期货占比: number; 长期债券占比: number }[] }) {
  return (
    <ResponsiveContainer width="100%" height={300}>
      <BarChart data={data}>
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey="规模档位" tick={{ fontSize: 11 }} />
        <YAxis />
        <Tooltip />
        <Legend />
        <Bar dataKey="回购占比" stackId="a" fill="#3b82f6" />
        <Bar dataKey="债券占比" stackId="a" fill="#10b981" />
        <Bar dataKey="基金占比" stackId="a" fill="#f59e0b" />
        <Bar dataKey="转债占比" stackId="a" fill="#ef4444" />
        <Bar dataKey="期货占比" stackId="a" fill="#8b5cf6" />
        <Bar dataKey="长期债券占比" stackId="a" fill="#06b6d4" />
      </BarChart>
    </ResponsiveContainer>
  );
}

function PositionTimeSeriesChart({ data, selectedAccount }: { data: PositionTimeSeries[]; selectedAccount?: string }) {
  const displayData = selectedAccount
    ? data.filter(d => d.账户代码 === selectedAccount)
    : data.slice(0, 5);

  if (displayData.length === 0) return <div className="text-muted-foreground text-sm">无时间序列数据</div>;

  const allDates = Array.from(new Set(displayData.flatMap(d => d.数据.map(p => p.日期)))).sort();

  const chartData = allDates.map(date => {
    const point: Record<string, number | string> = { 日期: date };
    for (const acct of displayData) {
      const record = acct.数据.find(p => p.日期 === date);
      point[`${acct.账户名称}_规模`] = record ? record.账户规模 : 0;
    }
    return point;
  });

  const colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#84cc16'];

  return (
    <ResponsiveContainer width="100%" height={350}>
      <LineChart data={chartData}>
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey="日期" tick={{ fontSize: 10 }} angle={-30} textAnchor="end" height={50} />
        <YAxis label={{ value: '规模(亿)', angle: -90, position: 'insideLeft', fontSize: 10 }} />
        <Tooltip />
        <Legend />
        {displayData.map((acct, i) => (
          <Line
            key={acct.账户代码}
            type="monotone"
            dataKey={`${acct.账户名称}_规模`}
            stroke={colors[i % colors.length]}
            strokeWidth={2}
            dot={false}
          />
        ))}
      </LineChart>
    </ResponsiveContainer>
  );
}

function PositionNavTimeSeriesChart({ data, selectedAccount }: { data: PositionTimeSeries[]; selectedAccount?: string }) {
  const displayData = selectedAccount
    ? data.filter(d => d.账户代码 === selectedAccount)
    : data.slice(0, 5).filter(d => d.数据.some(p => p.累计单位净值 !== undefined));

  if (displayData.length === 0) return <div className="text-muted-foreground text-sm">无净值时间序列数据</div>;

  const allDates = Array.from(new Set(displayData.flatMap(d => d.数据.filter(p => p.累计单位净值 !== undefined).map(p => p.日期)))).sort();

  const chartData = allDates.map(date => {
    const point: Record<string, number | string> = { 日期: date };
    for (const acct of displayData) {
      const record = acct.数据.find(p => p.日期 === date);
      point[`${acct.账户名称}_净值`] = record && record.累计单位净值 !== undefined ? record.累计单位净值 : 0;
    }
    return point;
  });

  const colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#84cc16'];

  return (
    <ResponsiveContainer width="100%" height={300}>
      <LineChart data={chartData}>
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey="日期" tick={{ fontSize: 10 }} angle={-30} textAnchor="end" height={50} />
        <YAxis domain={['auto', 'auto']} label={{ value: '单位净值', angle: -90, position: 'insideLeft', fontSize: 10 }} />
        <Tooltip />
        <Legend />
        {displayData.map((acct, i) => (
          <Line
            key={acct.账户代码}
            type="monotone"
            dataKey={`${acct.账户名称}_净值`}
            stroke={colors[i % colors.length]}
            strokeWidth={2}
            dot={false}
          />
        ))}
      </LineChart>
    </ResponsiveContainer>
  );
}


// ============================================================
// 主页面组件
// ============================================================

export default function Home() {
  const [_file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [searchCode, setSearchCode] = useState('');
  const [compareCodes, setCompareCodes] = useState<string[]>([]);
  const [showCompare, setShowCompare] = useState(false);
  const [selectedPositionAccount, setSelectedPositionAccount] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setFile(f);
    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const { hg, zq, position } = await parseExcel(f);
      const res = runAnalysis(hg, zq, position);
      setResult(res);
    } catch (err: any) {
      setError(err.message || '解析失败，请检查文件格式');
    } finally {
      setLoading(false);
    }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const f = e.dataTransfer.files[0];
    if (f && fileInputRef.current) {
      const dt = new DataTransfer();
      dt.items.add(f);
      fileInputRef.current.files = dt.files;
      handleFileChange({ target: fileInputRef.current } as any);
    }
  }, [handleFileChange]);

  const searchedAccount = useMemo(() => {
    if (!result || !searchCode) return null;
    return result.allAccountsDetail[searchCode] || null;
  }, [result, searchCode]);

  const compareAccounts = useMemo(() => {
    if (!result) return [];
    return compareCodes.map(c => result.allAccountsDetail[c]).filter(Boolean);
  }, [result, compareCodes]);

  const toggleCompare = useCallback((code: string) => {
    setCompareCodes(prev => {
      if (prev.includes(code)) return prev.filter(c => c !== code);
      if (prev.length >= 3) return prev;
      return [...prev, code];
    });
  }, []);

  if (!result) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 flex items-center justify-center p-4"
        onDragOver={(e) => e.preventDefault()} onDrop={handleDrop}>
        <div className="max-w-xl w-full">
          <div className="text-center mb-8">
            <h1 className="text-3xl font-bold text-slate-800 mb-2">交易数据分析平台</h1>
            <p className="text-slate-500">上传 Excel 文件，自动分析回购、债券、基金、持仓规模等多维数据</p>
          </div>
          <Card className="border-2 border-dashed border-slate-300 hover:border-blue-400 transition-colors cursor-pointer"
            onClick={() => fileInputRef.current?.click()}>
            <CardContent className="p-12 text-center">
              <Upload className="w-12 h-12 text-slate-400 mx-auto mb-4" />
              <p className="text-lg font-medium text-slate-700 mb-2">点击或拖拽上传 Excel 文件</p>
              <p className="text-sm text-slate-400">支持 .xlsx 格式，需包含「回购」「债券/其他」「持仓规模」三个 Sheet</p>
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFileChange} />
            </CardContent>
          </Card>
          {loading && (
            <div className="text-center mt-6">
              <div className="inline-block w-8 h-8 border-2 border-blue-500 border-t-transparent rounded-full animate-spin" />
              <p className="text-sm text-slate-500 mt-2">正在分析数据...</p>
            </div>
          )}
          {error && (
            <div className="mt-4 p-4 bg-red-50 border border-red-200 rounded-lg text-red-600 text-sm">
              <AlertTriangle className="w-4 h-4 inline mr-2" />{error}
            </div>
          )}
        </div>
      </div>
    );
  }


  return (
    <div className="min-h-screen bg-slate-50">
      {/* Header */}
      <div className="bg-white border-b sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <BarChart3 className="w-6 h-6 text-blue-600" />
            <h1 className="text-lg font-bold">交易数据分析平台</h1>
            <Badge variant="secondary" className="text-xs">
              {result.totalHgRecords + result.totalZqRecords} 笔交易 / {result.totalAccounts} 个账户
              {result.totalPositionRecords > 0 && ` / ${result.positionAccounts} 个有持仓数据`}
            </Badge>
          </div>
          <div className="flex items-center gap-2">
            <Input placeholder="搜索账户代码..." value={searchCode} onChange={e => setSearchCode(e.target.value)}
              className="w-48 text-sm h-8" />
            {searchCode && (
              <Button variant="ghost" size="sm" onClick={() => setSearchCode('')}><X className="w-4 h-4" /></Button>
            )}
            <Button variant="outline" size="sm" onClick={() => { setResult(null); setFile(null); setSearchCode(''); setCompareCodes([]); }}>
              <Upload className="w-4 h-4 mr-1" />重新上传
            </Button>
          </div>
        </div>
      </div>

      {/* Search Result */}
      {searchedAccount && (
        <div className="max-w-7xl mx-auto px-4 py-4">
          <div className="flex items-center justify-between mb-2">
            <h2 className="text-sm font-semibold flex items-center gap-2">
              <Search className="w-4 h-4" />搜索结果: {searchedAccount.账户名称}
            </h2>
            <Button variant="outline" size="sm" onClick={() => toggleCompare(searchedAccount.账户代码)}>
              {compareCodes.includes(searchedAccount.账户代码) ? '已加入对比' : '加入对比'}
            </Button>
          </div>
          <AccountDetailCard detail={searchedAccount} />
        </div>
      )}

      {/* Compare */}
      {compareCodes.length > 0 && (
        <div className="max-w-7xl mx-auto px-4 py-2">
          <div className="flex items-center gap-2 mb-2">
            <GitCompare className="w-4 h-4" />
            <span className="text-sm font-semibold">对比账户 ({compareCodes.length}/3):</span>
            {compareCodes.map(c => (
              <Badge key={c} variant="secondary" className="cursor-pointer" onClick={() => toggleCompare(c)}>
                {c} <X className="w-3 h-3 ml-1" />
              </Badge>
            ))}
            <Button variant="ghost" size="sm" onClick={() => setShowCompare(!showCompare)}>
              {showCompare ? '收起' : '展开'}
            </Button>
          </div>
          {showCompare && (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {compareAccounts.map(a => <AccountDetailCard key={a.账户代码} detail={a} />)}
            </div>
          )}
        </div>
      )}

      {/* Main Tabs */}
      <div className="max-w-7xl mx-auto px-4 py-4">
        <Tabs defaultValue="overview" className="space-y-4">
          <TabsList className="flex-wrap h-auto">
            <TabsTrigger value="overview"><Activity className="w-3 h-3 mr-1" />概览</TabsTrigger>
            <TabsTrigger value="hg"><TrendingUp className="w-3 h-3 mr-1" />回购业务</TabsTrigger>
            <TabsTrigger value="bond"><BarChart3 className="w-3 h-3 mr-1" />债券分析</TabsTrigger>
            <TabsTrigger value="fund"><Package className="w-3 h-3 mr-1" />基金/转债/期货</TabsTrigger>
            <TabsTrigger value="low"><AlertTriangle className="w-3 h-3 mr-1" />低活跃度</TabsTrigger>
            <TabsTrigger value="special"><FileText className="w-3 h-3 mr-1" />特殊品种</TabsTrigger>
            <TabsTrigger value="position"><Wallet className="w-3 h-3 mr-1" />持仓规模</TabsTrigger>
            <TabsTrigger value="viz"><Thermometer className="w-3 h-3 mr-1" />可视化</TabsTrigger>
          </TabsList>

          {/* Overview */}
          <TabsContent value="overview" className="space-y-4">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              <StatCard title="回购交易" value={`${result.totalHgRecords.toLocaleString()} 笔`} subtitle={`${result.hgAccounts} 个账户`} icon={TrendingUp} />
              <StatCard title="债券交易" value={`${result.totalZqRecords.toLocaleString()} 笔`} subtitle={`${result.zqAccounts} 个账户`} icon={BarChart3} color="green" />
              <StatCard title="总账户数" value={result.totalAccounts.toLocaleString()} icon={Activity} color="purple" />
              <StatCard title="异常交易" value={result.anomalies.length.toString()} subtitle="已标红提示" icon={AlertTriangle} color="red" />
            </div>
            {result.totalPositionRecords > 0 && (
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                <StatCard title="持仓账户" value={`${result.positionOverview.有规模数据的账户数} 个`} icon={Wallet} color="cyan" />
                <StatCard title="总持仓规模" value={`${result.positionOverview.总持仓规模} 亿`} icon={Scale} color="amber" />
                <StatCard title="平均规模" value={`${result.positionOverview.平均账户规模} 亿`} icon={Scale} color="amber" />
                <StatCard title="平均净值" value={result.positionOverview.平均单位净值} subtitle={`${result.positionOverview.有净值数据的账户数} 个有数据`} icon={TrendingUp} color="green" />
              </div>
            )}
            <Top10Section title="回购业务 Top10" byValue={result.hgTop10ByValue} byCount={result.hgTop10ByCount} color="#3b82f6" />
            <Top10Section title="长期债券(10年+) Top10" byValue={result.longTermTop10ByValue} byCount={result.longTermTop10ByCount} color="#10b981" />
            <Top10Section title="基金 Top10" byValue={result.fundTop10ByValue} byCount={result.fundTop10ByCount} color="#f59e0b" />
            <Top10Section title="可转债 Top10" byValue={result.zhaiquanTop10ByValue} byCount={result.zhaiquanTop10ByCount} color="#ef4444" />
            {result.totalPositionRecords > 0 && result.turnoverTop10.length > 0 && (
              <Card>
                <CardHeader><CardTitle className="text-sm">换手率 Top10（最活跃）</CardTitle></CardHeader>
                <CardContent>
                  <Table>
                    <TableHeader><TableRow><TableHead>排名</TableHead><TableHead>账户代码</TableHead><TableHead>账户名称</TableHead><TableHead>换手率</TableHead><TableHead>规模(亿)</TableHead><TableHead>交易金额(亿)</TableHead></TableRow></TableHeader>
                    <TableBody>
                      {result.turnoverTop10.map(item => (
                        <TableRow key={item.账户代码}>
                          <TableCell>{item.排名}</TableCell>
                          <TableCell className="font-mono text-xs">{item.账户代码}</TableCell>
                          <TableCell>{item.账户名称}</TableCell>
                          <TableCell className="text-right font-mono text-red-600">{item.换手率}x</TableCell>
                          <TableCell className="text-right font-mono">{item.账户规模}</TableCell>
                          <TableCell className="text-right font-mono">{item.交易总金额}</TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </CardContent>
              </Card>
            )}
          </TabsContent>

          {/* 回购 */}
          <TabsContent value="hg" className="space-y-4">
            <Top10Section title="回购业务 Top10（已过滤单笔超100亿）" byValue={result.hgTop10ByValue} byCount={result.hgTop10ByCount} color="#3b82f6" />
            <Top10Section title="协议回购 Top10" byValue={result.xyhgTop10ByValue} byCount={result.xyhgTop10ByCount} color="#8b5cf6" />
            <Card><CardHeader><CardTitle className="text-sm">协议回购账户画像（Top5）</CardTitle></CardHeader>
              <CardContent className="space-y-3">
                {result.xyhgAccounts.map(acct => (
                  <div key={acct.账户代码} className="bg-slate-50 rounded p-3">
                    <div className="font-semibold text-sm">{acct.账户名称} <span className="text-xs text-muted-foreground font-mono">({acct.账户代码})</span></div>
                    <div className="grid grid-cols-3 gap-2 text-xs mt-1">
                      <div>协议回购: {acct.协议回购金额}亿 ({acct.协议回购笔数}笔)</div>
                      <div>回购类型: {Object.entries(acct.全部回购类型).map(([k, v]) => `${k}(${v})`).join(', ')}</div>
                      <div>方向: {Object.entries(acct.交易方向分布).map(([k, v]) => `${k}(${v})`).join(', ')}</div>
                    </div>
                  </div>
                ))}
              </CardContent>
            </Card>
          </TabsContent>

          {/* 债券 */}
          <TabsContent value="bond" className="space-y-4">
            <Top10Section title="长期债券(10年+) Top10" byValue={result.longTermTop10ByValue} byCount={result.longTermTop10ByCount} color="#10b981" />
            <Card><CardHeader><CardTitle className="text-sm">长期债券账户深度分析（Top5）</CardTitle></CardHeader>
              <CardContent>
                <Accordion type="multiple" className="w-full">
                  {result.longTermAccounts.map(acct => (
                    <AccordionItem key={acct.账户代码} value={acct.账户代码}>
                      <AccordionTrigger className="text-sm">{acct.账户名称} ({acct.账户代码}) - {acct.总笔数}笔</AccordionTrigger>
                      <AccordionContent><AccountDetailCard detail={acct} /></AccordionContent>
                    </AccordionItem>
                  ))}
                </Accordion>
              </CardContent>
            </Card>
          </TabsContent>

          {/* 基金/转债/期货 */}
          <TabsContent value="fund" className="space-y-4">
            <Top10Section title="基金 Top10" byValue={result.fundTop10ByValue} byCount={result.fundTop10ByCount} color="#f59e0b" />
            <Top10Section title="可转债 Top10" byValue={result.zhaiquanTop10ByValue} byCount={result.zhaiquanTop10ByCount} color="#ef4444" />
            <Top10Section title="国债期货 Top10" byValue={result.qhTop10ByValue} byCount={result.qhTop10ByCount} color="#8b5cf6" />
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              <Card><CardHeader><CardTitle className="text-sm">基金账户分析（Top5）</CardTitle></CardHeader>
                <CardContent><Accordion type="multiple">{result.fundAccounts.map(acct => (
                  <AccordionItem key={acct.账户代码} value={acct.账户代码}>
                    <AccordionTrigger className="text-sm">{acct.账户名称}</AccordionTrigger>
                    <AccordionContent><AccountDetailCard detail={acct} /></AccordionContent>
                  </AccordionItem>))}</Accordion></CardContent>
              </Card>
              <Card><CardHeader><CardTitle className="text-sm">可转债账户分析（Top5）</CardTitle></CardHeader>
                <CardContent><Accordion type="multiple">{result.zhaiquanAccounts.map(acct => (
                  <AccordionItem key={acct.账户代码} value={acct.账户代码}>
                    <AccordionTrigger className="text-sm">{acct.账户名称}</AccordionTrigger>
                    <AccordionContent><AccountDetailCard detail={acct} /></AccordionContent>
                  </AccordionItem>))}</Accordion></CardContent>
              </Card>
            </div>
          </TabsContent>

          {/* 低活跃度 */}
          <TabsContent value="low" className="space-y-4">
            <Card><CardHeader><CardTitle className="text-sm">低活跃度账户（交易笔数 ≤ 2）</CardTitle></CardHeader>
              <CardContent>
                <p className="text-xs text-muted-foreground mb-2">共 {result.lowActivityAccounts.length} 个账户</p>
                <ScrollArea className="h-[500px]">
                  <Table>
                    <TableHeader><TableRow><TableHead>账户代码</TableHead><TableHead>账户名称</TableHead><TableHead>笔数</TableHead><TableHead>金额(亿)</TableHead><TableHead>持有品种</TableHead></TableRow></TableHeader>
                    <TableBody>
                      {result.lowActivityAccounts.map(acct => (
                        <TableRow key={acct.账户代码}>
                          <TableCell className="font-mono text-xs">{acct.账户代码}</TableCell>
                          <TableCell>{acct.账户名称}</TableCell>
                          <TableCell>{acct.总交易笔数}</TableCell>
                          <TableCell className="font-mono">{acct.总交易金额}</TableCell>
                          <TableCell className="text-xs">{acct.持有品种.join(', ')}</TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </ScrollArea>
              </CardContent>
            </Card>
            {result.zeroTradeAccounts.length > 0 && (
              <Card><CardHeader><CardTitle className="text-sm">零交易账户</CardTitle></CardHeader>
                <CardContent>
                  <Table>
                    <TableHeader><TableRow><TableHead>账户代码</TableHead><TableHead>账户名称</TableHead></TableRow></TableHeader>
                    <TableBody>{result.zeroTradeAccounts.map(a => (
                      <TableRow key={a.账户代码}><TableCell className="font-mono text-xs">{a.账户代码}</TableCell><TableCell>{a.账户名称}</TableCell></TableRow>
                    ))}</TableBody>
                  </Table>
                </CardContent>
              </Card>
            )}
          </TabsContent>

          {/* 特殊品种 */}
          <TabsContent value="special" className="space-y-4">
            <Top10Section title="ABS Top10" byValue={result.absTop10ByValue} byCount={result.absTop10ByCount} color="#06b6d4" />
            <Top10Section title="REITs Top10" byValue={result.reitsTop10ByValue} byCount={result.reitsTop10ByCount} color="#ec4899" />
            <Card><CardHeader><CardTitle className="text-sm">特殊业务品种</CardTitle></CardHeader>
              <CardContent>
                <Table>
                  <TableHeader><TableRow><TableHead>排名</TableHead><TableHead>账户代码</TableHead><TableHead>账户名称</TableHead><TableHead>业务品种</TableHead><TableHead>金额(亿)</TableHead><TableHead>笔数</TableHead></TableRow></TableHeader>
                  <TableBody>{result.specialTop10.map(item => (
                    <TableRow key={`${item.账户代码}-${item.业务品种}`}>
                      <TableCell>{item.排名}</TableCell><TableCell className="font-mono text-xs">{item.账户代码}</TableCell>
                      <TableCell>{item.账户名称}</TableCell><TableCell><Badge variant="outline">{item.业务品种}</Badge></TableCell>
                      <TableCell className="font-mono">{item.金额}</TableCell><TableCell>{item.笔数}</TableCell>
                    </TableRow>
                  ))}</TableBody>
                </Table>
              </CardContent>
            </Card>
          </TabsContent>

          {/* 持仓规模（新增） */}
            <TabsContent value="position" className="space-y-4">
              {/* 概览卡片 */}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                <StatCard title="持仓账户" value={`${result.positionOverview.有规模数据的账户数} 个`} icon={Wallet} color="cyan" />
                <StatCard title="总持仓规模" value={`${result.positionOverview.总持仓规模} 亿`} icon={Scale} color="amber" />
                <StatCard title="平均规模" value={`${result.positionOverview.平均账户规模} 亿`} subtitle={`中位数 ${result.positionOverview.中位数规模} 亿`} icon={Scale} color="amber" />
                <StatCard title="平均净值" value={result.positionOverview.平均单位净值} subtitle={`${result.positionOverview.有净值数据的账户数} 个有数据`} icon={TrendingUp} color="green" />
              </div>

              {/* 规模档位统计 */}
              <Card>
                <CardHeader><CardTitle className="text-sm">规模档位交易特征</CardTitle></CardHeader>
                <CardContent>
                  <SizeLayerChart data={result.sizeLayerStats} />
                  <Table className="mt-4">
                    <TableHeader>
                      <TableRow>
                        <TableHead>规模档位</TableHead><TableHead>账户数</TableHead>
                        <TableHead>总交易(亿)</TableHead><TableHead>总笔数</TableHead>
                        <TableHead>平均交易(亿)</TableHead><TableHead>回购%</TableHead>
                        <TableHead>债券%</TableHead><TableHead>基金%</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {result.sizeLayerStats.map(layer => (
                        <TableRow key={layer.规模档位}>
                          <TableCell><Badge variant="outline">{layer.规模档位}</Badge></TableCell>
                          <TableCell>{layer.账户数}</TableCell>
                          <TableCell className="font-mono">{layer.总交易金额}</TableCell>
                          <TableCell>{layer.总交易笔数}</TableCell>
                          <TableCell className="font-mono">{layer.平均交易金额}</TableCell>
                          <TableCell>{layer.回购占比}%</TableCell>
                          <TableCell>{layer.债券占比}%</TableCell>
                          <TableCell>{layer.基金占比}%</TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </CardContent>
              </Card>

              {/* 换手率 */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                <Card>
                  <CardHeader><CardTitle className="text-sm">换手率 Top10（最活跃）</CardTitle></CardHeader>
                  <CardContent>
                    <Table>
                      <TableHeader><TableRow><TableHead>排名</TableHead><TableHead>账户名称</TableHead><TableHead>换手率</TableHead><TableHead>规模(亿)</TableHead></TableRow></TableHeader>
                      <TableBody>
                        {result.turnoverTop10.map(item => (
                          <TableRow key={item.账户代码}>
                            <TableCell>{item.排名}</TableCell>
                            <TableCell>{item.账户名称}</TableCell>
                            <TableCell className="text-right font-mono text-red-600">{item.换手率}x</TableCell>
                            <TableCell className="text-right font-mono">{item.账户规模}</TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </CardContent>
                </Card>
                <Card>
                  <CardHeader><CardTitle className="text-sm">换手率倒数 Top10（最"躺平"）</CardTitle></CardHeader>
                  <CardContent>
                    <Table>
                      <TableHeader><TableRow><TableHead>排名</TableHead><TableHead>账户名称</TableHead><TableHead>换手率</TableHead><TableHead>规模(亿)</TableHead></TableRow></TableHeader>
                      <TableBody>
                        {result.turnoverBottom10.map(item => (
                          <TableRow key={item.账户代码}>
                            <TableCell>{item.排名}</TableCell>
                            <TableCell>{item.账户名称}</TableCell>
                            <TableCell className="text-right font-mono text-green-600">{item.换手率}x</TableCell>
                            <TableCell className="text-right font-mono">{item.账户规模}</TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </CardContent>
                </Card>
              </div>

              {/* 规模与品种关联 */}
              <Card>
                <CardHeader><CardTitle className="text-sm">规模档位与交易品种关联</CardTitle></CardHeader>
                <CardContent>
                  <ResponsiveContainer width="100%" height={280}>
                    <ComposedChart data={result.sizeBizRelation}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="规模档位" tick={{ fontSize: 11 }} />
                      <YAxis yAxisId="left" />
                      <YAxis yAxisId="right" orientation="right" />
                      <Tooltip />
                      <Legend />
                      <Bar yAxisId="left" dataKey="平均回购金额" fill="#3b82f6" />
                      <Bar yAxisId="left" dataKey="平均债券买入" fill="#10b981" />
                      <Bar yAxisId="left" dataKey="平均基金金额" fill="#f59e0b" />
                      <Line yAxisId="right" type="monotone" dataKey="账户数" stroke="#ef4444" strokeWidth={2} />
                    </ComposedChart>
                  </ResponsiveContainer>
                </CardContent>
              </Card>

              {/* 净值与交易行为 */}
              <Card>
                <CardHeader><CardTitle className="text-sm">净值与交易行为关联（按净值降序）</CardTitle></CardHeader>
                <CardContent>
                  <ScrollArea className="h-[400px]">
                    <Table>
                      <TableHeader>
                        <TableRow>
                          <TableHead>账户代码</TableHead><TableHead>账户名称</TableHead>
                          <TableHead>最新净值</TableHead><TableHead>最新规模(亿)</TableHead>
                          <TableHead>交易金额(亿)</TableHead><TableHead>笔数</TableHead>
                          <TableHead>买入(亿)</TableHead><TableHead>卖出(亿)</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {result.navTradeRelation.slice(0, 50).map(item => (
                          <TableRow key={item.账户代码}>
                            <TableCell className="font-mono text-xs">{item.账户代码}</TableCell>
                            <TableCell>{item.账户名称}</TableCell>
                            <TableCell className={`font-mono ${item.最新净值 && item.最新净值 > 1 ? 'text-green-600' : item.最新净值 && item.最新净值 < 1 ? 'text-red-600' : ''}`}>
                              {item.最新净值 ?? '-'}
                            </TableCell>
                            <TableCell className="font-mono">{item.最新规模}</TableCell>
                            <TableCell className="font-mono">{item.交易总金额}</TableCell>
                            <TableCell>{item.交易笔数}</TableCell>
                            <TableCell className="font-mono text-blue-600">{item.买入金额}</TableCell>
                            <TableCell className="font-mono text-red-600">{item.卖出金额}</TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </ScrollArea>
                </CardContent>
              </Card>

              {/* 规模时间序列 */}
              <Card>
                <CardHeader className="flex flex-row items-center justify-between">
                  <CardTitle className="text-sm">账户规模时间序列（规模Top20）</CardTitle>
                  <Select value={selectedPositionAccount} onValueChange={setSelectedPositionAccount}>
                    <SelectTrigger className="w-64 text-xs h-8">
                      <SelectValue placeholder="选择账户查看详情" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="">全部（Top5）</SelectItem>
                      {result.positionTimeSeries.map(ts => (
                        <SelectItem key={ts.账户代码} value={ts.账户代码}>{ts.账户名称} ({ts.账户代码})</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </CardHeader>
                <CardContent>
                  <PositionTimeSeriesChart data={result.positionTimeSeries} selectedAccount={selectedPositionAccount || undefined} />
                </CardContent>
              </Card>

              {/* 净值时间序列 */}
              <Card>
                <CardHeader><CardTitle className="text-sm">单位净值时间序列</CardTitle></CardHeader>
                <CardContent>
                  <PositionNavTimeSeriesChart data={result.positionTimeSeries} selectedAccount={selectedPositionAccount || undefined} />
                </CardContent>
              </Card>

              {/* 规模变动 */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                <Card>
                  <CardHeader><CardTitle className="text-sm">规模增长 Top10</CardTitle></CardHeader>
                  <CardContent>
                    <Table>
                      <TableHeader><TableRow><TableHead>账户</TableHead><TableHead>期初(亿)</TableHead><TableHead>期末(亿)</TableHead><TableHead>变动</TableHead></TableRow></TableHeader>
                      <TableBody>
                        {result.sizeChangeTop10.map(item => (
                          <TableRow key={item.账户代码}>
                            <TableCell>{item.账户名称}</TableCell>
                            <TableCell className="font-mono">{item.期初规模}</TableCell>
                            <TableCell className="font-mono">{item.期末规模}</TableCell>
                            <TableCell className={`font-mono ${item.变动比例 > 0 ? 'text-green-600' : 'text-red-600'}`}>
                              {item.变动比例 > 0 ? '+' : ''}{item.变动比例}% ({item.变动金额}亿)
                            </TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </CardContent>
                </Card>
                <Card>
                  <CardHeader><CardTitle className="text-sm">规模缩水 Top10</CardTitle></CardHeader>
                  <CardContent>
                    <Table>
                      <TableHeader><TableRow><TableHead>账户</TableHead><TableHead>期初(亿)</TableHead><TableHead>期末(亿)</TableHead><TableHead>变动</TableHead></TableRow></TableHeader>
                      <TableBody>
                        {result.sizeShrinkTop10.map(item => (
                          <TableRow key={item.账户代码}>
                            <TableCell>{item.账户名称}</TableCell>
                            <TableCell className="font-mono">{item.期初规模}</TableCell>
                            <TableCell className="font-mono">{item.期末规模}</TableCell>
                            <TableCell className="font-mono text-red-600">
                              {item.变动比例}% ({item.变动金额}亿)
                            </TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </CardContent>
                </Card>
              </div>

              {/* 持仓异常 */}
              {result.positionAnomalies.length > 0 && (
                <Card className="border-red-200">
                  <CardHeader className="flex items-center gap-2">
                    <AlertTriangle className="w-5 h-5 text-red-500" />
                    <CardTitle className="text-sm">持仓异常账户 ({result.positionAnomalies.length} 个)</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <Table>
                      <TableHeader><TableRow><TableHead>类型</TableHead><TableHead>账户</TableHead><TableHead>规模(亿)</TableHead><TableHead>换手率</TableHead><TableHead>描述</TableHead></TableRow></TableHeader>
                      <TableBody>
                        {result.positionAnomalies.map((item, i) => (
                          <TableRow key={i} className={item.类型 === '净值为负' ? 'bg-red-50' : item.类型 === '规模大交易少' ? 'bg-amber-50' : ''}>
                            <TableCell><Badge variant={item.类型 === '净值为负' ? 'destructive' : 'outline'}>{item.类型}</Badge></TableCell>
                            <TableCell>{item.账户名称}</TableCell>
                            <TableCell className="font-mono">{item.账户规模}</TableCell>
                            <TableCell className="font-mono">{item.换手率}x</TableCell>
                            <TableCell className="text-xs">{item.描述}</TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </CardContent>
                </Card>
              )}

              {/* 合并账户列表 */}
              <Card>
                <CardHeader><CardTitle className="text-sm">完整账户列表（含规模数据）</CardTitle></CardHeader>
                <CardContent>
                  <ScrollArea className="h-[500px]">
                    <Table>
                      <TableHeader>
                        <TableRow>
                          <TableHead>账户代码</TableHead><TableHead>账户名称</TableHead>
                          <TableHead>规模档位</TableHead><TableHead>最新规模(亿)</TableHead>
                          <TableHead>净值</TableHead><TableHead>换手率</TableHead>
                          <TableHead>交易金额(亿)</TableHead><TableHead>笔数</TableHead>
                          <TableHead>主要品种</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {result.mergedAccountList.map(item => (
                          <TableRow key={item.账户代码}>
                            <TableCell className="font-mono text-xs">{item.账户代码}</TableCell>
                            <TableCell>{item.账户名称}</TableCell>
                            <TableCell>{item.规模档位 ? <Badge variant="outline">{item.规模档位}</Badge> : '-'}</TableCell>
                            <TableCell className="font-mono">{item.最新规模 ?? '-'}</TableCell>
                            <TableCell className={`font-mono ${item.最新净值 && item.最新净值 > 1 ? 'text-green-600' : item.最新净值 && item.最新净值 < 1 ? 'text-red-600' : ''}`}>
                              {item.最新净值 ?? '-'}
                            </TableCell>
                            <TableCell className="font-mono">{item.换手率 !== undefined ? `${item.换手率}x` : '-'}</TableCell>
                            <TableCell className="font-mono">{item.总交易金额}</TableCell>
                            <TableCell>{item.总交易笔数}</TableCell>
                            <TableCell><Badge variant="secondary">{item.主要品种}</Badge></TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </ScrollArea>
                </CardContent>
              </Card>
            </TabsContent>

          {/* 可视化 */}
          <TabsContent value="viz" className="space-y-4">
            <Card>
              <CardHeader><CardTitle className="text-sm">月度交易趋势</CardTitle></CardHeader>
              <CardContent>
                <ResponsiveContainer width="100%" height={300}>
                  <AreaChart data={result.monthlyTrend}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="月份" />
                    <YAxis />
                    <Tooltip />
                    <Legend />
                    <Area type="monotone" dataKey="回购业务" stackId="1" stroke="#3b82f6" fill="#3b82f6" fillOpacity={0.6} />
                    <Area type="monotone" dataKey="债券业务" stackId="1" stroke="#10b981" fill="#10b981" fillOpacity={0.6} />
                    <Area type="monotone" dataKey="基金" stackId="1" stroke="#f59e0b" fill="#f59e0b" fillOpacity={0.6} />
                    <Area type="monotone" dataKey="可转债" stackId="1" stroke="#ef4444" fill="#ef4444" fillOpacity={0.6} />
                    <Area type="monotone" dataKey="国债期货" stackId="1" stroke="#8b5cf6" fill="#8b5cf6" fillOpacity={0.6} />
                  </AreaChart>
                </ResponsiveContainer>
              </CardContent>
            </Card>

            <Card>
              <CardHeader><CardTitle className="text-sm">账户交易热力图（Top30 账户 × 月份）</CardTitle></CardHeader>
              <CardContent>
                <ScrollArea className="w-full overflow-auto">
                  <div className="min-w-[800px]">
                    <Table>
                      <TableHeader>
                        <TableRow>
                          <TableHead className="sticky left-0 bg-white">账户</TableHead>
                          {result.heatmapMonths.map(m => <TableHead key={m} className="text-xs">{m}</TableHead>)}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {result.heatmapData.map(row => (
                          <TableRow key={row.账户代码}>
                            <TableCell className="sticky left-0 bg-white text-xs font-mono">{row.账户名称}</TableCell>
                            {result.heatmapMonths.map(m => {
                              const v = row[m] as number;
                              const intensity = Math.min(v / 10, 1);
                              return (
                                <TableCell key={m} className="text-xs font-mono p-1" style={{ backgroundColor: `rgba(59, 130, 246, ${intensity * 0.3})` }}>
                                  {v > 0 ? v.toFixed(2) : '-'}
                                </TableCell>
                              );
                            })}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                </ScrollArea>
              </CardContent>
            </Card>

            <Card>
              <CardHeader><CardTitle className="text-sm">账户画像雷达图（Top15）</CardTitle></CardHeader>
              <CardContent>
                <ResponsiveContainer width="100%" height={400}>
                  <RadarChart data={result.radarAccounts}>
                    <PolarGrid />
                    <PolarAngleAxis dataKey="账户名称" tick={{ fontSize: 10 }} />
                    <PolarRadiusAxis angle={30} domain={[0, 100]} />
                    {result.radarAccounts.slice(0, 5).map((acct, i) => (
                      <Radar key={acct.账户代码} name={acct.账户名称} dataKey={(data: any) => data.账户代码 === acct.账户代码 ? data.回购活跃度 : 0}
                        stroke={['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'][i]} fill={['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'][i]} fillOpacity={0.1} />
                    ))}
                    <Legend />
                    <Tooltip />
                  </RadarChart>
                </ResponsiveContainer>
              </CardContent>
            </Card>

            <Card>
              <CardHeader><CardTitle className="text-sm">债券期限结构</CardTitle></CardHeader>
              <CardContent>
                <ResponsiveContainer width="100%" height={300}>
                  <ComposedChart data={result.termStructure}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="期限段" />
                    <YAxis />
                    <Tooltip />
                    <Legend />
                    <Bar dataKey="买入金额" fill="#3b82f6" />
                    <Bar dataKey="卖出金额" fill="#ef4444" />
                    <Line type="monotone" dataKey="净买卖" stroke="#10b981" strokeWidth={2} />
                  </ComposedChart>
                </ResponsiveContainer>
              </CardContent>
            </Card>

            {result.anomalies.length > 0 && (
              <Card className="border-red-200">
                <CardHeader className="flex items-center gap-2">
                  <AlertTriangle className="w-5 h-5 text-red-500" />
                  <CardTitle className="text-sm">交易异常记录 ({result.anomalies.length} 条)</CardTitle>
                </CardHeader>
                <CardContent>
                  <ScrollArea className="h-[300px]">
                    <Table>
                      <TableHeader><TableRow><TableHead>类型</TableHead><TableHead>账户</TableHead><TableHead>品种</TableHead><TableHead>金额(亿)</TableHead><TableHead>日期</TableHead></TableRow></TableHeader>
                      <TableBody>
                        {result.anomalies.map((a, i) => (
                          <TableRow key={i} className={a.类型 === '超大额' ? 'bg-red-50' : 'bg-amber-50'}>
                            <TableCell><Badge variant={a.类型 === '超大额' ? 'destructive' : 'outline'}>{a.类型}</Badge></TableCell>
                            <TableCell>{a.账户名称}</TableCell>
                            <TableCell>{a.业务品种}</TableCell>
                            <TableCell className="font-mono">{a.金额}</TableCell>
                            <TableCell>{a.日期}</TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </ScrollArea>
                </CardContent>
              </Card>
            )}
          </TabsContent>
        </Tabs>
      </div>
    </div>
  );
}
