import { useState, useCallback, useRef, useMemo, useEffect } from 'react';
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
import { Dialog, DialogContent, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { ScrollArea } from '@/components/ui/scroll-area';
import { parseExcel, runAnalysis, type AnalysisResult, type BondAccountDetail } from '@/lib/analyzer';
import {
  Upload, TrendingUp, Activity, BarChart3, AlertTriangle, Package, FileText,
  Search, X, Zap, GitCompare, Thermometer, RadarIcon, Clock,
} from 'lucide-react';

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
      <div className="h-72">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={chartData} margin={{ top: 10, right: 10, left: 0, bottom: 30 }}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="name" angle={-30} textAnchor="end" height={60} fontSize={11} />
            <YAxis fontSize={11} />
            <Tooltip formatter={(v: number) => v.toFixed(4)} />
            <Legend />
            <Bar dataKey="金额" fill={color} name="交易金额(亿元)" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Card>
          <CardHeader className="pb-2"><CardTitle className="text-sm font-semibold">金额 Top10</CardTitle></CardHeader>
          <CardContent>
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className="w-12">排名</TableHead>
                  <TableHead className="text-xs">账户代码</TableHead>
                  <TableHead>账户名称</TableHead>
                  <TableHead className="text-right">金额(亿元)</TableHead>
                  <TableHead className="text-right">笔数</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {byValue.map((item, i) => (
                  <TableRow key={item.账户代码}>
                    <TableCell><Badge variant={i < 3 ? 'default' : 'secondary'}>{i + 1}</Badge></TableCell>
                    <TableCell className="text-xs font-mono text-muted-foreground">{item.账户代码}</TableCell>
                    <TableCell className="font-medium text-xs">{item.账户名称}</TableCell>
                    <TableCell className="text-right font-mono text-xs">{item.交易金额.toFixed(4)}</TableCell>
                    <TableCell className="text-right text-xs">{item.交易笔数}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
        <Card>
          <CardHeader className="pb-2"><CardTitle className="text-sm font-semibold">笔数 Top10</CardTitle></CardHeader>
          <CardContent>
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className="w-12">排名</TableHead>
                  <TableHead className="text-xs">账户代码</TableHead>
                  <TableHead>账户名称</TableHead>
                  <TableHead className="text-right">笔数</TableHead>
                  <TableHead className="text-right">金额(亿元)</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {byCount.map((item, i) => (
                  <TableRow key={item.账户代码}>
                    <TableCell><Badge variant={i < 3 ? 'default' : 'secondary'}>{i + 1}</Badge></TableCell>
                    <TableCell className="text-xs font-mono text-muted-foreground">{item.账户代码}</TableCell>
                    <TableCell className="font-medium text-xs">{item.账户名称}</TableCell>
                    <TableCell className="text-right text-xs">{item.交易笔数}</TableCell>
                    <TableCell className="text-right font-mono text-xs">{item.交易金额.toFixed(4)}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

function AccountDetailCard({ account, index }: { account: BondAccountDetail; index: number }) {
  return (
    <AccordionItem value={`acct-${index}`}>
      <AccordionTrigger className="text-sm hover:no-underline py-3">
        <div className="flex items-center gap-3 text-left w-full pr-4">
          <Badge variant="outline">{index + 1}</Badge>
          <span className="font-semibold">{account.账户名称}</span>
          <span className="text-xs text-muted-foreground font-mono">{account.账户代码}</span>
          <Badge variant="secondary" className="ml-auto text-xs">{account.总笔数} 笔</Badge>
        </div>
      </AccordionTrigger>
      <AccordionContent>
        <div className="space-y-4 pl-4 pb-2">
          {account.债券类别.length > 0 && (
            <div>
              <h4 className="text-sm font-semibold mb-2 flex items-center gap-2"><BarChart3 className="w-4 h-4" /> 债券类别分析</h4>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {account.债券类别.map((cat) => (
                  <Card key={cat.类别} className="bg-muted/30">
                    <CardContent className="pt-4 pb-3">
                      <div className="flex justify-between items-center mb-2">
                        <span className="font-semibold text-sm">{cat.类别}</span>
                        {cat.平均期限 !== undefined && <Badge variant="outline" className="text-xs">平均期限 {cat.平均期限} 年</Badge>}
                      </div>
                      <div className="grid grid-cols-3 gap-2 text-xs mb-2">
                        <div><div className="text-muted-foreground">买入</div><div className="font-mono font-medium text-green-600">{cat.买入金额.toFixed(4)}</div></div>
                        <div><div className="text-muted-foreground">卖出</div><div className="font-mono font-medium text-red-500">{cat.卖出金额.toFixed(4)}</div></div>
                        <div><div className="text-muted-foreground">净买卖</div><div className={`font-mono font-medium ${cat.净买卖 >= 0 ? 'text-green-600' : 'text-red-500'}`}>{cat.净买卖.toFixed(4)}</div></div>
                      </div>
                      {Object.keys(cat.品种分布).length > 0 && (
                        <div className="flex flex-wrap gap-1">
                          {Object.entries(cat.品种分布).map(([k, v]) => <Badge key={k} variant="secondary" className="text-xs">{k}: {v}</Badge>)}
                        </div>
                      )}
                    </CardContent>
                  </Card>
                ))}
              </div>
            </div>
          )}
          {account.基金分析 && (
            <div>
              <h4 className="text-sm font-semibold mb-2 flex items-center gap-2"><TrendingUp className="w-4 h-4" /> 基金交易分析</h4>
              <Card className="bg-muted/30">
                <CardContent className="pt-4 pb-3">
                  <div className="grid grid-cols-2 gap-4 text-xs mb-2">
                    <div><span className="text-muted-foreground">买入:</span> <span className="font-mono text-green-600">{account.基金分析.买入金额.toFixed(4)} 亿元 ({account.基金分析.买入笔数} 笔)</span></div>
                    <div><span className="text-muted-foreground">卖出:</span> <span className="font-mono text-red-500">{account.基金分析.卖出金额.toFixed(4)} 亿元 ({account.基金分析.卖出笔数} 笔)</span></div>
                  </div>
                  <div className="text-xs text-muted-foreground">涉及基金: {account.基金分析.涉及基金.join(', ')}</div>
                </CardContent>
              </Card>
            </div>
          )}
          {account.转债分析 && (
            <div>
              <h4 className="text-sm font-semibold mb-2 flex items-center gap-2"><Activity className="w-4 h-4" /> 可转债交易分析</h4>
              <Card className="bg-muted/30">
                <CardContent className="pt-4 pb-3">
                  <div className="grid grid-cols-2 gap-4 text-xs mb-2">
                    <div><span className="text-muted-foreground">买入:</span> <span className="font-mono text-green-600">{account.转债分析.买入金额.toFixed(4)} 亿元 ({account.转债分析.买入笔数} 笔)</span></div>
                    <div><span className="text-muted-foreground">卖出:</span> <span className="font-mono text-red-500">{account.转债分析.卖出金额.toFixed(4)} 亿元 ({account.转债分析.卖出笔数} 笔)</span></div>
                  </div>
                  <div className="text-xs text-muted-foreground">涉及转债: {account.转债分析.涉及转债.join(', ')}</div>
                </CardContent>
              </Card>
            </div>
          )}
          {account.期货分析 && (
            <div>
              <h4 className="text-sm font-semibold mb-2 flex items-center gap-2"><AlertTriangle className="w-4 h-4" /> 国债期货交易分析</h4>
              <Card className="bg-muted/30">
                <CardContent className="pt-4 pb-3">
                  <div className="grid grid-cols-2 gap-4 text-xs">
                    <div><span className="text-muted-foreground">交易笔数:</span> <span className="font-mono font-medium">{account.期货分析.笔数} 笔</span></div>
                    <div><span className="text-muted-foreground">交易总金额:</span> <span className="font-mono font-medium">{account.期货分析.总金额.toFixed(4)} 亿元</span></div>
                  </div>
                  <div className="text-xs text-muted-foreground mt-2">涉及品种: {account.期货分析.涉及品种.join(', ')}</div>
                </CardContent>
              </Card>
            </div>
          )}
          {Object.keys(account.买卖方向).length > 0 && (
            <div>
              <h4 className="text-sm font-semibold mb-2">全账户买卖方向总览</h4>
              <div className="flex flex-wrap gap-2">
                {Object.entries(account.买卖方向).map(([dir, amount]) => (
                  <Badge key={dir} variant="outline" className="text-xs">{dir}: {amount.toFixed(4)} 亿元</Badge>
                ))}
              </div>
            </div>
          )}
        </div>
      </AccordionContent>
    </AccordionItem>
  );
}

function HeatmapChart({ data, months }: { data: import('@/lib/analyzer').HeatmapData[]; months: string[] }) {
  if (data.length === 0 || months.length === 0) return null;
  let maxVal = 0;
  for (const row of data) {
    for (const m of months) {
      const v = typeof row[m] === 'number' ? (row[m] as number) : 0;
      if (v > maxVal) maxVal = v;
    }
  }
  if (maxVal === 0) maxVal = 1;
  const getColor = (v: number) => {
    const ratio = Math.min(v / maxVal, 1);
    const r = Math.round(239 - ratio * 180);
    const g = Math.round(246 - ratio * 120);
    const b = Math.round(255 - ratio * 40);
    return `rgb(${r},${g},${b})`;
  };
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-xs border-collapse">
        <thead>
          <tr>
            <th className="text-left p-2 border-b sticky left-0 bg-background z-10 min-w-[160px]">账户</th>
            {months.map(m => <th key={m} className="text-center p-2 border-b min-w-[70px]">{m}</th>)}
          </tr>
        </thead>
        <tbody>
          {data.map((row) => (
            <tr key={row.账户代码}>
              <td className="p-2 border-b sticky left-0 bg-background z-10">
                <div className="font-medium truncate max-w-[120px]" title={row.账户名称 as string}>{row.账户名称 as string}</div>
                <div className="text-muted-foreground font-mono text-[10px]">{row.账户代码}</div>
              </td>
              {months.map(m => {
                const v = typeof row[m] === 'number' ? (row[m] as number) : 0;
                return (
                  <td key={m} className="p-1 border-b text-center" style={{ backgroundColor: getColor(v) }}>
                    <span className={`font-mono ${v > maxVal * 0.5 ? 'text-white' : 'text-foreground'}`}>{v > 0 ? v.toFixed(2) : ''}</span>
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
      <div className="flex items-center gap-2 mt-2 text-xs text-muted-foreground">
        <span>低</span>
        <div className="h-3 w-32 rounded" style={{ background: 'linear-gradient(to right, rgb(239,246,255), rgb(59,130,246))' }} />
        <span>高 ({maxVal.toFixed(2)} 亿元)</span>
      </div>
    </div>
  );
}

function AccountSearchDialog({ open, onClose, accountCode, allAccountsDetail }: {
  open: boolean; onClose: () => void; accountCode: string;
  allAccountsDetail: Record<string, BondAccountDetail>;
}) {
  const account = allAccountsDetail[accountCode];
  if (!account) return null;
  return (
    <Dialog open={open} onOpenChange={(v) => !v && onClose()}>
      <DialogContent className="max-w-3xl max-h-[80vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2">
            <span>{account.账户名称}</span>
            <Badge variant="outline" className="font-mono text-xs">{account.账户代码}</Badge>
          </DialogTitle>
        </DialogHeader>
        <div className="space-y-4 mt-2">
          <div className="grid grid-cols-3 gap-3">
            <Card className="bg-muted/30">
              <CardContent className="pt-4 pb-3 text-center">
                <div className="text-2xl font-bold text-primary">{account.总笔数}</div>
                <div className="text-xs text-muted-foreground">总交易笔数</div>
              </CardContent>
            </Card>
            <Card className="bg-muted/30">
              <CardContent className="pt-4 pb-3 text-center">
                <div className="text-2xl font-bold text-primary">{Object.values(account.买卖方向).reduce((a, b) => a + b, 0).toFixed(4)}</div>
                <div className="text-xs text-muted-foreground">总交易金额(亿元)</div>
              </CardContent>
            </Card>
            <Card className="bg-muted/30">
              <CardContent className="pt-4 pb-3 text-center">
                <div className="text-2xl font-bold text-primary">{account.债券类别.length}</div>
                <div className="text-xs text-muted-foreground">债券类别数</div>
              </CardContent>
            </Card>
          </div>
          {account.债券类别.length > 0 && (
            <div>
              <h4 className="text-sm font-semibold mb-2">债券类别分析</h4>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {account.债券类别.map((cat) => (
                  <Card key={cat.类别} className="bg-muted/30">
                    <CardContent className="pt-4 pb-3">
                      <div className="flex justify-between items-center mb-2">
                        <span className="font-semibold text-sm">{cat.类别}</span>
                        {cat.平均期限 !== undefined && <Badge variant="outline" className="text-xs">平均期限 {cat.平均期限} 年</Badge>}
                      </div>
                      <div className="grid grid-cols-3 gap-2 text-xs mb-2">
                        <div><div className="text-muted-foreground">买入</div><div className="font-mono font-medium text-green-600">{cat.买入金额.toFixed(4)}</div></div>
                        <div><div className="text-muted-foreground">卖出</div><div className="font-mono font-medium text-red-500">{cat.卖出金额.toFixed(4)}</div></div>
                        <div><div className="text-muted-foreground">净买卖</div><div className={`font-mono font-medium ${cat.净买卖 >= 0 ? 'text-green-600' : 'text-red-500'}`}>{cat.净买卖.toFixed(4)}</div></div>
                      </div>
                    </CardContent>
                  </Card>
                ))}
              </div>
            </div>
          )}
          {account.基金分析 && (
            <div>
              <h4 className="text-sm font-semibold mb-2">基金交易分析</h4>
              <Card className="bg-muted/30">
                <CardContent className="pt-4 pb-3">
                  <div className="grid grid-cols-2 gap-4 text-xs">
                    <div>买入: <span className="font-mono text-green-600">{account.基金分析.买入金额.toFixed(4)} 亿元 ({account.基金分析.买入笔数}笔)</span></div>
                    <div>卖出: <span className="font-mono text-red-500">{account.基金分析.卖出金额.toFixed(4)} 亿元 ({account.基金分析.卖出笔数}笔)</span></div>
                  </div>
                </CardContent>
              </Card>
            </div>
          )}
          {account.转债分析 && (
            <div>
              <h4 className="text-sm font-semibold mb-2">可转债交易分析</h4>
              <Card className="bg-muted/30">
                <CardContent className="pt-4 pb-3">
                  <div className="grid grid-cols-2 gap-4 text-xs">
                    <div>买入: <span className="font-mono text-green-600">{account.转债分析.买入金额.toFixed(4)} 亿元 ({account.转债分析.买入笔数}笔)</span></div>
                    <div>卖出: <span className="font-mono text-red-500">{account.转债分析.卖出金额.toFixed(4)} 亿元 ({account.转债分析.卖出笔数}笔)</span></div>
                  </div>
                </CardContent>
              </Card>
            </div>
          )}
          {account.期货分析 && (
            <div>
              <h4 className="text-sm font-semibold mb-2">国债期货交易分析</h4>
              <Card className="bg-muted/30">
                <CardContent className="pt-4 pb-3">
                  <div className="grid grid-cols-2 gap-4 text-xs">
                    <div>交易笔数: <span className="font-mono">{account.期货分析.笔数} 笔</span></div>
                    <div>总金额: <span className="font-mono">{account.期货分析.总金额.toFixed(4)} 亿元</span></div>
                  </div>
                </CardContent>
              </Card>
            </div>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
}


export default function Home() {
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string>('');
  const [activeTab, setActiveTab] = useState('hg');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [searchQuery, setSearchQuery] = useState('');
  const [searchOpen, setSearchOpen] = useState(false);
  const [selectedAccount, setSelectedAccount] = useState<string>('');
  const searchRef = useRef<HTMLDivElement>(null);

  const [compareCodes, setCompareCodes] = useState<string[]>(['', '', '']);

  const handleFileUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    setError(null);
    setFileName(file.name);
    try {
      const { hg, zq } = await parseExcel(file);
      const analysis = runAnalysis(hg, zq);
      setResult(analysis);
    } catch (err) {
      setError(err instanceof Error ? err.message : '解析失败');
    } finally {
      setLoading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  }, []);

  const triggerFileSelect = () => fileInputRef.current?.click();

  const searchResults = useMemo(() => {
    if (!result || !searchQuery.trim()) return [];
    const q = searchQuery.trim().toLowerCase();
    return result.accountList
      .filter(a => a.账户代码.toLowerCase().includes(q) || a.账户名称.toLowerCase().includes(q))
      .slice(0, 10);
  }, [result, searchQuery]);

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (searchRef.current && !searchRef.current.contains(e.target as Node)) {
        setSearchOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const radarData = useMemo(() => {
    if (!result) return [];
    const dims = ['回购活跃度', '债券活跃度', '基金活跃度', '期货活跃度', '长期债券偏好', '交易频次'];
    return dims.map(dim => {
      const row: Record<string, string | number> = { subject: dim };
      for (const acct of result.radarAccounts) {
        row[acct.账户名称] = acct[dim as keyof typeof acct] as number;
      }
      return row;
    });
  }, [result]);

  const radarColors = ['#3b82f6', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#06b6d4', '#f97316', '#84cc16', '#ef4444', '#6366f1'];

  const compareAccounts = useMemo(() => {
    if (!result) return [];
    return compareCodes
      .filter(c => c && result.allAccountsDetail[c])
      .map(c => result.allAccountsDetail[c]);
  }, [result, compareCodes]);

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b bg-card sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3 shrink-0">
            <BarChart3 className="w-7 h-7 text-primary" />
            <div>
              <h1 className="text-xl font-bold">交易数据分析平台</h1>
              <p className="text-xs text-muted-foreground">回购业务 · 债券业务 · 深度账户画像</p>
            </div>
          </div>

          {result && (
            <div className="relative flex-1 max-w-md" ref={searchRef}>
              <div className="relative">
                <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
                <Input
                  placeholder="搜索账户代码或名称..."
                  className="pl-9 pr-8"
                  value={searchQuery}
                  onChange={(e) => { setSearchQuery(e.target.value); setSearchOpen(!!e.target.value); }}
                  onFocus={() => searchQuery && setSearchOpen(true)}
                />
                {searchQuery && (
                  <button className="absolute right-2.5 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground" onClick={() => { setSearchQuery(''); setSearchOpen(false); }}>
                    <X className="w-4 h-4" />
                  </button>
                )}
              </div>
              {searchOpen && searchResults.length > 0 && (
                <div className="absolute top-full left-0 right-0 mt-1 bg-popover border rounded-md shadow-lg z-50 max-h-64 overflow-y-auto">
                  {searchResults.map(a => (
                    <button
                      key={a.账户代码}
                      className="w-full text-left px-3 py-2 hover:bg-muted text-sm flex items-center justify-between"
                      onClick={() => { setSelectedAccount(a.账户代码); setSearchOpen(false); setSearchQuery(''); }}
                    >
                      <span><span className="font-medium">{a.账户名称}</span><span className="text-muted-foreground font-mono text-xs ml-2">{a.账户代码}</span></span>
                      <Badge variant="secondary" className="text-xs">{a.总笔数}笔</Badge>
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}

          <div className="flex items-center gap-2 shrink-0">
            {fileName && <Badge variant="outline" className="text-xs hidden md:inline-flex"><FileText className="w-3 h-3 mr-1" />{fileName}</Badge>}
            {result && (
              <>
                <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{ display: 'none' }} />
                <Button variant="outline" size="sm" onClick={triggerFileSelect}><Upload className="w-4 h-4 mr-1" /> 重新上传</Button>
              </>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-6 space-y-6">
        {!result && (
          <Card className="border-dashed border-2">
            <CardContent className="pt-8 pb-8">
              <div className="flex flex-col items-center justify-center text-center space-y-4">
                <div className="w-16 h-16 rounded-full bg-primary/10 flex items-center justify-center"><Upload className="w-8 h-8 text-primary" /></div>
                <div>
                  <h3 className="text-lg font-semibold">上传交易流水Excel文件</h3>
                  <p className="text-sm text-muted-foreground mt-1">支持包含「回购业务」和「债券和其他业务」两个Sheet的Excel文件 · 数据仅在本地分析不上传服务器</p>
                </div>
                <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{ display: 'none' }} />
                <Button disabled={loading} onClick={triggerFileSelect}>{loading ? '分析中...' : '选择文件'}</Button>
                {error && <div className="text-sm text-red-500 bg-red-50 px-4 py-2 rounded-md">{error}</div>}
              </div>
            </CardContent>
          </Card>
        )}

        {result && (
          <>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <Card><CardContent className="pt-5 pb-3"><div className="text-2xl font-bold text-primary">{result.totalHgRecords}</div><div className="text-xs text-muted-foreground">回购业务记录</div></CardContent></Card>
              <Card><CardContent className="pt-5 pb-3"><div className="text-2xl font-bold text-primary">{result.totalZqRecords}</div><div className="text-xs text-muted-foreground">债券及其他记录</div></CardContent></Card>
              <Card><CardContent className="pt-5 pb-3"><div className="text-2xl font-bold text-primary">{result.totalAccounts}</div><div className="text-xs text-muted-foreground">涉及账户总数</div></CardContent></Card>
              <Card><CardContent className="pt-5 pb-3"><div className="text-2xl font-bold text-primary">{result.lowActivityAccounts.length}</div><div className="text-xs text-muted-foreground">低活跃度账户(≤2笔)</div></CardContent></Card>
            </div>

            <AccountSearchDialog open={!!selectedAccount} onClose={() => setSelectedAccount('')} accountCode={selectedAccount} allAccountsDetail={result.allAccountsDetail} />

            <Tabs value={activeTab} onValueChange={setActiveTab} className="space-y-4">
              <ScrollArea className="w-full whitespace-nowrap">
                <TabsList className="inline-flex h-auto gap-1 flex-wrap">
                  <TabsTrigger value="hg">回购业务</TabsTrigger>
                  <TabsTrigger value="xyhg">协议回购</TabsTrigger>
                  <TabsTrigger value="longterm">长期债券</TabsTrigger>
                  <TabsTrigger value="fund">基金</TabsTrigger>
                  <TabsTrigger value="zhaiquan">可转债</TabsTrigger>
                  <TabsTrigger value="qh">国债期货</TabsTrigger>
                  <TabsTrigger value="trend"><TrendingUp className="w-3 h-3 mr-1" />趋势分析</TabsTrigger>
                  <TabsTrigger value="heatmap"><Thermometer className="w-3 h-3 mr-1" />热力图</TabsTrigger>
                  <TabsTrigger value="radar"><RadarIcon className="w-3 h-3 mr-1" />雷达画像</TabsTrigger>
                  <TabsTrigger value="term"><Clock className="w-3 h-3 mr-1" />期限结构</TabsTrigger>
                  <TabsTrigger value="compare"><GitCompare className="w-3 h-3 mr-1" />账户对比</TabsTrigger>
                  <TabsTrigger value="anomaly"><Zap className="w-3 h-3 mr-1" />异常检测</TabsTrigger>
                  <TabsTrigger value="low">低活跃度</TabsTrigger>
                  <TabsTrigger value="special">ABS/REITs</TabsTrigger>
                </TabsList>
              </ScrollArea>

              <TabsContent value="hg" className="space-y-4">
                <Top10Section title="回购业务总体" byValue={result.hgTop10ByValue} byCount={result.hgTop10ByCount} color="#3b82f6" />
              </TabsContent>

              <TabsContent value="xyhg" className="space-y-4">
                <Top10Section title="协议回购" byValue={result.xyhgTop10ByValue} byCount={result.xyhgTop10ByCount} color="#10b981" />
                <Card>
                  <CardHeader><CardTitle className="text-base">协议回购账户特征</CardTitle></CardHeader>
                  <CardContent>
                    <Accordion type="single" collapsible className="w-full">
                      {result.xyhgAccounts.map((acct, i) => (
                        <AccordionItem key={acct.账户代码} value={`xyhg-${i}`}>
                          <AccordionTrigger className="text-sm hover:no-underline">
                            <div className="flex items-center gap-3 text-left w-full pr-4">
                              <Badge variant="outline">{i + 1}</Badge>
                              <span className="font-semibold">{acct.账户名称}</span>
                              <Badge variant="secondary" className="ml-auto text-xs">{acct.协议回购金额.toFixed(4)} 亿元</Badge>
                            </div>
                          </AccordionTrigger>
                          <AccordionContent>
                            <div className="space-y-3 pl-4 pb-2">
                              <div className="text-sm"><span className="text-muted-foreground">协议回购笔数:</span> <span className="font-mono">{acct.协议回购笔数} 笔</span></div>
                              <div className="text-sm"><span className="text-muted-foreground">全部回购类型分布:</span><div className="flex flex-wrap gap-1 mt-1">{Object.entries(acct.全部回购类型).map(([k, v]) => <Badge key={k} variant="secondary" className="text-xs">{k}: {v}</Badge>)}</div></div>
                              <div className="text-sm"><span className="text-muted-foreground">交易方向分布:</span><div className="flex flex-wrap gap-1 mt-1">{Object.entries(acct.交易方向分布).map(([k, v]) => <Badge key={k} variant="outline" className="text-xs">{k}: {v}</Badge>)}</div></div>
                            </div>
                          </AccordionContent>
                        </AccordionItem>
                      ))}
                    </Accordion>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="longterm" className="space-y-4">
                <Top10Section title="长期债券(10-30年、30年以上)" byValue={result.longTermTop10ByValue} byCount={result.longTermTop10ByCount} color="#f59e0b" />
                <Card>
                  <CardHeader><CardTitle className="text-base">长期债券账户深度画像</CardTitle></CardHeader>
                  <CardContent>
                    <Accordion type="single" collapsible className="w-full">
                      {result.longTermAccounts.map((acct, i) => <AccountDetailCard key={acct.账户代码} account={acct} index={i} />)}
                    </Accordion>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="fund" className="space-y-4">
                <Top10Section title="基金买卖" byValue={result.fundTop10ByValue} byCount={result.fundTop10ByCount} color="#8b5cf6" />
                <Card>
                  <CardHeader><CardTitle className="text-base">基金交易账户深度画像</CardTitle></CardHeader>
                  <CardContent>
                    <Accordion type="single" collapsible className="w-full">
                      {result.fundAccounts.map((acct, i) => <AccountDetailCard key={acct.账户代码} account={acct} index={i} />)}
                    </Accordion>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="zhaiquan" className="space-y-4">
                <Top10Section title="可转债交易" byValue={result.zhaiquanTop10ByValue} byCount={result.zhaiquanTop10ByCount} color="#ec4899" />
                <Card>
                  <CardHeader><CardTitle className="text-base">可转债交易账户深度画像</CardTitle></CardHeader>
                  <CardContent>
                    <Accordion type="single" collapsible className="w-full">
                      {result.zhaiquanAccounts.map((acct, i) => <AccountDetailCard key={acct.账户代码} account={acct} index={i} />)}
                    </Accordion>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="qh" className="space-y-4">
                <Top10Section title="国债期货交易" byValue={result.qhTop10ByValue} byCount={result.qhTop10ByCount} color="#06b6d4" />
                <Card>
                  <CardHeader><CardTitle className="text-base">国债期货交易账户深度画像</CardTitle></CardHeader>
                  <CardContent>
                    <Accordion type="single" collapsible className="w-full">
                      {result.qhAccounts.map((acct, i) => <AccountDetailCard key={acct.账户代码} account={acct} index={i} />)}
                    </Accordion>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="low" className="space-y-4">
                <Card>
                  <CardHeader>
                    <CardTitle className="text-base flex items-center gap-2">
                      <AlertTriangle className="w-4 h-4 text-amber-500" /> 低活跃度账户（交易笔数≤2笔，无金额限制）
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-sm text-muted-foreground mb-4">共计 {result.lowActivityAccounts.length} 个账户</div>
                    <Accordion type="single" collapsible className="w-full">
                      {result.lowActivityAccounts.map((acct, i) => (
                        <AccordionItem key={acct.账户代码} value={`low-${i}`}>
                          <AccordionTrigger className="text-sm hover:no-underline">
                            <div className="flex items-center gap-3 text-left w-full pr-4">
                              <Badge variant="outline">{i + 1}</Badge>
                              <span className="font-semibold">{acct.账户名称}</span>
                              <span className="text-xs text-muted-foreground font-mono">{acct.账户代码}</span>
                              <Badge variant="secondary" className="ml-auto text-xs">{acct.总交易笔数} 笔</Badge>
                            </div>
                          </AccordionTrigger>
                          <AccordionContent>
                            <div className="space-y-3 pl-4 pb-2">
                              <div className="text-sm"><span className="text-muted-foreground">总交易金额:</span> <span className="font-mono">{acct.总交易金额.toFixed(4)} 亿元</span></div>
                              {acct.持有品种.length > 0 && (
                                <div className="text-sm">
                                  <span className="text-muted-foreground">交易/持有品种:</span>
                                  <div className="flex flex-wrap gap-1 mt-1">{acct.持有品种.map(h => <Badge key={h} variant="secondary" className="text-xs">{h}</Badge>)}</div>
                                </div>
                              )}
                              <div className="text-sm">
                                <span className="text-muted-foreground">交易明细:</span>
                                <div className="mt-1 space-y-1">
                                  {acct.交易明细.map((d, j) => (
                                    <div key={j} className="text-xs bg-muted/50 px-2 py-1 rounded">[{d.业务品种}] {d.交易方向} {d.名称} | {d.金额.toFixed(4)} 亿元</div>
                                  ))}
                                </div>
                              </div>
                            </div>
                          </AccordionContent>
                        </AccordionItem>
                      ))}
                    </Accordion>
                  </CardContent>
                </Card>
                {result.zeroTradeAccounts.length > 0 && (
                  <Card>
                    <CardHeader><CardTitle className="text-base">无交易账户</CardTitle></CardHeader>
                    <CardContent>
                      <div className="flex flex-wrap gap-2">{result.zeroTradeAccounts.map(acct => <Badge key={acct.账户代码} variant="outline">{acct.账户名称} ({acct.账户代码})</Badge>)}</div>
                    </CardContent>
                  </Card>
                )}
              </TabsContent>

              <TabsContent value="special" className="space-y-4">
                <Top10Section title="ABS（资产支持证券）" byValue={result.absTop10ByValue} byCount={result.absTop10ByCount} color="#84cc16" />
                <Top10Section title="REITs" byValue={result.reitsTop10ByValue} byCount={result.reitsTop10ByCount} color="#f97316" />
                <Card>
                  <CardHeader>
                    <CardTitle className="text-base flex items-center gap-2"><Package className="w-4 h-4" /> 其他特殊品种（分销、远期、转托管、债券回售等）</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <Table>
                      <TableHeader>
                        <TableRow>
                          <TableHead>排名</TableHead>
                          <TableHead className="text-xs">账户代码</TableHead>
                          <TableHead>账户名称</TableHead>
                          <TableHead>业务品种</TableHead>
                          <TableHead className="text-right">金额(亿元)</TableHead>
                          <TableHead className="text-right">笔数</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {result.specialTop10.map(item => (
                          <TableRow key={`${item.账户代码}-${item.业务品种}`}>
                            <TableCell><Badge variant="outline">{item.排名}</Badge></TableCell>
                            <TableCell className="text-xs font-mono text-muted-foreground">{item.账户代码}</TableCell>
                            <TableCell className="text-xs">{item.账户名称}</TableCell>
                            <TableCell><Badge variant="secondary" className="text-xs">{item.业务品种}</Badge></TableCell>
                            <TableCell className="text-right font-mono text-xs">{item.金额.toFixed(4)}</TableCell>
                            <TableCell className="text-right text-xs">{item.笔数}</TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </CardContent>
                </Card>
              </TabsContent>

              {/* ===== 趋势分析 ===== */}
              <TabsContent value="trend" className="space-y-4">
                <Card>
                  <CardHeader><CardTitle className="text-base flex items-center gap-2"><TrendingUp className="w-4 h-4" /> 月度交易趋势</CardTitle></CardHeader>
                  <CardContent>
                    <div className="h-96">
                      <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={result.monthlyTrend} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="月份" />
                          <YAxis />
                          <Tooltip formatter={(v: number) => v.toFixed(4)} />
                          <Legend />
                          <Area type="monotone" dataKey="回购业务" stackId="1" stroke="#3b82f6" fill="#3b82f6" fillOpacity={0.3} />
                          <Area type="monotone" dataKey="债券业务" stackId="1" stroke="#10b981" fill="#10b981" fillOpacity={0.3} />
                          <Area type="monotone" dataKey="基金" stackId="1" stroke="#8b5cf6" fill="#8b5cf6" fillOpacity={0.3} />
                          <Area type="monotone" dataKey="可转债" stackId="1" stroke="#ec4899" fill="#ec4899" fillOpacity={0.3} />
                          <Area type="monotone" dataKey="国债期货" stackId="1" stroke="#06b6d4" fill="#06b6d4" fillOpacity={0.3} />
                        </AreaChart>
                      </ResponsiveContainer>
                    </div>
                  </CardContent>
                </Card>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <Card>
                    <CardHeader><CardTitle className="text-sm">回购 vs 债券 金额对比</CardTitle></CardHeader>
                    <CardContent>
                      <div className="h-72">
                        <ResponsiveContainer width="100%" height="100%">
                          <ComposedChart data={result.monthlyTrend}>
                            <CartesianGrid strokeDasharray="3 3" />
                            <XAxis dataKey="月份" fontSize={11} />
                            <YAxis fontSize={11} />
                            <Tooltip formatter={(v: number) => v.toFixed(4)} />
                            <Legend />
                            <Bar dataKey="回购业务" fill="#3b82f6" name="回购业务" />
                            <Bar dataKey="债券业务" fill="#10b981" name="债券业务" />
                            <Line type="monotone" dataKey="基金" stroke="#8b5cf6" name="基金" dot={false} />
                          </ComposedChart>
                        </ResponsiveContainer>
                      </div>
                    </CardContent>
                  </Card>
                  <Card>
                    <CardHeader><CardTitle className="text-sm">衍生品交易趋势（可转债+期货）</CardTitle></CardHeader>
                    <CardContent>
                      <div className="h-72">
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={result.monthlyTrend}>
                            <CartesianGrid strokeDasharray="3 3" />
                            <XAxis dataKey="月份" fontSize={11} />
                            <YAxis fontSize={11} />
                            <Tooltip formatter={(v: number) => v.toFixed(4)} />
                            <Legend />
                            <Line type="monotone" dataKey="可转债" stroke="#ec4899" strokeWidth={2} dot />
                            <Line type="monotone" dataKey="国债期货" stroke="#06b6d4" strokeWidth={2} dot />
                            <Line type="monotone" dataKey="基金" stroke="#8b5cf6" strokeWidth={2} dot />
                          </LineChart>
                        </ResponsiveContainer>
                      </div>
                    </CardContent>
                  </Card>
                </div>
              </TabsContent>

              {/* ===== 热力图 ===== */}
              <TabsContent value="heatmap" className="space-y-4">
                <Card>
                  <CardHeader><CardTitle className="text-base flex items-center gap-2"><Thermometer className="w-4 h-4" /> 账户月度交易热力图（Top 30 活跃账户）</CardTitle></CardHeader>
                  <CardContent><HeatmapChart data={result.heatmapData} months={result.heatmapMonths} /></CardContent>
                </Card>
              </TabsContent>

              {/* ===== 雷达画像 ===== */}
              <TabsContent value="radar" className="space-y-4">
                <Card>
                  <CardHeader><CardTitle className="text-base flex items-center gap-2"><RadarIcon className="w-4 h-4" /> 账户多维画像雷达图（Top 15 账户）</CardTitle></CardHeader>
                  <CardContent>
                    <div className="h-[500px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <RadarChart cx="50%" cy="50%" outerRadius="70%" data={radarData}>
                          <PolarGrid />
                          <PolarAngleAxis dataKey="subject" tick={{ fontSize: 12 }} />
                          <PolarRadiusAxis angle={30} domain={[0, 100]} tick={{ fontSize: 10 }} />
                          {result.radarAccounts.slice(0, 6).map((acct, i) => (
                            <Radar key={acct.账户代码} name={acct.账户名称} dataKey={acct.账户名称} stroke={radarColors[i % radarColors.length]} fill={radarColors[i % radarColors.length]} fillOpacity={0.1} strokeWidth={2} />
                          ))}
                          <Legend />
                          <Tooltip />
                        </RadarChart>
                      </ResponsiveContainer>
                    </div>
                  </CardContent>
                </Card>
                <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-3">
                  {result.radarAccounts.map((acct) => (
                    <Card key={acct.账户代码} className="cursor-pointer hover:bg-muted/50 transition-colors" onClick={() => setSelectedAccount(acct.账户代码)}>
                      <CardContent className="pt-4 pb-3">
                        <div className="text-xs font-medium truncate" title={acct.账户名称}>{acct.账户名称}</div>
                        <div className="text-[10px] text-muted-foreground font-mono">{acct.账户代码}</div>
                        <div className="mt-2 space-y-1">
                          {(['回购活跃度', '债券活跃度', '基金活跃度', '期货活跃度', '长期债券偏好', '交易频次'] as const).map(dim => (
                            <div key={dim} className="flex items-center gap-1">
                              <div className="text-[10px] text-muted-foreground w-16 truncate">{dim}</div>
                              <div className="flex-1 h-1.5 bg-muted rounded-full overflow-hidden"><div className="h-full bg-primary rounded-full" style={{ width: `${acct[dim]}%` }} /></div>
                              <div className="text-[10px] font-mono w-6 text-right">{acct[dim]}</div>
                            </div>
                          ))}
                        </div>
                      </CardContent>
                    </Card>
                  ))}
                </div>
              </TabsContent>

              {/* ===== 期限结构 ===== */}
              <TabsContent value="term" className="space-y-4">
                <Card>
                  <CardHeader><CardTitle className="text-base flex items-center gap-2"><Clock className="w-4 h-4" /> 债券期限结构分布</CardTitle></CardHeader>
                  <CardContent>
                    <div className="h-80">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={result.termStructure} margin={{ top: 10, right: 30, left: 0, bottom: 5 }}>
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="期限段" />
                          <YAxis />
                          <Tooltip formatter={(v: number) => v.toFixed(4)} />
                          <Legend />
                          <Bar dataKey="买入金额" fill="#10b981" name="买入金额" radius={[4, 4, 0, 0]} />
                          <Bar dataKey="卖出金额" fill="#ef4444" name="卖出金额" radius={[4, 4, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </CardContent>
                </Card>
                <Card>
                  <CardHeader><CardTitle className="text-sm">期限结构净买卖</CardTitle></CardHeader>
                  <CardContent>
                    <div className="h-64">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={result.termStructure}>
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="期限段" fontSize={11} />
                          <YAxis fontSize={11} />
                          <Tooltip formatter={(v: number) => v.toFixed(4)} />
                          <Bar dataKey="净买卖" fill="#3b82f6" name="净买卖" radius={[4, 4, 0, 0]} fillOpacity={0.8} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </CardContent>
                </Card>
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>期限段</TableHead>
                      <TableHead className="text-right">买入金额(亿元)</TableHead>
                      <TableHead className="text-right">卖出金额(亿元)</TableHead>
                      <TableHead className="text-right">净买卖(亿元)</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {result.termStructure.map(row => (
                      <TableRow key={row.期限段}>
                        <TableCell className="font-medium">{row.期限段}</TableCell>
                        <TableCell className="text-right font-mono text-green-600">{row.买入金额.toFixed(4)}</TableCell>
                        <TableCell className="text-right font-mono text-red-500">{row.卖出金额.toFixed(4)}</TableCell>
                        <TableCell className={`text-right font-mono font-medium ${row.净买卖 >= 0 ? 'text-green-600' : 'text-red-500'}`}>{row.净买卖.toFixed(4)}</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </TabsContent>

              {/* ===== 账户对比 ===== */}
              <TabsContent value="compare" className="space-y-4">
                <Card>
                  <CardHeader><CardTitle className="text-base flex items-center gap-2"><GitCompare className="w-4 h-4" /> 账户对比分析</CardTitle></CardHeader>
                  <CardContent>
                    <div className="flex flex-wrap gap-3 mb-6">
                      {[0, 1, 2].map(idx => (
                        <div key={idx} className="w-64">
                          <label className="text-xs text-muted-foreground mb-1 block">选择账户 {idx + 1}</label>
                          <Select value={compareCodes[idx]} onValueChange={(v) => { const next = [...compareCodes]; next[idx] = v; setCompareCodes(next); }}>
                            <SelectTrigger><SelectValue placeholder={`选择账户 ${idx + 1}`} /></SelectTrigger>
                            <SelectContent>
                              {result.accountList.map(a => <SelectItem key={a.账户代码} value={a.账户代码}>{a.账户名称} ({a.账户代码})</SelectItem>)}
                            </SelectContent>
                          </Select>
                        </div>
                      ))}
                    </div>
                    {compareAccounts.length > 0 && (
                      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                        {compareAccounts.map(acct => (
                          <Card key={acct.账户代码} className="border-2">
                            <CardHeader className="pb-2">
                              <CardTitle className="text-sm flex items-center justify-between">
                                <span className="truncate">{acct.账户名称}</span>
                                <Badge variant="outline" className="font-mono text-xs shrink-0">{acct.账户代码}</Badge>
                              </CardTitle>
                            </CardHeader>
                            <CardContent className="space-y-3">
                              <div className="grid grid-cols-2 gap-2 text-xs">
                                <div className="bg-muted/50 rounded p-2 text-center">
                                  <div className="text-lg font-bold text-primary">{acct.总笔数}</div>
                                  <div className="text-muted-foreground">总笔数</div>
                                </div>
                                <div className="bg-muted/50 rounded p-2 text-center">
                                  <div className="text-lg font-bold text-primary">{Object.values(acct.买卖方向).reduce((a, b) => a + b, 0).toFixed(2)}</div>
                                  <div className="text-muted-foreground">总金额(亿元)</div>
                                </div>
                              </div>
                              {acct.债券类别.length > 0 && (
                                <div>
                                  <div className="text-xs font-semibold mb-1">债券类别</div>
                                  {acct.债券类别.map(cat => (
                                    <div key={cat.类别} className="text-xs grid grid-cols-3 gap-1 mb-1">
                                      <span className="text-muted-foreground">{cat.类别}</span>
                                      <span className="text-green-600 font-mono">买 {cat.买入金额.toFixed(2)}</span>
                                      <span className="text-red-500 font-mono">卖 {cat.卖出金额.toFixed(2)}</span>
                                    </div>
                                  ))}
                                </div>
                              )}
                              {acct.基金分析 && <div className="text-xs"><span className="font-semibold">基金:</span><span className="text-green-600 ml-1">买 {acct.基金分析.买入金额.toFixed(2)}</span><span className="text-red-500 ml-2">卖 {acct.基金分析.卖出金额.toFixed(2)}</span></div>}
                              {acct.转债分析 && <div className="text-xs"><span className="font-semibold">转债:</span><span className="text-green-600 ml-1">买 {acct.转债分析.买入金额.toFixed(2)}</span><span className="text-red-500 ml-2">卖 {acct.转债分析.卖出金额.toFixed(2)}</span></div>}
                              {acct.期货分析 && <div className="text-xs"><span className="font-semibold">期货:</span><span className="ml-1">{acct.期货分析.笔数} 笔 / {acct.期货分析.总金额.toFixed(2)} 亿元</span></div>}
                            </CardContent>
                          </Card>
                        ))}
                      </div>
                    )}
                  </CardContent>
                </Card>
              </TabsContent>

              {/* ===== 异常检测 ===== */}
              <TabsContent value="anomaly" className="space-y-4">
                <Card>
                  <CardHeader>
                    <CardTitle className="text-base flex items-center gap-2"><Zap className="w-4 h-4 text-amber-500" /> 异常交易检测</CardTitle>
                  </CardHeader>
                  <CardContent>
                    {result.anomalies.length === 0 ? (
                      <div className="text-center py-8 text-muted-foreground">未检测到异常交易</div>
                    ) : (
                      <>
                        <div className="text-sm text-muted-foreground mb-4">
                          共检测到 {result.anomalies.filter(a => a.类型 === '超大额').length} 笔超大额交易、{result.anomalies.filter(a => a.类型 === '高频率').length} 个高频账户
                        </div>
                        <Table>
                          <TableHeader>
                            <TableRow>
                              <TableHead>类型</TableHead>
                              <TableHead>账户</TableHead>
                              <TableHead>业务品种</TableHead>
                              <TableHead>交易方向</TableHead>
                              <TableHead>名称</TableHead>
                              <TableHead className="text-right">金额(亿元)</TableHead>
                              <TableHead>日期/月份</TableHead>
                            </TableRow>
                          </TableHeader>
                          <TableBody>
                            {result.anomalies.map((a, i) => (
                              <TableRow key={i} className={a.类型 === '超大额' ? 'bg-red-50/50' : 'bg-amber-50/50'}>
                                <TableCell><Badge variant={a.类型 === '超大额' ? 'destructive' : 'secondary'} className="text-xs">{a.类型}</Badge></TableCell>
                                <TableCell className="text-xs"><div className="font-medium">{a.账户名称}</div><div className="text-muted-foreground font-mono">{a.账户代码}</div></TableCell>
                                <TableCell className="text-xs">{a.业务品种}</TableCell>
                                <TableCell className="text-xs">{a.交易方向}</TableCell>
                                <TableCell className="text-xs max-w-[120px] truncate">{a.名称}</TableCell>
                                <TableCell className={`text-right font-mono text-xs ${a.金额 >= 50 ? 'text-red-600 font-bold' : ''}`}>{a.金额 > 0 ? a.金额.toFixed(4) : '-'}</TableCell>
                                <TableCell className="text-xs">{a.日期 || '-'}</TableCell>
                              </TableRow>
                            ))}
                          </TableBody>
                        </Table>
                      </>
                    )}
                  </CardContent>
                </Card>
              </TabsContent>
            </Tabs>
          </>
        )}
      </main>

      <footer className="border-t mt-12 py-4 text-center text-xs text-muted-foreground">
        交易数据分析平台 · 金额单位：亿元 · 数据仅在本地分析不上传服务器
      </footer>
    </div>
  );
}
