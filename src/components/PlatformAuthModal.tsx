import React, { useState, useEffect } from 'react';
import {
  X,
  KeyRound,
  Cookie,
  ShieldCheck,
  CheckCircle2,
  AlertCircle,
  Loader2,
  LogOut,
  ExternalLink,
  HelpCircle,
  Store,
  Sparkles,
  Lock,
  Unlock,
  Check
} from 'lucide-react';
import { PriceSourceSite, PlatformAuthStatus } from '../types';

interface PlatformAuthModalProps {
  isOpen: boolean;
  onClose: () => void;
  initialPlatform?: PriceSourceSite;
  onSessionUpdated?: () => void;
}

interface PlatformConfig {
  id: PriceSourceSite;
  name: string;
  domain: string;
  colorClass: string;
  activeColorClass: string;
  badgeBg: string;
  badgeText: string;
  requiresAuth: boolean;
  description: string;
  loginTip: string;
  authNotice: string;
  usernameLabel: string;
  usernamePlaceholder: string;
  officialLoginUrl: string;
}

const PLATFORM_CONFIGS: PlatformConfig[] = [
  {
    id: 'meloni',
    name: 'Meloni Store',
    domain: 'melonistore.com',
    colorClass: 'border-purple-500/40 text-purple-400 bg-purple-950/20',
    activeColorClass: 'bg-purple-600 text-white shadow-lg shadow-purple-950/50',
    badgeBg: 'bg-purple-950/80 border-purple-800 text-purple-300',
    badgeText: 'B2B商户批发',
    requiresAuth: true,
    description: '日化洗护与百货批发商城。常规商品公开，大宗采购与批发底价需商户登录。',
    loginTip: '输入您的 Meloni Store 商户账号或登录后的 Cookie，提取时直接解锁批发专享底价。',
    authNotice: '若遇 "Accedi per visualizzare i prezzi" (登录查看价格)，配置此授权即可破除限制！',
    usernameLabel: 'Meloni 注册邮箱 (Email)',
    usernamePlaceholder: 'name@example.com',
    officialLoginUrl: 'https://www.melonistore.com/login.php',
  },
  {
    id: 'megacedi',
    name: 'MegaCedi',
    domain: 'megacedi.com',
    colorClass: 'border-emerald-500/40 text-emerald-400 bg-emerald-950/20',
    activeColorClass: 'bg-emerald-600 text-white shadow-lg shadow-emerald-950/50',
    badgeBg: 'bg-emerald-950/80 border-emerald-800 text-emerald-300',
    badgeText: 'B2B批发仓储',
    requiresAuth: true,
    description: '意大利商超百货 B2B 批发订货平台。全量支持装箱规格，核心采购底价需商户授权。',
    loginTip: '输入 MegaCedi 账号密码或 Cookie，即可直接获取商超内部专属采购底价。',
    authNotice: '系统支持账号密码直接提交或 Cookie 导入，自动保持长效 Session。',
    usernameLabel: 'MegaCedi 用户名 / 账号',
    usernamePlaceholder: '输入您的商户账号',
    officialLoginUrl: 'https://www.megacedi.com/GestioneUtentiPubblici/home.aspx',
  },
  {
    id: 'maurys',
    name: "Maury's Online",
    domain: 'maurysonline.it',
    colorClass: 'border-cyan-500/40 text-cyan-400 bg-cyan-950/20',
    activeColorClass: 'bg-cyan-600 text-white shadow-lg shadow-cyan-950/50',
    badgeBg: 'bg-cyan-950/80 border-cyan-800 text-cyan-300',
    badgeText: '公开零售大卖场',
    requiresAuth: false,
    description: '意大利连锁百货电商。100% 免登录，全网所有商品原价、红字特价完全公开透明。',
    loginTip: 'Maury\'s 默认无需任何登录即可直接提取所有价格！如需绑定会员专享价也可在此填入 Cookie。',
    authNotice: '🟢 100% 免登录：不配置任何账号也可正常并发比价和导出！',
    usernameLabel: '会员账号 (选填)',
    usernamePlaceholder: '可选会员用户名',
    officialLoginUrl: 'https://maurysonline.it/login',
  },
  {
    id: 'risparmiocasa',
    name: 'Risparmio Casa',
    domain: 'shop.risparmiocasa.com',
    colorClass: 'border-amber-500/40 text-amber-400 bg-amber-950/20',
    activeColorClass: 'bg-amber-600 text-slate-950 shadow-lg shadow-amber-950/50 font-bold',
    badgeBg: 'bg-amber-950/80 border-amber-800 text-amber-300',
    badgeText: '公开连锁超市',
    requiresAuth: false,
    description: '意大利国民日化连锁。100% 免登录，直营门店价与活动折扣全网免登录抓取。',
    loginTip: 'Risparmio Casa 默认无需任何登录即可直接提取！若有会员特权也可在此配置专属 Cookie。',
    authNotice: '🟢 100% 免登录：无需账号密码，即开即用！',
    usernameLabel: '会员账号 (选填)',
    usernamePlaceholder: '可选会员用户名',
    officialLoginUrl: 'https://shop.risparmiocasa.com/customer/account/login/',
  },
  {
    id: 'piume',
    name: 'PiùMe (piume.it)',
    domain: 'piume.it',
    colorClass: 'border-rose-500/40 text-rose-400 bg-rose-950/20',
    activeColorClass: 'bg-rose-600 text-white shadow-lg shadow-rose-950/50',
    badgeBg: 'bg-rose-950/80 border-rose-800 text-rose-300',
    badgeText: '公开美妆日化连锁',
    requiresAuth: false,
    description: '意大利知名个人护理与家居日化连锁电商。100% 免登录，数万种日化大牌价格全公开。',
    loginTip: 'PiùMe 默认 100% 免登录即可提取！若持有 PiùMe 会员卡或专属特权，可在此填入账号或 Cookie。',
    authNotice: '🟢 100% 免登录：无需账号密码，已自动支持极速高精匹配！',
    usernameLabel: 'PiùMe 账号/邮箱 (选填)',
    usernamePlaceholder: '可选会员邮箱',
    officialLoginUrl: 'https://piume.it/login.php',
  },
];

export const PlatformAuthModal: React.FC<PlatformAuthModalProps> = ({
  isOpen,
  onClose,
  initialPlatform = 'meloni',
  onSessionUpdated,
}) => {
  const [selectedPlatform, setSelectedPlatform] = useState<PriceSourceSite>(initialPlatform);
  const [activeTab, setActiveTab] = useState<'credentials' | 'cookie'>('credentials');
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [temporaryCode, setTemporaryCode] = useState('');
  const [cookieString, setCookieString] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [statuses, setStatuses] = useState<Record<string, PlatformAuthStatus>>({});
  const [feedback, setFeedback] = useState<{ type: 'success' | 'error' | 'info'; message: string } | null>(null);

  // Sync selectedPlatform when initialPlatform changes
  useEffect(() => {
    if (initialPlatform) {
      setSelectedPlatform(initialPlatform);
    }
  }, [initialPlatform]);

  // Fetch status for all platforms
  const fetchAllStatuses = async () => {
    try {
      const res = await fetch('/api/auth/status');
      const data = await res.json();
      if (data.success && data.statuses) {
        setStatuses(data.statuses);
      }
    } catch {}
  };

  useEffect(() => {
    if (isOpen) {
      fetchAllStatuses();
      setFeedback(null);
    }
  }, [isOpen]);

  if (!isOpen) return null;

  const currentConfig = PLATFORM_CONFIGS.find((p) => p.id === selectedPlatform) || PLATFORM_CONFIGS[0];
  const currentStatus = statuses[selectedPlatform];

  const handleLoginWithCredentials = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!username.trim() || !password.trim()) {
      setFeedback({ type: 'error', message: '请填写完整的账号和密码' });
      return;
    }

    setIsLoading(true);
    setFeedback({ type: 'info', message: `正在向 ${currentConfig.domain} 发起登录认证...` });

    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          platform: selectedPlatform,
          username: username.trim(),
          password: password.trim(),
          temporaryCode: temporaryCode.trim(),
        }),
      });
      const data = await res.json();
      if (data.success) {
        setFeedback({
          type: 'success',
          message: data.message || `${currentConfig.name} 登录成功！已保存认证会话。`,
        });
        await fetchAllStatuses();
        onSessionUpdated?.();
      } else {
        setFeedback({
          type: 'error',
          message: data.error || '登录失败，请检查账号密码，或直接切换到【Cookie 模式】导入。',
        });
      }
    } catch (err: any) {
      setFeedback({ type: 'error', message: '网络错误: ' + err.message });
    } finally {
      setIsLoading(false);
    }
  };

  const handleSaveCookie = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!cookieString.trim()) {
      setFeedback({ type: 'error', message: '请输入从浏览器中复制的 Cookie 字符串' });
      return;
    }

    setIsLoading(true);
    setFeedback({ type: 'info', message: `正在保存并验证 ${currentConfig.name} Cookie 凭证...` });

    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          platform: selectedPlatform,
          cookie: cookieString.trim(),
          username: username.trim() || undefined,
        }),
      });
      const data = await res.json();
      if (data.success) {
        setFeedback({ type: 'success', message: `${currentConfig.name} Cookie 保存成功！后续请求将自动携带该会话。` });
        await fetchAllStatuses();
        onSessionUpdated?.();
      } else {
        setFeedback({ type: 'error', message: data.error || 'Cookie 保存失败，请确认格式是否正确。' });
      }
    } catch (err: any) {
      setFeedback({ type: 'error', message: '网络错误: ' + err.message });
    } finally {
      setIsLoading(false);
    }
  };

  const handleLogout = async () => {
    setIsLoading(true);
    try {
      const res = await fetch('/api/auth/logout', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ platform: selectedPlatform }),
      });
      const data = await res.json();
      if (data.success) {
        setFeedback({ type: 'info', message: `已清除 ${currentConfig.name} 的登录会话凭证。` });
        await fetchAllStatuses();
        onSessionUpdated?.();
      }
    } catch (err: any) {
      setFeedback({ type: 'error', message: '清除会话失败: ' + err.message });
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/85 backdrop-blur-md animate-in fade-in duration-200">
      <div className="bg-slate-900 border border-slate-800 w-full max-w-2xl rounded-2xl shadow-2xl overflow-hidden relative flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="p-5 border-b border-slate-800 flex items-center justify-between bg-slate-950/60 shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-amber-500/15 border border-amber-500/30 flex items-center justify-center text-amber-400">
              <KeyRound className="w-5 h-5" />
            </div>
            <div>
              <h3 className="font-black text-white text-base flex items-center gap-2">
                <span>多平台账号与 Cookie 授权管理中心</span>
                <span className="text-[10px] bg-amber-500 text-slate-950 font-black px-2 py-0.5 rounded-full">
                  多平台支持
                </span>
              </h3>
              <p className="text-xs text-slate-400">
                配置 Meloni、MegaCedi 等 B2B 批发账号或 Cookie，解锁商超专属批发底价与全权数据
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Platform Selector Tabs */}
        <div className="p-4 bg-slate-950/40 border-b border-slate-800/80 shrink-0">
          <div className="text-[11px] font-bold text-slate-400 mb-2 flex items-center justify-between">
            <span>选择要配置授权的网站：</span>
            <span className="text-[10px] text-slate-500">点击切换不同平台配置</span>
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
            {PLATFORM_CONFIGS.map((plat) => {
              const isPlatActive = selectedPlatform === plat.id;
              const isAuthed = statuses[plat.id]?.authenticated;
              return (
                <button
                  key={plat.id}
                  type="button"
                  onClick={() => {
                    setSelectedPlatform(plat.id);
                    setFeedback(null);
                  }}
                  className={`p-2.5 rounded-xl border text-left transition relative cursor-pointer ${
                    isPlatActive
                      ? `${plat.activeColorClass} border-transparent ring-2 ring-white/20`
                      : 'bg-slate-900 border-slate-800 text-slate-300 hover:border-slate-700 hover:text-white'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-xs">{plat.name}</span>
                    {isAuthed ? (
                      <span className="w-2 h-2 rounded-full bg-emerald-400 ring-2 ring-emerald-400/30" title="已认证" />
                    ) : plat.requiresAuth ? (
                      <span title="需授权"><Lock className="w-3 h-3 text-amber-400/80 opacity-70" /></span>
                    ) : (
                      <span title="免登录"><Unlock className="w-3 h-3 text-cyan-400/80 opacity-70" /></span>
                    )}
                  </div>
                  <div className="text-[10px] opacity-75 font-mono truncate mt-0.5">{plat.domain}</div>
                  <div className="mt-1 flex items-center justify-between text-[9px]">
                    <span className={`px-1 py-0.2 rounded ${isPlatActive ? 'bg-black/30' : 'bg-slate-800 text-slate-400'}`}>
                      {isAuthed ? '🟢 已认证' : plat.requiresAuth ? '需登录' : '免登录'}
                    </span>
                  </div>
                </button>
              );
            })}
          </div>
        </div>

        {/* Modal Scrollable Body */}
        <div className="p-5 overflow-y-auto space-y-4 flex-1">
          {/* Platform Info & Status Card */}
          <div className="bg-slate-950/70 border border-slate-800 rounded-xl p-4 space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-800/80 pb-3">
              <div className="flex items-center gap-2">
                <Store className="w-4 h-4 text-amber-400" />
                <span className="font-bold text-sm text-white">{currentConfig.name}</span>
                <code className="text-xs text-slate-400 font-mono">({currentConfig.domain})</code>
                <span className={`text-[10px] px-2 py-0.5 rounded-full border ${currentConfig.badgeBg}`}>
                  {currentConfig.badgeText}
                </span>
              </div>

              {currentStatus?.authenticated ? (
                <div className="flex items-center gap-2">
                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-950 border border-emerald-500/40 text-emerald-300">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                    <span>已激活认证 ({currentStatus.username || 'Cookie Session'})</span>
                  </span>
                  <button
                    type="button"
                    onClick={handleLogout}
                    disabled={isLoading}
                    className="text-xs text-red-400 hover:text-red-300 bg-red-950/40 hover:bg-red-950 border border-red-800/40 px-2 py-1 rounded-lg transition flex items-center gap-1 cursor-pointer"
                  >
                    <LogOut className="w-3 h-3" />
                    <span>清除会话</span>
                  </button>
                </div>
              ) : (
                <div className="flex items-center gap-1.5 text-xs text-slate-400">
                  <span className="w-2 h-2 rounded-full bg-slate-500" />
                  <span>当前状态：{currentConfig.requiresAuth ? '未登录（公开模式）' : '公开零售（免登录即查）'}</span>
                </div>
              )}
            </div>

            <p className="text-xs text-slate-300 leading-relaxed">
              {currentConfig.description}
            </p>

            <div className="flex items-center justify-between text-xs text-amber-300/90 bg-amber-950/20 border border-amber-800/30 rounded-lg p-2.5">
              <span className="flex items-center gap-1.5 text-[11px]">
                <HelpCircle className="w-3.5 h-3.5 text-amber-400 shrink-0" />
                <span>{currentConfig.authNotice}</span>
              </span>
              <a
                href={currentConfig.officialLoginUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="text-[11px] text-cyan-400 hover:underline flex items-center gap-1 shrink-0 ml-2"
              >
                <span>在浏览器中打开登录页</span>
                <ExternalLink className="w-3 h-3" />
              </a>
            </div>
          </div>

          {/* Feedback Alert */}
          {feedback && (
            <div
              className={`p-3 rounded-xl border flex items-start gap-2.5 text-xs animate-in fade-in ${
                feedback.type === 'success'
                  ? 'bg-emerald-950/70 border-emerald-500/40 text-emerald-300'
                  : feedback.type === 'error'
                  ? 'bg-red-950/70 border-red-500/40 text-red-300'
                  : 'bg-cyan-950/70 border-cyan-500/40 text-cyan-300'
              }`}
            >
              {feedback.type === 'success' ? (
                <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
              ) : feedback.type === 'error' ? (
                <AlertCircle className="w-4 h-4 text-red-400 shrink-0 mt-0.5" />
              ) : (
                <Loader2 className="w-4 h-4 text-cyan-400 shrink-0 mt-0.5 animate-spin" />
              )}
              <span>{feedback.message}</span>
            </div>
          )}

          {/* Mode Switcher: Account/Password vs Cookie */}
          <div className="border border-slate-800 rounded-xl p-4 bg-slate-950/40 space-y-4">
            <div className="flex border-b border-slate-800 pb-3 gap-2">
              <button
                type="button"
                onClick={() => setActiveTab('credentials')}
                className={`flex-1 py-2 px-3 rounded-lg text-xs font-bold transition flex items-center justify-center gap-1.5 cursor-pointer ${
                  activeTab === 'credentials'
                    ? 'bg-slate-800 text-white shadow'
                    : 'text-slate-400 hover:text-white hover:bg-slate-900'
                }`}
              >
                <KeyRound className="w-3.5 h-3.5 text-amber-400" />
                <span>方式一：账号密码自动登录</span>
              </button>
              <button
                type="button"
                onClick={() => setActiveTab('cookie')}
                className={`flex-1 py-2 px-3 rounded-lg text-xs font-bold transition flex items-center justify-center gap-1.5 cursor-pointer ${
                  activeTab === 'cookie'
                    ? 'bg-slate-800 text-white shadow'
                    : 'text-slate-400 hover:text-white hover:bg-slate-900'
                }`}
              >
                <Cookie className="w-3.5 h-3.5 text-cyan-400" />
                <span>方式二：Cookie 极速导入 (推荐 100% 成功)</span>
              </button>
            </div>

            {/* Form 1: Username & Password */}
            {activeTab === 'credentials' && (
              <form onSubmit={handleLoginWithCredentials} className="space-y-3.5">
                <div>
                  <label className="block text-xs font-bold text-slate-300 mb-1">
                    {currentConfig.usernameLabel} <span className="text-red-400">*</span>
                  </label>
                  <input
                    type="text"
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                    placeholder={currentConfig.usernamePlaceholder}
                    className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-amber-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-300 mb-1">
                    登录密码 (Password) <span className="text-red-400">*</span>
                  </label>
                  <input
                    type="password"
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    placeholder="输入密码"
                    className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-amber-500"
                  />
                </div>

                {selectedPlatform === 'megacedi' && (
                  <div>
                    <label className="block text-xs font-medium text-slate-400 mb-1">
                      激活码 / 临时验证码 (Codice Attivazione，无则留空)
                    </label>
                    <input
                      type="text"
                      value={temporaryCode}
                      onChange={(e) => setTemporaryCode(e.target.value)}
                      placeholder="选填"
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs text-white placeholder-slate-600 focus:outline-none focus:border-slate-700"
                    />
                  </div>
                )}

                <button
                  type="submit"
                  disabled={isLoading}
                  className="w-full bg-gradient-to-r from-amber-500 via-amber-600 to-amber-700 hover:from-amber-400 hover:to-amber-600 text-slate-950 font-black text-xs py-2.5 rounded-lg transition shadow-md flex items-center justify-center gap-2 cursor-pointer disabled:opacity-50"
                >
                  {isLoading ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin text-slate-950" />
                      <span>正在向 {currentConfig.name} 发起登录认证...</span>
                    </>
                  ) : (
                    <>
                      <ShieldCheck className="w-4 h-4 text-slate-950" />
                      <span>立即提交 {currentConfig.name} 登录并保存认证</span>
                    </>
                  )}
                </button>
              </form>
            )}

            {/* Form 2: Cookie Mode */}
            {activeTab === 'cookie' && (
              <form onSubmit={handleSaveCookie} className="space-y-3.5">
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-xs font-bold text-slate-300">
                      浏览器已登录 Cookie 字符串 <span className="text-red-400">*</span>
                    </label>
                    <span className="text-[10px] text-cyan-400 font-mono">从浏览器 Network 请求头直接复制</span>
                  </div>
                  <textarea
                    rows={4}
                    value={cookieString}
                    onChange={(e) => setCookieString(e.target.value)}
                    placeholder={`在浏览器登录 ${currentConfig.domain} 后，按 F12 -> Network -> 复制 Request Headers 里的 Cookie 字符串粘贴在此处...\n例: SHOP_SESSION_TOKEN=xxx; PHPSESSID=yyy; ...`}
                    className="w-full bg-slate-900 border border-slate-700 rounded-lg p-2.5 text-xs text-white font-mono placeholder-slate-600 focus:outline-none focus:border-cyan-500"
                  />
                </div>

                <div className="bg-slate-900/90 border border-slate-800 rounded-lg p-3 text-[11px] text-slate-400 space-y-1.5">
                  <div className="font-bold text-slate-300 flex items-center gap-1.5">
                    <Sparkles className="w-3.5 h-3.5 text-cyan-400" />
                    <span>30秒快速获取 Cookie 教程：</span>
                  </div>
                  <ol className="list-decimal list-inside space-y-1 text-[10px] text-slate-400 leading-relaxed pl-1">
                    <li>在 Chrome/Edge 浏览器中打开 <a href={currentConfig.officialLoginUrl} target="_blank" rel="noopener noreferrer" className="text-cyan-400 underline">{currentConfig.domain}</a> 并登录您的账号；</li>
                    <li>按键盘快捷键 <code className="text-white bg-slate-800 px-1 py-0.5 rounded">F12</code> 打开开发者工具，切换到 <strong>Network (网络)</strong> 面板；</li>
                    <li>刷新网页，点击列表第一条请求，在右侧 <strong>Request Headers (请求标头)</strong> 中找到 <strong>Cookie:</strong> 并完整复制；</li>
                    <li>粘贴到上方文本框，点击下方【保存 Cookie 凭据】即可永久解锁！</li>
                  </ol>
                </div>

                <button
                  type="submit"
                  disabled={isLoading}
                  className="w-full bg-gradient-to-r from-cyan-600 to-blue-600 hover:from-cyan-500 hover:to-blue-500 text-white font-black text-xs py-2.5 rounded-lg transition shadow-md flex items-center justify-center gap-2 cursor-pointer disabled:opacity-50"
                >
                  {isLoading ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>正在保存并验证 Cookie...</span>
                    </>
                  ) : (
                    <>
                      <Check className="w-4 h-4" />
                      <span>保存并激活 {currentConfig.name} Cookie 凭证</span>
                    </>
                  )}
                </button>
              </form>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="p-4 border-t border-slate-800 bg-slate-950/80 flex items-center justify-between text-xs text-slate-400 shrink-0">
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-4 h-4 text-emerald-400" />
            <span>凭据加密保存在本地服务器内存中，绝不上传第三方</span>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-white rounded-lg transition cursor-pointer font-bold text-xs"
          >
            完成并关闭
          </button>
        </div>
      </div>
    </div>
  );
};
