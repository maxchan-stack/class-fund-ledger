import React, { useState, useEffect } from 'react';
import { Lock, ShieldCheck, AlertCircle, BookOpen, CreditCard, Receipt, Loader2, ArrowUpRight, ArrowDownRight, Wallet } from 'lucide-react';

// 沿用主系統的儲存服務：優先使用 Claude 環境的 window.storage，否則降級為 localStorage
const StorageService = {
  async get(key, shared = true) {
    if (window.storage && typeof window.storage.get === 'function') {
      try {
        const res = await window.storage.get(key, shared);
        if (res && res.value !== undefined) return res;
      } catch (e) {
        console.warn('window.storage.get 失敗，降級使用 localStorage', e);
      }
    }
    const val = localStorage.getItem(key);
    return val ? { value: val } : null;
  }
};

// 從網址列取得雲端試算表位址：支援 ?api=腳本ID 或 ?url=完整網址，跟主系統邏輯一致
function resolveSheetUrlFromLocation() {
  const params = new URLSearchParams(window.location.search);
  const apiId = params.get('api') || params.get('id');
  const urlParam = params.get('url');
  if (urlParam) return urlParam;
  if (apiId) {
    if (apiId.startsWith('http://') || apiId.startsWith('https://')) {
      const match = apiId.match(/macros\/s\/([^/]+)\/exec/);
      return match ? `https://script.google.com/macros/s/${match[1]}/exec` : apiId;
    }
    return `https://script.google.com/macros/s/${apiId}/exec`;
  }
  return '';
}

export default function ParentPortal() {
  const [sheetUrl, setSheetUrl] = useState('');
  const [seat, setSeat] = useState('');
  const [pin, setPin] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [view, setView] = useState(null);
  const [justCreatedPin, setJustCreatedPin] = useState(false);

  useEffect(() => {
    (async () => {
      const DEFAULT_SHEET_URL = 'https://script.google.com/macros/s/AKfycbyttGtb7zSbR8_Z7eErIHBcA3OBmk8ylghKzOybZ0TUD5zT3O_LuTtUpXNjzAEylj3Xvg/exec';
      const fromUrl = resolveSheetUrlFromLocation();
      if (fromUrl) {
        setSheetUrl(fromUrl);
        return;
      }
      const s = await StorageService.get('settings', true);
      if (s) {
        try {
          const parsed = JSON.parse(s.value);
          setSheetUrl(parsed.sheetUrl || DEFAULT_SHEET_URL);
          return;
        } catch (e) {
          console.error('讀取現有設定失敗', e);
        }
      }
      setSheetUrl(DEFAULT_SHEET_URL);
    })();
  }, []);

  async function handleLogin(e) {
    e.preventDefault();
    setError('');

    if (!sheetUrl) {
      setError('尚未取得雲端試算表連結，請確認開啟的網址是否正確（需包含 ?api=... 參數）');
      return;
    }
    if (!seat.trim() || !pin.trim()) {
      setError('請輸入座號與 PIN 碼');
      return;
    }
    if (pin.trim().length !== 4) {
      setError('PIN 碼必須為 4 位數字');
      return;
    }

    setLoading(true);
    try {
      const queryParams = new URLSearchParams({
        action: 'parentAuth',
        seat: seat.trim(),
        pin: pin.trim()
      });
      const separator = sheetUrl.includes('?') ? '&' : '?';
      const res = await fetch(`${sheetUrl}${separator}${queryParams.toString()}`, {
        method: 'GET',
        mode: 'cors'
      });
      const json = await res.json();
      if (!json.success) {
        setError(json.error || '登入失敗，請確認座號與 PIN 碼是否正確');
        setLoading(false);
        return;
      }
      setView(json.data);
      setJustCreatedPin(!!json.isNewSetup);
    } catch (err) {
      console.error(err);
      setError('連線失敗，請確認網路連線或稍後再試');
    } finally {
      setLoading(false);
    }
  }

  function handleLogout() {
    setView(null);
    setSeat('');
    setPin('');
    setJustCreatedPin(false);
  }

  if (view) {
    return <StudentLedgerView view={view} justCreatedPin={justCreatedPin} onLogout={handleLogout} />;
  }

  return (
    <div style={styles.page}>
      <div style={styles.card}>
        <div style={styles.headerRow}>
          <div style={styles.iconCircle}>
            <BookOpen size={24} color="#2563eb" />
          </div>
          <div>
            <h1 style={styles.title}>教材費專戶查詢</h1>
            <p style={styles.subtitle}>學生個人教材費專款專用查詢系統</p>
          </div>
        </div>

        <div style={styles.bannerBox}>
          <ShieldCheck size={18} color="#2563eb" style={{ flexShrink: 0 }} />
          <span>本系統採學生專戶專款專用制，您僅能查閱貴子弟之教材扣款明細與專戶餘額。</span>
        </div>

        <form onSubmit={handleLogin} style={styles.form}>
          <label style={styles.label}>
            學生座號
            <input
              style={styles.input}
              type="text"
              inputMode="numeric"
              value={seat}
              onChange={(ev) => setSeat(ev.target.value.replace(/\D/g, ''))}
              placeholder="例如：13"
              autoComplete="off"
            />
          </label>
          <label style={styles.label}>
            PIN 碼 (4 位數字)
            <input
              style={styles.input}
              type="password"
              inputMode="numeric"
              maxLength={4}
              value={pin}
              onChange={(ev) => {
                const val = ev.target.value.replace(/\D/g, '');
                if (val.length <= 4) {
                  setPin(val);
                }
              }}
              placeholder="初次登入將自動建立 (請牢記)"
              autoComplete="off"
            />
          </label>

          {error && (
            <div style={styles.errorBox}>
              <AlertCircle size={16} style={{ flexShrink: 0 }} />
              <span>{error}</span>
            </div>
          )}

          <button type="submit" style={styles.submitBtn} disabled={loading}>
            {loading ? <Loader2 size={18} className="spin" /> : <Lock size={18} />}
            {loading ? '驗證查詢中…' : '進入專戶存摺'}
          </button>
        </form>
      </div>
    </div>
  );
}

function StudentLedgerView({ view, justCreatedPin, onLogout }) {
  const [activeTab, setActiveTab] = useState('expenses');
  const payment = view.myChild?.payment || {};
  const studentName = view.myChild?.name || '學生';
  const studentSeat = view.myChild?.seat || '';

  const amountPaid = payment.amountPaid || 0;
  const amountSpent = payment.amountSpent || 0;
  const balance = view.balance !== undefined ? view.balance : (amountPaid - amountSpent);
  const amountDue = payment.amountDue || 0;

  const expenses = view.expenses || [];
  const incomes = view.incomes || [];

  const isOverdrawn = balance < 0;

  return (
    <div style={styles.page}>
      <div style={styles.card}>
        {justCreatedPin && (
          <div style={styles.noticeBox}>
            <ShieldCheck size={18} color="#16a34a" />
            <span>PIN 碼設定成功！請妥善保存您的 4 位數 PIN 碼，日後查詢皆需使用。</span>
          </div>
        )}

        {/* 頂部學生身份與登出 */}
        <div style={styles.profileHeader}>
          <div>
            <div style={styles.seatTag}>座號 {studentSeat} 號</div>
            <h1 style={styles.studentName}>{studentName} 的教材費存摺</h1>
          </div>
          <button type="button" onClick={onLogout} style={styles.logoutBtn}>
            登出
          </button>
        </div>

        {/* 個人專戶核心資產卡 */}
        <div style={{
          ...styles.balanceHeroCard,
          borderColor: isOverdrawn ? '#fca5a5' : '#bfdbfe',
          background: isOverdrawn ? 'linear-gradient(135deg, #fff1f2 0%, #fee2e2 100%)' : 'linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%)'
        }}>
          <div style={styles.balanceHeaderRow}>
            <div style={styles.balanceLabelWrap}>
              <Wallet size={18} color={isOverdrawn ? '#dc2626' : '#2563eb'} />
              <span style={{ ...styles.balanceLabel, color: isOverdrawn ? '#991b1b' : '#1e40af' }}>專戶可用餘額</span>
            </div>
            {isOverdrawn ? (
              <span style={styles.badgeDanger}>餘額不足（需補繳 NT$ {Math.abs(balance).toLocaleString()}）</span>
            ) : (
              <span style={styles.badgeSuccess}>專戶餘額充足</span>
            )}
          </div>
          
          <div style={{ ...styles.balanceAmount, color: isOverdrawn ? '#dc2626' : '#1d4ed8' }}>
            NT$ {balance.toLocaleString()}
          </div>

          <div style={styles.balanceSubGrid}>
            <div style={styles.subItem}>
              <div style={styles.subItemLabel}>
                <ArrowDownRight size={14} color="#16a34a" />
                <span>累計預繳 (儲值)</span>
              </div>
              <div style={styles.subItemValGreen}>NT$ {amountPaid.toLocaleString()}</div>
            </div>
            <div style={styles.subItem}>
              <div style={styles.subItemLabel}>
                <ArrowUpRight size={14} color="#dc2626" />
                <span>教材累計扣款</span>
              </div>
              <div style={styles.subItemValRed}>NT$ {amountSpent.toLocaleString()}</div>
            </div>
            {amountDue > 0 && (
              <div style={styles.subItem}>
                <div style={styles.subItemLabel}>
                  <CreditCard size={14} color="#6b7280" />
                  <span>期初應繳標準</span>
                </div>
                <div style={styles.subItemValMuted}>NT$ {amountDue.toLocaleString()}</div>
              </div>
            )}
          </div>
        </div>

        {/* 分頁按鈕：教材扣款明細 vs 繳費歷史 */}
        <div style={styles.tabNav}>
          <button
            type="button"
            style={{ ...styles.tabItem, ...(activeTab === 'expenses' ? styles.tabItemActive : {}) }}
            onClick={() => setActiveTab('expenses')}
          >
            <Receipt size={16} />
            教材扣款明細 ({expenses.length})
          </button>
          <button
            type="button"
            style={{ ...styles.tabItem, ...(activeTab === 'incomes' ? styles.tabItemActive : {}) }}
            onClick={() => setActiveTab('incomes')}
          >
            <CreditCard size={16} />
            繳費紀錄 ({incomes.length})
          </button>
        </div>

        {/* 明細清單 */}
        {activeTab === 'expenses' ? (
          <div>
            {expenses.length === 0 ? (
              <div style={styles.emptyState}>
                <Receipt size={32} color="#9ca3af" />
                <p>目前尚無教材扣款紀錄</p>
              </div>
            ) : (
              <div style={styles.listContainer}>
                {expenses.map((exp) => (
                  <div key={exp.id || exp.item + exp.date} style={styles.ledgerRow}>
                    <div style={styles.ledgerLeft}>
                      <div style={styles.ledgerTitleRow}>
                        <span style={styles.categoryBadge}>{exp.category || '教材'}</span>
                        <strong style={styles.itemTitle}>{exp.item}</strong>
                      </div>
                      <div style={styles.metaRow}>
                        <span>{exp.date}</span>
                        {exp.term && <span>· {exp.term}</span>}
                        {exp.note && <span style={styles.noteText}>· {exp.note}</span>}
                      </div>
                    </div>
                    <div style={styles.expenseAmount}>
                      - NT$ {Number(exp.amount || 0).toLocaleString()}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        ) : (
          <div>
            {incomes.length === 0 ? (
              <div style={styles.emptyState}>
                <CreditCard size={32} color="#9ca3af" />
                <p>目前尚無繳費儲值紀錄</p>
              </div>
            ) : (
              <div style={styles.listContainer}>
                {incomes.map((inc) => (
                  <div key={inc.id || inc.date} style={styles.ledgerRow}>
                    <div style={styles.ledgerLeft}>
                      <div style={styles.ledgerTitleRow}>
                        <span style={styles.incomeBadge}>繳費儲值</span>
                        <strong style={styles.itemTitle}>{inc.term || '本學期'} 教材預繳費</strong>
                      </div>
                      <div style={styles.metaRow}>
                        <span>{inc.date}</span>
                        {inc.note && <span style={styles.noteText}>· {inc.note}</span>}
                      </div>
                    </div>
                    <div style={styles.incomeAmount}>
                      + NT$ {Number(inc.amount || 0).toLocaleString()}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        <div style={styles.footerNote}>
          <ShieldCheck size={14} color="#6b7280" />
          <span>本存摺紀錄僅供家長查閱個別子女專戶明細，若對扣款品項有疑問請洽導師。</span>
        </div>
      </div>
    </div>
  );
}

const styles = {
  page: {
    minHeight: '100vh',
    background: '#f8fafc',
    padding: '2rem 1rem',
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'flex-start',
    fontFamily: '"Noto Sans TC", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
    color: '#1e293b'
  },
  card: {
    background: '#ffffff',
    borderRadius: '1.25rem',
    border: '1px solid #e2e8f0',
    boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.05), 0 8px 10px -6px rgba(0, 0, 0, 0.05)',
    width: '100%',
    maxWidth: '560px',
    padding: '2rem',
    boxSizing: 'border-box'
  },
  headerRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '1rem',
    marginBottom: '1rem'
  },
  iconCircle: {
    width: '48px',
    height: '48px',
    borderRadius: '12px',
    background: '#eff6ff',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0
  },
  title: {
    fontSize: '1.35rem',
    fontWeight: '800',
    margin: 0,
    color: '#0f172a',
    letterSpacing: '-0.02em'
  },
  subtitle: {
    fontSize: '0.875rem',
    color: '#64748b',
    margin: '0.25rem 0 0 0'
  },
  bannerBox: {
    background: '#f0fdf4',
    border: '1px solid #bbf7d0',
    borderRadius: '0.75rem',
    padding: '0.75rem 1rem',
    fontSize: '0.825rem',
    color: '#166534',
    display: 'flex',
    alignItems: 'center',
    gap: '0.625rem',
    marginBottom: '1.5rem',
    lineHeight: '1.45'
  },
  form: {
    display: 'flex',
    flexDirection: 'column',
    gap: '1.25rem'
  },
  label: {
    display: 'flex',
    flexDirection: 'column',
    gap: '0.4rem',
    fontSize: '0.875rem',
    fontWeight: '600',
    color: '#334155'
  },
  input: {
    padding: '0.75rem 1rem',
    borderRadius: '0.625rem',
    border: '1px solid #cbd5e1',
    fontSize: '1rem',
    outline: 'none',
    transition: 'border-color 0.2s',
    background: '#f8fafc',
    color: '#0f172a'
  },
  errorBox: {
    background: '#fef2f2',
    border: '1px solid #fecaca',
    borderRadius: '0.625rem',
    padding: '0.75rem 1rem',
    color: '#dc2626',
    fontSize: '0.85rem',
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem'
  },
  noticeBox: {
    background: '#f0fdf4',
    border: '1px solid #bbf7d0',
    borderRadius: '0.75rem',
    padding: '0.875rem 1rem',
    color: '#15803d',
    fontSize: '0.85rem',
    display: 'flex',
    alignItems: 'center',
    gap: '0.625rem',
    marginBottom: '1.25rem'
  },
  submitBtn: {
    background: '#2563eb',
    color: '#ffffff',
    border: 'none',
    borderRadius: '0.625rem',
    padding: '0.875rem',
    fontSize: '1rem',
    fontWeight: '700',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '0.5rem',
    transition: 'background 0.2s'
  },
  profileHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    marginBottom: '1.25rem',
    paddingBottom: '1rem',
    borderBottom: '1px solid #f1f5f9'
  },
  seatTag: {
    display: 'inline-block',
    fontSize: '0.75rem',
    fontWeight: '700',
    color: '#2563eb',
    background: '#eff6ff',
    padding: '0.2rem 0.5rem',
    borderRadius: '0.375rem',
    marginBottom: '0.25rem'
  },
  studentName: {
    fontSize: '1.25rem',
    fontWeight: '800',
    margin: 0,
    color: '#0f172a'
  },
  logoutBtn: {
    background: '#f1f5f9',
    border: 'none',
    borderRadius: '0.5rem',
    padding: '0.4rem 0.75rem',
    fontSize: '0.825rem',
    fontWeight: '600',
    color: '#475569',
    cursor: 'pointer'
  },
  balanceHeroCard: {
    border: '1.5px solid',
    borderRadius: '1rem',
    padding: '1.25rem',
    marginBottom: '1.5rem'
  },
  balanceHeaderRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '0.5rem'
  },
  balanceLabelWrap: {
    display: 'flex',
    alignItems: 'center',
    gap: '0.375rem'
  },
  balanceLabel: {
    fontSize: '0.875rem',
    fontWeight: '700'
  },
  badgeSuccess: {
    fontSize: '0.75rem',
    fontWeight: '700',
    color: '#15803d',
    background: '#dcfce7',
    padding: '0.2rem 0.5rem',
    borderRadius: '0.375rem'
  },
  badgeDanger: {
    fontSize: '0.75rem',
    fontWeight: '700',
    color: '#b91c1c',
    background: '#fee2e2',
    padding: '0.2rem 0.5rem',
    borderRadius: '0.375rem'
  },
  balanceAmount: {
    fontSize: '2rem',
    fontWeight: '900',
    fontFamily: '"Space Mono", monospace',
    letterSpacing: '-0.03em',
    marginBottom: '1rem'
  },
  balanceSubGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))',
    gap: '0.75rem',
    borderTop: '1px solid rgba(0, 0, 0, 0.06)',
    paddingTop: '0.875rem'
  },
  subItem: {
    display: 'flex',
    flexDirection: 'column',
    gap: '0.25rem'
  },
  subItemLabel: {
    display: 'flex',
    alignItems: 'center',
    gap: '0.25rem',
    fontSize: '0.75rem',
    color: '#64748b'
  },
  subItemValGreen: {
    fontSize: '1rem',
    fontWeight: '700',
    color: '#16a34a',
    fontFamily: '"Space Mono", monospace'
  },
  subItemValRed: {
    fontSize: '1rem',
    fontWeight: '700',
    color: '#dc2626',
    fontFamily: '"Space Mono", monospace'
  },
  subItemValMuted: {
    fontSize: '1rem',
    fontWeight: '700',
    color: '#475569',
    fontFamily: '"Space Mono", monospace'
  },
  tabNav: {
    display: 'flex',
    gap: '0.5rem',
    borderBottom: '2px solid #f1f5f9',
    marginBottom: '1rem'
  },
  tabItem: {
    background: 'none',
    border: 'none',
    borderBottom: '2px solid transparent',
    marginBottom: '-2px',
    padding: '0.625rem 0.75rem',
    fontSize: '0.875rem',
    fontWeight: '600',
    color: '#64748b',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: '0.375rem'
  },
  tabItemActive: {
    color: '#2563eb',
    borderBottomColor: '#2563eb',
    fontWeight: '700'
  },
  listContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '0.75rem',
    maxHeight: '400px',
    overflowY: 'auto'
  },
  ledgerRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '0.75rem 1rem',
    background: '#f8fafc',
    borderRadius: '0.75rem',
    border: '1px solid #f1f5f9'
  },
  ledgerLeft: {
    display: 'flex',
    flexDirection: 'column',
    gap: '0.25rem'
  },
  ledgerTitleRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem'
  },
  categoryBadge: {
    fontSize: '0.7rem',
    fontWeight: '700',
    color: '#4338ca',
    background: '#e0e7ff',
    padding: '0.15rem 0.4rem',
    borderRadius: '0.25rem'
  },
  incomeBadge: {
    fontSize: '0.7rem',
    fontWeight: '700',
    color: '#15803d',
    background: '#dcfce7',
    padding: '0.15rem 0.4rem',
    borderRadius: '0.25rem'
  },
  itemTitle: {
    fontSize: '0.925rem',
    fontWeight: '700',
    color: '#1e293b'
  },
  metaRow: {
    fontSize: '0.75rem',
    color: '#64748b'
  },
  noteText: {
    color: '#94a3b8'
  },
  expenseAmount: {
    fontSize: '1rem',
    fontWeight: '800',
    color: '#dc2626',
    fontFamily: '"Space Mono", monospace'
  },
  incomeAmount: {
    fontSize: '1rem',
    fontWeight: '800',
    color: '#16a34a',
    fontFamily: '"Space Mono", monospace'
  },
  emptyState: {
    textAlign: 'center',
    padding: '2.5rem 1rem',
    color: '#94a3b8',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '0.5rem',
    fontSize: '0.875rem'
  },
  footerNote: {
    marginTop: '1.5rem',
    paddingTop: '1rem',
    borderTop: '1px solid #f1f5f9',
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem',
    fontSize: '0.75rem',
    color: '#94a3b8',
    lineHeight: '1.4'
  }
};
