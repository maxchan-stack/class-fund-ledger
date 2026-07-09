import React, { useState, useEffect } from 'react';
import { Lock, ShieldCheck, AlertCircle, TrendingUp, TrendingDown, Loader2 } from 'lucide-react';

/**
 * ==== 資料型別（JSDoc，供編輯器型別提示用，等同 TypeScript interface） ====
 *
 * @typedef {Object} DerivedPaymentStatus
 * @property {string} seat
 * @property {string} term
 * @property {number} amountDue
 * @property {number} amountPaid
 * @property {"paid"|"partial"|"unpaid"} status
 *
 * @typedef {Object} AnonymizedLedgerEntry
 * @property {string} id
 * @property {"income"|"expense"} type
 * @property {string} date
 * @property {number} amount
 * @property {string} term
 * @property {string} [note]
 * @property {string} [category]
 * @property {string} [item]
 *
 * @typedef {Object} ParentPortalView
 * @property {{seat: string, name: string, payment: DerivedPaymentStatus}} myChild
 * @property {AnonymizedLedgerEntry[]} classLedger
 * @property {number} classBalance
 */

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
  if (apiId) return `https://script.google.com/macros/s/${apiId}/exec`;
  return '';
}

const STATUS_LABEL = {
  paid: { text: '已繳清', color: 'var(--green)' },
  partial: { text: '部分繳交', color: 'var(--amber)' },
  unpaid: { text: '尚未繳交', color: 'var(--red)' }
};

export default function ParentPortal() {
  const [sheetUrl, setSheetUrl] = useState('');
  const [seat, setSeat] = useState('');
  const [pin, setPin] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [view, setView] = useState(/** @type {ParentPortalView|null} */ (null));
  const [justCreatedPin, setJustCreatedPin] = useState(false);

  useEffect(() => {
    (async () => {
      const fromUrl = resolveSheetUrlFromLocation();
      if (fromUrl) {
        setSheetUrl(fromUrl);
        return;
      }
      const s = await StorageService.get('settings', true);
      if (s) {
        try {
          const parsed = JSON.parse(s.value);
          if (parsed.sheetUrl) setSheetUrl(parsed.sheetUrl);
        } catch (e) {
          console.error('讀取現有設定失敗', e);
        }
      }
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
    return <PaymentStatusView view={view} justCreatedPin={justCreatedPin} onLogout={handleLogout} />;
  }

  return (
    <div style={styles.page}>
      <div style={styles.card}>
        <div style={styles.headerRow}>
          <Lock size={22} color="var(--primary)" />
          <h1 style={styles.title}>班費查詢</h1>
        </div>
        <p style={styles.subtitle}>輸入您孩子的座號與 4 位數 PIN 碼查看繳費狀態。第一次登入將自動建立 PIN，請務必記住。</p>

        <form onSubmit={handleLogin} style={styles.form}>
          <label style={styles.label}>
            座號
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
              placeholder="第一次登入將自動建立 (4 位數)"
              autoComplete="off"
            />
          </label>

          {error && (
            <div style={styles.errorBox}>
              <AlertCircle size={16} />
              <span>{error}</span>
            </div>
          )}

          <button type="submit" style={styles.submitBtn} disabled={loading}>
            {loading ? <Loader2 size={16} className="spin" /> : <ShieldCheck size={16} />}
            {loading ? '查詢中…' : '查詢繳費狀態'}
          </button>
        </form>
      </div>
    </div>
  );
}

function PaymentStatusView({ view, justCreatedPin, onLogout }) {
  const [ledgerTab, setLedgerTab] = React.useState('income');
  const status = STATUS_LABEL[view.myChild.payment.status] || STATUS_LABEL.unpaid;
  const { amountDue, amountPaid } = view.myChild.payment;

  const incomeEntries = view.classLedger.filter(e => e.type === 'income');
  const expenseEntries = view.classLedger.filter(e => e.type === 'expense');
  const visibleEntries = ledgerTab === 'income' ? incomeEntries : expenseEntries;

  const incomeTotal = incomeEntries.reduce((s, e) => s + e.amount, 0);
  const expenseTotal = expenseEntries.reduce((s, e) => s + e.amount, 0);

  return (
    <div style={styles.page}>
      <div style={styles.card}>
        {justCreatedPin && (
          <div style={styles.noticeBox}>
            <ShieldCheck size={16} />
            <span>PIN 碼已建立，請記住這組 4 位數 PIN 碼，之後查詢會用到。</span>
          </div>
        )}

        <div style={styles.headerRow}>
          <h1 style={styles.title}>{view.myChild.name} 同學（座號 {view.myChild.seat}）</h1>
        </div>

        <div style={{ ...styles.statusCard, borderColor: status.color }}>
          <span style={{ ...styles.statusBadge, background: status.color }}>{status.text}</span>
          <div style={styles.statusRow}>
            <span>本學期應繳</span>
            <strong>{amountDue > 0 ? `NT$ ${amountDue.toLocaleString()}` : '尚未設定'}</strong>
          </div>
          <div style={styles.statusRow}>
            <span>已繳金額</span>
            <strong>NT$ {amountPaid.toLocaleString()}</strong>
          </div>
        </div>

        {/* 全班統計儀表板 */}
        {view.classPaymentStats && (
          <div style={styles.dashboardCard}>
            <div style={styles.dashboardRow}>
              <span>全班繳費率</span>
              <strong style={{ color: 'var(--text)' }}>
                {view.classPaymentStats.paid} / {view.classPaymentStats.total} 人 ({
                  view.classPaymentStats.total > 0 
                    ? Math.round((view.classPaymentStats.paid / view.classPaymentStats.total) * 100) 
                    : 0
                }%)
              </strong>
            </div>
            {/* 進度條 */}
            <div style={styles.progressBarBg}>
              <div style={{
                ...styles.progressBarFill,
                width: `${view.classPaymentStats.total > 0 ? (view.classPaymentStats.paid / view.classPaymentStats.total) * 100 : 0}%`
              }} />
            </div>
            
            {/* 累計收支與餘額迷你字卡 */}
            <div style={styles.miniStatsGrid}>
              <div style={styles.miniStatBox}>
                <span style={styles.miniStatLabel}>累計總收入</span>
                <strong style={{ ...styles.miniStatVal, color: 'var(--green)' }}>
                  NT$ {incomeTotal.toLocaleString()}
                </strong>
              </div>
              <div style={styles.miniStatBox}>
                <span style={styles.miniStatLabel}>累計總支出</span>
                <strong style={{ ...styles.miniStatVal, color: 'var(--red)' }}>
                  NT$ {expenseTotal.toLocaleString()}
                </strong>
              </div>
            </div>
          </div>
        )}

        <h2 style={styles.sectionTitle}>班費收支明細</h2>
        <p style={styles.sectionHint}>僅顯示金額與項目，不顯示個人資訊</p>

        <div style={styles.tabBar}>
          <button
            type="button"
            style={{ ...styles.tabBtn, ...(ledgerTab === 'income' ? styles.tabBtnActive : {}) }}
            onClick={() => setLedgerTab('income')}
          >
            <TrendingUp size={14} />
            收入 · NT$ {incomeTotal.toLocaleString()}
          </button>
          <button
            type="button"
            style={{ ...styles.tabBtn, ...(ledgerTab === 'expense' ? styles.tabBtnActiveExpense : {}) }}
            onClick={() => setLedgerTab('expense')}
          >
            <TrendingDown size={14} />
            支出 · NT$ {expenseTotal.toLocaleString()}
          </button>
        </div>

        <div style={styles.ledgerList}>
          {visibleEntries.length === 0 && (
            <p style={styles.emptyText}>
              {ledgerTab === 'income' ? '目前尚無收入紀錄' : '目前尚無支出紀錄'}
            </p>
          )}
          {visibleEntries.map((entry) => (
            <div key={entry.id} style={styles.ledgerRow}>
              <div style={styles.ledgerIcon}>
                {entry.type === 'income'
                  ? <TrendingUp size={16} color="var(--green)" />
                  : <TrendingDown size={16} color="var(--red)" />}
              </div>
              <div style={styles.ledgerMain}>
                <div style={styles.ledgerTitle}>
                  {entry.type === 'income' ? '班費收入' : (entry.category || entry.item || '支出')}
                </div>
                <div style={styles.ledgerDate}>{entry.date}{entry.note ? ` · ${entry.note}` : ''}</div>
              </div>
              <div style={{ ...styles.ledgerAmount, color: entry.type === 'income' ? 'var(--green)' : 'var(--red)' }}>
                {entry.type === 'income' ? '+' : '−'}NT$ {entry.amount.toLocaleString()}
              </div>
            </div>
          ))}
        </div>

        <div style={styles.balanceRow}>
          <span>班費結餘</span>
          <strong>NT$ {view.classBalance.toLocaleString()}</strong>
        </div>

        <button type="button" style={styles.logoutBtn} onClick={onLogout}>返回查詢頁</button>
      </div>
    </div>
  );
}

const styles = {
  page: {
    minHeight: '100vh',
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'center',
    background: 'var(--bg)',
    padding: '48px 16px',
    fontFamily: 'var(--font-main)',
    color: 'var(--text)'
  },
  card: {
    width: '100%',
    maxWidth: 440,
    background: 'var(--bg-card)',
    borderRadius: 'var(--radius)',
    boxShadow: 'var(--shadow-lg)',
    border: '1px solid var(--border)',
    padding: 28
  },
  headerRow: { display: 'flex', alignItems: 'center', gap: 10, marginBottom: 8 },
  title: { fontSize: 19, fontWeight: 700, color: 'var(--primary)', margin: 0 },
  subtitle: { color: 'var(--text-soft)', fontSize: 14, lineHeight: 1.6, marginBottom: 20 },
  form: { display: 'flex', flexDirection: 'column', gap: 14 },
  label: { display: 'flex', flexDirection: 'column', gap: 6, fontSize: 13, color: 'var(--text-soft)', fontWeight: 600 },
  input: {
    padding: '10px 12px',
    borderRadius: 'var(--radius-sm)',
    border: '1px solid var(--border)',
    fontSize: 15,
    fontFamily: 'var(--font-mono)',
    outline: 'none',
    background: 'var(--bg)'
  },
  errorBox: {
    display: 'flex', alignItems: 'center', gap: 8,
    background: 'rgba(184,50,50,0.08)', color: 'var(--red)',
    padding: '10px 12px', borderRadius: 'var(--radius-sm)', fontSize: 13
  },
  noticeBox: {
    display: 'flex', alignItems: 'center', gap: 8,
    background: 'var(--accent-light)', color: 'var(--accent)',
    padding: '10px 12px', borderRadius: 'var(--radius-sm)', fontSize: 13, marginBottom: 16
  },
  submitBtn: {
    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
    background: 'var(--primary-mid)', color: '#fff', border: 'none',
    borderRadius: 'var(--radius-sm)', padding: '12px 16px', fontSize: 15, fontWeight: 600,
    cursor: 'pointer', marginTop: 4
  },
  statusCard: {
    border: '2px solid var(--border)', borderRadius: 'var(--radius)',
    padding: 18, marginBottom: 24, position: 'relative'
  },
  statusBadge: {
    display: 'inline-block', color: '#fff', fontSize: 12, fontWeight: 700,
    padding: '4px 10px', borderRadius: 999, marginBottom: 12
  },
  statusRow: { display: 'flex', justifyContent: 'space-between', fontSize: 14, padding: '4px 0', color: 'var(--text-soft)' },
  sectionTitle: { fontSize: 15, fontWeight: 700, color: 'var(--text)', margin: '0 0 2px' },
  sectionHint: { fontSize: 12, color: 'var(--text-soft)', margin: '0 0 10px' },
  tabBar: {
    display: 'flex', gap: 8, marginBottom: 12
  },
  tabBtn: {
    flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
    padding: '9px 12px', borderRadius: 'var(--radius-sm)',
    border: '1.5px solid var(--border)', background: 'var(--bg)',
    fontSize: 13, fontWeight: 600, color: 'var(--text-soft)',
    cursor: 'pointer', transition: 'all 0.15s'
  },
  tabBtnActive: {
    borderColor: 'var(--green)', background: 'rgba(45,106,79,0.08)', color: 'var(--green)'
  },
  tabBtnActiveExpense: {
    borderColor: 'var(--red)', background: 'rgba(184,50,50,0.08)', color: 'var(--red)'
  },
  ledgerList: { display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 300, overflowY: 'auto' },
  emptyText: { fontSize: 13, color: 'var(--text-soft)', textAlign: 'center', padding: '20px 0' },
  ledgerRow: {
    display: 'flex', alignItems: 'center', gap: 10,
    padding: '10px 8px', borderBottom: '1px solid var(--border)'
  },
  ledgerIcon: { flexShrink: 0 },
  ledgerMain: { flex: 1, minWidth: 0 },
  ledgerTitle: { fontSize: 13, fontWeight: 600, color: 'var(--text)' },
  ledgerDate: { fontSize: 12, color: 'var(--text-soft)' },
  ledgerAmount: { fontSize: 13, fontWeight: 700, fontFamily: 'var(--font-mono)', whiteSpace: 'nowrap' },
  balanceRow: {
    display: 'flex', justifyContent: 'space-between', alignItems: 'center',
    marginTop: 16, paddingTop: 16, borderTop: '1px solid var(--border)',
    fontSize: 15, fontWeight: 700, color: 'var(--primary)'
  },
  logoutBtn: {
    marginTop: 20, width: '100%', padding: '10px 16px',
    borderRadius: 'var(--radius-sm)', border: '1px solid var(--border)',
    background: 'transparent', color: 'var(--text-soft)', fontSize: 13, cursor: 'pointer'
  },
  dashboardCard: {
    border: '1px solid var(--border)',
    borderRadius: 'var(--radius)',
    padding: 16,
    marginBottom: 20,
    background: 'var(--bg-card)'
  },
  dashboardRow: {
    display: 'flex',
    justifyContent: 'space-between',
    fontSize: 13,
    fontWeight: 600,
    color: 'var(--text-soft)',
    marginBottom: 8
  },
  progressBarBg: {
    width: '100%',
    height: 8,
    borderRadius: 999,
    background: 'var(--border)',
    overflow: 'hidden',
    marginBottom: 16
  },
  progressBarFill: {
    height: '100%',
    borderRadius: 999,
    background: 'var(--green)',
    transition: 'width 0.4s ease-out'
  },
  miniStatsGrid: {
    display: 'flex',
    gap: 10
  },
  miniStatBox: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    gap: 4,
    padding: '8px 10px',
    borderRadius: 'var(--radius-sm)',
    border: '1px solid var(--border)',
    background: 'var(--bg)'
  },
  miniStatLabel: {
    fontSize: 11,
    color: 'var(--text-soft)'
  },
  miniStatVal: {
    fontSize: 13,
    fontWeight: 700,
    fontFamily: 'var(--font-mono)'
  }
};
