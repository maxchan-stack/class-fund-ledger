import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { 
  Stamp, 
  Plus, 
  Trash2, 
  X, 
  TrendingUp, 
  TrendingDown, 
  AlertCircle, 
  RefreshCw, 
  Database 
} from 'lucide-react';

const EXPENSE_CATEGORIES = ['冷氣費', '簿本費', '學科學資', '練習卷／單冊', '班級活動', '清潔費', '其他'];

// 安全相容的儲存服務：相容原 window.storage 與標準 localStorage
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
  },
  async set(key, value, shared = true) {
    if (window.storage && typeof window.storage.set === 'function') {
      try {
        const res = await window.storage.set(key, value, shared);
        if (res) return true;
      } catch (e) {
        console.warn('window.storage.set 失敗，降級使用 localStorage', e);
      }
    }
    try {
      localStorage.setItem(key, value);
      return true;
    } catch (e) {
      console.error('localStorage.set 失敗', e);
      return false;
    }
  },
  async delete(key, shared = true) {
    if (window.storage && typeof window.storage.delete === 'function') {
      try {
        await window.storage.delete(key, shared);
        return;
      } catch (e) {
        console.warn('window.storage.delete 失敗，降級使用 localStorage', e);
      }
    }
    localStorage.removeItem(key);
  }
};

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function money(n) {
  const v = Number(n) || 0;
  return v.toLocaleString('zh-TW');
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

export default function ClassFundLedger() {
  const [loading, setLoading] = useState(true);
  const [transactions, setTransactions] = useState([]);
  const [roster, setRoster] = useState([]); // 學生名冊：{ seat, name }
  const [settings, setSettings] = useState({ className: '', pin: '', sheetUrl: '', terms: [], currentTerm: '' });
  const [teacherMode, setTeacherMode] = useState(false);
  const [deviceRole, setDeviceRole] = useState('loading');
  const [tab, setTab] = useState('income'); // 'income' | 'expense' | 'unpaid'
  const [termFilter, setTermFilter] = useState('all');
  const [error, setError] = useState('');

  // 雲端同步狀態：'idle' | 'syncing' | 'synced' | 'pending_push' | 'error'
  const [syncStatus, setSyncStatus] = useState('idle');
  const [syncConflictModal, setSyncConflictModal] = useState(null); // 'conflict' | null
  const [cloudDataTemp, setCloudDataTemp] = useState(null);

  const tapCountRef = useRef(0);
  const tapTimerRef = useRef(null);

  const [modal, setModal] = useState(null); // 'setup' | 'unlock' | null
  const [pinInput, setPinInput] = useState('');
  const [pinConfirm, setPinConfirm] = useState('');
  const [pinError, setPinError] = useState('');
  const [stamping, setStamping] = useState(false);

  const [showForm, setShowForm] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [editingClassName, setEditingClassName] = useState(false);
  const [classNameInput, setClassNameInput] = useState('');
  const [editingBackup, setEditingBackup] = useState(false);
  const [sheetUrlInput, setSheetUrlInput] = useState('');
  const [backupHelp, setBackupHelp] = useState(false);
  const [editingTerms, setEditingTerms] = useState(false);
  const [newTermInput, setNewTermInput] = useState('');
  
  // 學生名冊編輯狀態
  const [newStudentSeat, setNewStudentSeat] = useState('');
  const [newStudentName, setNewStudentName] = useState('');

  // 智慧座號輸入模式：'select' (從名冊選) | 'manual' (手動輸入)
  const [seatInputMode, setSeatInputMode] = useState('select');

  const blankIncome = { date: todayStr(), source: '', seat: '', amount: '', term: '', note: '' };
  const blankExpense = { date: todayStr(), category: EXPENSE_CATEGORIES[0], item: '', unitPrice: '', qty: '1', payee: '', term: '', note: '' };
  const [incomeForm, setIncomeForm] = useState(blankIncome);
  const [expenseForm, setExpenseForm] = useState(blankExpense);

  useEffect(() => {
    loadAll();
  }, []);

  async function loadAll() {
    setLoading(true);
    let loadedSettings = { className: '', pin: '', sheetUrl: '', terms: [], currentTerm: '' };
    let loadedTransactions = [];
    let loadedRoster = [];

    try {
      const s = await StorageService.get('settings', true);
      if (s) {
        loadedSettings = JSON.parse(s.value);
        setSettings(loadedSettings);
        setClassNameInput(loadedSettings.className || '');
        setSheetUrlInput(loadedSettings.sheetUrl || '');
      }
    } catch (e) {
      console.error('載入 settings 失敗', e);
    }

    try {
      const t = await StorageService.get('ledger', true);
      if (t) {
        loadedTransactions = JSON.parse(t.value);
        setTransactions(loadedTransactions);
      }
    } catch (e) {
      console.error('載入 ledger 失敗', e);
    }

    try {
      const r = await StorageService.get('roster', true);
      if (r) {
        loadedRoster = JSON.parse(r.value);
        setRoster(loadedRoster);
      }
    } catch (e) {
      console.error('載入 roster 失敗', e);
    }

    try {
      const d = await StorageService.get('device-role', false); // 個人層級，不共享
      setDeviceRole(d && d.value === 'teacher' ? 'teacher' : 'viewer');
    } catch {
      setDeviceRole('viewer');
    }

    setLoading(false);

    // 如果有設定試算表網址，載入時自動執行一次同步（拉取雲端）
    if (loadedSettings.sheetUrl) {
      autoPullData(loadedSettings.sheetUrl, loadedTransactions, loadedRoster, loadedSettings);
    }
  }

  // 雲端雙向同步：拉取資料
  const autoPullData = async (url, localTrans, localRoster, localSettings) => {
    setSyncStatus('syncing');
    try {
      const res = await fetch(url, { method: 'GET', mode: 'cors' });
      if (!res.ok) throw new Error('Network response was not ok');
      const result = await res.json();
      
      if (result.success && result.data) {
        const cloudTrans = result.data.transactions || [];
        const cloudRoster = result.data.roster || [];
        const cloudSettings = result.data.settings || {};

        // 如果雲端有資料，且本地與雲端不一致
        const hasDiff = 
          JSON.stringify(cloudTrans) !== JSON.stringify(localTrans) ||
          JSON.stringify(cloudRoster) !== JSON.stringify(localRoster) ||
          cloudSettings.className !== localSettings.className ||
          cloudSettings.currentTerm !== localSettings.currentTerm;

        if (hasDiff) {
          // 如果本地沒有任何交易與名冊資料，直接套用雲端資料
          if (localTrans.length === 0 && localRoster.length === 0) {
            await applyCloudData(cloudTrans, cloudRoster, { ...localSettings, ...cloudSettings });
            setSyncStatus('synced');
          } else {
            // 本地與雲端皆有資料且不一致，記錄雲端資料並提示衝突 Modal
            setCloudDataTemp({ transactions: cloudTrans, roster: cloudRoster, settings: cloudSettings });
            setSyncConflictModal('conflict');
            setSyncStatus('pending_push');
          }
        } else {
          setSyncStatus('synced');
        }
      } else {
        setSyncStatus('error');
      }
    } catch (err) {
      console.error('自動拉取雲端資料失敗：', err);
      setSyncStatus('error');
    }
  };

  // 套用雲端資料到本地
  async function applyCloudData(cloudTrans, cloudRoster, mergedSettings) {
    setTransactions(cloudTrans);
    setRoster(cloudRoster);
    setSettings(mergedSettings);
    setClassNameInput(mergedSettings.className || '');
    
    await StorageService.set('ledger', JSON.stringify(cloudTrans), true);
    await StorageService.set('roster', JSON.stringify(cloudRoster), true);
    await StorageService.set('settings', JSON.stringify(mergedSettings), true);
  }

  // 手動觸發雲端同步（合併與推送）
  async function handleManualSync() {
    if (!settings.sheetUrl) {
      setError('請先在設定中填寫試算表網址');
      return;
    }
    setSyncStatus('syncing');
    setError('');
    
    try {
      // 1. 先嘗試拉取最新資料
      const res = await fetch(settings.sheetUrl, { method: 'GET', mode: 'cors' });
      if (!res.ok) throw new Error('Fetch failed');
      const result = await res.json();
      
      if (result.success && result.data) {
        const cloudTrans = result.data.transactions || [];
        const cloudRoster = result.data.roster || [];
        const cloudSettings = result.data.settings || {};
        
        // 判斷是否有一致性衝突
        const hasDiff = 
          JSON.stringify(cloudTrans) !== JSON.stringify(transactions) ||
          JSON.stringify(cloudRoster) !== JSON.stringify(roster) ||
          cloudSettings.className !== settings.className ||
          cloudSettings.currentTerm !== settings.currentTerm;
          
        if (hasDiff) {
          // 儲存雲端資料，彈出衝突詢問視窗
          setCloudDataTemp({ transactions: cloudTrans, roster: cloudRoster, settings: cloudSettings });
          setSyncConflictModal('conflict');
          setSyncStatus('pending_push');
        } else {
          // 資料已是一致，直接標記同步完成
          setSyncStatus('synced');
        }
      } else {
        // 如果雲端尚未初始化或為空，直接將本地推送上去
        await pushLocalToCloud(settings.sheetUrl, transactions, roster, settings);
      }
    } catch (err) {
      console.error('同步讀取失敗，嘗試直接推送到雲端：', err);
      // 網路或 CORS 限制，嘗試直接寫入
      await pushLocalToCloud(settings.sheetUrl, transactions, roster, settings);
    }
  }

  // 將本地資料強行推送到試算表
  async function pushLocalToCloud(url, trans, rost, sett) {
    setSyncStatus('syncing');
    try {
      // 排除安全密碼 pin，避免明文上傳
      const settingsToPush = { ...sett };
      delete settingsToPush.pin;

      const response = await fetch(url, {
        method: 'POST',
        mode: 'cors',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          action: 'sync',
          transactions: trans,
          roster: rost,
          settings: settingsToPush
        })
      });
      
      if (!response.ok) throw new Error('Post failed');
      const resJson = await response.json();
      
      if (resJson.success) {
        setSyncStatus('synced');
        setSyncConflictModal(null);
      } else {
        setSyncStatus('error');
        setError('雲端寫入失敗：' + (resJson.error || '未知錯誤'));
      }
    } catch (err) {
      console.error('推送雲端失敗：', err);
      // 降級處理：若為 CORS 限制，採用 no-cors 盡力而為模式
      try {
        await fetch(url, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain' },
          body: JSON.stringify({
            action: 'sync',
            transactions: trans,
            roster: rost,
            settings: sett
          })
        });
        setSyncStatus('synced'); // 視作成功但無法驗證
        setSyncConflictModal(null);
      } catch (e) {
        setSyncStatus('error');
        setError('同步連線失敗，請檢查網路狀態');
      }
    }
  }

  // 隱藏手勢：快速點擊班級名稱 5 下才會叫出教師密碼視窗
  function handleSecretTap() {
    tapCountRef.current += 1;
    if (tapTimerRef.current) clearTimeout(tapTimerRef.current);
    tapTimerRef.current = setTimeout(() => { tapCountRef.current = 0; }, 2500);
    if (tapCountRef.current >= 5) {
      tapCountRef.current = 0;
      openStamp();
    }
  }

  async function forgetDevice() {
    setTeacherMode(false);
    setShowForm(false);
    setDeviceRole('viewer');
    try {
      await StorageService.delete('device-role', false);
    } catch {
      /* ignore */
    }
  }

  const saveTransactions = useCallback(async (next) => {
    setTransactions(next);
    try {
      const res = await StorageService.set('ledger', JSON.stringify(next), true);
      if (!res) setError('儲存失敗，請檢查儲存空間');
      else {
        setSyncStatus('pending_push');
        // 異步嘗試推送到雲端
        if (settings.sheetUrl) {
          pushLocalToCloud(settings.sheetUrl, next, roster, settings);
        }
      }
    } catch {
      setError('儲存失敗，請檢查儲存空間');
    }
  }, [roster, settings]);

  const saveSettings = useCallback(async (next) => {
    setSettings(next);
    try {
      const res = await StorageService.set('settings', JSON.stringify(next), true);
      if (!res) setError('儲存失敗，請檢查儲存空間');
      else {
        setSyncStatus('pending_push');
        if (next.sheetUrl) {
          pushLocalToCloud(next.sheetUrl, transactions, roster, next);
        }
      }
    } catch {
      setError('儲存失敗，請檢查儲存空間');
    }
  }, [transactions, roster]);

  const saveRoster = useCallback(async (next) => {
    setRoster(next);
    try {
      const res = await StorageService.set('roster', JSON.stringify(next), true);
      if (!res) setError('儲存失敗，請檢查儲存空間');
      else {
        setSyncStatus('pending_push');
        if (settings.sheetUrl) {
          pushLocalToCloud(settings.sheetUrl, transactions, next, settings);
        }
      }
    } catch {
      setError('儲存失敗，請檢查儲存空間');
    }
  }, [transactions, settings]);

  const totals = useMemo(() => {
    let income = 0, expense = 0;
    for (const t of transactions) {
      if (t.type === 'income') income += Number(t.amount) || 0;
      else expense += Number(t.amount) || 0;
    }
    return { income, expense, balance: income - expense };
  }, [transactions]);

  const categoryBreakdown = useMemo(() => {
    const map = {};
    for (const t of transactions) {
      if (t.type !== 'expense') continue;
      map[t.category] = (map[t.category] || 0) + (Number(t.amount) || 0);
    }
    const arr = Object.entries(map).map(([category, amount]) => ({ category, amount }));
    arr.sort((a, b) => b.amount - a.amount);
    return arr.slice(0, 6);
  }, [transactions]);

  const termsList = useMemo(() => {
    const set = new Set(settings.terms || []);
    transactions.forEach((t) => { if (t.term) set.add(t.term); });
    return Array.from(set).sort();
  }, [settings.terms, transactions]);

  const filteredTotals = useMemo(() => {
    if (termFilter === 'all') return null;
    let income = 0, expense = 0;
    for (const t of transactions) {
      if (t.term !== termFilter) continue;
      if (t.type === 'income') income += Number(t.amount) || 0;
      else expense += Number(t.amount) || 0;
    }
    return { income, expense };
  }, [transactions, termFilter]);

  const unpaidInfo = useMemo(() => {
    if (termFilter === 'all') return null;
    const paidSeats = new Set(
      transactions
        .filter((t) => t.type === 'income' && t.term === termFilter && String(t.seat || '').trim())
        .map((t) => String(t.seat).trim())
    );
    const sortedRoster = roster.slice().sort((a, b) => Number(a.seat) - Number(b.seat));
    const unpaid = sortedRoster.filter((s) => !paidSeats.has(String(s.seat).trim()));
    const paidCount = sortedRoster.length - unpaid.length;
    return { unpaid, paidCount, total: sortedRoster.length };
  }, [transactions, roster, termFilter]);

  const sortedList = useMemo(() => {
    return transactions
      .filter((t) => t.type === tab)
      .filter((t) => termFilter === 'all' || t.term === termFilter)
      .slice()
      .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  }, [transactions, tab, termFilter]);

  function openStamp() {
    if (teacherMode) {
      setTeacherMode(false);
      setShowForm(false);
      return;
    }
    setPinInput('');
    setPinConfirm('');
    setPinError('');
    setModal(settings.pin ? 'unlock' : 'setup');
  }

  async function submitSetup() {
    if (pinInput.length < 4) { setPinError('密碼至少 4 碼'); return; }
    if (pinInput !== pinConfirm) { setPinError('兩次輸入不一致'); return; }
    await saveSettings({ ...settings, pin: pinInput });
    doStamp();
  }

  function submitUnlock() {
    if (pinInput !== settings.pin) { setPinError('密碼錯誤'); return; }
    doStamp();
  }

  async function doStamp() {
    setStamping(true);
    setModal(null);
    try { 
      await StorageService.set('device-role', 'teacher', false); 
    } catch { 
      /* ignore */ 
    }
    setDeviceRole('teacher');
    setTimeout(() => {
      setTeacherMode(true);
      setStamping(false);
    }, 380);
  }

  // 智慧座號選擇器變更處理
  const handleRosterSeatSelect = (e) => {
    const selectedSeat = e.target.value;
    if (selectedSeat === '__manual__') {
      setSeatInputMode('manual');
      setIncomeForm({ ...incomeForm, seat: '' });
      return;
    }
    
    const student = roster.find(s => String(s.seat) === String(selectedSeat));
    const termLabel = incomeForm.term || settings.currentTerm || '';
    const autoSource = student ? `${termLabel} 班費 - ${student.name}` : incomeForm.source;
    
    setIncomeForm({
      ...incomeForm,
      seat: selectedSeat,
      source: incomeForm.source ? incomeForm.source : autoSource
    });
  };

  async function addIncome(e) {
    e.preventDefault();
    if (!incomeForm.source || !incomeForm.amount) return;
    const entry = {
      id: uid(), 
      type: 'income', 
      date: incomeForm.date,
      source: incomeForm.source, 
      seat: incomeForm.seat,
      amount: Number(incomeForm.amount),
      term: incomeForm.term || settings.currentTerm || '',
      note: incomeForm.note || '',
    };
    await saveTransactions([...transactions, entry]);
    setIncomeForm({ ...blankIncome, term: settings.currentTerm || '' });
    setSeatInputMode('select'); // 重置輸入模式
    setShowForm(false);
  }

  async function addExpense(e) {
    e.preventDefault();
    if (!expenseForm.item || !expenseForm.unitPrice) return;
    const amount = (Number(expenseForm.unitPrice) || 0) * (Number(expenseForm.qty) || 1);
    const entry = {
      id: uid(), 
      type: 'expense', 
      date: expenseForm.date,
      category: expenseForm.category, 
      item: expenseForm.item,
      unitPrice: Number(expenseForm.unitPrice), 
      qty: Number(expenseForm.qty) || 1,
      payee: expenseForm.payee, 
      amount,
      term: expenseForm.term || settings.currentTerm || '',
      note: expenseForm.note || '',
    };
    await saveTransactions([...transactions, entry]);
    setExpenseForm({ ...blankExpense, term: settings.currentTerm || '' });
    setShowForm(false);
  }

  async function doDelete(id) {
    await saveTransactions(transactions.filter((t) => t.id !== id));
    setConfirmDelete(null);
  }

  async function saveClassName() {
    await saveSettings({ ...settings, className: classNameInput });
    setEditingClassName(false);
  }

  async function saveBackupUrl() {
    const formattedUrl = sheetUrlInput.trim();
    await saveSettings({ ...settings, sheetUrl: formattedUrl });
    setEditingBackup(false);
    if (formattedUrl) {
      autoPullData(formattedUrl, transactions, roster, settings);
    }
  }

  async function addTerm() {
    const label = newTermInput.trim();
    if (!label) return;
    const terms = Array.from(new Set([...(settings.terms || []), label]));
    await saveSettings({ ...settings, terms, currentTerm: label });
    setNewTermInput('');
    setTermFilter(label);
  }

  async function setCurrentTerm(label) {
    await saveSettings({ ...settings, currentTerm: label });
  }

  async function addStudent(e) {
    e.preventDefault();
    const seat = newStudentSeat.trim();
    const name = newStudentName.trim();
    if (!seat || !name) return;
    if (roster.some((s) => String(s.seat).trim() === seat)) { 
      setError('這個座號已經在名冊裡了'); 
      return; 
    }
    const next = [...roster, { seat, name }];
    await saveRoster(next);
    setNewStudentSeat('');
    setNewStudentName('');
  }

  async function deleteStudent(seat) {
    await saveRoster(roster.filter((s) => String(s.seat).trim() !== String(seat).trim()));
  }

  const maxCat = Math.max(1, ...categoryBreakdown.map((c) => c.amount));
  const activeRosterSorted = useMemo(() => {
    return roster.slice().sort((a, b) => Number(a.seat) - Number(b.seat));
  }, [roster]);

  return (
    <div className="cfl-root">
      {/* 頂部 Sticky Header (參考提案平台的樣式與行為) */}
      <header>
        <div className="header-inner">
          <div className="brand" onClick={() => window.location.reload()}>
            班費紀錄系統
            <small>{settings.className ? `${settings.className} ｜ ` : ''}即時同步帳本 ｜ 智慧未繳比對</small>
            
            {/* 雲端同步狀態標記放在 Header 左側 */}
            {settings.sheetUrl && (
              <div className="cfl-sync-status-row">
                <span className={`cfl-sync-badge ${
                  syncStatus === 'synced' ? 'synced' : 
                  syncStatus === 'syncing' ? 'pending' : 
                  syncStatus === 'pending_push' ? 'pending' : 'error'
                }`} />
                <span>
                  {syncStatus === 'synced' && '雲端已同步'}
                  {syncStatus === 'syncing' && '雲端同步中...'}
                  {syncStatus === 'pending_push' && '本地有未儲存變更'}
                  {syncStatus === 'error' && '同步失敗，點選重試'}
                  {syncStatus === 'idle' && '已連接試算表'}
                </span>
                <button 
                  onClick={handleManualSync} 
                  style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', display: 'flex', alignItems: 'center', color: '#fff', marginLeft: 4 }}
                  title="立即手動雙向同步"
                  disabled={syncStatus === 'syncing'}
                >
                  <RefreshCw size={11} className={syncStatus === 'syncing' ? 'spin-icon animate-spin' : ''} style={{ animation: syncStatus === 'syncing' ? 'spin 1s linear infinite' : 'none' }} />
                </button>
              </div>
            )}
          </div>
          <nav>
            <button className={tab === 'income' ? 'active' : ''} onClick={() => { setTab('income'); setShowForm(false); }}>收入紀錄</button>
            <button className={tab === 'expense' ? 'active' : ''} onClick={() => { setTab('expense'); setShowForm(false); }}>支出紀錄</button>
            <button className={tab === 'unpaid' ? 'active' : ''} onClick={() => { setTab('unpaid'); setShowForm(false); }}>未繳費名單</button>
          </nav>
        </div>
      </header>

      <div className="cfl-wrap">
        {error && (
          <div className="cfl-error">
            <AlertCircle size={15} />
            <span>{error}</span>
            <button className="cfl-row-del" style={{ marginLeft: 'auto', color: 'inherit' }} onClick={() => setError('')}><X size={14} /></button>
          </div>
        )}

        <div className="cfl-cover">
          <div className="cfl-cover-top">
            <div>
              {editingClassName ? (
                <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                  <input
                    className="cfl-classname-input"
                    value={classNameInput}
                    onChange={(e) => setClassNameInput(e.target.value)}
                    placeholder="例：209 班"
                    autoFocus
                  />
                  <button className="cfl-classname-edit-btn" onClick={saveClassName}>儲存</button>
                </div>
              ) : (
                <div className="cfl-classname" onClick={deviceRole !== 'teacher' ? handleSecretTap : undefined}>
                  <span>{settings.className || '點此設定班級名稱'}</span>
                  {teacherMode && (
                    <>
                      <button className="cfl-classname-edit-btn" style={{ marginLeft: 8 }} onClick={() => setEditingClassName(true)}>編輯名稱</button>
                      <button className="cfl-classname-edit-btn" style={{ marginLeft: 8 }} onClick={() => setEditingBackup(true)}>
                        {settings.sheetUrl ? '已連接雲端' : '連接試算表'}
                      </button>
                    </>
                  )}
                </div>
              )}

              {teacherMode && editingBackup && (
                <div className="cfl-backup-panel" style={{ marginTop: 12 }}>
                  <div className="cfl-backup-row">
                    <input
                      className="cfl-classname-input"
                      style={{ flex: 1 }}
                      placeholder="貼上 Apps Script 網址（.../exec）"
                      value={sheetUrlInput}
                      onChange={(e) => setSheetUrlInput(e.target.value)}
                    />
                    <button className="cfl-classname-edit-btn" onClick={saveBackupUrl}>儲存</button>
                    <button className="cfl-classname-edit-btn" onClick={() => setEditingBackup(false)}>關閉</button>
                  </div>
                  <div className="cfl-backup-help-toggle" onClick={() => setBackupHelp((v) => !v)}>
                    {backupHelp ? '隱藏說明 ▲' : '如何設定同步網址？ ▼'}
                  </div>
                  {backupHelp && (
                    <div className="cfl-backup-help">
                      在 Google 試算表中，點選「擴充功能 → Apps Script」，建立一個新檔案並貼上本專案根目錄的 `google-apps-script.js` 程式碼。<br />
                      接著部署為「網頁應用程式」，將「執行身分」選為自己、「存取權限」設為「所有人」，最後將產生的 API 網址複製並貼在上方即可。
                    </div>
                  )}
                </div>
              )}
            </div>
            
            {deviceRole === 'teacher' && (
              <div className="cfl-stamp-btn-container">
                <button
                  className={`cfl-stamp-btn ${teacherMode ? 'active' : ''} ${stamping ? 'stamping' : ''}`}
                  onClick={openStamp}
                  title={teacherMode ? '點一下鎖回檢視模式' : '蓋章解鎖教師模式'}
                >
                  <Stamp size={22} />
                </button>
                <div className="cfl-stamp-label">{teacherMode ? '教師模式' : '檢視模式'}</div>
              </div>
            )}
          </div>

          <div className="cfl-balance-row">
            <div className="cfl-balance-label">班費總餘額</div>
            <div className={`cfl-balance-value cfl-mono ${totals.balance < 0 ? 'neg' : ''}`}>
              NT$ {money(totals.balance)}
            </div>
          </div>

          <div className="cfl-stat-row">
            <div className="cfl-stat">
              <div className="cfl-stat-top"><TrendingUp size={13} color="var(--green)" />收入總額</div>
              <div className="cfl-stat-num income cfl-mono">{money(totals.income)}</div>
            </div>
            <div className="cfl-stat">
              <div className="cfl-stat-top"><TrendingDown size={13} color="var(--red)" />支出總額</div>
              <div className="cfl-stat-num expense cfl-mono">{money(totals.expense)}</div>
            </div>
          </div>
        </div>

        {deviceRole === 'teacher' ? (
          <div className="cfl-note">
            {teacherMode
              ? <>目前為教師模式，可新增或刪除紀錄。再按一次印章即可鎖回檢視模式。<br /><span className="cfl-forget" onClick={forgetDevice}>不是自己的裝置？點此忘記此裝置</span></>
              : '目前為檢視模式，按印章即可解鎖編輯。'}
          </div>
        ) : (
          <div className="cfl-note">資料即時同步，內容僅供查看。</div>
        )}

        <div className="cfl-term-row">
          <label>學期</label>
          <select value={termFilter} onChange={(e) => setTermFilter(e.target.value)}>
            <option value="all">全部</option>
            {termsList.map((tm) => <option key={tm} value={tm}>{tm}</option>)}
          </select>
          {termFilter !== 'all' && filteredTotals && (
            <span className="cfl-term-subtotal cfl-mono">
              本學期收入 {money(filteredTotals.income)}｜支出 {money(filteredTotals.expense)}
            </span>
          )}
          {teacherMode && (
            <button className="cfl-classname-edit-btn" style={{ marginLeft: 'auto' }} onClick={() => setEditingTerms((v) => !v)}>管理學期</button>
          )}
        </div>

        {teacherMode && editingTerms && (
          <div className="cfl-backup-panel" style={{ marginBottom: 12 }}>
            <div className="cfl-backup-row">
              <input
                className="cfl-classname-input"
                style={{ flex: 1 }}
                placeholder="新增學期，例：113-2"
                value={newTermInput}
                onChange={(e) => setNewTermInput(e.target.value)}
              />
              <button className="cfl-classname-edit-btn" onClick={addTerm}>新增並設為目前學期</button>
            </div>
            {termsList.length > 0 && (
              <div className="cfl-term-chip-row">
                目前學期：
                {termsList.map((tm) => (
                  <span
                    key={tm}
                    className={`cfl-term-chip ${settings.currentTerm === tm ? 'active' : ''}`}
                    onClick={() => setCurrentTerm(tm)}
                  >{tm}</span>
                ))}
              </div>
            )}
            <div className="cfl-backup-help">新增紀錄時會預設帶入目前學期，您仍可以在表單裡改成別的學期。班費總餘額永遠是累計金額，學期只是用來篩選查看、跟未繳費名單比對，不會把餘額歸零重算。</div>
          </div>
        )}

        {tab === 'unpaid' ? (
          <div className="cfl-table-wrap">
            {termFilter === 'all' ? (
              <div className="cfl-empty">請先在上方選擇一個學期，才能比對未繳費名單。</div>
            ) : roster.length === 0 ? (
              <div className="cfl-empty">名冊還是空的，{teacherMode ? '請在下方新增學生。' : '請導師先建立學生名冊。'}</div>
            ) : (
              <>
                <div className="cfl-unpaid-summary">
                  {termFilter} 已繳 {unpaidInfo.paidCount}／{unpaidInfo.total} 人
                </div>
                {unpaidInfo.unpaid.length === 0 ? (
                  <div className="cfl-empty">全班都繳齊了 🎉</div>
                ) : (
                  unpaidInfo.unpaid.map((s) => (
                    <div className="cfl-row" key={s.seat} style={{ gridTemplateColumns: '78px 1fr 26px' }}>
                      <div className="cfl-row-date cfl-mono">座號 {s.seat}</div>
                      <div className="cfl-row-main"><div className="cfl-row-title">{s.name}</div></div>
                      {teacherMode && <button className="cfl-row-del" onClick={() => deleteStudent(s.seat)}><Trash2 size={15} /></button>}
                    </div>
                  ))
                )}
              </>
            )}

            {teacherMode && (
              <div className="cfl-addbar">
                <form className="cfl-form" onSubmit={addStudent} style={{ marginTop: 0 }}>
                  <div className="cfl-form-grid">
                    <div className="cfl-field"><label>座號</label><input type="number" min="1" placeholder="例：5" value={newStudentSeat} onChange={(e) => setNewStudentSeat(e.target.value)} required /></div>
                    <div className="cfl-field"><label>姓名</label><input placeholder="例：王小明" value={newStudentName} onChange={(e) => setNewStudentName(e.target.value)} required /></div>
                  </div>
                  <div className="cfl-form-actions">
                    <button type="submit" className="cfl-btn-primary">加入名冊</button>
                  </div>
                </form>
                {roster.length > 0 && (
                  <div className="cfl-roster-list">
                    {activeRosterSorted.map((s) => (
                      <span key={s.seat} className="cfl-roster-chip">
                        {s.seat}．{s.name}
                        <X size={11} onClick={() => deleteStudent(s.seat)} />
                      </span>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>
        ) : (
          <div className="cfl-table-wrap">
            {loading ? (
              <div className="cfl-empty">載入中…</div>
            ) : sortedList.length === 0 ? (
              <div className="cfl-empty">{tab === 'income' ? '目前沒有收入紀錄' : '目前沒有支出紀錄'}</div>
            ) : (
              sortedList.map((t) => (
                <div className="cfl-row" key={t.id}>
                  <div className="cfl-row-date cfl-mono">{t.date}</div>
                  <div className="cfl-row-main">
                    {t.type === 'income' ? (
                      <>
                        <div className="cfl-row-title">{t.source}</div>
                        <div className="cfl-row-sub">
                          {t.seat && <span className="cfl-mono">座號 {t.seat}</span>}
                          {t.term && <span>{t.term}</span>}
                          {t.note && <span>備註：{t.note}</span>}
                        </div>
                      </>
                    ) : (
                      <>
                        <div className="cfl-row-title">{t.item}</div>
                        <div className="cfl-row-sub">
                          <span>{t.category}</span>
                          {t.term && <span>{t.term}</span>}
                          <span>{t.payee ? `經手：${t.payee}` : `${t.unitPrice} × ${t.qty}`}</span>
                          {t.note && <span>備註：{t.note}</span>}
                        </div>
                      </>
                    )}
                  </div>
                  <div className={`cfl-row-amount cfl-mono ${t.type}`}>{t.type === 'income' ? '+' : '−'}{money(t.amount)}</div>
                  {teacherMode ? (
                    confirmDelete === t.id ? (
                      <div style={{ gridColumn: '1 / -1' }} className="cfl-confirm-box">
                        <span>確定刪除此紀錄？</span>
                        <div style={{ display: 'flex', gap: 6 }}>
                          <button className="cfl-btn-primary" style={{ padding: '4px 10px', fontSize: 11 }} onClick={() => doDelete(t.id)}>刪除</button>
                          <button className="cfl-btn-ghost" style={{ padding: '4px 10px', fontSize: 11, background: '#fff' }} onClick={() => setConfirmDelete(null)}>取消</button>
                        </div>
                      </div>
                    ) : (
                      <button className="cfl-row-del" onClick={() => setConfirmDelete(t.id)}><Trash2 size={15} /></button>
                    )
                  ) : <div />}
                </div>
              ))
            )}

            {teacherMode && (
              <div className="cfl-addbar">
                {!showForm ? (
                  <button className="cfl-add-btn" onClick={() => { setIncomeForm((f) => ({ ...f, term: f.term || settings.currentTerm || '' })); setExpenseForm((f) => ({ ...f, term: f.term || settings.currentTerm || '' })); setShowForm(true); }}><Plus size={16} />新增一筆{tab === 'income' ? '收入' : '支出'}</button>
                ) : tab === 'income' ? (
                  <form className="cfl-form" onSubmit={addIncome}>
                    <div className="cfl-form-grid">
                      <div className="cfl-field"><label>日期</label><input type="date" value={incomeForm.date} onChange={(e) => setIncomeForm({ ...incomeForm, date: e.target.value })} required /></div>
                      <div className="cfl-field"><label>金額</label><input type="number" min="0" placeholder="例：3000" value={incomeForm.amount} onChange={(e) => setIncomeForm({ ...incomeForm, amount: e.target.value })} required /></div>
                      
                      {/* 智慧座號選擇器 */}
                      <div className="cfl-field">
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <label>學生座號（選填）</label>
                          {roster.length > 0 && (
                            <button 
                              type="button" 
                              onClick={() => { setSeatInputMode(seatInputMode === 'select' ? 'manual' : 'select'); }} 
                              style={{ background: 'none', border: 'none', color: 'var(--accent)', fontSize: 11, cursor: 'pointer', padding: 0 }}
                            >
                              {seatInputMode === 'select' ? '手動輸入座號' : '從名冊選擇'}
                            </button>
                          )}
                        </div>
                        {roster.length > 0 && seatInputMode === 'select' ? (
                          <select 
                            value={incomeForm.seat} 
                            onChange={handleRosterSeatSelect}
                          >
                            <option value="">（不指定學生）</option>
                            {activeRosterSorted.map(s => (
                              <option key={s.seat} value={s.seat}>
                                {s.seat}號 — {s.name}
                              </option>
                            ))}
                            <option value="__manual__">[手動輸入座號]</option>
                          </select>
                        ) : (
                          <input 
                            type="number" 
                            min="1" 
                            placeholder="例：5" 
                            value={incomeForm.seat} 
                            onChange={(e) => setIncomeForm({ ...incomeForm, seat: e.target.value })} 
                          />
                        )}
                      </div>

                      <div className="cfl-field">
                        <label>學期</label>
                        <select value={incomeForm.term} onChange={(e) => setIncomeForm({ ...incomeForm, term: e.target.value })}>
                          <option value="">（不指定）</option>
                          {termsList.map((tm) => <option key={tm} value={tm}>{tm}</option>)}
                        </select>
                      </div>

                      <div className="cfl-field" style={{ gridColumn: '1 / -1' }}><label>內容（收入來源）</label><input placeholder="例：113-1 班費" value={incomeForm.source} onChange={(e) => setIncomeForm({ ...incomeForm, source: e.target.value })} required /></div>
                      <div className="cfl-field" style={{ gridColumn: '1 / -1' }}><label>備註（選填）</label><input placeholder="例：補繳" value={incomeForm.note} onChange={(e) => setIncomeForm({ ...incomeForm, note: e.target.value })} /></div>
                    </div>
                    <div className="cfl-form-actions">
                      <button type="submit" className="cfl-btn-primary">儲存</button>
                      <button type="button" className="cfl-btn-ghost" onClick={() => setShowForm(false)}>取消</button>
                    </div>
                  </form>
                ) : (
                  <form className="cfl-form" onSubmit={addExpense}>
                    <div className="cfl-form-grid">
                      <div className="cfl-field"><label>日期</label><input type="date" value={expenseForm.date} onChange={(e) => setExpenseForm({ ...expenseForm, date: e.target.value })} required /></div>
                      <div className="cfl-field">
                        <label>支出類別</label>
                        <select value={expenseForm.category} onChange={(e) => setExpenseForm({ ...expenseForm, category: e.target.value })}>
                          {EXPENSE_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>
                      <div className="cfl-field" style={{ gridColumn: '1 / -1' }}><label>支出項目</label><input placeholder="例：冷氣卡儲值" value={expenseForm.item} onChange={(e) => setExpenseForm({ ...expenseForm, item: e.target.value })} required /></div>
                      <div className="cfl-field"><label>單價</label><input type="number" min="0" value={expenseForm.unitPrice} onChange={(e) => setExpenseForm({ ...expenseForm, unitPrice: e.target.value })} required /></div>
                      <div className="cfl-field"><label>數量</label><input type="number" min="1" value={expenseForm.qty} onChange={(e) => setExpenseForm({ ...expenseForm, qty: e.target.value })} required /></div>
                      <div className="cfl-field"><label>取款人（選填）</label><input placeholder="例：曾美蓮" value={expenseForm.payee} onChange={(e) => setExpenseForm({ ...expenseForm, payee: e.target.value })} /></div>
                      <div className="cfl-field">
                        <label>學期</label>
                        <select value={expenseForm.term} onChange={(e) => setExpenseForm({ ...expenseForm, term: e.target.value })}>
                          <option value="">（不指定）</option>
                          {termsList.map((tm) => <option key={tm} value={tm}>{tm}</option>)}
                        </select>
                      </div>
                      <div className="cfl-field" style={{ gridColumn: '1 / -1' }}><label>備註（選填）</label><input placeholder="例：分兩批購買" value={expenseForm.note} onChange={(e) => setExpenseForm({ ...expenseForm, note: e.target.value })} /></div>
                    </div>
                    
                    {/* 金額即時小計顯示 */}
                    <div style={{ fontSize: 13, fontWeight: 700, color: 'var(--text-soft)', background: 'var(--primary-light)', padding: '10px 14px', borderRadius: 'var(--radius-sm)', marginBottom: 12 }}>
                      小計金額：NT$ {money((Number(expenseForm.unitPrice) || 0) * (Number(expenseForm.qty) || 1))}
                    </div>
                    
                    <div className="cfl-form-actions">
                      <button type="submit" className="cfl-btn-primary">儲存</button>
                      <button type="button" className="cfl-btn-ghost" onClick={() => setShowForm(false)}>取消</button>
                    </div>
                  </form>
                )}
              </div>
            )}
          </div>
        )}

        {categoryBreakdown.length > 0 && (
          <div className="cfl-breakdown">
            <div className="cfl-breakdown-title">支出類別分布</div>
            {categoryBreakdown.map((c) => (
              <div className="cfl-bar-row" key={c.category}>
                <div className="cfl-bar-label">{c.category}</div>
                <div className="cfl-bar-track">
                  <div className="cfl-bar-fill" style={{ width: `${(c.amount / maxCat) * 100}%` }} />
                </div>
                <div className="cfl-bar-amount cfl-mono">{money(c.amount)}</div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* 密碼解鎖/設定 Modal */}
      {modal && (
        <div className="cfl-overlay" onClick={() => setModal(null)}>
          <div className="cfl-modal" onClick={(e) => e.stopPropagation()}>
            {modal === 'setup' ? (
              <>
                <div className="cfl-modal-title">設定教師密碼</div>
                <div className="cfl-modal-sub">第一次使用，請設定 4 碼以上密碼。這只是防止家長或學生誤觸修改，並非嚴格的帳號安全機制，請自行妥善保管。</div>
                <input type="password" inputMode="numeric" placeholder="輸入密碼" value={pinInput} onChange={(e) => setPinInput(e.target.value)} />
                <input type="password" inputMode="numeric" placeholder="再輸入一次" value={pinConfirm} onChange={(e) => setPinConfirm(e.target.value)} />
                {pinError && <div className="cfl-modal-err">{pinError}</div>}
                <div className="cfl-modal-actions">
                  <button className="cfl-btn-primary" style={{ flex: 1 }} onClick={submitSetup}>設定並解鎖</button>
                  <button className="cfl-btn-ghost" onClick={() => setModal(null)}>取消</button>
                </div>
              </>
            ) : (
              <>
                <div className="cfl-modal-title">輸入教師密碼</div>
                <div className="cfl-modal-sub">解鎖後即可新增、編輯或刪除紀錄。</div>
                <input type="password" inputMode="numeric" placeholder="密碼" value={pinInput} onChange={(e) => setPinInput(e.target.value)} autoFocus />
                {pinError && <div className="cfl-modal-err">{pinError}</div>}
                <div className="cfl-modal-actions">
                  <button className="cfl-btn-primary" style={{ flex: 1 }} onClick={submitUnlock}>解鎖</button>
                  <button className="cfl-btn-ghost" onClick={() => setModal(null)}>取消</button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* 雲端與本地衝突處理解決 Modal */}
      {syncConflictModal === 'conflict' && cloudDataTemp && (
        <div className="cfl-overlay">
          <div className="cfl-modal" style={{ maxWidth: '400px' }}>
            <div className="cfl-modal-title" style={{ color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: 6 }}>
              <Database size={18} />
              <span>同步衝突提示</span>
            </div>
            <div className="cfl-modal-sub">
              偵測到雲端試算表與本地瀏覽器的資料不一致。<br />
              請選擇您要採用的資料版本：
            </div>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10, margin: '16px 0' }}>
              <button 
                className="cfl-btn-ghost" 
                style={{ textAlign: 'left', padding: '12px', border: '1px solid var(--primary)', borderRadius: 'var(--radius)', background: 'var(--primary-light)' }}
                onClick={async () => {
                  const merged = { ...settings, ...cloudDataTemp.settings };
                  await applyCloudData(cloudDataTemp.transactions, cloudDataTemp.roster, merged);
                  setSyncConflictModal(null);
                  setCloudDataTemp(null);
                  setSyncStatus('synced');
                }}
              >
                <div style={{ fontWeight: 700, color: 'var(--primary-dark)' }}>▼ 下載雲端覆蓋本地（以雲端為準）</div>
                <div style={{ fontSize: 11, color: 'var(--text-soft)', marginTop: 4 }}>
                  雲端帳目筆數：{cloudDataTemp.transactions.length} 筆<br />
                  雲端學生名冊：{cloudDataTemp.roster.length} 人
                </div>
              </button>

              <button 
                className="cfl-btn-ghost" 
                style={{ textAlign: 'left', padding: '12px', border: '1px solid var(--accent)', borderRadius: 'var(--radius)', background: 'rgba(200, 134, 42, 0.05)' }}
                onClick={() => {
                  pushLocalToCloud(settings.sheetUrl, transactions, roster, settings);
                }}
              >
                <div style={{ fontWeight: 700, color: 'var(--accent)' }}>▲ 上傳本地覆蓋雲端（以本地為準）</div>
                <div style={{ fontSize: 11, color: 'var(--text-soft)', marginTop: 4 }}>
                  本地帳目筆數：{transactions.length} 筆<br />
                  本地學生名冊：{roster.length} 人
                </div>
              </button>
            </div>

            <div className="cfl-modal-actions">
              <button className="cfl-btn-ghost" style={{ flex: 1 }} onClick={() => { setSyncConflictModal(null); setCloudDataTemp(null); }}>暫不處理（保持現狀）</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
