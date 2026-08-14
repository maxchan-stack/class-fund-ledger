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
  Database,
  Key,
  Users,
  CreditCard,
  Receipt,
  FileSpreadsheet,
  Printer,
  ChevronRight,
  ShieldCheck,
  Wallet
} from 'lucide-react';
import * as XLSX from 'xlsx';

const EXPENSE_CATEGORIES = ['學科學資', '練習卷', '刊物'];

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

function formatBrowserDate(val) {
  if (!val) return '';
  const str = String(val).trim();
  if (str.indexOf('T') !== -1) {
    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      const m = ("0" + (d.getMonth() + 1)).slice(-2);
      const day = ("0" + d.getDate()).slice(-2);
      return y + "-" + m + "-" + day;
    }
  }
  return str.replace(/\//g, '-');
}

function parseRosterFromRows(rows) {
  if (!rows || rows.length === 0) return [];
  let seatColIndex = -1;
  let nameColIndex = -1;
  for (let r = 0; r < Math.min(rows.length, 3); r++) {
    const row = rows[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = String(row[c] || '').trim();
      if (
        val.includes('座號') || 
        val === '座' || 
        val === '號' || 
        val === '座号' || 
        val === 'No' || 
        val === 'no' || 
        val === 'seat' || 
        val === 'Seat'
      ) {
        seatColIndex = c;
      }
      if (
        val.includes('姓名') || 
        val === '名' || 
        val === '學生' || 
        val === '学生' || 
        val === 'name' || 
        val === 'Name' || 
        val === '姓名(學生)'
      ) {
        nameColIndex = c;
      }
    }
    if (seatColIndex !== -1 && nameColIndex !== -1) {
      rows = rows.slice(r + 1);
      break;
    }
  }
  if (seatColIndex === -1) seatColIndex = 0;
  if (nameColIndex === -1) nameColIndex = 1;
  const result = [];
  const processedSeats = new Set();
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const seatVal = String(row[seatColIndex] || '').trim();
    const nameVal = String(row[nameColIndex] || '').trim();
    const seatNum = parseInt(seatVal, 10);
    if (!isNaN(seatNum) && seatNum > 0 && nameVal) {
      const seatKey = String(seatNum);
      if (!processedSeats.has(seatKey)) {
        processedSeats.add(seatKey);
        result.push({ seat: seatNum, name: nameVal });
      }
    }
  }
  return result.sort((a, b) => a.seat - b.seat);
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

function areRostersEqual(local, cloud) {
  const localList = (local || []).map(r => ({ seat: String(r.seat || '').trim(), name: r.name || '' }));
  const cloudList = (cloud || []).map(r => ({ seat: String(r.seat || '').trim(), name: r.name || '' }));
  if (localList.length !== cloudList.length) return false;
  const localMap = {};
  localList.forEach(r => { localMap[r.seat] = r.name; });
  return cloudList.every(r => localMap[r.seat] === r.name);
}

function areSettingsEqual(local, cloud) {
  if (!local || !cloud) return false;
  if ((local.className || '') !== (cloud.className || '')) return false;
  if ((local.currentTerm || '') !== (cloud.currentTerm || '')) return false;
  const localTerms = Array.from(new Set(local.terms || [])).sort().join(',');
  const cloudTerms = Array.from(new Set(cloud.terms || [])).sort().join(',');
  if (localTerms !== cloudTerms) return false;
  return true;
}

function areTransactionsEqual(local, cloud) {
  const localIncomes = (local || []).filter(t => t.type === 'income').map(t => ({
    date: formatBrowserDate(t.date),
    source: t.source || t.item || '',
    amount: Number(t.amount) || 0,
    seat: String(t.seat || ''),
    term: t.term || '',
    note: t.note || ''
  }));
  const cloudIncomes = (cloud || []).filter(t => t.type === 'income').map(t => ({
    date: formatBrowserDate(t.date),
    source: t.source || t.item || '',
    amount: Number(t.amount) || 0,
    seat: String(t.seat || ''),
    term: t.term || '',
    note: t.note || ''
  }));

  const localExpenses = (local || []).filter(t => t.type === 'expense').map(t => ({
    date: formatBrowserDate(t.date),
    category: t.category || '',
    item: t.item || t.source || '',
    unitPrice: Number(t.unitPrice) || 0,
    qty: Number(t.qty) || 1,
    amount: Number(t.amount) || 0,
    payee: String(t.payee || ''),
    seat: String(t.seat || ''),
    term: t.term || '',
    note: t.note || ''
  }));
  const cloudExpenses = (cloud || []).filter(t => t.type === 'expense').map(t => ({
    date: formatBrowserDate(t.date),
    category: t.category || '',
    item: t.item || t.source || '',
    unitPrice: Number(t.unitPrice) || 0,
    qty: Number(t.qty) || 1,
    amount: Number(t.amount) || 0,
    payee: String(t.payee || ''),
    seat: String(t.seat || ''),
    term: t.term || '',
    note: t.note || ''
  }));

  if (localIncomes.length !== cloudIncomes.length) return false;
  for (let i = 0; i < localIncomes.length; i++) {
    const l = localIncomes[i];
    const c = cloudIncomes[i];
    if (l.date !== c.date || l.source !== c.source || l.amount !== c.amount || l.seat !== c.seat || l.term !== c.term || l.note !== c.note) {
      return false;
    }
  }

  if (localExpenses.length !== cloudExpenses.length) return false;
  for (let i = 0; i < localExpenses.length; i++) {
    const l = localExpenses[i];
    const c = cloudExpenses[i];
    if (l.date !== c.date || l.category !== c.category || l.item !== c.item || l.unitPrice !== c.unitPrice || l.qty !== c.qty || l.amount !== c.amount || l.payee !== c.payee || l.seat !== c.seat || l.term !== c.term || l.note !== c.note) {
      return false;
    }
  }

  return true;
}

export default function App() {
  const [deviceRole, setDeviceRole] = useState('viewer'); // 'teacher' | 'viewer'
  const [teacherMode, setTeacherMode] = useState(false);
  const [transactions, setTransactions] = useState([]);
  const [roster, setRoster] = useState([]);
  const [settings, setSettings] = useState({
    className: '214 班',
    pin: '',
    sheetUrl: '',
    spreadsheetUrl: '',
    terms: ['114-1'],
    currentTerm: '114-1',
    duesConfig: {}
  });
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState('students'); // 'students' | 'income' | 'expense'
  const [termFilter, setTermFilter] = useState('all');
  const [error, setError] = useState('');

  // 雲端同步狀態：'idle' | 'syncing' | 'synced' | 'pending_push' | 'error'
  const [syncStatus, setSyncStatus] = useState('idle');
  const [syncConflictModal, setSyncConflictModal] = useState(null);
  const [cloudDataTemp, setCloudDataTemp] = useState(null);

  const tapCountRef = useRef(0);
  const tapTimerRef = useRef(null);
  const fileInputRef = useRef(null);

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
  const [shareLinkCopied, setShareLinkCopied] = useState(false);
  const [sheetUrlInput, setSheetUrlInput] = useState('');
  const [spreadsheetUrlInput, setSpreadsheetUrlInput] = useState('');
  const [backupHelp, setBackupHelp] = useState(false);
  const [editingTerms, setEditingTerms] = useState(false);
  const [newTermInput, setNewTermInput] = useState('');
  
  // 學生名冊編輯狀態
  const [newStudentSeat, setNewStudentSeat] = useState('');
  const [newStudentName, setNewStudentName] = useState('');

  // 學生個別存摺明細檢視 Modal
  const [selectedStudentForModal, setSelectedStudentForModal] = useState(null);

  // 記帳表單模式
  const [incomeMode, setIncomeMode] = useState('individual'); // 'individual' | 'batch'
  const [expenseMode, setExpenseMode] = useState('batch'); // 'batch' | 'individual'

  // 批次選擇的座號陣列（預設名冊全選）
  const [batchSelectedSeats, setBatchSelectedSeats] = useState([]);

  // 表單資料
  const blankIncome = { date: todayStr(), source: '', seat: '', amount: '', term: '', note: '' };
  const blankExpense = { date: todayStr(), category: EXPENSE_CATEGORIES[0], item: '', unitPrice: '', qty: '1', payee: '', seat: '', term: '', note: '' };
  const [incomeForm, setIncomeForm] = useState(blankIncome);
  const [expenseForm, setExpenseForm] = useState(blankExpense);

  useEffect(() => {
    loadAll();
  }, []);

  // 當名冊變動時，預設勾選全部學生進行批次扣款
  useEffect(() => {
    if (roster.length > 0) {
      setBatchSelectedSeats(roster.map(s => String(s.seat)));
    }
  }, [roster]);

  async function loadAll() {
    setLoading(true);
    const DEFAULT_SHEET_URL = 'https://script.google.com/macros/s/AKfycbwBoo653bMsvkZceaXks8x2Ul2GuFJWI5ctXQMkh3vq_YLolAerNJWIv9gRyEnOmvN_Bw/exec';
    const DEFAULT_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/18CZv9TNJHrUBiAGPUTti2pNuDEMMXwniXU9aBs_lsyQ/edit';

    let loadedSettings = { 
      className: '214 班', 
      pin: '', 
      sheetUrl: DEFAULT_SHEET_URL, 
      spreadsheetUrl: DEFAULT_SPREADSHEET_URL, 
      terms: ['114-1'], 
      currentTerm: '114-1',
      duesConfig: {}
    };
    let loadedTransactions = [];
    let loadedRoster = [];

    const params = new URLSearchParams(window.location.search);
    const apiId = params.get('api') || params.get('id');
    const urlParam = params.get('url');
    let querySheetUrl = '';
    if (urlParam) {
      querySheetUrl = urlParam;
    } else if (apiId) {
      querySheetUrl = `https://script.google.com/macros/s/${apiId}/exec`;
    }

    try {
      const s = await StorageService.get('settings', true);
      if (s) {
        const parsed = JSON.parse(s.value);
        loadedSettings = {
          ...loadedSettings,
          ...parsed,
          sheetUrl: parsed.sheetUrl || DEFAULT_SHEET_URL,
          spreadsheetUrl: parsed.spreadsheetUrl || DEFAULT_SPREADSHEET_URL,
          className: parsed.className || '214 班'
        };
      }
    } catch (e) {
      console.error('載入 settings 失敗', e);
    }

    if (querySheetUrl && querySheetUrl !== loadedSettings.sheetUrl) {
      loadedSettings = { ...loadedSettings, sheetUrl: querySheetUrl };
      try {
        await StorageService.set('settings', JSON.stringify(loadedSettings), true);
      } catch (e) {
        console.error('儲存網址列參數設定失敗', e);
      }
    }

    setSettings(loadedSettings);
    setClassNameInput(loadedSettings.className || '214 班');
    setSheetUrlInput(loadedSettings.sheetUrl || DEFAULT_SHEET_URL);
    setSpreadsheetUrlInput(loadedSettings.spreadsheetUrl || DEFAULT_SPREADSHEET_URL);

    try {
      const t = await StorageService.get('ledger', true);
      if (t) {
        loadedTransactions = JSON.parse(t.value).map(tx => ({
          ...tx,
          date: formatBrowserDate(tx.date),
          source: tx.type === 'income' ? (tx.source || tx.item || '') : '',
          item: tx.type === 'expense' ? (tx.item || tx.source || '') : ''
        }));
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
      const d = await StorageService.get('device-role', false);
      setDeviceRole(d && d.value === 'teacher' ? 'teacher' : 'viewer');
    } catch {
      setDeviceRole('viewer');
    }

    setLoading(false);

    if (loadedSettings.sheetUrl) {
      autoPullData(loadedSettings.sheetUrl, loadedTransactions, loadedRoster, loadedSettings);
    }
  }

  // 雲端雙向同步：拉取資料
  const autoPullData = async (url, localTrans, localRoster, localSettings) => {
    setSyncStatus('syncing');
    try {
      const queryParams = new URLSearchParams();
      if (localSettings.spreadsheetUrl) queryParams.set('url', localSettings.spreadsheetUrl);
      if (localSettings.pin) queryParams.set('auth', localSettings.pin);
      const queryString = queryParams.toString() ? `?${queryParams.toString()}` : '';
      const res = await fetch(`${url}${queryString}`, { method: 'GET', mode: 'cors' });
      
      const text = await res.text();
      let result;
      try {
        result = JSON.parse(text);
      } catch (jsonErr) {
        console.error('雲端回應無法解析為 JSON：', text.substring(0, 200));
        setSyncStatus('error');
        if (text.includes('<!doctype') || text.includes('<html')) {
          setError('雲端回應授權頁面：請確認 Apps Script 部署設定『誰有存取權』已設為『所有人』');
        } else {
          setError('雲端回應格式錯誤，請檢查試算表設定');
        }
        return;
      }

      if (!result.success) {
        console.error('雲端拉取錯誤：', result.error);
        setSyncStatus('error');
        setError(`雲端同步失敗：${result.error}`);
        return;
      }

      const cloudData = result.data;
      const isTransEqual = areTransactionsEqual(localTrans, cloudData.transactions);
      const isRosterEqual = areRostersEqual(localRoster, cloudData.roster);
      const isSettingsEqual = areSettingsEqual(localSettings, cloudData.settings);

      if (isTransEqual && isRosterEqual && isSettingsEqual) {
        setSyncStatus('synced');
        return;
      }

      const isLocalEmpty = (!localTrans || localTrans.length === 0) && (!localRoster || localRoster.length === 0);
      const isCloudEmpty = (!cloudData.transactions || cloudData.transactions.length === 0) && (!cloudData.roster || cloudData.roster.length === 0);

      if (isLocalEmpty && !isCloudEmpty) {
        const merged = { ...localSettings, ...cloudData.settings };
        await applyCloudData(cloudData.transactions, cloudData.roster, merged);
        setSyncStatus('synced');
        return;
      }

      if (!isLocalEmpty && isCloudEmpty) {
        await pushLocalToCloud(url, localTrans, localRoster, localSettings);
        return;
      }

      setCloudDataTemp(cloudData);
      setSyncConflictModal('conflict');
      setSyncStatus('pending_push');
    } catch (err) {
      console.error('autoPullData exception:', err);
      setSyncStatus('error');
    }
  };

  const applyCloudData = async (cloudTrans, cloudRoster, cloudSettings) => {
    const formattedTrans = (cloudTrans || []).map(t => ({
      ...t,
      date: formatBrowserDate(t.date),
      source: t.type === 'income' ? (t.source || t.item || '') : '',
      item: t.type === 'expense' ? (t.item || t.source || '') : ''
    }));

    setTransactions(formattedTrans);
    await StorageService.set('ledger', JSON.stringify(formattedTrans), true);

    if (cloudRoster) {
      setRoster(cloudRoster);
      await StorageService.set('roster', JSON.stringify(cloudRoster), true);
    }

    if (cloudSettings) {
      setSettings(cloudSettings);
      setClassNameInput(cloudSettings.className || '214 班');
      await StorageService.set('settings', JSON.stringify(cloudSettings), true);
    }
  };

  const pushLocalToCloud = async (url, curTrans, curRoster, curSettings) => {
    setSyncStatus('syncing');
    try {
      const payload = {
        action: 'sync',
        auth: curSettings.pin || '',
        sheetUrl: curSettings.spreadsheetUrl || '',
        transactions: curTrans,
        roster: curRoster,
        settings: curSettings
      };

      const res = await fetch(url, {
        method: 'POST',
        mode: 'cors',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify(payload)
      });

      const json = await res.json();
      if (json.success) {
        setSyncStatus('synced');
        setSyncConflictModal(null);
        setCloudDataTemp(null);
      } else {
        console.error('推送到雲端失敗：', json.error);
        setSyncStatus('error');
        setError(`推送到雲端失敗：${json.error}`);
      }
    } catch (err) {
      console.error('pushLocalToCloud exception:', err);
      setSyncStatus('error');
      setError('推送到雲端連線失敗，請檢查網路或稍後再試');
    }
  };

  const handleManualSync = () => {
    if (!settings.sheetUrl) {
      setEditingBackup(true);
      return;
    }
    autoPullData(settings.sheetUrl, transactions, roster, settings);
  };

  async function saveTransactions(next) {
    setTransactions(next);
    await StorageService.set('ledger', JSON.stringify(next), true);
    if (settings.sheetUrl) {
      pushLocalToCloud(settings.sheetUrl, next, roster, settings);
    }
  }

  async function saveRoster(next) {
    setRoster(next);
    await StorageService.set('roster', JSON.stringify(next), true);
    if (settings.sheetUrl) {
      pushLocalToCloud(settings.sheetUrl, transactions, next, settings);
    }
  }

  async function saveSettings(next) {
    setSettings(next);
    await StorageService.set('settings', JSON.stringify(next), true);
    if (next.sheetUrl) {
      pushLocalToCloud(next.sheetUrl, transactions, roster, next);
    }
  }

  const handleSecretTap = () => {
    tapCountRef.current += 1;
    if (tapTimerRef.current) clearTimeout(tapTimerRef.current);
    tapTimerRef.current = setTimeout(() => {
      tapCountRef.current = 0;
    }, 2000);

    if (tapCountRef.current >= 5) {
      tapCountRef.current = 0;
      setDeviceRole('teacher');
      StorageService.set('device-role', 'teacher', false);
      openStamp();
    }
  };

  const forgetDevice = async () => {
    await StorageService.delete('device-role', false);
    setDeviceRole('viewer');
    setTeacherMode(false);
    setShowForm(false);
  };

  const activeRosterSorted = useMemo(() => {
    return roster.slice().sort((a, b) => Number(a.seat) - Number(b.seat));
  }, [roster]);

  // 學期列表
  const termsList = useMemo(() => {
    const set = new Set(settings.terms || []);
    transactions.forEach((t) => { if (t.term) set.add(t.term); });
    return Array.from(set).sort();
  }, [settings.terms, transactions]);

  // 全局收支與餘額
  const totals = useMemo(() => {
    let income = 0, expense = 0;
    for (const t of transactions) {
      if (t.type === 'income') income += Number(t.amount) || 0;
      else expense += Number(t.amount) || 0;
    }
    return { income, expense, balance: income - expense };
  }, [transactions]);

  // 篩選學期收支
  const filteredTotals = useMemo(() => {
    if (termFilter === 'all') return null;
    let income = 0, expense = 0;
    for (const t of transactions) {
      if (t.term !== termFilter) continue;
      if (t.type === 'income') income += Number(t.amount) || 0;
      else expense += Number(t.amount) || 0;
    }
    return { income, expense, balance: income - expense };
  }, [transactions, termFilter]);

  // 類別統計（圈餅圖）
  const categoryBreakdown = useMemo(() => {
    const map = {};
    for (const t of transactions) {
      if (t.type !== 'expense') continue;
      if (termFilter !== 'all' && t.term !== termFilter) continue;
      const cat = t.category || '其他';
      map[cat] = (map[cat] || 0) + (Number(t.amount) || 0);
    }
    return Object.keys(map).map((k) => ({ category: k, amount: map[k] })).sort((a, b) => b.amount - a.amount);
  }, [transactions, termFilter]);

  // 結餘趨勢折線圖
  const balanceTrend = useMemo(() => {
    const sorted = transactions.slice().sort((a, b) => (a.date > b.date ? 1 : a.date < b.date ? -1 : 0));
    let current = 0;
    const points = [{ date: '起點', balance: 0 }];
    for (const t of sorted) {
      if (t.type === 'income') {
        current += Number(t.amount) || 0;
      } else {
        current -= Number(t.amount) || 0;
      }
      points.push({ date: t.date, balance: current });
    }
    return points;
  }, [transactions]);

  // 學生個人專戶資產負債計算（Student Sub-accounts）
  const studentAccounts = useMemo(() => {
    const map = {};
    activeRosterSorted.forEach(s => {
      map[String(s.seat)] = {
        seat: String(s.seat),
        name: s.name,
        income: 0,
        expense: 0,
        balance: 0,
        incomes: [],
        expenses: []
      };
    });

    transactions.forEach(t => {
      if (termFilter !== 'all' && t.term !== termFilter) return;
      const seatKey = String(t.seat || '').trim();
      if (!seatKey || !map[seatKey]) return;

      const amt = Number(t.amount) || 0;
      if (t.type === 'income') {
        map[seatKey].income += amt;
        map[seatKey].incomes.push(t);
      } else {
        map[seatKey].expense += amt;
        map[seatKey].expenses.push(t);
      }
    });

    Object.values(map).forEach(acc => {
      acc.balance = acc.income - acc.expense;
    });

    return Object.values(map).sort((a, b) => Number(a.seat) - Number(b.seat));
  }, [activeRosterSorted, transactions, termFilter]);

  // 學生專戶健康度統計
  const accountStats = useMemo(() => {
    const total = studentAccounts.length;
    const okCount = studentAccounts.filter(a => a.balance >= 0).length;
    const dueCount = total - okCount;
    return { total, okCount, dueCount };
  }, [studentAccounts]);

  const sortedList = useMemo(() => {
    if (tab === 'students') return [];
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
    if (String(pinInput) !== String(settings.pin || '')) { setPinError('密碼錯誤'); return; }
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

  // 批次選擇勾選/取消
  const toggleBatchSeat = (seatStr) => {
    setBatchSelectedSeats(prev => 
      prev.includes(seatStr) ? prev.filter(s => s !== seatStr) : [...prev, seatStr]
    );
  };

  const selectAllBatchSeats = () => {
    setBatchSelectedSeats(roster.map(s => String(s.seat)));
  };

  const clearBatchSeats = () => {
    setBatchSelectedSeats([]);
  };

  // 新增預繳 (Income)
  async function handleAddIncome(e) {
    e.preventDefault();
    if (!incomeForm.amount) return;

    const termVal = incomeForm.term || settings.currentTerm || '';
    const dateVal = incomeForm.date || todayStr();
    const noteVal = incomeForm.note || '';

    if (incomeMode === 'batch') {
      if (batchSelectedSeats.length === 0) {
        setError('請至少選擇一位學生進行批次預繳登記');
        return;
      }
      const unitAmt = Number(incomeForm.amount);
      const newEntries = batchSelectedSeats.map(seatStr => {
        const student = roster.find(s => String(s.seat) === seatStr);
        const namePart = student ? ` - ${student.name}` : '';
        return {
          id: uid(),
          type: 'income',
          date: dateVal,
          source: incomeForm.source || `${termVal} 教材預繳費${namePart}`,
          seat: seatStr,
          amount: unitAmt,
          term: termVal,
          note: noteVal
        };
      });

      await saveTransactions([...transactions, ...newEntries]);
      setIncomeForm({ ...blankIncome, term: settings.currentTerm || '' });
      setShowForm(false);
    } else {
      if (!incomeForm.seat) {
        setError('請選擇或指定預繳學生座號');
        return;
      }
      const student = roster.find(s => String(s.seat) === String(incomeForm.seat));
      const namePart = student ? ` - ${student.name}` : '';
      const entry = {
        id: uid(),
        type: 'income',
        date: dateVal,
        source: incomeForm.source || `${termVal} 教材預繳費${namePart}`,
        seat: String(incomeForm.seat),
        amount: Number(incomeForm.amount),
        term: termVal,
        note: noteVal
      };

      await saveTransactions([...transactions, entry]);
      setIncomeForm({ ...blankIncome, term: settings.currentTerm || '' });
      setShowForm(false);
    }
  }

  // 新增教材扣款 (Expense)
  async function handleAddExpense(e) {
    e.preventDefault();
    if (!expenseForm.item || !expenseForm.unitPrice) return;

    const termVal = expenseForm.term || settings.currentTerm || '';
    const dateVal = expenseForm.date || todayStr();
    const noteVal = expenseForm.note || '';
    const unitPriceNum = Number(expenseForm.unitPrice) || 0;

    if (expenseMode === 'batch') {
      if (batchSelectedSeats.length === 0) {
        setError('請至少選擇一位扣款學生');
        return;
      }
      const newEntries = batchSelectedSeats.map(seatStr => ({
        id: uid(),
        type: 'expense',
        date: dateVal,
        category: expenseForm.category,
        item: expenseForm.item,
        unitPrice: unitPriceNum,
        qty: 1,
        amount: unitPriceNum,
        payee: expenseForm.payee || '',
        seat: seatStr,
        term: termVal,
        note: noteVal
      }));

      await saveTransactions([...transactions, ...newEntries]);
      setExpenseForm({ ...blankExpense, term: settings.currentTerm || '' });
      setShowForm(false);
    } else {
      const qtyNum = Number(expenseForm.qty) || 1;
      const totalAmount = unitPriceNum * qtyNum;
      const entry = {
        id: uid(),
        type: 'expense',
        date: dateVal,
        category: expenseForm.category,
        item: expenseForm.item,
        unitPrice: unitPriceNum,
        qty: qtyNum,
        amount: totalAmount,
        payee: expenseForm.payee || '',
        seat: expenseForm.seat ? String(expenseForm.seat) : '',
        term: termVal,
        note: noteVal
      };

      await saveTransactions([...transactions, entry]);
      setExpenseForm({ ...blankExpense, term: settings.currentTerm || '' });
      setShowForm(false);
    }
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
    const formattedSpreadsheetUrl = spreadsheetUrlInput.trim();
    const nextSettings = { ...settings, sheetUrl: formattedUrl, spreadsheetUrl: formattedSpreadsheetUrl };
    await saveSettings(nextSettings);
    setEditingBackup(false);
    if (formattedUrl) {
      autoPullData(formattedUrl, transactions, roster, nextSettings);
    }
  }

  const copyParentShareUrl = () => {
    if (!settings.sheetUrl) return;
    const match = settings.sheetUrl.match(/macros\/s\/([^\/]+)\/exec/);
    const apiId = match ? match[1] : '';
    if (!apiId) return;
    const shareUrl = `${window.location.origin}${window.location.pathname}?view=parent&api=${apiId}`;
    
    navigator.clipboard.writeText(shareUrl)
      .then(() => {
        setShareLinkCopied(true);
        setTimeout(() => setShareLinkCopied(false), 3000);
      })
      .catch((err) => {
        console.error('複製連結失敗', err);
        setError('複製連結失敗，請手動複製網址');
      });
  };

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

  async function handleResetParentPin(seat, name) {
    if (!settings.sheetUrl) {
      setError('請先在設定中填寫試算表網址');
      return;
    }
    if (!window.confirm(`確定要重設座號 ${seat}（${name}）的家長密碼 (PIN 碼) 嗎？\n這將刪除該座號在雲端試算表的註冊紀錄，允許家長重新設定。`)) return;

    setSyncStatus('syncing');
    setError('');
    try {
      const response = await fetch(settings.sheetUrl, {
        method: 'POST',
        mode: 'cors',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          action: 'resetParentPin',
          auth: settings.pin || '',
          sheetUrl: settings.spreadsheetUrl || '',
          seat: String(seat)
        })
      });
      
      if (!response.ok) throw new Error('Network response was not ok');
      const resJson = await response.json();
      if (resJson.success) {
        alert(`座號 ${seat}（${name}）家長 PIN 碼已重設，家長下次登入時可重新設定。`);
        setSyncStatus('synced');
      } else {
        setError(`重設失敗：${resJson.error || '未知錯誤'}`);
        setSyncStatus('error');
      }
    } catch (err) {
      console.error('Reset PIN error:', err);
      setError('連線失敗，請檢查網路後再試');
      setSyncStatus('error');
    }
  }

  const handleRosterImport = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        
        const parsedRoster = parseRosterFromRows(rows);
        if (parsedRoster.length === 0) {
          setError('未辨識出學生座號或姓名，請確認 Excel 首列包含「座號」與「姓名」欄位');
          return;
        }

        const confirmMsg = `已辨識出 ${parsedRoster.length} 位學生名冊資料。\n\n點選「確定」將以此名冊覆蓋現有名冊。\n（若有既有資料將自動同步至雲端）`;
        if (window.confirm(confirmMsg)) {
          await saveRoster(parsedRoster);
        }
      } catch (err) {
        console.error('名冊解析失敗：', err);
        setError('名冊檔案解析失敗，請確認檔案格式是否正確（.xlsx, .xls 或 .csv）');
      } finally {
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // 匯出 Excel 對帳報表
  const handleExportStatement = () => {
    const exportData = studentAccounts.map(acc => ({
      '座號': acc.seat,
      '學生姓名': acc.name,
      '本學期預繳總額': acc.income,
      '教材累計扣款': acc.expense,
      '專戶目前餘額': acc.balance,
      '專戶狀態': acc.balance >= 0 ? '餘額充足' : '需補繳'
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '教材費專戶對帳表');
    XLSX.writeFile(wb, `${settings.className || '班級'}_${termFilter === 'all' ? '歷年' : termFilter}_教材費專戶對帳表.xlsx`);
  };

  if (loading) {
    return <div className="cfl-root" style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', minHeight: '100vh', color: 'var(--text-soft)' }}>系統載入中...</div>;
  }

  return (
    <div className="cfl-root">
      <header>
        <div className="header-inner">
          <div className="brand" onClick={deviceRole !== 'teacher' ? handleSecretTap : undefined}>
            教材費紀錄系統
            <small>學生專戶專款專用管理 · {settings.className || '214 班'}</small>
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
            <button className={tab === 'students' ? 'active' : ''} onClick={() => { setTab('students'); setShowForm(false); }}>
              學生專戶總覽
            </button>
            <button className={tab === 'income' ? 'active' : ''} onClick={() => { setTab('income'); setShowForm(false); }}>
              預繳紀錄
            </button>
            <button className={tab === 'expense' ? 'active' : ''} onClick={() => { setTab('expense'); setShowForm(false); }}>
              教材扣款
            </button>
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
                    placeholder="例：214 班"
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
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                    <div className="cfl-field" style={{ margin: 0 }}>
                      <label style={{ fontSize: '12px', fontWeight: 700, color: 'var(--primary-dark)' }}>Google Apps Script 網頁應用程式網址（必填）</label>
                      <input
                        className="cfl-classname-input"
                        style={{ width: '100%', boxSizing: 'border-box' }}
                        placeholder="貼上 Apps Script 網址（https://script.google.com/macros/s/.../exec）"
                        value={sheetUrlInput}
                        onChange={(e) => setSheetUrlInput(e.target.value)}
                      />
                    </div>
                    <div className="cfl-field" style={{ margin: 0 }}>
                      <label style={{ fontSize: '12px', fontWeight: 700, color: 'var(--primary-dark)' }}>Google 試算表網址（選填，若 Apps Script 為獨立版則必填）</label>
                      <input
                        className="cfl-classname-input"
                        style={{ width: '100%', boxSizing: 'border-box' }}
                        placeholder="貼上您的 Google 試算表網址（https://docs.google.com/spreadsheets/d/.../edit）"
                        value={spreadsheetUrlInput}
                        onChange={(e) => setSpreadsheetUrlInput(e.target.value)}
                      />
                    </div>
                    <div style={{ display: 'flex', gap: '8px', justifyContent: 'flex-end', marginTop: '6px' }}>
                      <button className="cfl-btn-primary" style={{ padding: '8px 20px', fontSize: '13px' }} onClick={saveBackupUrl}>儲存設定</button>
                      <button className="cfl-btn-ghost" style={{ padding: '8px 16px', fontSize: '13px' }} onClick={() => setEditingBackup(false)}>關閉</button>
                    </div>
                    {settings.sheetUrl && (
                      <div style={{ borderTop: '1px solid rgba(255,255,255,0.08)', paddingTop: 10, marginTop: 4 }}>
                        <button 
                          type="button"
                          className="cfl-btn-ghost" 
                          style={{ width: '100%', padding: '8px', fontSize: '13px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6 }}
                          onClick={copyParentShareUrl}
                        >
                          <Plus size={14} /> 複製家長專戶查詢連結（獨立存摺視圖）
                        </button>
                        {shareLinkCopied && (
                          <div style={{ fontSize: 11, color: 'var(--accent)', marginTop: 4, textAlign: 'center' }}>
                            已複製家長查詢連結！家長登入後僅能看到自身子女之專戶扣款明細與餘額。
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                  <div className="cfl-backup-help-toggle" onClick={() => setBackupHelp((v) => !v)}>
                    {backupHelp ? '隱藏說明 ▲' : '如何設定同步網址？ ▼'}
                  </div>
                  {backupHelp && (
                    <div className="cfl-backup-help">
                      <strong>【推薦方式】繫結型試算表 Apps Script：</strong><br />
                      1. 在您的 Google 試算表中，點選「擴充功能」→「Apps Script」。<br />
                      2. 將本專案根目錄的 <code>google-apps-script.js</code> 內容複製貼入並儲存。<br />
                      3. 點選右上角「部署」→「管理部署」/「新增部署」，類型選擇「網頁應用程式」。<br />
                      4. 設定「執行身分」為「我」，「誰有存取權」為「所有人」，點選「部署」並授權。<br />
                      5. 複製產生的「網頁應用程式 URL」並貼在上方的「網頁應用程式網址」即可。
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
                <div className="cfl-stamp-label">{teacherMode ? '教師管理模式' : '檢視模式'}</div>
              </div>
            )}
          </div>

          {/* 專戶儀表板 */}
          <div className="cfl-dashboard">
            <div className={`cfl-dash-main ${totals.balance < 0 ? 'neg' : ''}`}>
              <div className="cfl-dash-main-label">教材專戶總結餘</div>
              <div className="cfl-dash-main-value cfl-mono">
                NT$ {money(totals.balance)}
              </div>
              {totals.income > 0 && (
                <div className="cfl-dash-progress-wrap">
                  <div
                    className="cfl-dash-progress-bar"
                    style={{ width: `${Math.min(100, Math.round((totals.expense / totals.income) * 100))}%` }}
                  />
                  <div className="cfl-dash-progress-label">
                    總教材款已扣抵 {Math.min(100, Math.round((totals.expense / totals.income) * 100))}%
                  </div>
                </div>
              )}
            </div>

            <div className="cfl-dash-side-row">
              <div className="cfl-dash-side income">
                <div className="cfl-dash-side-label">
                  <TrendingUp size={12} />總預繳 (儲值)
                </div>
                <div className="cfl-dash-side-value cfl-mono">+{money(totals.income)}</div>
              </div>
              <div className="cfl-dash-side expense">
                <div className="cfl-dash-side-label">
                  <TrendingDown size={12} />總教材支出
                </div>
                <div className="cfl-dash-side-value cfl-mono">-{money(totals.expense)}</div>
              </div>
            </div>
          </div>
        </div>

        {teacherMode && (
          <div className="cfl-note">
            目前為教師模式，可登記預繳、進行批次或個別教材扣款。再按一次印章即可鎖回檢視模式。<br />
            <span className="cfl-forget" onClick={forgetDevice}>不是自己的裝置？點此忘記此裝置</span>
          </div>
        )}

        {/* 學期切換欄 */}
        <div className="cfl-term-row">
          <label>學期</label>
          <select value={termFilter} onChange={(e) => setTermFilter(e.target.value)}>
            <option value="all">全部學期</option>
            {termsList.map((tm) => <option key={tm} value={tm}>{tm}</option>)}
          </select>
          {termFilter !== 'all' && filteredTotals && (
            <span className="cfl-term-subtotal cfl-mono">
              本學期預繳 {money(filteredTotals.income)}｜扣款 {money(filteredTotals.expense)}｜結餘 {money(filteredTotals.balance)}
            </span>
          )}
          {teacherMode && (
            <button className="cfl-classname-edit-btn" style={{ marginLeft: 'auto' }} onClick={() => setEditingTerms((v) => !v)}>
              管理學期
            </button>
          )}
        </div>

        {teacherMode && editingTerms && (
          <div className="cfl-backup-panel" style={{ marginBottom: 12 }}>
            <div className="cfl-backup-row">
              <input
                className="cfl-classname-input"
                style={{ flex: 1 }}
                placeholder="新增學期，例：114-2"
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
          </div>
        )}

        {/* TAB 1: 學生個人專戶總覽 */}
        {tab === 'students' && (
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: 'var(--primary-dark)', display: 'flex', alignItems: 'center', gap: 8 }}>
                <Users size={16} />
                <span>全班學生教材專戶 ({studentAccounts.length} 人)</span>
                <span style={{ fontSize: 12, fontWeight: 600, color: 'var(--green)', background: '#dcfce7', padding: '2px 8px', borderRadius: 99 }}>
                  正常 {accountStats.okCount} 人
                </span>
                {accountStats.dueCount > 0 && (
                  <span style={{ fontSize: 12, fontWeight: 600, color: 'var(--red)', background: '#fee2e2', padding: '2px 8px', borderRadius: 99 }}>
                    需補繳 {accountStats.dueCount} 人
                  </span>
                )}
              </div>
              <button 
                type="button" 
                className="cfl-btn-ghost" 
                style={{ padding: '6px 12px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: 4 }}
                onClick={handleExportStatement}
              >
                <FileSpreadsheet size={14} /> 匯出對帳報表
              </button>
            </div>

            {studentAccounts.length === 0 ? (
              <div className="cfl-empty">
                目前尚無學生名冊，請{teacherMode ? '在下方新增學生或匯入 Excel 名冊。' : '導師先建立學生名冊。'}
              </div>
            ) : (
              <div className="cfl-student-grid">
                {studentAccounts.map(acc => {
                  const isOk = acc.balance >= 0;
                  return (
                    <div key={acc.seat} className="cfl-student-card">
                      <div>
                        <div className="cfl-student-card-header">
                          <div>
                            <span className="cfl-student-seat">座號 {acc.seat}</span>
                            <div className="cfl-student-name">{acc.name}</div>
                          </div>
                          <span className={`cfl-student-badge ${isOk ? 'ok' : 'due'}`}>
                            {isOk ? '餘額充足' : `需補繳 NT$ ${money(Math.abs(acc.balance))}`}
                          </span>
                        </div>

                        <div className="cfl-student-stats" style={{ marginTop: 12 }}>
                          <div className="cfl-student-stat-item">
                            <span className="cfl-student-stat-label">累計預繳</span>
                            <span className="cfl-student-stat-val" style={{ color: 'var(--green)' }}>+{money(acc.income)}</span>
                          </div>
                          <div className="cfl-student-stat-item">
                            <span className="cfl-student-stat-label">教材扣款</span>
                            <span className="cfl-student-stat-val" style={{ color: 'var(--red)' }}>-{money(acc.expense)}</span>
                          </div>
                          <div className="cfl-student-stat-item">
                            <span className="cfl-student-stat-label">專戶結餘</span>
                            <span className="cfl-student-stat-val" style={{ color: isOk ? 'var(--text)' : 'var(--red)' }}>
                              ${money(acc.balance)}
                            </span>
                          </div>
                        </div>
                      </div>

                      <div className="cfl-student-actions">
                        <button
                          type="button"
                          className="cfl-classname-edit-btn"
                          style={{ padding: '4px 10px', fontSize: '11px' }}
                          onClick={() => setSelectedStudentForModal(acc)}
                        >
                          查閱個人存摺明細
                        </button>
                        {teacherMode && (
                          <button
                            type="button"
                            className="cfl-row-del"
                            title="重設家長 PIN 碼"
                            onClick={() => handleResetParentPin(acc.seat, acc.name)}
                            style={{ color: 'var(--accent)', padding: 4 }}
                          >
                            <Key size={14} />
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}

            {/* 名冊管理區塊 (教師模式) */}
            {teacherMode && (
              <div className="cfl-addbar" style={{ marginTop: 24 }}>
                <div style={{ display: 'flex', gap: 8, marginBottom: 16, borderBottom: '1px solid rgba(255,255,255,0.08)', paddingBottom: 16 }}>
                  <button 
                    type="button" 
                    className="cfl-btn-ghost" 
                    style={{ flex: 1, padding: '8px 12px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, fontSize: '0.85rem' }}
                    onClick={() => fileInputRef.current && fileInputRef.current.click()}
                  >
                    <Plus size={14} /> 匯入 Excel／CSV 名冊
                  </button>
                  <input 
                    type="file" 
                    ref={fileInputRef} 
                    style={{ display: 'none' }} 
                    accept=".xlsx, .xls, .csv" 
                    onChange={handleRosterImport} 
                  />
                </div>
                <form className="cfl-form" onSubmit={addStudent} style={{ marginTop: 0 }}>
                  <div className="cfl-form-grid">
                    <div className="cfl-field">
                      <label>座號</label>
                      <input type="number" min="1" placeholder="例：5" value={newStudentSeat} onChange={(e) => setNewStudentSeat(e.target.value)} required />
                    </div>
                    <div className="cfl-field">
                      <label>姓名</label>
                      <input placeholder="例：王小明" value={newStudentName} onChange={(e) => setNewStudentName(e.target.value)} required />
                    </div>
                  </div>
                  <div className="cfl-form-actions">
                    <button type="submit" className="cfl-btn-primary">加入名冊</button>
                  </div>
                </form>
                {roster.length > 0 && (
                  <div className="cfl-roster-list" style={{ marginTop: 14 }}>
                    {activeRosterSorted.map((s) => (
                      <span key={s.seat} className="cfl-roster-chip">
                        {s.seat}．{s.name}
                        <Key size={11} className="cfl-reset-key-btn" title="重設家長密碼" onClick={() => handleResetParentPin(s.seat, s.name)} style={{ marginLeft: 4 }} />
                        <X size={11} title="刪除學生" onClick={() => deleteStudent(s.seat)} />
                      </span>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* TAB 2 & 3: 預繳紀錄 (income) 或 教材扣款 (expense) */}
        {tab !== 'students' && (
          <div className="cfl-table-wrap">
            {sortedList.length === 0 ? (
              <div className="cfl-empty">目前尚無{tab === 'income' ? '預繳' : '扣款'}紀錄。</div>
            ) : (
              sortedList.map((t) => {
                const isIncome = t.type === 'income';
                const student = t.seat ? roster.find(s => String(s.seat) === String(t.seat)) : null;
                return (
                  <div className="cfl-row" key={t.id}>
                    <div className="cfl-row-date cfl-mono">{t.date}</div>
                    <div className="cfl-row-main">
                      <div className="cfl-row-title">
                        {isIncome ? (
                          <>
                            {t.source}
                            {t.seat && (
                              <span style={{ fontSize: 11, fontWeight: 700, color: 'var(--primary-mid)', background: 'var(--primary-light)', padding: '1px 6px', borderRadius: 4, marginLeft: 6 }}>
                                座號 {t.seat} {student ? `(${student.name})` : ''}
                              </span>
                            )}
                          </>
                        ) : (
                          <>
                            <span style={{ fontSize: 11, fontWeight: 700, color: 'var(--accent)', background: 'var(--accent-light)', padding: '1px 6px', borderRadius: 4, marginRight: 6 }}>
                              {t.category || '教材'}
                            </span>
                            {t.item}
                            {t.seat && (
                              <span style={{ fontSize: 11, fontWeight: 700, color: 'var(--primary-mid)', background: 'var(--primary-light)', padding: '1px 6px', borderRadius: 4, marginLeft: 6 }}>
                                座號 {t.seat} {student ? `(${student.name})` : ''}
                              </span>
                            )}
                          </>
                        )}
                      </div>
                      <div className="cfl-row-sub">
                        {t.term && <span>{t.term}</span>}
                        {t.note && <span> · {t.note}</span>}
                      </div>
                    </div>
                    <div className={`cfl-row-amount cfl-mono ${isIncome ? 'income' : 'expense'}`}>
                      {isIncome ? `+${money(t.amount)}` : `-${money(t.amount)}`}
                    </div>
                    {teacherMode && (
                      confirmDelete === t.id ? (
                        <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
                          <button className="cfl-row-del-confirm" onClick={() => doDelete(t.id)}>確定</button>
                          <button className="cfl-row-del" onClick={() => setConfirmDelete(null)}><X size={14} /></button>
                        </div>
                      ) : (
                        <button className="cfl-row-del" onClick={() => setConfirmDelete(t.id)}><Trash2 size={15} /></button>
                      )
                    )}
                  </div>
                );
              })
            )}

            {/* 新增紀錄按鈕與表單 */}
            {teacherMode && (
              <div className="cfl-addbar">
                {!showForm ? (
                  <button 
                    className="cfl-add-btn" 
                    onClick={() => {
                      setIncomeForm((f) => ({ ...f, term: f.term || settings.currentTerm || '' }));
                      setExpenseForm((f) => ({ ...f, term: f.term || settings.currentTerm || '' }));
                      setShowForm(true);
                    }}
                  >
                    <Plus size={16} />新增一筆{tab === 'income' ? '預繳 (儲值)' : '教材扣款'}
                  </button>
                ) : tab === 'income' ? (
                  /* 新增預繳表單 */
                  <form className="cfl-form" onSubmit={handleAddIncome}>
                    <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
                      <button
                        type="button"
                        className={incomeMode === 'individual' ? 'cfl-btn-primary' : 'cfl-btn-ghost'}
                        style={{ flex: 1, padding: '6px 10px', fontSize: '12px' }}
                        onClick={() => setIncomeMode('individual')}
                      >
                        個別學生預繳
                      </button>
                      <button
                        type="button"
                        className={incomeMode === 'batch' ? 'cfl-btn-primary' : 'cfl-btn-ghost'}
                        style={{ flex: 1, padding: '6px 10px', fontSize: '12px' }}
                        onClick={() => setIncomeMode('batch')}
                      >
                        全班批次預繳登記
                      </button>
                    </div>

                    <div className="cfl-form-grid">
                      <div className="cfl-field">
                        <label>預繳日期</label>
                        <input type="date" value={incomeForm.date} onChange={(e) => setIncomeForm({ ...incomeForm, date: e.target.value })} required />
                      </div>
                      <div className="cfl-field">
                        <label>學期</label>
                        <select value={incomeForm.term} onChange={(e) => setIncomeForm({ ...incomeForm, term: e.target.value })}>
                          <option value="">（不指定）</option>
                          {termsList.map((tm) => <option key={tm} value={tm}>{tm}</option>)}
                        </select>
                      </div>

                      {incomeMode === 'individual' ? (
                        <div className="cfl-field" style={{ gridColumn: '1 / -1' }}>
                          <label>學生座號</label>
                          <select 
                            value={incomeForm.seat} 
                            onChange={(e) => setIncomeForm({ ...incomeForm, seat: e.target.value })}
                            required
                          >
                            <option value="">-- 請選擇學生 --</option>
                            {activeRosterSorted.map(s => (
                              <option key={s.seat} value={s.seat}>
                                座號 {s.seat} — {s.name}
                              </option>
                            ))}
                          </select>
                        </div>
                      ) : (
                        <div className="cfl-field" style={{ gridColumn: '1 / -1' }}>
                          <div className="cfl-batch-selector">
                            <div className="cfl-batch-header">
                              <span>登記對象（已選 {batchSelectedSeats.length} / {roster.length} 人）</span>
                              <div style={{ display: 'flex', gap: 6 }}>
                                <button type="button" className="cfl-classname-edit-btn" onClick={selectAllBatchSeats}>全選</button>
                                <button type="button" className="cfl-classname-edit-btn" onClick={clearBatchSeats}>清空</button>
                              </div>
                            </div>
                            <div className="cfl-batch-chips">
                              {activeRosterSorted.map(s => {
                                const isSel = batchSelectedSeats.includes(String(s.seat));
                                return (
                                  <span
                                    key={s.seat}
                                    className={`cfl-batch-chip ${isSel ? 'selected' : ''}`}
                                    onClick={() => toggleBatchSeat(String(s.seat))}
                                  >
                                    {s.seat} {s.name}
                                  </span>
                                );
                              })}
                            </div>
                          </div>
                        </div>
                      )}

                      <div className="cfl-field">
                        <label>{incomeMode === 'batch' ? '每人預繳金額' : '預繳金額'}</label>
                        <input type="number" min="1" placeholder="例：1000" value={incomeForm.amount} onChange={(e) => setIncomeForm({ ...incomeForm, amount: e.target.value })} required />
                      </div>
                      <div className="cfl-field">
                        <label>項目名稱（選填）</label>
                        <input placeholder="例：114-1 期初教材預繳費" value={incomeForm.source} onChange={(e) => setIncomeForm({ ...incomeForm, source: e.target.value })} />
                      </div>
                      <div className="cfl-field" style={{ gridColumn: '1 / -1' }}>
                        <label>備註（選填）</label>
                        <input placeholder="例：現金繳交" value={incomeForm.note} onChange={(e) => setIncomeForm({ ...incomeForm, note: e.target.value })} />
                      </div>
                    </div>

                    <div className="cfl-form-actions">
                      <button type="submit" className="cfl-btn-primary">
                        {incomeMode === 'batch' ? `批次為 ${batchSelectedSeats.length} 位學生登記預繳` : '儲存'}
                      </button>
                      <button type="button" className="cfl-btn-ghost" onClick={() => setShowForm(false)}>取消</button>
                    </div>
                  </form>
                ) : (
                  /* 新增扣款表單 */
                  <form className="cfl-form" onSubmit={handleAddExpense}>
                    <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
                      <button
                        type="button"
                        className={expenseMode === 'batch' ? 'cfl-btn-primary' : 'cfl-btn-ghost'}
                        style={{ flex: 1, padding: '6px 10px', fontSize: '12px' }}
                        onClick={() => setExpenseMode('batch')}
                      >
                        全班／多選批次扣款
                      </button>
                      <button
                        type="button"
                        className={expenseMode === 'individual' ? 'cfl-btn-primary' : 'cfl-btn-ghost'}
                        style={{ flex: 1, padding: '6px 10px', fontSize: '12px' }}
                        onClick={() => setExpenseMode('individual')}
                      >
                        個別學生加扣款
                      </button>
                    </div>

                    <div className="cfl-form-grid">
                      <div className="cfl-field">
                        <label>扣款日期</label>
                        <input type="date" value={expenseForm.date} onChange={(e) => setExpenseForm({ ...expenseForm, date: e.target.value })} required />
                      </div>
                      <div className="cfl-field">
                        <label>教材類別</label>
                        <select value={expenseForm.category} onChange={(e) => setExpenseForm({ ...expenseForm, category: e.target.value })}>
                          {EXPENSE_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div className="cfl-field" style={{ gridColumn: '1 / -1' }}>
                        <label>教材／品項名稱</label>
                        <input placeholder="例：數學隨堂測驗卷、自然材料包" value={expenseForm.item} onChange={(e) => setExpenseForm({ ...expenseForm, item: e.target.value })} required />
                      </div>

                      {expenseMode === 'batch' ? (
                        <div className="cfl-field" style={{ gridColumn: '1 / -1' }}>
                          <div className="cfl-batch-selector">
                            <div className="cfl-batch-header">
                              <span>扣款對象（已選 {batchSelectedSeats.length} / {roster.length} 人）</span>
                              <div style={{ display: 'flex', gap: 6 }}>
                                <button type="button" className="cfl-classname-edit-btn" onClick={selectAllBatchSeats}>全選</button>
                                <button type="button" className="cfl-classname-edit-btn" onClick={clearBatchSeats}>清空</button>
                              </div>
                            </div>
                            <div className="cfl-batch-chips">
                              {activeRosterSorted.map(s => {
                                const isSel = batchSelectedSeats.includes(String(s.seat));
                                return (
                                  <span
                                    key={s.seat}
                                    className={`cfl-batch-chip ${isSel ? 'selected' : ''}`}
                                    onClick={() => toggleBatchSeat(String(s.seat))}
                                  >
                                    {s.seat} {s.name}
                                  </span>
                                );
                              })}
                            </div>
                          </div>
                        </div>
                      ) : (
                        <div className="cfl-field" style={{ gridColumn: '1 / -1' }}>
                          <label>指定扣款學生座號</label>
                          <select 
                            value={expenseForm.seat} 
                            onChange={(e) => setExpenseForm({ ...expenseForm, seat: e.target.value })}
                            required
                          >
                            <option value="">-- 請選擇學生 --</option>
                            {activeRosterSorted.map(s => (
                              <option key={s.seat} value={s.seat}>
                                座號 {s.seat} — {s.name}
                              </option>
                            ))}
                          </select>
                        </div>
                      )}

                      <div className="cfl-field">
                        <label>{expenseMode === 'batch' ? '每人扣款金額' : '單價'}</label>
                        <input type="number" min="0" value={expenseForm.unitPrice} onChange={(e) => setExpenseForm({ ...expenseForm, unitPrice: e.target.value })} required />
                      </div>

                      {expenseMode === 'individual' ? (
                        <div className="cfl-field">
                          <label>數量</label>
                          <input type="number" min="1" value={expenseForm.qty} onChange={(e) => setExpenseForm({ ...expenseForm, qty: e.target.value })} required />
                        </div>
                      ) : (
                        <div className="cfl-field">
                          <label>學期</label>
                          <select value={expenseForm.term} onChange={(e) => setExpenseForm({ ...expenseForm, term: e.target.value })}>
                            <option value="">（不指定）</option>
                            {termsList.map((tm) => <option key={tm} value={tm}>{tm}</option>)}
                          </select>
                        </div>
                      )}

                      <div className="cfl-field" style={{ gridColumn: '1 / -1' }}>
                        <label>備註（選填）</label>
                        <input placeholder="例：全班統一訂購" value={expenseForm.note} onChange={(e) => setExpenseForm({ ...expenseForm, note: e.target.value })} />
                      </div>
                    </div>
                    
                    <div style={{ fontSize: 13, fontWeight: 700, color: 'var(--text-soft)', background: 'var(--primary-light)', padding: '10px 14px', borderRadius: 'var(--radius-sm)', marginBottom: 12 }}>
                      {expenseMode === 'batch' 
                        ? `總計扣款：${batchSelectedSeats.length} 人 × NT$ ${money(expenseForm.unitPrice)} = NT$ ${money(batchSelectedSeats.length * (Number(expenseForm.unitPrice) || 0))}`
                        : `小計金額：NT$ ${money((Number(expenseForm.unitPrice) || 0) * (Number(expenseForm.qty) || 1))}`
                      }
                    </div>
                    
                    <div className="cfl-form-actions">
                      <button type="submit" className="cfl-btn-primary">
                        {expenseMode === 'batch' ? `執行批次扣款 (${batchSelectedSeats.length} 人)` : '儲存扣款'}
                      </button>
                      <button type="button" className="cfl-btn-ghost" onClick={() => setShowForm(false)}>取消</button>
                    </div>
                  </form>
                )}
              </div>
            )}
          </div>
        )}

        {/* 統合分析區塊：圈餅圖 + 趨勢折線圖 */}
        {(categoryBreakdown.length > 0 || balanceTrend.length > 1) && (() => {
          const PIE_COLORS = ['#1e4232','#c8860e','#2a6e30','#b26a00','#1976d2','#7b1fa2'];
          const total = categoryBreakdown.reduce((s, c) => s + c.amount, 0);
          const CX = 90, CY = 90, R = 72, IR = 44;
          let angle = -Math.PI / 2;
          const slices = categoryBreakdown.map((c, i) => {
            const frac = total > 0 ? c.amount / total : 0;
            const startAngle = angle;
            angle += frac * 2 * Math.PI;
            const endAngle = angle;
            const x1 = CX + R * Math.cos(startAngle), y1 = CY + R * Math.sin(startAngle);
            const x2 = CX + R * Math.cos(endAngle),   y2 = CY + R * Math.sin(endAngle);
            const xi1 = CX + IR * Math.cos(startAngle), yi1 = CY + IR * Math.sin(startAngle);
            const xi2 = CX + IR * Math.cos(endAngle),   yi2 = CY + IR * Math.sin(endAngle);
            const large = frac > 0.5 ? 1 : 0;
            const d = `M ${xi1} ${yi1} L ${x1} ${y1} A ${R} ${R} 0 ${large} 1 ${x2} ${y2} L ${xi2} ${yi2} A ${IR} ${IR} 0 ${large} 0 ${xi1} ${yi1} Z`;
            return { ...c, d, color: PIE_COLORS[i % PIE_COLORS.length], pct: Math.round(frac * 100) };
          });

          const balances = balanceTrend.map(p => p.balance);
          const maxBal = Math.max(...balances, 1000);
          const minBal = Math.min(...balances, 0);
          const range = maxBal - minBal || 1;
          const thresholdY = 180 - 25 - ((1000 - minBal) / range) * 130;
          const trendPts = balanceTrend.map((p, idx) => {
            const x = 40 + (idx / Math.max(1, balanceTrend.length - 1)) * 420;
            const y = 180 - 25 - ((p.balance - minBal) / range) * 130;
            return { x, y, ...p };
          });

          return (
            <div className="cfl-analysis-card">
              <div className="cfl-analysis-title">教材費分析總覽</div>

              {categoryBreakdown.length > 0 && (
                <div className="cfl-analysis-section">
                  <div className="cfl-analysis-section-label">教材類別分布</div>
                  <div className="cfl-pie-layout">
                    <svg viewBox="0 0 180 180" className="cfl-pie-svg">
                      {slices.map((s, i) => (
                        <path key={i} d={s.d} fill={s.color} stroke="#fff" strokeWidth="2" />
                      ))}
                      <text x={CX} y={CY - 6} textAnchor="middle" style={{ fontSize: '11px', fill: 'var(--text-soft)', fontWeight: 600 }}>總教材支出</text>
                      <text x={CX} y={CY + 11} textAnchor="middle" style={{ fontSize: '13px', fill: 'var(--text)', fontWeight: 800 }}>${money(total)}</text>
                    </svg>
                    <div className="cfl-pie-legend">
                      {slices.map((s, i) => (
                        <div key={i} className="cfl-pie-legend-item">
                          <span className="cfl-pie-dot" style={{ background: s.color }} />
                          <span className="cfl-pie-legend-label">{s.category}</span>
                          <span className="cfl-pie-legend-pct cfl-mono">{s.pct}%</span>
                          <span className="cfl-pie-legend-amt cfl-mono">${money(s.amount)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {balanceTrend.length > 1 && (
                <div className="cfl-analysis-section" style={{ borderTop: categoryBreakdown.length > 0 ? '1px solid var(--border)' : 'none', paddingTop: categoryBreakdown.length > 0 ? 18 : 0 }}>
                  <div className="cfl-analysis-section-label">專戶結餘變化趨勢</div>
                  <div style={{ overflowX: 'auto' }}>
                    <svg viewBox="0 0 500 180" width="100%" style={{ minWidth: '380px', display: 'block' }}>
                      <defs>
                        <linearGradient id="trendGrad2" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="var(--primary-mid)" stopOpacity="0.2" />
                          <stop offset="100%" stopColor="var(--primary-mid)" stopOpacity="0.0" />
                        </linearGradient>
                      </defs>
                      <line x1="40" y1="155" x2="460" y2="155" stroke="rgba(0,0,0,0.08)" strokeWidth="1.5" />
                      <path
                        d={`M 40,155 ${trendPts.map(p => `L ${p.x},${p.y}`).join(' ')} L 460,155 Z`}
                        fill="url(#trendGrad2)"
                      />
                      <polyline
                        fill="none"
                        stroke="var(--primary-mid)"
                        strokeWidth="2.5"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        points={trendPts.map(p => `${p.x},${p.y}`).join(' ')}
                      />
                      {trendPts.map((p, idx) => {
                        if (balanceTrend.length > 8 && idx > 0 && idx < balanceTrend.length - 1 && idx % Math.ceil(balanceTrend.length / 5) !== 0) return null;
                        return (
                          <g key={idx}>
                            <circle cx={p.x} cy={p.y} r="4" fill="#fff" stroke="var(--primary-mid)" strokeWidth="2" />
                            <text x={p.x} y={p.y - 8} fill="var(--text)" style={{ fontSize: '9px', fontWeight: 700 }} textAnchor="middle">${money(p.balance)}</text>
                            <text x={p.x} y={170} fill="var(--text-soft)" style={{ fontSize: '9px' }} textAnchor="middle">{p.date === '起點' ? '起點' : p.date.substring(5)}</text>
                          </g>
                        );
                      })}
                    </svg>
                  </div>
                </div>
              )}
            </div>
          );
        })()}
      </div>

      {/* 學生個別存摺明細檢視 Modal */}
      {selectedStudentForModal && (
        <div className="cfl-overlay" onClick={() => setSelectedStudentForModal(null)}>
          <div className="cfl-modal" style={{ maxWidth: '600px', width: '95%' }} onClick={(e) => e.stopPropagation()}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', borderBottom: '1px solid var(--border)', paddingBottom: 12 }}>
              <div>
                <span className="cfl-student-seat">座號 {selectedStudentForModal.seat}</span>
                <div className="cfl-modal-title" style={{ marginTop: 4 }}>{selectedStudentForModal.name} 同學的教材費存摺</div>
              </div>
              <button className="cfl-row-del" onClick={() => setSelectedStudentForModal(null)}><X size={18} /></button>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10, background: 'var(--bg)', padding: 12, borderRadius: 'var(--radius-sm)', margin: '14px 0' }}>
              <div>
                <div style={{ fontSize: 11, color: 'var(--text-soft)' }}>累計預繳</div>
                <div style={{ fontSize: 15, fontWeight: 800, color: 'var(--green)' }}>+{money(selectedStudentForModal.income)}</div>
              </div>
              <div>
                <div style={{ fontSize: 11, color: 'var(--text-soft)' }}>教材扣款</div>
                <div style={{ fontSize: 15, fontWeight: 800, color: 'var(--red)' }}>-{money(selectedStudentForModal.expense)}</div>
              </div>
              <div>
                <div style={{ fontSize: 11, color: 'var(--text-soft)' }}>專戶結餘</div>
                <div style={{ fontSize: 15, fontWeight: 800, color: selectedStudentForModal.balance >= 0 ? 'var(--text)' : 'var(--red)' }}>
                  ${money(selectedStudentForModal.balance)}
                </div>
              </div>
            </div>

            <div style={{ maxHeight: '320px', overflowY: 'auto' }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: 'var(--primary-dark)', marginBottom: 8 }}>教材扣款明細清單 ({selectedStudentForModal.expenses.length} 筆)</div>
              {selectedStudentForModal.expenses.length === 0 ? (
                <div style={{ textAlign: 'center', padding: 20, color: 'var(--text-soft)', fontSize: 13 }}>尚無教材扣款紀錄</div>
              ) : (
                selectedStudentForModal.expenses.map(exp => (
                  <div key={exp.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 10px', background: '#fff', border: '1px solid var(--border)', borderRadius: 6, marginBottom: 6 }}>
                    <div>
                      <div style={{ fontSize: 13, fontWeight: 700 }}>
                        <span style={{ fontSize: 10, color: 'var(--accent)', background: 'var(--accent-light)', padding: '1px 4px', borderRadius: 3, marginRight: 4 }}>{exp.category}</span>
                        {exp.item}
                      </div>
                      <div style={{ fontSize: 11, color: 'var(--text-soft)' }}>{exp.date} {exp.term ? `· ${exp.term}` : ''} {exp.note ? `· ${exp.note}` : ''}</div>
                    </div>
                    <div style={{ fontSize: 13, fontWeight: 800, color: 'var(--red)' }}>-{money(exp.amount)}</div>
                  </div>
                ))
              )}
            </div>

            <div className="cfl-modal-actions" style={{ marginTop: 14 }}>
              <button 
                type="button" 
                className="cfl-btn-ghost" 
                style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6 }}
                onClick={() => window.print()}
              >
                <Printer size={15} /> 列印此存摺
              </button>
              <button type="button" className="cfl-btn-primary" onClick={() => setSelectedStudentForModal(null)}>關閉</button>
            </div>
          </div>
        </div>
      )}

      {/* 密碼解鎖/設定 Modal */}
      {modal && (
        <div className="cfl-overlay" onClick={() => setModal(null)}>
          <div className="cfl-modal" onClick={(e) => e.stopPropagation()}>
            {modal === 'setup' ? (
              <>
                <div className="cfl-modal-title">設定教師密碼</div>
                <div className="cfl-modal-sub">第一次使用，請設定 4 碼以上密碼，防止誤觸修改。</div>
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
                <div className="cfl-modal-sub">解鎖後即可登記預繳、進行教材扣款與管理學生名冊。</div>
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
          <div className="cfl-modal" style={{ maxWidth: '640px', width: '95%' }}>
            <div className="cfl-modal-title" style={{ color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: 6, marginBottom: 12 }}>
              <Database size={18} />
              <span>同步衝突提示</span>
            </div>
            <div className="cfl-modal-sub" style={{ marginBottom: 20 }}>
              偵測到雲端試算表與本地瀏覽器的資料不一致。請對照下方兩側的版本數據，並選擇您要採用的版本：
            </div>
            
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 16, margin: '16px 0' }}>
              <div style={{ 
                border: '1.5px solid var(--border)', 
                borderRadius: 'var(--radius)', 
                padding: '16px', 
                background: 'rgba(200, 134, 42, 0.03)',
                display: 'flex',
                flexDirection: 'column',
                justifyContent: 'space-between'
              }}>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--accent)', borderBottom: '1px solid var(--border)', paddingBottom: 8, marginBottom: 12 }}>
                    <span>💻 本地本機版本</span>
                  </div>
                  <div style={{ fontSize: 13, display: 'flex', flexDirection: 'column', gap: 8, color: 'var(--text-soft)', marginBottom: 20 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <span>收支交易：</span>
                      <strong style={{ color: 'var(--text)' }}>{transactions.length} 筆</strong>
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <span>名冊人數：</span>
                      <strong style={{ color: 'var(--text)' }}>{roster.length} 人</strong>
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <span>班級名稱：</span>
                      <strong style={{ color: 'var(--text)' }}>{settings.className || '未設定'}</strong>
                    </div>
                  </div>
                </div>
                
                <button 
                  type="button"
                  className="cfl-btn-primary" 
                  style={{ width: '100%', padding: '10px', background: 'var(--accent)', borderColor: 'var(--accent)' }}
                  onClick={() => pushLocalToCloud(settings.sheetUrl, transactions, roster, settings)}
                >
                  ▲ 用本地覆蓋雲端
                </button>
              </div>

              <div style={{ 
                border: '1.5px solid var(--primary-mid)', 
                borderRadius: 'var(--radius)', 
                padding: '16px', 
                background: 'var(--primary-light)',
                display: 'flex',
                flexDirection: 'column',
                justifyContent: 'space-between'
              }}>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--primary-dark)', borderBottom: '1px solid var(--border)', paddingBottom: 8, marginBottom: 12 }}>
                    <span>☁️ 雲端備份版本</span>
                  </div>
                  <div style={{ fontSize: 13, display: 'flex', flexDirection: 'column', gap: 8, color: 'var(--text-soft)', marginBottom: 20 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <span>收支交易：</span>
                      <strong style={{ color: 'var(--text)' }}>{cloudDataTemp.transactions.length} 筆</strong>
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <span>名冊人數：</span>
                      <strong style={{ color: 'var(--text)' }}>{cloudDataTemp.roster.length} 人</strong>
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <span>班級名稱：</span>
                      <strong style={{ color: 'var(--text)' }}>{cloudDataTemp.settings.className || '未設定'}</strong>
                    </div>
                  </div>
                </div>
                
                <button 
                  type="button"
                  className="cfl-btn-primary" 
                  style={{ width: '100%', padding: '10px' }}
                  onClick={async () => {
                    const merged = { ...settings, ...cloudDataTemp.settings };
                    await applyCloudData(cloudDataTemp.transactions, cloudDataTemp.roster, merged);
                    setSyncConflictModal(null);
                    setCloudDataTemp(null);
                    setSyncStatus('synced');
                  }}
                >
                  ▼ 下載雲端覆蓋本地
                </button>
              </div>
            </div>

            <div className="cfl-modal-actions" style={{ marginTop: 12 }}>
              <button 
                type="button"
                className="cfl-btn-ghost" 
                style={{ flex: 1, padding: '10px' }} 
                onClick={() => { setSyncConflictModal(null); setCloudDataTemp(null); }}
              >
                暫不處理（保持現狀）
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
