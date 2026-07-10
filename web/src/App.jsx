import React, { useState, useEffect, useCallback, createContext } from 'react';

/* ============================================================
   Icons — 内联 SVG（不依赖外部图标库，符合 ui_ux_pro_max 规范）
   ============================================================ */
const Icons = {
  home: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>,
  landmark: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="2" x2="12" y2="22"/><path d="M5 12H2l10-9 10 9h-3"/></svg>,
  receipt: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M4 4v16c0 1.1.9 2 2 2h12a2 2 0 002-2V8l-6-6H6a2 2 0 00-2 2z"/><path d="M14 2v6h6"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>,
  table: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/><line x1="15" y1="3" x2="15" y2="21"/></svg>,
  sun: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>,
  moon: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/></svg>,
  chevronLeft: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="15 18 9 12 15 6"/></svg>,
  chevronRight: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="9 18 15 12 9 6"/></svg>,
  upload: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>,
  file: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>,
  folderOpen: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M22 19a2 2 0 01-2 2H4a2 2 0 01-2-2V5a2 2 0 012-2h5l2 3h9a2 2 0 012 2z"/></svg>,
  download: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>,
  clipboard: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M16 4h2a2 2 0 012 2v14a2 2 0 01-2 2H6a2 2 0 01-2-2V6a2 2 0 012-2h2"/><rect x="8" y="2" width="8" height="4" rx="1" ry="1"/></svg>,
  refresh: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/></svg>,
  settings: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 112.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"/></svg>,
  x: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  check: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>,
  alertCircle: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>,
  copy: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>,
  trash: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/></svg>,
  search: <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>,
};

/* ============================================================
   API 桥接 — 调用 window.pywebview.api.* 
   ============================================================ */
function callApi(method, ...args) {
  if (!window.pywebview || !window.pywebview.api) {
    // Dev mode fallback: return mock or reject
    return Promise.reject(new Error('pywebview bridge not available'));
  }
  const api = window.pywebview.api;
  // 动态调用方法
  const fn = api[method];
  if (!fn) {
    return Promise.reject(new Error(`API method "${method}" not found`));
  }
  return fn.apply(api, args);
}

/* ============================================================
   Contexts
   ============================================================ */
export const ApiContext = createContext(callApi);
export const ThemeContext = createContext({ theme: 'light', toggle: () => {} });
export const ToastContext = createContext(() => {});

/* ============================================================
   Toast 系统
   ============================================================ */
let toastId = 0;

function useToasts() {
  const [toasts, setToasts] = useState([]);
  
  const addToast = useCallback((title, message, type = 'ok') => {
    const id = ++toastId;
    setToasts(prev => [...prev, { id, title, message, type }]);
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, 3500);
  }, []);

  return { toasts, addToast };
}

/* ============================================================
   Sidebar 组件
   ============================================================ */
function Sidebar({ collapsed, setCollapsed, currentView, navigate, features }) {
  return (
    <aside className={`sidebar ${collapsed ? 'collapsed' : ''}`}>
      <div className="sidebar-brand">
        <span style={{ fontSize: 20 }}>💰</span>
        <span>财务助手</span>
      </div>
      <ul className="nav-list">
        {features.map(f => (
          <li key={f.name}
            className={`nav-item ${currentView === f.name ? 'active' : ''}`}
            onClick={() => navigate(f.name)}
          >
            <span className="nav-icon">{Icons[f.icon] || Icons.home}</span>
            <span className="nav-label">{f.title}</span>
          </li>
        ))}
      </ul>
      <button className="toggle-btn" onClick={() => setCollapsed(!collapsed)} title={collapsed ? "展开侧边栏" : "折叠侧边栏"}>
        {collapsed ? Icons.chevronRight : Icons.chevronLeft}
      </button>
    </aside>
  );
}

/* ============================================================
   TopBar 组件
   ============================================================ */
function TopBar({ title, onThemeToggle, isDark }) {
  return (
    <div className="top-bar">
      <div className="brand">
        <span>{title || '我的财务助手'}</span>
      </div>
      <div className="actions">
        <button className="btn btn-ghost btn-sm" onClick={onThemeToggle} title="切换主题">
          {isDark ? Icons.sun : Icons.moon}
        </button>
      </div>
    </div>
  );
}

/* ============================================================
   Toast 容器
   ============================================================ */
function ToastContainer({ toasts }) {
  if (!toasts.length) return null;
  return (
    <div className="toast-container">
      {toasts.map(t => (
        <div key={t.id} className={`toast toast-${t.type}`}>
          <strong>{t.title}</strong> {t.message && `— ${t.message}`}
        </div>
      ))}
    </div>
  );
}

/* ============================================================
   Lazy-loaded Views
   ============================================================ */

// Bank Classify View
function BankClassifyView() {
  const api = React.useContext(ApiContext);
  const toast = React.useContext(ToastContext);
  const [text, setText] = useState('');
  const [preview, setPreview] = useState(null);
  const [loading, setLoading] = useState(false);
  const [whitelist, setWhitelist] = useState(null);

  useEffect(() => {
    api('bank_whitelist').then(setWhitelist).catch(() => {});
  }, [api]);

  const handlePreview = async () => {
    setLoading(true);
    try {
      const res = await api('bank_preview', text);
      setPreview(res.error ? { error: res.error } : res);
    } catch (e) {
      toast('预览失败', e.message, 'err');
    } finally { setLoading(false); }
  };

  const handlePaste = async () => {
    try {
      const clip = await api('read_clipboard');
      if (clip.startsWith('__ERR__')) {
        toast('剪贴板读取失败', clip.slice(6), 'err');
        return;
      }
      setText(clip);
      toast('已粘贴', `${clip.split('\n').filter(Boolean).length} 条银行名称`, 'ok');
    } catch (e) { toast('错误', e.message, 'err'); }
  };

  const handleLoadFile = async () => {
    const paths = await api('pick_files', 'bank');
    if (!paths?.length) return;
    setLoading(true);
    try {
      const res = await api('bank_load_file', paths[0]);
      if (res.error) throw new Error(res.error);
      setText(res.banks.join('\n'));
      setPreview(res);
    } catch (e) { toast('加载失败', e.message, 'err'); }
    finally { setLoading(false); }
  };

  const handleExport = async () => {
    const out = await api('save_file_dialog', '银行承兑汇票分类结果.xlsx');
    if (!out) return;
    setLoading(true);
    try {
      const res = await api('bank_run', text, out);
      if (res.error) throw new Error(res.error);
      toast('导出成功', `已导出 ${res.total} 条（21家 ${res.n_yes} / 其他 ${res.n_no}）`, 'ok');
    } catch (e) { toast('导出失败', e.message, 'err'); }
    finally { setLoading(false); }
  };

  return (
    <div className="page-content">
      <div className="page-header">
        <h1>🏦 票据银行分类</h1>
        <p>自动识别银行全称，按 21 家白名单分类，输出 Excel 报表</p>
      </div>

      {/* Whitelist summary */}
      {whitelist && (
        <div className="card mb-md">
          <div className="card-title">白名单概览 ({whitelist.total} 家)</div>
          {whitelist.groups.map(g => (
            <div key={g[0]} style={{ marginBottom: 8 }}>
              <strong style={{ fontSize: 13 }}>{g[0]}</strong>
              <span style={{ color: 'var(--text-muted)', marginLeft: 8 }}>
                {g[1].join(' / ')}
              </span>
            </div>
          ))}
        </div>
      )}

      <div className="card">
        <div className="form-group">
          <label>输入银行名称（每行一个）</label>
          <textarea rows={8} value={text} onChange={e => setText(e.target.value)}
            placeholder="如：&#10;工商银行股份有限公司某支行&#10;招商银行南京分行..." />
        </div>
        <div className="flex gap-sm mb-md">
          <button className="btn btn-secondary" onClick={handlePaste}>
            {Icons.clipboard} 从剪贴板粘贴
          </button>
          <button className="btn btn-secondary" onClick={handleLoadFile}>
            {Icons.folderOpen} 加载文件
          </button>
          <button className="btn btn-primary" onClick={handlePreview} disabled={loading || !text.trim()}>
            {loading ? <span className="spinner" /> : Icons.search} 预览分类
          </button>
        </div>

        {preview && preview.error ? (
          <div className="badge badge-red" style={{ display: 'inline-block' }}>{preview.error}</div>
        ) : preview && preview.rows ? (
          <>
            <div className="stats-row">
              <div className="stat-card"><div className="stat-value">{preview.total}</div><div className="stat-label">总计</div></div>
              <div className="stat-card"><div className="stat-value" style={{color:'var(--success)'}}>{preview.n_yes}</div><div className="stat-label">21家承兑汇票</div></div>
              <div className="stat-card"><div className="stat-value" style={{color:'var(--danger)'}}>{preview.n_no}</div><div className="stat-label">其他</div></div>
            </div>
            <div className="table-wrap">
              <table className="data-table">
                <thead><tr><th>#</th><th>银行名称</th><th>简称</th><th>分类</th></tr></thead>
                <tbody>
                  {preview.rows.map(r => (
                    <tr key={r.seq}>
                      <td>{r.seq}</td>
                      <td>{r.name}</td>
                      <td>{r.short}</td>
                      <td><span className={`badge ${r.cat === '21银行承兑汇票' ? 'badge-green' : 'badge-orange'}`}>{r.cat}</span></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="mt-md flex gap-sm">
              <button className="btn btn-success" onClick={handleExport} disabled={loading}>
                {Icons.download} 导出 Excel
              </button>
            </div>
          </>
        ) : null}
      </div>
    </div>
  );
}

// JS Statement View
function JsStatementView() {
  const api = React.useContext(ApiContext);
  const toast = React.useContext(ToastContext);
  const [filePath, setFilePath] = useState('');
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [sortCol, setSortCol] = useState(-1);
  const [sortAsc, setSortAsc] = useState(true);

  const handleOpen = async () => {
    const paths = await api('pick_files', 'excel');
    if (!paths?.length) return;
    setFilePath(paths[0]);
    setLoading(true);
    try {
      const res = await api('js_load', paths[0]);
      if (res.error) throw new Error(res.error);
      setData(res);
    } catch (e) { toast('读取失败', e.message, 'err'); }
    finally { setLoading(false); }
  };

  const handleCopyAll = () => {
    if (!data) return;
    const lines = [data.headers.join('\t'), ...data.rows.map(r => r.join('\t'))].join('\n');
    navigator.clipboard.writeText(lines).then(
      () => toast('已复制', `${data.rows.length} 行数据已复制到剪贴板`, 'ok'),
      () => toast('复制失败', '请手动选择并复制', 'err')
    );
  };

  const handleSort = (colIdx) => {
    if (sortCol === colIdx) { setSortAsc(!sortAsc); }
    else { setSortCol(colIdx); setSortAsc(true); }
  };

  let sortedRows = data?.rows || [];
  if (sortCol >= 0 && sortedRows.length > 0) {
    sortedRows = [...sortedRows].sort((a, b) => {
      const va = a[sortCol] ?? '';
      const vb = b[sortCol] ?? '';
      try { return sortAsc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va)); }
      catch { return sortAsc ? (va > vb ? 1 : -1) : (vb > va ? 1 : -1); }
    });
  }

  return (
    <div className="page-content">
      <div className="page-header">
        <h1>📄 江苏银行对账单复制</h1>
        <p>读取对账单 Excel，表格展示、排序后一键复制到剪贴板</p>
      </div>

      <div className="card">
        <div className="flex gap-sm items-center justify-between mb-md">
          <button className="btn btn-primary" onClick={handleOpen} disabled={loading}>
            {loading ? <span className="spinner" /> : Icons.folderOpen} 选择文件
          </button>
          {data && (
            <>
              <span className="text-muted" style={{fontSize:13}}>{data.name} — {data.rows.length} 行</span>
              <button className="btn btn-secondary btn-sm" onClick={handleCopyAll}>{Icons.copy} 复制全部</button>
            </>
          )}
        </div>

        {!data && !loading ? (
          <div className="empty-state">
            {Icons.file}
            <p>点击上方按钮选择江苏银行对账单文件</p>
          </div>
        ) : loading ? (
          <div className="empty-state"><span className="spinner" /><p>正在读取...</p></div>
        ) : (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  {data.headers.map((h, i) => (
                    <th key={i} onClick={() => handleSort(i)} style={{ cursor: 'pointer', userSelect: 'none' }}>
                      {h} {sortCol === i ? (sortAsc ? '▲' : '▼') : ''}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sortedRows.map((row, ri) => (
                  <tr key={ri}>
                    {row.map((cell, ci) => <td key={ci}>{cell}</td>)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

// Notes Import View (with mapping dialog)
function NotesImportView() {
  const api = React.useContext(ApiContext);
  const toast = React.useContext(ToastContext);
  const [files, setFiles] = useState([]);
  const [recvOrg, setRecvOrg] = useState('');
  const [result, setResult] = useState(null);     // import result
  const [editingRows, setEditingRows] = useState([]); // editable table data
  const [importing, setImporting] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [showMapping, setShowMapping] = useState(false);
  const [templateInfo, setTemplateInfo] = useState(null);

  // Load template fields on mount
  useEffect(() => {
    api('notes_get_template_fields').then(tf => {
      setTemplateInfo(tf);
    }).catch(() => {});
  }, [api]);

  // File management
  const addFiles = async () => {
    const paths = await api('pick_files', 'excel');
    if (!paths?.length) return;
    const r = await api('notes_add_files', paths);
    setFiles(r.files || []);
    toast('已添加', `${paths.length} 个文件`, 'ok');
  };

  const addDropped = async (fileList) => {
    for (const f of fileList) {
      if (!/\.(xlsx|xlsm|xls)$/i.test(f.name)) continue;
      const b64 = await new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result.split(',')[1] || '');
        reader.readAsDataURL(f);
      });
      const r = await api('notes_add_dropped', f.name, b64);
      if (r.files) setFiles(r.files || []);
    }
  };

  const removeFile = async (name) => {
    const r = await api('notes_remove_file', name);
    setFiles(r.files || []);
  };

  const clearData = async () => {
    const r = await api('notes_clear');
    setFiles(r.files || []);
    setResult(null);
    setEditingRows([]);
  };

  // Import
  const doImport = async () => {
    if (!recvOrg.trim()) { toast('提示', '请先填写收款组织', 'warn'); return; }
    setImporting(true);
    try {
      const res = await api('notes_import', recvOrg);
      if (res.error) throw new Error(res.error);
      setResult(res);
      setEditingRows(res.table_data || []);
    } catch (e) { toast('导入失败', e.message, 'err'); }
    finally { setImporting(false); }
  };

  // Export
  const doExport = async () => {
    if (!editingRows.length) { toast('提示', '没有可导出的数据', 'warn'); return; }
    const out = await api('save_file_dialog', '应收票据导入结果.xlsx');
    if (!out) return;
    setExporting(true);
    try {
      const res = await api('notes_export', editingRows, out);
      if (res.error) throw new Error(res.error);
      toast('导出成功', `已写入 ${res.written || editingRows.length} 行`, 'ok');
    } catch (e) { toast('导出失败', e.message, 'err'); }
    finally { setExporting(false); }
  };

  // Mapping dialog
  const saveMapping = async (maps) => {
    const r = await api('notes_save_mapping', maps);
    if (r.error) throw new Error(r.error);
    setResult(r);
    setEditingRows(r.table_data || []);
  };

  // Cell edit handler
  const updateCell = (rowIdx, field, value) => {
    setEditingRows(prev => prev.map((r, i) => i === rowIdx ? { ...r, [field]: value } : r));
  };

  const readonlySet = new Set(templateInfo?.readonly || []);
  const requiredSet = new Set((templateInfo?.fields || []).filter(f => f.required).map(f => f.header));

  return (
    <div className="page-content"
      onDragOver={e => { e.preventDefault(); e.currentTarget.querySelector('.drop-zone')?.classList.add('active'); }}
      onDragLeave={e => { e.preventDefault(); e.currentTarget.querySelector('.drop-zone')?.classList.remove('active'); }}
      onDrop={async e => { e.preventDefault(); await addDropped(e.dataTransfer.files); }}
    >
      <div className="page-header">
        <h1>📋 应收票据批量导入</h1>
        <p>多来源 Excel → 自动列名匹配 → 可编辑预览 → 导出固定模板</p>
      </div>

      {/* Step 1: Input */}
      <div className="card mb-md">
        <div className="card-title">1️⃣ 数据来源</div>
        
        <div className="drop-zone" id="notesDropZone">
          {Icons.upload}
          <p style={{ marginTop: 8, fontWeight: 500 }}>拖拽 Excel 文件到此处</p>
          <p style={{ fontSize: 12 }}>或点击下方按钮选择文件</p>
        </div>

        <div className="flex gap-sm mt-md">
          <button className="btn btn-secondary" onClick={addFiles}>{Icons.file} 添加文件</button>
          <button className="btn btn-ghost btn-sm" onClick={clearData} disabled={!files.length}>
            {Icons.trash} 清除
          </button>
        </div>

        {files.length > 0 && (
          <div className="file-pill-list mt-md">
            {files.map(f => (
              <div key={f.path} className="file-pill">
                <span style={{ color: 'var(--primary)' }}>{Icons.file}</span>
                <span className="name">{f.name}</span>
                <button className="btn btn-ghost btn-sm" style={{ padding: '2px 6px' }}
                  onClick={() => removeFile(f.name)} title="移除">{Icons.x}</button>
              </div>
            ))}
          </div>
        )}

        <div className="form-group mt-md">
          <label>收款组织 <span className="text-danger">*</span></label>
          <input type="text" placeholder="如：XX有限公司"
            value={recvOrg} onChange={e => setRecvOrg(e.target.value)} />
        </div>

        <button className="btn btn-primary btn-lg" onClick={doImport}
          disabled={importing || !files.length || !recvOrg.trim()}
          style={{ width: '100%', marginTop: 12 }}>
          {importing ? <><span className="spinner" /> 正在导入…</> : <>{Icons.refresh} 导入 / 刷新</>}
        </button>
      </div>

      {/* Step 2: Result Table */}
      {result && (
        <div className="card mb-md">
          <div className="flex items-center justify-between mb-md">
            <div className="card-title">2️⃣ 预览与编辑 ({editingRows.length} 行)</div>
            <div className="flex gap-sm">
              <button className="btn btn-secondary btn-sm"
                onClick={() => setShowMapping(true)}
                disabled={!result.files?.length}>
                {Icons.settings} 配置映射
              </button>
              <button className="btn btn-success" onClick={doExport} disabled={exporting}>
                {exporting ? <span className="spinner" /> : Icons.download} 导出模板
              </button>
            </div>
          </div>

          {/* Stats */}
          <div className="stats-row" style={{ marginBottom: 12 }}>
            <div className="stat-card"><div className="stat-value">{result.row_count}</div><div className="stat-label">总行数</div></div>
            <div className="stat-card"><div className="stat-value">{result.files?.length || 0}</div><div className="stat-label">来源文件</div></div>
            {result.has_missing !== undefined && (
              <div className="stat-card">
                <div className="stat-value" style={{color: result.has_missing ? 'var(--warning)' : 'var(--success)'}}>
                  {result.has_missing ? '有缺失' : '完整'}
                </div>
                <div className="stat-label">必录字段</div>
              </div>
            )}
          </div>

          {/* Log */}
          {result.log_lines?.length > 0 && (
            <div className="log-panel mb-md">
              {result.log_lines.map((line, i) => (
                <div key={i} className="log-line">{line}</div>
              ))}
            </div>
          )}

          {/* Editable table */}
          <div className="table-wrap">
            <table className="data-table" id="notesGrid">
              <thead>
                <tr>
                  {(result.headers || []).map(h => (
                    <th key={h} className={requiredSet.has(h) ? 'req' : ''}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {editingRows.map((rec, ri) => (
                  <tr key={ri}>
                    {(result.headers || []).map(h => {
                      const val = rec[h] ?? '';
                      const isReadonly = readonlySet.has(h);
                      if (isReadonly) {
                        return <td key={h}><span className="text-muted">{val}</span></td>;
                      }
                      return (
                        <td key={h} className="editable">
                          <input value={val} onChange={e => updateCell(ri, h, e.target.value)} />
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Mapping Dialog */}
      {showMapping && templateInfo && (
        <MappingDialog
          files={result?.files || []}
          templateFields={templateInfo.fields || []}
          readonlyHeaders={templateInfo.readonly || []}
          onClose={() => setShowMapping(false)}
          onSave={saveMapping}
        />
      )}
    </div>
  );
}

/* ============================================================
   Mapping Dialog Component
   ============================================================ */
function MappingDialog({ files, templateFields, readonlyHeaders, onClose, onSave }) {
  const [activeFileIdx, setActiveFileIdx] = useState(0);
  const [mapping, setMapping] = useState({});
  const activeFile = files[activeFileIdx];

  // Initialize mapping from current_map of first file
  useEffect(() => {
    if (!activeFile) return;
    const initial = {};
    // Invert current_map from {target: source} to our UI state
    Object.entries(activeFile.current_map || {}).forEach(([tgt, src]) => {
      initial[tgt] = src;
    });
    setMapping(initial);
  }, [activeFileIdx, activeFile]);

  // Non-readonly fields for dialog
  const dialogFields = templateFields.filter(f => !readonlyHeaders.includes(f.header));

  const sourceHeaders = activeFile?.source_headers || [];
  
  // Conflict detection
  const usedSources = new Set(Object.values(mapping).filter(Boolean));
  const hasConflict = usedSources.size !== [...usedSources].length; // duplicate values

  const missingRequired = dialogFields
    .filter(f => f.required && !mapping[f.header])
    .map(f => f.header);

  const handleSave = () => {
    // Build per-file maps dict
    const allMaps = {};
    files.forEach(file => {
      const fm = activeFileIdx === files.indexOf(file) ? mapping :
        (file.current_map || {}); // keep other files as-is
      allMaps[file.name] = fm;
    });
    onSave(allMaps);
    onClose();
  };

  const resetAuto = () => {
    const autoMap = {};
    Object.entries(activeFile.auto_map || {}).forEach(([target, source]) => {
      autoMap[target] = source;
    });
    setMapping(autoMap);
  };

  return (
    <div className="overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3>⚙️ 配置列名映射 — {activeFile?.name}</h3>
          <button className="btn btn-ghost btn-sm" onClick={onClose}>{Icons.x}</button>
        </div>
        <div className="modal-body">

          {/* Multi-file tabs */}
          {files.length > 1 && (
            <div className="flex gap-sm mb-md">
              <select value={activeFileIdx} onChange={e => setActiveFileIdx(Number(e.target.value))}
                style={{ padding: '6px 10px', borderRadius: 'var(--r-xs)', border: '1px solid var(--border)' }}>
                {files.map((f, i) => <option key={i} value={i}>{f.name}</option>)}
              </select>
            </div>
          )}

          {/* Warnings */}
          {hasConflict && (
            <div className="badge badge-red mb-md" style={{ display: 'block', textAlign: 'center' }}>
              ⚠️ 检测到重复的来源列映射！同一列不能映射到多个目标字段。
            </div>
          )}
          {missingRequired.length > 0 && (
            <div className="badge badge-orange mb-md" style={{ display: 'block', textAlign: 'center' }}>
              ⚠️ 必录字段未映射：{missingRequired.join(', ')}
            </div>
          )}

          {/* Mapping table */}
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr><th style={{width:180}}>模板字段</th><th>来源列</th></tr>
              </thead>
              <tbody>
                {dialogFields.map(field => {
                  const hdr = field.header;
                  const val = mapping[hdr] || '';
                  const isDup = val && [...Object.entries(mapping)]
                    .filter(([k, v]) => k !== hdr && v === val).length > 0;
                  const isMissing = field.required && !val;

                  return (
                    <tr key={hdr} style={{ background: isMissing ? 'rgba(227,116,0,.05)' : '' }}>
                      <td style={{ fontWeight: field.required ? 600 : 400 }}>
                        {field.required && <span className="text-danger" title="必录">*</span>} {hdr}
                      </td>
                      <td>
                        <select value={val}
                          onChange={e => setMapping(prev => ({ ...prev, [hdr]: e.target.value || '' }))}
                          style={{
                            width: '100%',
                            padding: '5px 8px',
                            border: `1px solid ${isDup ? 'var(--danger)' : isMissing ? 'var(--warning)' : 'var(--border)'}`,
                            borderRadius: 'var(--r-xs)',
                            background: 'var(--surface-1)',
                            color: 'var(--text)',
                            outline: 'none',
                          }}>
                          <option value="">〈 不匹配 〉</option>
                          {sourceHeaders.map(sh => (
                            <option key={sh} value={sh}>{sh}</option>
                          ))}
                        </select>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-ghost" onClick={resetAuto}>恢复自动匹配</button>
          <button className="btn btn-ghost" onClick={onClose}>取消</button>
          <button className="btn btn-primary" onClick={handleSave}>确定</button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   App 主组件
   ============================================================ */
export default function App() {
  const [features, setFeatures] = useState([
    { name: 'bank_classify', icon: 'landmark', title: '票据银行分类' },
    { name: 'js_statement', icon: 'receipt', title: '江苏银行对账单复制' },
    { name: 'notes_import', icon: 'table', title: '应收票据批量导入' },
  ]);
  const [currentView, setCurrentView] = useState('bank_classify');
  const [collapsed, setCollapsed] = useState(false);
  const [theme, setTheme] = useState(() =>
    localStorage.getItem('theme') || 'light'
  );
  const { toasts, addToast } = useToasts();

  // Apply theme to document
  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
  }, [theme]);

  // Fetch meta on mount
  useEffect(() => {
    callApi('get_meta').then(meta => {
      if (meta.features) setFeatures(meta.features);
    }).catch(() => {});
  }, []);

  const toggleTheme = () => setTheme(prev => prev === 'light' ? 'dark' : 'light');

  const renderView = () => {
    switch (currentView) {
      case 'bank_classify': return <BankClassifyView />;
      case 'js_statement': return <JsStatementView />;
      case 'notes_import': return <NotesImportView />;
      default: return <BankClassifyView />;
    }
  };

  const currentTitle = features.find(f => f.name === currentView)?.title || '';

  return (
    <ApiContext.Provider value={callApi}>
      <ThemeContext.Provider value={{ theme, toggle: toggleTheme }}>
        <ToastContext.Provider value={addToast}>
          <div className="app-layout" data-theme={theme}>
            <Sidebar collapsed={collapsed} setCollapsed={setCollapsed}
              currentView={currentView} navigate={setCurrentView} features={features} />
            <div className="main-area">
              <TopBar title={currentTitle} onThemeToggle={toggleTheme} isDark={theme === 'dark'} />
              {renderView()}
            </div>
          </div>
          <ToastContainer toasts={toasts} />
        </ToastContext.Provider>
      </ThemeContext.Provider>
    </ApiContext.Provider>
  );
}
