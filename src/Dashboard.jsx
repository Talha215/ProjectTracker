import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const PRIORITY_COLORS = {
  high: { bg: "#FF6B6B20", border: "#FF6B6B", text: "#FF6B6B", label: "High" },
  medium: { bg: "#F0A84020", border: "#F0A840", text: "#F0A840", label: "Med" },
  low: { bg: "#4ECDC420", border: "#4ECDC4", text: "#4ECDC4", label: "Low" },
};

const STATUS_OPTIONS = ["Not Started", "In Progress", "Blocked", "Done"];
const STATUS_COLORS = {
  "Not Started": "#6B7280",
  "In Progress": "#5B8DEF",
  Blocked: "#FF6B6B",
  Done: "#4ECDC4",
};

const TILE_COLORS = ["#5B8DEF", "#4ECDC4", "#F0A840", "#C084FC", "#FF6B6B", "#F472B6", "#A78BFA", "#34D399", "#FB923C", "#38BDF8"];

const uid = () => crypto.randomUUID().slice(0, 8);

const daysUntil = (d) => {
  if (!d) return null;
  const now = new Date(); now.setHours(0, 0, 0, 0);
  return Math.ceil((new Date(d + "T00:00:00") - now) / 86400000);
};

const dlInfo = (d) => {
  const days = daysUntil(d);
  if (days === null) return null;
  if (days < 0) return { label: `${Math.abs(days)}d over`, color: "#FF6B6B" };
  if (days === 0) return { label: "Today", color: "#FF6B6B" };
  if (days <= 3) return { label: `${days}d`, color: "#F0A840" };
  if (days <= 7) return { label: `${days}d`, color: "#5B8DEF" };
  return { label: `${days}d`, color: "#6B7280" };
};

const STORAGE_KEY = "pm-dashboard-data";

function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function save(d) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(d));
  } catch (e) {
    console.error("Save failed:", e);
  }
}

const DEFAULTS = {
  projects: [
    { id: uid(), name: "Workforce", color: "#5B8DEF", tasks: [] },
    { id: uid(), name: "Seniors", color: "#4ECDC4", tasks: [] },
    { id: uid(), name: "DOH", color: "#F0A840", tasks: [] },
    { id: uid(), name: "CSC", color: "#C084FC", tasks: [] },
    { id: uid(), name: "MOPD Employment", color: "#FF6B6B", tasks: [] },
    { id: uid(), name: "LIRI", color: "#F472B6", tasks: [] },
    { id: uid(), name: "MOPD", color: "#A78BFA", tasks: [] },
    { id: uid(), name: "HOP", color: "#34D399", tasks: [] },
    { id: uid(), name: "DOF", color: "#FB923C", tasks: [] },
    { id: uid(), name: "CDPH", color: "#38BDF8", tasks: [] },
    { id: uid(), name: "CSLSI", color: "#5B8DEF", tasks: [] },
  ],
};

// ── CSV helpers ────────────────────────────────────────────────────────────
function parseCsv(text) {
  const rows = []; let cols = []; let cur = ''; let inQ = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') { inQ = !inQ; }
    else if (ch === ',' && !inQ) { cols.push(cur.trim()); cur = ''; }
    else if ((ch === '\n' || ch === '\r') && !inQ) {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      cols.push(cur.trim()); cur = '';
      if (cols.some(c => c)) rows.push(cols);
      cols = [];
    } else { cur += ch; }
  }
  if (cur || cols.length) { cols.push(cur.trim()); if (cols.some(c => c)) rows.push(cols); }
  return rows;
}

function parseDate(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m4 = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m4) return `${m4[3]}-${m4[1].padStart(2, '0')}-${m4[2].padStart(2, '0')}`;
  const m2 = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2})(?!\d)/);
  if (m2) { const yr = parseInt(m2[3]) > 50 ? '19' + m2[3] : '20' + m2[3]; return `${yr}-${m2[1].padStart(2, '0')}-${m2[2].padStart(2, '0')}`; }
  try { const d = new Date(raw); if (!isNaN(d)) return d.toISOString().slice(0, 10); } catch {}
  return '';
}

function smartsheetToTasks(rows) {
  if (rows.length < 2) return null;
  const headers = rows[0].map(h => String(h ?? '').toLowerCase().replace(/[^a-z0-9 #]/g, '').trim());
  const col = (...names) => { for (const n of names) { const i = headers.findIndex(h => h === n || h.includes(n)); if (i >= 0) return i; } return -1; };

  const titleCol = col('task name', 'name', 'row name', 'summary', 'title', 'item');
  if (titleCol < 0) return null;
  const statusCol = col('status');
  const priorityCol = col('priority');
  const dateCol = col('completion date', 'due date', 'deadline', 'end date', 'finish');
  const programCol = col('program');
  const ticketCol = col('ticket');

  const tasks = rows.slice(1).map(r => {
    const get = (i) => (i >= 0 && i < r.length ? String(r[i] ?? '').trim() : '');
    const rawTitle = get(titleCol); if (!rawTitle) return null;
    const rawStatus = get(statusCol).toLowerCase();
    const rawPriority = get(priorityCol).toLowerCase();
    const ticket = get(ticketCol);
    const program = get(programCol);

    let status = 'Not Started';
    if (/complete|done|finish|verified|closed/.test(rawStatus)) status = 'Done';
    else if (/progress|start|active|ongoing/.test(rawStatus)) status = 'In Progress';
    else if (/block|hold|stuck|waiting/.test(rawStatus)) status = 'Blocked';
    // "to-do", "to do", "not started", etc. → stays Not Started

    let priority = 'medium';
    if (/high|critical|urgent/.test(rawPriority)) priority = 'high';
    else if (/low/.test(rawPriority)) priority = 'low';

    const deadline = parseDate(get(dateCol));
    const title = ticket ? `[#${ticket}] ${rawTitle}` : rawTitle;

    return { id: uid(), title, status, priority, deadline, _program: program || null };
  }).filter(Boolean);

  const hasProgram = programCol >= 0 && tasks.some(t => t._program);
  return { tasks, hasProgram };
}

function ImportModal({ parsed, projects, onImport, onClose }) {
  const [targetId, setTargetId] = useState(parsed.hasProgram ? '__program__' : '__new__');
  const active = parsed.tasks.filter(t => t.status !== 'Done').length;
  const done = parsed.tasks.length - active;

  // Group by program for preview
  const programGroups = {};
  if (parsed.hasProgram) {
    parsed.tasks.forEach(t => {
      const p = t._program || 'Uncategorized';
      programGroups[p] = (programGroups[p] || 0) + 1;
    });
  }

  return (
    <div style={{
      position: 'fixed', inset: 0, background: '#000000aa', zIndex: 100,
      display: 'flex', alignItems: 'center', justifyContent: 'center',
    }} onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{
        background: 'var(--c-surface)', borderRadius: 16, border: '1px solid var(--c-border)',
        padding: '24px', width: 440, maxWidth: '90vw',
      }}>
        <h3 style={{ margin: '0 0 4px', fontSize: 16, fontWeight: 700 }}>Import from SmartSheet</h3>
        <p style={{ margin: '0 0 16px', fontSize: 12, color: 'var(--c-muted)' }}>
          <strong style={{ color: 'var(--c-text)' }}>{parsed.tasks.length} tasks</strong> found in <em>{parsed.fileName}</em>
          {done > 0 && <> &nbsp;·&nbsp; {active} active, {done} completed</>}
        </p>

        {parsed.hasProgram && (
          <div style={{ marginBottom: 16, padding: '10px 12px', background: 'var(--c-bg)', borderRadius: 8, border: '1px solid var(--c-border)' }}>
            <div style={{ fontSize: 11, color: 'var(--c-muted)', marginBottom: 6, fontWeight: 600 }}>Programs detected</div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
              {Object.entries(programGroups).map(([name, count]) => (
                <span key={name} style={{ fontSize: 11, padding: '2px 8px', borderRadius: 4, background: 'var(--c-surface)', border: '1px solid var(--c-border)', color: 'var(--c-text)' }}>
                  {name} <span style={{ color: 'var(--c-muted)' }}>({count})</span>
                </span>
              ))}
            </div>
          </div>
        )}

        <label style={{ fontSize: 12, color: 'var(--c-muted)', display: 'block', marginBottom: 6 }}>Import into</label>
        <select value={targetId} onChange={e => setTargetId(e.target.value)} style={{
          width: '100%', padding: '8px 10px', borderRadius: 8, fontSize: 13,
          background: 'var(--c-bg)', border: '1px solid var(--c-border)',
          color: 'var(--c-text)', fontFamily: 'inherit', marginBottom: 20,
        }}>
          {parsed.hasProgram && <option value="__program__">Auto-sort by Program column</option>}
          <option value="__new__">New project: "{parsed.projectName}"</option>
          {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>

        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
          <button onClick={onClose} style={{
            padding: '8px 16px', borderRadius: 8, border: '1px solid var(--c-border)',
            background: 'none', color: 'var(--c-muted)', fontSize: 13, cursor: 'pointer', fontFamily: 'inherit',
          }}>Cancel</button>
          <button onClick={() => onImport(targetId, parsed)} style={{
            padding: '8px 18px', borderRadius: 8, border: 'none',
            background: '#5B8DEF', color: '#fff', fontSize: 13, fontWeight: 600, cursor: 'pointer', fontFamily: 'inherit',
          }}>Import {parsed.tasks.length} tasks</button>
        </div>
      </div>
    </div>
  );
}

// ── Flags (Today's Focus) — stored separately so they survive Clear All ──────
const FLAGS_KEY = "pm-dashboard-flags";
function loadFlags() { try { return new Set(JSON.parse(localStorage.getItem(FLAGS_KEY) || "[]")); } catch { return new Set(); } }
function saveFlags(s) { try { localStorage.setItem(FLAGS_KEY, JSON.stringify([...s])); } catch {} }

function TaskRow({ task, accentColor, flagged, onToggleFlag }) {
  const pri = PRIORITY_COLORS[task.priority];
  const dl = dlInfo(task.deadline);
  const isDone = task.status === "Done";

  return (
    <div style={{
      display: "flex", alignItems: "center", gap: 7,
      padding: "5px 8px", borderRadius: 6, marginBottom: 3,
      background: isDone ? "transparent" : flagged ? "#F0A84010" : "var(--c-row)",
      opacity: isDone ? 0.5 : 1, transition: "background .15s",
      border: flagged && !isDone ? "1px solid #F0A84033" : "1px solid transparent",
    }}>
      <button onClick={() => onToggleFlag(task.id)} title={flagged ? "Remove from focus" : "Add to today's focus"} style={{
        background: "none", border: "none", cursor: "pointer", padding: 0,
        fontSize: 13, lineHeight: 1, flexShrink: 0,
        color: flagged ? "#F0A840" : "var(--c-muted)", opacity: flagged ? 1 : 0.35,
        transition: "all .15s",
      }}>★</button>

      <span style={{
        flex: 1, fontSize: 12, color: "var(--c-text)",
        textDecoration: isDone ? "line-through" : "none",
        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
      }}>{task.title}</span>

      {pri && <span style={{ fontSize: 9, fontWeight: 700, color: pri.text, letterSpacing: 0.3, flexShrink: 0 }}>{pri.label.toUpperCase()}</span>}

      <span style={{
        display: "inline-block", width: 7, height: 7, borderRadius: 99, flexShrink: 0,
        background: STATUS_COLORS[task.status],
      }} title={task.status} />

      {dl && <span style={{ fontSize: 10, fontWeight: 600, color: dl.color, flexShrink: 0, minWidth: 46, textAlign: "right" }}>{dl.label}</span>}
    </div>
  );
}

function FocusPanel({ flaggedIds, projects, onToggleFlag }) {
  const flaggedTasks = projects.flatMap(p =>
    p.tasks.filter(t => flaggedIds.has(t.id)).map(t => ({ ...t, pName: p.name, pColor: p.color }))
  );
  if (flaggedTasks.length === 0) return null;
  return (
    <div style={{
      marginBottom: 20, borderRadius: 12, overflow: "hidden",
      border: "1px solid #F0A84033", background: "var(--c-surface)",
    }}>
      <div style={{ padding: "10px 16px", borderBottom: "1px solid #F0A84022", display: "flex", alignItems: "center", gap: 8 }}>
        <span style={{ fontSize: 11, fontWeight: 700, color: "#F0A840", letterSpacing: 0.8 }}>TODAY'S FOCUS</span>
        <span style={{ fontSize: 11, color: "#F0A840", opacity: 0.7 }}>{flaggedTasks.length} item{flaggedTasks.length !== 1 ? "s" : ""}</span>
      </div>
      {flaggedTasks.map(t => {
        const dl = dlInfo(t.deadline);
        return (
          <div key={t.id} style={{
            display: "flex", alignItems: "center", gap: 10,
            padding: "8px 16px", borderBottom: "1px solid var(--c-border)",
          }}>
            <button onClick={() => onToggleFlag(t.id)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 13, color: "#F0A840", padding: 0, flexShrink: 0 }}>★</button>
            <div style={{ width: 8, height: 8, borderRadius: 99, background: t.pColor, flexShrink: 0 }} />
            <span style={{ flex: 1, fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.title}</span>
            <span style={{ fontSize: 11, color: "var(--c-muted)", flexShrink: 0 }}>{t.pName}</span>
            {PRIORITY_COLORS[t.priority] && <span style={{ fontSize: 9, fontWeight: 700, color: PRIORITY_COLORS[t.priority].text, flexShrink: 0 }}>{PRIORITY_COLORS[t.priority].label.toUpperCase()}</span>}
            {dl && <span style={{ fontSize: 11, fontWeight: 600, color: dl.color, flexShrink: 0 }}>{dl.label}</span>}
          </div>
        );
      })}
    </div>
  );
}

function ProjectDetail({ project, flagged, onToggleFlag, onBack }) {
  const [showCompleted, setShowCompleted] = useState(false);
  const activeTasks = project.tasks.filter(t => t.status !== "Done");
  const doneTasks = project.tasks.filter(t => t.status === "Done");
  const total = project.tasks.length;
  const pct = total ? Math.round((doneTasks.length / total) * 100) : 0;
  const overdue = activeTasks.filter(t => daysUntil(t.deadline) !== null && daysUntil(t.deadline) < 0).length;

  return (
    <div style={{ maxWidth: 720, margin: "0 auto" }}>
      <button onClick={onBack} style={{
        display: "inline-flex", alignItems: "center", gap: 6,
        background: "none", border: "1px solid var(--c-border)", borderRadius: 8,
        color: "var(--c-muted)", fontSize: 12, cursor: "pointer",
        padding: "5px 12px", marginBottom: 20, fontFamily: "inherit",
      }}
        onMouseEnter={e => { e.currentTarget.style.color = "var(--c-text)"; e.currentTarget.style.borderColor = "var(--c-text)"; }}
        onMouseLeave={e => { e.currentTarget.style.color = "var(--c-muted)"; e.currentTarget.style.borderColor = "var(--c-border)"; }}
      >← Back to Projects</button>

      <div style={{ background: "var(--c-surface)", borderRadius: 16, border: "1px solid var(--c-border)", overflow: "hidden" }}>
        <div style={{ height: 5, background: `linear-gradient(90deg, ${project.color}, ${project.color}55)` }} />

        <div style={{ padding: "20px 24px 16px", borderBottom: "1px solid var(--c-border)" }}>
          <h2 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: "var(--c-text)" }}>{project.name}</h2>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 8 }}>
            <span style={{ fontSize: 12, color: "var(--c-muted)" }}>{doneTasks.length} of {total} done</span>
            {pct > 0 && <span style={{ fontSize: 12, fontWeight: 600, color: project.color }}>{pct}%</span>}
            {overdue > 0 && <span style={{ fontSize: 12, fontWeight: 600, color: "#FF6B6B" }}>⚠ {overdue} overdue</span>}
          </div>
          <div style={{ height: 4, background: "var(--c-border)", borderRadius: 2, marginTop: 10, maxWidth: 300 }}>
            <div style={{ height: "100%", width: `${pct}%`, background: project.color, borderRadius: 2, transition: "width .3s" }} />
          </div>
        </div>

        <div style={{ padding: "12px 16px" }}>
          {activeTasks.length === 0 && doneTasks.length === 0 && (
            <p style={{ fontSize: 13, color: "var(--c-muted)", textAlign: "center", padding: "32px 0", margin: 0 }}>No tasks</p>
          )}
          {activeTasks.length === 0 && doneTasks.length > 0 && (
            <p style={{ fontSize: 13, color: project.color, textAlign: "center", padding: "24px 0 12px", margin: 0, fontWeight: 600 }}>All tasks completed ✓</p>
          )}
          {activeTasks.map(t => (
            <TaskRow key={t.id} task={t} accentColor={project.color} flagged={flagged.has(t.id)} onToggleFlag={onToggleFlag} />
          ))}
        </div>

        {doneTasks.length > 0 && (
          <div style={{ borderTop: "1px solid var(--c-border)" }}>
            <button onClick={() => setShowCompleted(v => !v)} style={{
              width: "100%", padding: "10px 16px", background: "none", border: "none",
              cursor: "pointer", display: "flex", alignItems: "center", gap: 6,
              color: "var(--c-muted)", fontSize: 12, fontFamily: "inherit",
            }}>
              <span style={{ fontSize: 10, transition: "transform .15s", transform: showCompleted ? "rotate(90deg)" : "none", display: "inline-block" }}>▶</span>
              Completed ({doneTasks.length})
            </button>
            {showCompleted && (
              <div style={{ padding: "4px 16px 12px" }}>
                {doneTasks.map(t => (
                  <TaskRow key={t.id} task={t} accentColor={project.color} flagged={flagged.has(t.id)} onToggleFlag={onToggleFlag} />
                ))}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

function Tile({ project, flagged, onToggleFlag, onFocus }) {
  const activeTasks = project.tasks.filter(t => t.status !== "Done");
  const total = project.tasks.length;
  const done = total - activeTasks.length;
  const pct = total ? Math.round((done / total) * 100) : 0;
  const overdue = activeTasks.filter(t => daysUntil(t.deadline) !== null && daysUntil(t.deadline) < 0).length;

  return (
    <div style={{
      background: "var(--c-surface)", borderRadius: 16,
      border: "1px solid var(--c-border)", display: "flex", flexDirection: "column",
      aspectRatio: "1 / 1", overflow: "hidden",
    }}>
      <div style={{ height: 4, background: `linear-gradient(90deg, ${project.color}, ${project.color}66)`, flexShrink: 0 }} />

      <div onClick={onFocus} style={{ padding: "14px 16px 8px", flexShrink: 0, cursor: "pointer" }}>
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 6 }}>
          <h2 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "var(--c-text)", lineHeight: 1.3, flex: 1, minWidth: 0 }}>{project.name}</h2>
          <span style={{ fontSize: 11, color: "var(--c-muted)", opacity: 0.5, flexShrink: 0, marginTop: 2 }}>⤢</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 6 }}>
          <span style={{ fontSize: 11, color: "var(--c-muted)" }}>{done}/{total}</span>
          {pct > 0 && <span style={{ fontSize: 11, color: project.color, fontWeight: 600 }}>{pct}%</span>}
          {overdue > 0 && <span style={{ fontSize: 10, fontWeight: 600, color: "#FF6B6B" }}>⚠ {overdue}</span>}
        </div>
        <div style={{ height: 3, background: "var(--c-border)", borderRadius: 2, marginTop: 6 }}>
          <div style={{ height: "100%", width: `${pct}%`, background: project.color, borderRadius: 2, transition: "width .3s" }} />
        </div>
      </div>

      <div style={{ flex: 1, overflowY: "auto", padding: "4px 10px", minHeight: 0 }}>
        {activeTasks.length === 0 && done === 0 && (
          <p style={{ fontSize: 12, color: "var(--c-muted)", textAlign: "center", padding: "20px 0", margin: 0, opacity: 0.6 }}>No tasks</p>
        )}
        {activeTasks.length === 0 && done > 0 && (
          <p style={{ fontSize: 12, color: project.color, textAlign: "center", padding: "20px 0", margin: 0, fontWeight: 600 }}>All done ✓</p>
        )}
        {activeTasks.map(t => (
          <TaskRow key={t.id} task={t} accentColor={project.color} flagged={flagged.has(t.id)} onToggleFlag={onToggleFlag} />
        ))}
      </div>
    </div>
  );
}

export default function Dashboard() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [view, setView] = useState("tiles");
  const [focusedProjectId, setFocusedProjectId] = useState(null);
  const [importModal, setImportModal] = useState(null);
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [flagged, setFlagged] = useState(() => loadFlags());
  const fileInputRef = useRef(null);

  useEffect(() => {
    const loaded = load();
    setData(loaded || { projects: [] });
    setLoading(false);
  }, []);

  const toggleFlag = (taskId) => {
    setFlagged(prev => {
      const next = new Set(prev);
      next.has(taskId) ? next.delete(taskId) : next.add(taskId);
      saveFlags(next);
      return next;
    });
  };

  const persist = useCallback((d) => { setData(d); save(d); }, []);
  const updateProject = (p) => persist({ ...data, projects: data.projects.map(x => x.id === p.id ? p : x) });

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    e.target.value = '';
    const isXls = /\.xlsx?$/i.test(file.name);
    const reader = new FileReader();
    reader.onload = (ev) => {
      let rows;
      if (isXls) {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
      } else {
        rows = parseCsv(ev.target.result);
      }
      const result = smartsheetToTasks(rows);
      if (!result || result.tasks.length === 0) { alert('No tasks found. Make sure the file has a column named "Task Name", "Name", or "Title".'); return; }
      const projectName = file.name.replace(/\.(csv|xlsx?)$/i, '');
      setImportModal({ ...result, fileName: file.name, projectName });
    };
    if (isXls) reader.readAsArrayBuffer(file); else reader.readAsText(file);
  };

  const handleImport = (targetId, parsed) => {
    const cleanTasks = (tasks) => tasks.map(({ _program, ...t }) => t);
    let newData;
    if (targetId === '__program__') {
      // Auto-sort by Program column — match existing projects by name (case-insensitive), create new for unmatched
      const groups = {};
      parsed.tasks.forEach(t => { const p = t._program || 'Uncategorized'; (groups[p] = groups[p] || []).push(t); });
      let projects = [...data.projects];
      for (const [progName, tasks] of Object.entries(groups)) {
        const existing = projects.find(p => p.name.toLowerCase() === progName.toLowerCase());
        if (existing) {
          projects = projects.map(p => p.id === existing.id ? { ...p, tasks: [...p.tasks, ...cleanTasks(tasks)] } : p);
        } else {
          projects.push({ id: uid(), name: progName, color: TILE_COLORS[projects.length % TILE_COLORS.length], tasks: cleanTasks(tasks) });
        }
      }
      newData = { ...data, projects };
    } else if (targetId === '__new__') {
      const color = TILE_COLORS[data.projects.length % TILE_COLORS.length];
      newData = { ...data, projects: [...data.projects, { id: uid(), name: parsed.projectName, color, tasks: cleanTasks(parsed.tasks) }] };
    } else {
      newData = { ...data, projects: data.projects.map(p => p.id === targetId ? { ...p, tasks: [...p.tasks, ...cleanTasks(parsed.tasks)] } : p) };
    }
    persist(newData);
    setImportModal(null);
  };

  if (loading) return <div style={{ padding: 60, textAlign: "center", color: "#6B7280", fontFamily: "system-ui" }}>Loading...</div>;

  const allTasks = data.projects.flatMap(p => p.tasks.map(t => ({ ...t, pName: p.name, pColor: p.color })));
  const totalTasks = allTasks.length;
  const doneTasks = allTasks.filter(t => t.status === "Done").length;
  const overdueTasks = allTasks.filter(t => t.status !== "Done" && daysUntil(t.deadline) !== null && daysUntil(t.deadline) < 0);
  const upcoming7 = allTasks.filter(t => t.status !== "Done" && daysUntil(t.deadline) !== null && daysUntil(t.deadline) >= 0 && daysUntil(t.deadline) <= 7);
  const filtered = search
    ? data.projects.map(p => ({ ...p, tasks: p.tasks.filter(t => t.title.toLowerCase().includes(search.toLowerCase())) }))
        .filter(p => p.name.toLowerCase().includes(search.toLowerCase()) || p.tasks.length > 0)
    : data.projects;
  const deadlineTasks = allTasks.filter(t => t.status !== "Done" && t.deadline).sort((a, b) => a.deadline.localeCompare(b.deadline));

  return (
    <div style={{
      "--c-bg": "#0D0F14", "--c-surface": "#161921", "--c-row": "#1C1F2A",
      "--c-border": "#252A37", "--c-text": "#E2E5ED", "--c-muted": "#6B7280",
      fontFamily: "'DM Sans', 'Segoe UI', system-ui, sans-serif",
      background: "var(--c-bg)", color: "var(--c-text)", minHeight: "100vh", padding: "20px 24px",
    }}>

      {/* ── Toolbar ── */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 14 }}>
          <h1 style={{ margin: 0, fontSize: 20, fontWeight: 700, letterSpacing: -0.5 }}>Projects</h1>
          {totalTasks > 0 && <span style={{ fontSize: 12, color: "var(--c-muted)" }}>{doneTasks}/{totalTasks} done</span>}
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          {data.projects.length > 0 && (
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search..."
              style={{ fontSize: 12, padding: "6px 12px", borderRadius: 8, border: "1px solid var(--c-border)", background: "var(--c-surface)", color: "var(--c-text)", width: 140, fontFamily: "inherit" }} />
          )}
          {data.projects.length > 0 && (
            <button onClick={() => setView(view === "tiles" ? "deadlines" : "tiles")} style={{
              padding: "6px 12px", borderRadius: 8, border: "1px solid var(--c-border)",
              background: view === "deadlines" ? "var(--c-row)" : "var(--c-surface)",
              color: "var(--c-text)", fontSize: 12, cursor: "pointer", fontFamily: "inherit",
            }}>{view === "tiles" ? "⏰ Deadlines" : "◻ Tiles"}</button>
          )}
          {data.projects.length > 0 && (
            <button onClick={() => setShowClearConfirm(true)} style={{
              padding: "6px 12px", borderRadius: 8, border: "1px solid var(--c-border)",
              background: "none", color: "var(--c-muted)", fontSize: 12, cursor: "pointer", fontFamily: "inherit",
            }}>Clear All</button>
          )}
          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleFileChange} style={{ display: "none" }} />
          <button onClick={() => fileInputRef.current.click()} style={{
            padding: "6px 14px", borderRadius: 8, border: "none",
            background: "#5B8DEF", color: "#fff", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit",
          }}>↑ Import</button>
        </div>
      </div>

      {/* ── Modals ── */}
      {importModal && <ImportModal parsed={importModal} projects={data.projects} onImport={handleImport} onClose={() => setImportModal(null)} />}
      {showClearConfirm && (
        <div style={{ position: "fixed", inset: 0, background: "#000000aa", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" }}
          onClick={e => e.target === e.currentTarget && setShowClearConfirm(false)}>
          <div style={{ background: "var(--c-surface)", borderRadius: 16, border: "1px solid var(--c-border)", padding: "24px", width: 360, maxWidth: "90vw" }}>
            <h3 style={{ margin: "0 0 8px", fontSize: 16, fontWeight: 700 }}>Clear all data?</h3>
            <p style={{ fontSize: 13, color: "var(--c-muted)", margin: "0 0 20px" }}>All projects and tasks will be removed. Import a fresh SmartSheet file to reload.</p>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setShowClearConfirm(false)} style={{ padding: "8px 16px", borderRadius: 8, border: "1px solid var(--c-border)", background: "none", color: "var(--c-muted)", fontSize: 13, cursor: "pointer", fontFamily: "inherit" }}>Cancel</button>
              <button onClick={() => { persist({ projects: [] }); setShowClearConfirm(false); setFocusedProjectId(null); }} style={{ padding: "8px 16px", borderRadius: 8, border: "none", background: "#FF6B6B", color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>Clear All</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Status pills ── */}
      {(overdueTasks.length > 0 || upcoming7.length > 0) && (
        <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
          {overdueTasks.length > 0 && <div style={{ padding: "5px 12px", borderRadius: 6, background: "#FF6B6B14", border: "1px solid #FF6B6B33", fontSize: 11, color: "#FF6B6B", fontWeight: 600 }}>{overdueTasks.length} overdue</div>}
          {upcoming7.length > 0 && <div style={{ padding: "5px 12px", borderRadius: 6, background: "#F0A84014", border: "1px solid #F0A84033", fontSize: 11, color: "#F0A840", fontWeight: 600 }}>{upcoming7.length} due this week</div>}
        </div>
      )}

      {/* ── Today's Focus panel ── */}
      <FocusPanel flaggedIds={flagged} projects={data.projects} onToggleFlag={toggleFlag} />

      {/* ── Empty state ── */}
      {data.projects.length === 0 && (
        <div style={{ textAlign: "center", padding: "100px 0", color: "var(--c-muted)" }}>
          <div style={{ fontSize: 36, marginBottom: 16, opacity: 0.4 }}>📊</div>
          <p style={{ fontSize: 15, margin: "0 0 6px", color: "var(--c-text)" }}>No data yet</p>
          <p style={{ fontSize: 12, margin: "0 0 24px" }}>Import your SmartSheet file to get started</p>
          <button onClick={() => fileInputRef.current.click()} style={{ padding: "10px 24px", borderRadius: 8, border: "none", background: "#5B8DEF", color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>↑ Import SmartSheet</button>
        </div>
      )}

      {/* ── Tiles view ── */}
      {view === "tiles" && data.projects.length > 0 && (() => {
        const focused = focusedProjectId && data.projects.find(p => p.id === focusedProjectId);
        if (focused) return (
          <ProjectDetail project={focused} flagged={flagged} onToggleFlag={toggleFlag} onBack={() => setFocusedProjectId(null)} />
        );
        return (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 14 }}>
            {filtered.map(p => (
              <Tile key={p.id} project={p} flagged={flagged} onToggleFlag={toggleFlag} onFocus={() => setFocusedProjectId(p.id)} />
            ))}
          </div>
        );
      })()}

      {/* ── Deadlines view ── */}
      {view === "deadlines" && data.projects.length > 0 && (
        <div style={{ maxWidth: 720 }}>
          {deadlineTasks.length === 0 && <p style={{ color: "var(--c-muted)", fontSize: 12, textAlign: "center", padding: 40 }}>No upcoming deadlines.</p>}
          {deadlineTasks.map(t => {
            const dl = dlInfo(t.deadline);
            return (
              <div key={t.id} style={{
                display: "flex", alignItems: "center", gap: 10, padding: "8px 12px",
                borderRadius: 8, background: "var(--c-surface)", border: "1px solid var(--c-border)", marginBottom: 4,
              }}>
                <button onClick={() => toggleFlag(t.id)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 13, color: flagged.has(t.id) ? "#F0A840" : "var(--c-muted)", opacity: flagged.has(t.id) ? 1 : 0.35, padding: 0, flexShrink: 0 }}>★</button>
                <div style={{ width: 8, height: 8, borderRadius: 99, background: t.pColor, flexShrink: 0 }} />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 12, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.title}</div>
                  <div style={{ fontSize: 10, color: "var(--c-muted)" }}>{t.pName}</div>
                </div>
                {PRIORITY_COLORS[t.priority] && <span style={{ fontSize: 10, fontWeight: 700, color: PRIORITY_COLORS[t.priority].text }}>{PRIORITY_COLORS[t.priority].label.toUpperCase()}</span>}
                <span style={{ fontSize: 11, fontWeight: 600, color: dl?.color, whiteSpace: "nowrap" }}>{t.deadline}</span>
                {dl && <span style={{ fontSize: 10, fontWeight: 700, color: dl.color, whiteSpace: "nowrap" }}>{dl.label}</span>}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
