import { useState, useEffect, useCallback } from "react";

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

function TaskRow({ task, onUpdate, onDelete, accentColor }) {
  const [editing, setEditing] = useState(false);
  const [title, setTitle] = useState(task.title);
  const pri = PRIORITY_COLORS[task.priority];
  const dl = dlInfo(task.deadline);
  const isDone = task.status === "Done";

  const commit = () => { if (title.trim()) onUpdate({ ...task, title: title.trim() }); setEditing(false); };

  return (
    <div style={{
      display: "flex", alignItems: "center", gap: 7,
      padding: "5px 8px", borderRadius: 6, marginBottom: 3,
      background: isDone ? "transparent" : "var(--c-row)",
      opacity: isDone ? 0.45 : 1, transition: "all .15s",
    }}>
      <button onClick={() => onUpdate({ ...task, status: isDone ? "In Progress" : "Done" })} style={{
        width: 16, height: 16, borderRadius: 4, flexShrink: 0, cursor: "pointer",
        border: `1.5px solid ${isDone ? STATUS_COLORS.Done : accentColor + "88"}`,
        background: isDone ? STATUS_COLORS.Done : "transparent",
        color: "#fff", fontSize: 10, lineHeight: 1,
        display: "flex", alignItems: "center", justifyContent: "center",
      }}>{isDone ? "✓" : ""}</button>

      <div style={{ flex: 1, minWidth: 0 }}>
        {editing ? (
          <input value={title} onChange={e => setTitle(e.target.value)}
            onBlur={commit} onKeyDown={e => e.key === "Enter" && commit()} autoFocus
            style={{ width: "100%", background: "var(--c-bg)", border: "1px solid var(--c-border)", borderRadius: 4, padding: "2px 6px", color: "var(--c-text)", fontSize: 12, fontFamily: "inherit" }} />
        ) : (
          <span onClick={() => setEditing(true)} style={{
            fontSize: 12, color: "var(--c-text)", cursor: "text",
            textDecoration: isDone ? "line-through" : "none",
            display: "block", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
          }}>{task.title}</span>
        )}
      </div>

      <button onClick={() => {
        const k = ["low", "medium", "high"];
        onUpdate({ ...task, priority: k[(k.indexOf(task.priority) + 1) % 3] });
      }} style={{ background: "none", border: "none", cursor: "pointer", padding: 0 }}>
        <span style={{ fontSize: 9, fontWeight: 700, color: pri.text, letterSpacing: 0.3 }}>{pri.label.toUpperCase()}</span>
      </button>

      <button onClick={() => {
        const next = STATUS_OPTIONS[(STATUS_OPTIONS.indexOf(task.status) + 1) % STATUS_OPTIONS.length];
        onUpdate({ ...task, status: next });
      }} style={{ background: "none", border: "none", cursor: "pointer", padding: 0 }}>
        <span style={{
          display: "inline-block", width: 7, height: 7, borderRadius: 99,
          background: STATUS_COLORS[task.status],
        }} title={task.status} />
      </button>

      <div style={{ position: "relative", flexShrink: 0, width: 50, textAlign: "right" }}>
        {dl ? (
          <span style={{ fontSize: 10, fontWeight: 600, color: dl.color }}>{dl.label}</span>
        ) : (
          <input type="date" value="" onChange={e => onUpdate({ ...task, deadline: e.target.value })}
            style={{ width: 20, opacity: 0.3, background: "none", border: "none", cursor: "pointer", fontSize: 10 }} title="Set deadline" />
        )}
        {task.deadline && (
          <input type="date" value={task.deadline} onChange={e => onUpdate({ ...task, deadline: e.target.value })}
            style={{ position: "absolute", inset: 0, opacity: 0, cursor: "pointer" }} />
        )}
      </div>

      <button onClick={() => onDelete(task.id)} style={{
        background: "none", border: "none", cursor: "pointer", color: "var(--c-muted)",
        fontSize: 13, padding: "0 2px", lineHeight: 1, opacity: 0.4,
      }} onMouseEnter={e => { e.currentTarget.style.opacity = 1; e.currentTarget.style.color = "#FF6B6B"; }}
         onMouseLeave={e => { e.currentTarget.style.opacity = 0.4; e.currentTarget.style.color = "var(--c-muted)"; }}>×</button>
    </div>
  );
}

function Tile({ project, onUpdate, onDelete }) {
  const [newTask, setNewTask] = useState("");
  const [editingName, setEditingName] = useState(false);
  const [name, setName] = useState(project.name);

  const total = project.tasks.length;
  const done = project.tasks.filter(t => t.status === "Done").length;
  const pct = total ? Math.round((done / total) * 100) : 0;
  const overdue = project.tasks.filter(t => t.status !== "Done" && daysUntil(t.deadline) !== null && daysUntil(t.deadline) < 0).length;

  const addTask = () => {
    if (!newTask.trim()) return;
    onUpdate({ ...project, tasks: [...project.tasks, { id: uid(), title: newTask.trim(), priority: "medium", status: "Not Started", deadline: "" }] });
    setNewTask("");
  };
  const updateTask = (t) => onUpdate({ ...project, tasks: project.tasks.map(x => x.id === t.id ? t : x) });
  const deleteTask = (tid) => onUpdate({ ...project, tasks: project.tasks.filter(x => x.id !== tid) });
  const commitName = () => { if (name.trim()) onUpdate({ ...project, name: name.trim() }); setEditingName(false); };

  return (
    <div style={{
      background: "var(--c-surface)",
      borderRadius: 16,
      border: "1px solid var(--c-border)",
      display: "flex", flexDirection: "column",
      aspectRatio: "1 / 1",
      position: "relative",
      overflow: "hidden",
    }}>
      <div style={{ height: 4, background: `linear-gradient(90deg, ${project.color}, ${project.color}66)`, flexShrink: 0 }} />

      <div style={{ padding: "14px 16px 8px", flexShrink: 0, display: "flex", alignItems: "flex-start", gap: 10 }}>
        <div style={{ flex: 1, minWidth: 0 }}>
          {editingName ? (
            <input value={name} onChange={e => setName(e.target.value)}
              onBlur={commitName} onKeyDown={e => e.key === "Enter" && commitName()} autoFocus
              style={{ width: "100%", fontSize: 15, fontWeight: 700, background: "var(--c-bg)", border: "1px solid var(--c-border)", borderRadius: 6, padding: "2px 8px", color: "var(--c-text)", fontFamily: "inherit" }} />
          ) : (
            <h2 onClick={() => setEditingName(true)} style={{
              margin: 0, fontSize: 15, fontWeight: 700, cursor: "text",
              color: "var(--c-text)", lineHeight: 1.3,
            }}>{project.name}</h2>
          )}

          <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 6 }}>
            <span style={{ fontSize: 11, color: "var(--c-muted)" }}>{done}/{total}</span>
            {pct > 0 && <span style={{ fontSize: 11, color: project.color, fontWeight: 600 }}>{pct}%</span>}
            {overdue > 0 && <span style={{ fontSize: 10, fontWeight: 600, color: "#FF6B6B" }}>⚠ {overdue}</span>}
          </div>

          <div style={{ height: 3, background: "var(--c-border)", borderRadius: 2, marginTop: 6 }}>
            <div style={{ height: "100%", width: `${pct}%`, background: project.color, borderRadius: 2, transition: "width .3s" }} />
          </div>
        </div>

        <div style={{ display: "flex", gap: 4, flexShrink: 0, marginTop: 2 }}>
          <button onClick={() => {
            const next = TILE_COLORS[(TILE_COLORS.indexOf(project.color) + 1) % TILE_COLORS.length];
            onUpdate({ ...project, color: next });
          }} style={{
            width: 18, height: 18, borderRadius: 99, background: project.color,
            border: "2px solid var(--c-border)", cursor: "pointer",
          }} title="Change color" />
          <button onClick={() => { if (confirm("Delete this project?")) onDelete(project.id); }}
            style={{ background: "none", border: "none", cursor: "pointer", color: "var(--c-muted)", fontSize: 14, padding: "0 2px", opacity: 0.5 }}
            onMouseEnter={e => { e.currentTarget.style.opacity = 1; e.currentTarget.style.color = "#FF6B6B"; }}
            onMouseLeave={e => { e.currentTarget.style.opacity = 0.5; e.currentTarget.style.color = "var(--c-muted)"; }}>×</button>
        </div>
      </div>

      <div style={{ flex: 1, overflowY: "auto", padding: "4px 10px", minHeight: 0 }}>
        {project.tasks.map(t => (
          <TaskRow key={t.id} task={t} onUpdate={updateTask} onDelete={deleteTask} accentColor={project.color} />
        ))}
        {total === 0 && (
          <p style={{ fontSize: 12, color: "var(--c-muted)", textAlign: "center", padding: "20px 0", margin: 0, opacity: 0.6 }}>No tasks yet</p>
        )}
      </div>

      <div style={{ padding: "8px 10px 12px", borderTop: "1px solid var(--c-border)", flexShrink: 0 }}>
        <div style={{ display: "flex", gap: 4 }}>
          <input value={newTask} onChange={e => setNewTask(e.target.value)}
            onKeyDown={e => e.key === "Enter" && addTask()}
            placeholder="+ Add task"
            style={{
              flex: 1, fontSize: 12, padding: "6px 10px", borderRadius: 6,
              border: "1px solid var(--c-border)", background: "var(--c-bg)",
              color: "var(--c-text)", fontFamily: "inherit",
            }} />
          {newTask.trim() && (
            <button onClick={addTask} style={{
              padding: "6px 12px", borderRadius: 6, border: "none",
              background: project.color, color: "#fff", fontSize: 12,
              fontWeight: 600, cursor: "pointer", fontFamily: "inherit",
            }}>Add</button>
          )}
        </div>
      </div>
    </div>
  );
}

export default function Dashboard() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [view, setView] = useState("tiles");

  useEffect(() => {
    const loaded = load();
    setData(loaded || DEFAULTS);
    setLoading(false);
  }, []);

  const persist = useCallback((d) => { setData(d); save(d); }, []);
  const addProject = () => {
    persist({ ...data, projects: [...data.projects, {
      id: uid(), name: "New Project", color: TILE_COLORS[data.projects.length % TILE_COLORS.length], tasks: [],
    }] });
  };
  const updateProject = (p) => persist({ ...data, projects: data.projects.map(x => x.id === p.id ? p : x) });
  const deleteProject = (pid) => persist({ ...data, projects: data.projects.filter(x => x.id !== pid) });

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
      "--c-bg": "#0D0F14",
      "--c-surface": "#161921",
      "--c-row": "#1C1F2A",
      "--c-border": "#252A37",
      "--c-text": "#E2E5ED",
      "--c-muted": "#6B7280",
      fontFamily: "'DM Sans', 'Segoe UI', system-ui, sans-serif",
      background: "var(--c-bg)", color: "var(--c-text)",
      minHeight: "100vh", padding: "20px 24px",
    }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 14 }}>
          <h1 style={{ margin: 0, fontSize: 20, fontWeight: 700, letterSpacing: -0.5 }}>Projects</h1>
          <span style={{ fontSize: 12, color: "var(--c-muted)" }}>
            {doneTasks}/{totalTasks} tasks done
          </span>
        </div>

        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search..."
            style={{ fontSize: 12, padding: "6px 12px", borderRadius: 8, border: "1px solid var(--c-border)", background: "var(--c-surface)", color: "var(--c-text)", width: 140, fontFamily: "inherit" }} />
          <button onClick={() => setView(view === "tiles" ? "deadlines" : "tiles")} style={{
            padding: "6px 12px", borderRadius: 8, border: "1px solid var(--c-border)",
            background: view === "deadlines" ? "var(--c-row)" : "var(--c-surface)",
            color: "var(--c-text)", fontSize: 12, cursor: "pointer", fontFamily: "inherit",
          }}>{view === "tiles" ? "⏰ Deadlines" : "◻ Tiles"}</button>
          <button onClick={addProject} style={{
            padding: "6px 14px", borderRadius: 8, border: "none",
            background: "#5B8DEF", color: "#fff", fontSize: 12, fontWeight: 600,
            cursor: "pointer", fontFamily: "inherit",
          }}>+ Project</button>
        </div>
      </div>

      {(overdueTasks.length > 0 || upcoming7.length > 0) && (
        <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
          {overdueTasks.length > 0 && (
            <div style={{ padding: "5px 12px", borderRadius: 6, background: "#FF6B6B14", border: "1px solid #FF6B6B33", fontSize: 11, color: "#FF6B6B", fontWeight: 600 }}>
              {overdueTasks.length} overdue
            </div>
          )}
          {upcoming7.length > 0 && (
            <div style={{ padding: "5px 12px", borderRadius: 6, background: "#F0A84014", border: "1px solid #F0A84033", fontSize: 11, color: "#F0A840", fontWeight: 600 }}>
              {upcoming7.length} due this week
            </div>
          )}
        </div>
      )}

      {view === "tiles" && (
        <div style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))",
          gap: 14,
        }}>
          {filtered.map(p => (
            <Tile key={p.id} project={p} onUpdate={updateProject} onDelete={deleteProject} />
          ))}
        </div>
      )}

      {view === "deadlines" && (
        <div style={{ maxWidth: 720 }}>
          {deadlineTasks.length === 0 && (
            <p style={{ color: "var(--c-muted)", fontSize: 12, textAlign: "center", padding: 40 }}>No deadlined tasks.</p>
          )}
          {deadlineTasks.map(t => {
            const dl = dlInfo(t.deadline);
            return (
              <div key={t.id} style={{
                display: "flex", alignItems: "center", gap: 10, padding: "8px 12px",
                borderRadius: 8, background: "var(--c-surface)", border: "1px solid var(--c-border)", marginBottom: 4,
              }}>
                <div style={{ width: 8, height: 8, borderRadius: 99, background: t.pColor, flexShrink: 0 }} />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 12, fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.title}</div>
                  <div style={{ fontSize: 10, color: "var(--c-muted)" }}>{t.pName}</div>
                </div>
                <span style={{ fontSize: 10, fontWeight: 700, color: PRIORITY_COLORS[t.priority].text }}>{PRIORITY_COLORS[t.priority].label.toUpperCase()}</span>
                <span style={{ fontSize: 11, fontWeight: 600, color: dl?.color, whiteSpace: "nowrap", minWidth: 50, textAlign: "right" }}>{t.deadline}</span>
                {dl && <span style={{ fontSize: 10, fontWeight: 700, color: dl.color, whiteSpace: "nowrap" }}>{dl.label}</span>}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
