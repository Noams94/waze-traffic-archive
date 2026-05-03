import { useState, useMemo } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend } from "recharts";

const RAW = {
  irregularities: [
    { name: "כפר סבא", from: "", to: "הפועל", jam: 4, time: 2488, hist: 103, len: 535, type: "DYNAMIC" },
    { name: "הציפורנים, קריית טבעון", from: "", to: "הסיגליות", jam: 4, time: 1401, hist: 84, len: 506, type: "DYNAMIC" },
    { name: "60, Al-Ram", from: "Al-Ram", to: "Al-Ram", jam: 4, time: 2208, hist: 159, len: 519, type: "DYNAMIC" },
    { name: "60, Bani Na'im", from: "Bani Na'im", to: "Hebron", jam: 4, time: 472, hist: 43, len: 803, type: "DYNAMIC" },
    { name: "90, עין תמר (א)", from: "עין תמר", to: "עין תמר", jam: 4, time: 1438, hist: 146, len: 3620, type: "DYNAMIC" },
    { name: "16 דרום, ירושלים", from: "", to: "גבעת שאול", jam: 4, time: 703, hist: 84, len: 1885, type: "DYNAMIC" },
    { name: "90, עין תמר (ב)", from: "עין תמר", to: "עין תמר", jam: 4, time: 1021, hist: 126, len: 3623, type: "DYNAMIC" },
    { name: "4 צפון, עכו", from: "", to: "4 צפון", jam: 4, time: 611, hist: 76, len: 1753, type: "DYNAMIC" },
    { name: "60, Jaba'", from: "Jaba'", to: "Al-Ram", jam: 4, time: 1166, hist: 169, len: 2250, type: "DYNAMIC" },
    { name: "22 צפון, חיפה", from: "חיפה", to: "חיפה", jam: 4, time: 627, hist: 102, len: 2489, type: "DYNAMIC" },
    { name: "3866, מחסיה", from: "מחסיה", to: "נס הרים", jam: 4, time: 645, hist: 105, len: 1733, type: "DYNAMIC" },
    { name: "חיפה (עיר)", from: "", to: "צביה ויצחק", jam: 4, time: 809, hist: 133, len: 585, type: "DYNAMIC" },
    { name: "40 צפון, באר שבע", from: "באר שבע", to: "באר שבע", jam: 4, time: 616, hist: 106, len: 2824, type: "DYNAMIC" },
    { name: "רמת השרון", from: "רמת השרון", to: "רמת השרון", jam: 4, time: 541, hist: 105, len: 618, type: "DYNAMIC" },
    { name: "417, Jahalin", from: "Jahalin", to: "Jahalin", jam: 4, time: 844, hist: 166, len: 953, type: "DYNAMIC" },
    { name: "899, יפתח", from: "יפתח", to: "יפתח", jam: 3, time: 588, hist: 147, len: 2611, type: "DYNAMIC" },
    { name: "באר יעקב", from: "באר יעקב", to: "באר יעקב", jam: 4, time: 715, hist: 215, len: 1084, type: "DYNAMIC" },
    { name: "6 דרום, נחלה", from: "נחלה", to: "שדה משה", jam: 3, time: 1172, hist: 373, len: 10107, type: "DYNAMIC" },
    { name: "466, Ein Siniya", from: "Ein Siniya", to: "Ein Siniya", jam: 3, time: 973, hist: 524, len: 4900, type: "DYNAMIC" },
    { name: "מירון 2026 (1)", from: "", to: "", jam: 5, time: 0, hist: 0, len: 919, type: "STATIC" },
    { name: "מירון 2026 (2)", from: "", to: "", jam: 5, time: 0, hist: 0, len: 3262, type: "STATIC" },
    { name: "מירון 2026 (3)", from: "", to: "", jam: 5, time: 0, hist: 0, len: 7803, type: "STATIC" },
    { name: "NTA", from: "", to: "", jam: 5, time: 0, hist: 0, len: 498, type: "STATIC" },
    { name: "תאונה קטלנית", from: "", to: "", jam: 5, time: 0, hist: 0, len: 6160, type: "STATIC" },
    { name: "טרם נפתח (1)", from: "", to: "", jam: 5, time: 0, hist: 0, len: 1226, type: "STATIC" },
    { name: "טרם נפתח (2)", from: "", to: "", jam: 5, time: 0, hist: 0, len: 1221, type: "STATIC" },
    { name: "מירון 2026 (4)", from: "", to: "", jam: 5, time: 0, hist: 0, len: 2812, type: "STATIC" },
    { name: "8655 Roadworks", from: "", to: "", jam: 5, time: 0, hist: 0, len: 1599, type: "STATIC" },
  ],
  subRoutes: [
    { from: "שד' בן-גוריון", to: "כפר שמואל", time: 1286, hist: 1196, len: 32122, jam: 0 },
    { from: "כפר שמואל", to: "6 צפון", time: 324, hist: 303, len: 4836, jam: 3 },
    { from: "6 צפון", to: "1 מערב", time: 756, hist: 711, len: 19721, jam: 0 },
    { from: "איילון", to: "20 איילון צפון", time: 102, hist: 102, len: 1085, jam: 3 },
  ],
  usersOnJams: [
    { count: 124528, level: 0 },
    { count: 2856, level: 1 },
    { count: 10133, level: 2 },
    { count: 37602, level: 3 },
    { count: 6384, level: 4 },
  ],
  lengthOfJams: [
    { level: 1, km: 96.6 },
    { level: 2, km: 195.4 },
    { level: 3, km: 298.2 },
    { level: 4, km: 108.3 },
    { level: 5, km: 49.5 },
  ],
  route: { name: "כביש 1 מערב + נתיב מהיר", from: "שד' בן-גוריון", to: "כביש 1 מערב", time: 2468, hist: 2312, len: 57764 },
};

const JAM_COLORS = {
  0: "#4ade80",
  1: "#a3e635",
  2: "#facc15",
  3: "#f97316",
  4: "#ef4444",
  5: "#991b1b",
};
const JAM_LABELS = { 0: "חופשי", 1: "קל", 2: "בינוני", 3: "כבד", 4: "כבד מאוד", 5: "חסום" };

const fmt = (s) => {
  const m = Math.floor(s / 60);
  const sec = s % 60;
  return m > 0 ? `${m}:${String(sec).padStart(2, "0")}` : `${sec}s`;
};
const pct = (t, h) => (h > 0 ? Math.round(((t / h) - 1) * 100) : null);

const TABS = ["סקירה כללית", "חריגות", "כביש 1"];

export default function WazeDashboard() {
  const [tab, setTab] = useState(0);
  const [sortBy, setSortBy] = useState("delay");

  const dynamicIrr = useMemo(
    () => RAW.irregularities.filter((i) => i.type === "DYNAMIC"),
    []
  );
  const staticIrr = useMemo(
    () => RAW.irregularities.filter((i) => i.type === "STATIC"),
    []
  );

  const sortedIrr = useMemo(() => {
    const arr = [...dynamicIrr];
    if (sortBy === "delay") arr.sort((a, b) => pct(b.time, b.hist) - pct(a.time, a.hist));
    else if (sortBy === "time") arr.sort((a, b) => b.time - a.time);
    else arr.sort((a, b) => b.len - a.len);
    return arr;
  }, [dynamicIrr, sortBy]);

  const totalUsers = RAW.usersOnJams.reduce((s, u) => s + u.count, 0);
  const totalJamKm = RAW.lengthOfJams.reduce((s, l) => s + l.km, 0);
  const congestedUsers = RAW.usersOnJams.filter((u) => u.level > 0).reduce((s, u) => s + u.count, 0);

  const pieData = RAW.usersOnJams.map((u) => ({
    name: JAM_LABELS[u.level],
    value: u.count,
    level: u.level,
  }));

  const jamBarData = RAW.lengthOfJams.map((l) => ({
    name: JAM_LABELS[l.level],
    km: Math.round(l.km),
    level: l.level,
  }));

  return (
    <div style={{
      fontFamily: "'Segoe UI', 'Noto Sans Hebrew', Tahoma, sans-serif",
      direction: "rtl",
      background: "linear-gradient(135deg, #0f172a 0%, #1e293b 100%)",
      color: "#e2e8f0",
      minHeight: "100vh",
      padding: "0",
    }}>
      {/* Header */}
      <div style={{
        background: "linear-gradient(90deg, #1e3a5f 0%, #0f172a 100%)",
        padding: "20px 28px",
        borderBottom: "2px solid #334155",
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        flexWrap: "wrap",
        gap: 12,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{
            width: 42, height: 42, borderRadius: 10,
            background: "linear-gradient(135deg, #38bdf8, #06b6d4)",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: 22, fontWeight: 800, color: "#0f172a",
          }}>W</div>
          <div>
            <div style={{ fontSize: 20, fontWeight: 700, letterSpacing: "-0.3px" }}>
              דשבורד תנועה — Waze
            </div>
            <div style={{ fontSize: 12, color: "#94a3b8" }}>
              עדכון: יום א׳, 03.05.2026 · 07:21
            </div>
          </div>
        </div>
        <div style={{
          display: "flex", alignItems: "center", gap: 6,
          background: "#0f172a", borderRadius: 8, padding: "3px 4px",
        }}>
          {TABS.map((t, i) => (
            <button key={i} onClick={() => setTab(i)} style={{
              padding: "7px 16px", borderRadius: 6, border: "none", cursor: "pointer",
              fontSize: 13, fontWeight: 600, transition: "all .2s",
              background: tab === i ? "#38bdf8" : "transparent",
              color: tab === i ? "#0f172a" : "#94a3b8",
            }}>{t}</button>
          ))}
        </div>
      </div>

      <div style={{ padding: "20px 28px", maxWidth: 1100, margin: "0 auto" }}>
        {/* TAB 0: Overview */}
        {tab === 0 && (
          <div>
            {/* KPI cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 14, marginBottom: 24 }}>
              {[
                { label: "משתמשים פעילים", value: totalUsers.toLocaleString(), accent: "#38bdf8" },
                { label: "בפקקים", value: congestedUsers.toLocaleString(), accent: "#f97316" },
                { label: 'ק"מ פקקים', value: `${Math.round(totalJamKm)}`, accent: "#ef4444" },
                { label: "חריגות דינמיות", value: dynamicIrr.length, accent: "#a855f7" },
                { label: "חסימות סטטיות", value: staticIrr.length, accent: "#64748b" },
              ].map((kpi, i) => (
                <div key={i} style={{
                  background: "#1e293b", borderRadius: 12, padding: "18px 20px",
                  borderRight: `4px solid ${kpi.accent}`,
                  boxShadow: "0 2px 8px rgba(0,0,0,.25)",
                }}>
                  <div style={{ fontSize: 12, color: "#94a3b8", marginBottom: 4 }}>{kpi.label}</div>
                  <div style={{ fontSize: 28, fontWeight: 800, color: kpi.accent }}>{kpi.value}</div>
                </div>
              ))}
            </div>

            {/* Two charts side by side */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18, marginBottom: 24 }}>
              {/* Users by jam level - Pie */}
              <div style={{ background: "#1e293b", borderRadius: 12, padding: 20, boxShadow: "0 2px 8px rgba(0,0,0,.25)" }}>
                <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 14, color: "#cbd5e1" }}>
                  משתמשים לפי רמת פקק
                </div>
                <ResponsiveContainer width="100%" height={220}>
                  <PieChart>
                    <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%"
                      innerRadius={50} outerRadius={85} paddingAngle={2} strokeWidth={0}>
                      {pieData.map((d, i) => <Cell key={i} fill={JAM_COLORS[d.level]} />)}
                    </Pie>
                    <Tooltip
                      formatter={(v) => v.toLocaleString()}
                      contentStyle={{ background: "#0f172a", border: "1px solid #334155", borderRadius: 8, direction: "rtl", fontSize: 12 }}
                      itemStyle={{ color: "#e2e8f0" }}
                    />
                    <Legend
                      formatter={(v) => <span style={{ color: "#94a3b8", fontSize: 11 }}>{v}</span>}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </div>

              {/* Jam length bar */}
              <div style={{ background: "#1e293b", borderRadius: 12, padding: 20, boxShadow: "0 2px 8px rgba(0,0,0,.25)" }}>
                <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 14, color: "#cbd5e1" }}>
                  אורך פקקים (ק״מ) לפי חומרה
                </div>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={jamBarData} layout="vertical" margin={{ right: 10 }}>
                    <XAxis type="number" tick={{ fill: "#64748b", fontSize: 11 }} />
                    <YAxis type="category" dataKey="name" width={70} tick={{ fill: "#94a3b8", fontSize: 12 }} />
                    <Tooltip
                      formatter={(v) => `${v} ק"מ`}
                      contentStyle={{ background: "#0f172a", border: "1px solid #334155", borderRadius: 8, direction: "rtl", fontSize: 12 }}
                      itemStyle={{ color: "#e2e8f0" }}
                    />
                    <Bar dataKey="km" radius={[0, 6, 6, 0]} barSize={22}>
                      {jamBarData.map((d, i) => <Cell key={i} fill={JAM_COLORS[d.level]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Static irregularities */}
            <div style={{ background: "#1e293b", borderRadius: 12, padding: 20, boxShadow: "0 2px 8px rgba(0,0,0,.25)" }}>
              <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 12, color: "#cbd5e1" }}>
                חסימות וסגירות סטטיות
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                {staticIrr.map((s, i) => (
                  <div key={i} style={{
                    background: "#0f172a", borderRadius: 8, padding: "10px 14px",
                    border: "1px solid #334155", fontSize: 13,
                    display: "flex", alignItems: "center", gap: 8,
                  }}>
                    <span style={{
                      display: "inline-block", width: 8, height: 8, borderRadius: "50%",
                      background: JAM_COLORS[s.jam],
                    }} />
                    <span style={{ fontWeight: 600 }}>{s.name}</span>
                    <span style={{ color: "#64748b", fontSize: 11 }}>
                      {(s.len / 1000).toFixed(1)} ק״מ
                    </span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* TAB 1: Irregularities detail */}
        {tab === 1 && (
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 18 }}>
              <span style={{ fontSize: 14, color: "#94a3b8" }}>מיון:</span>
              {[
                { key: "delay", label: "% עיכוב" },
                { key: "time", label: "זמן נוכחי" },
                { key: "length", label: "אורך" },
              ].map((s) => (
                <button key={s.key} onClick={() => setSortBy(s.key)} style={{
                  padding: "5px 14px", borderRadius: 6, border: "none", cursor: "pointer",
                  fontSize: 12, fontWeight: 600,
                  background: sortBy === s.key ? "#38bdf8" : "#334155",
                  color: sortBy === s.key ? "#0f172a" : "#94a3b8",
                }}>{s.label}</button>
              ))}
            </div>

            {/* Bar chart of delays */}
            <div style={{ background: "#1e293b", borderRadius: 12, padding: 20, marginBottom: 20, boxShadow: "0 2px 8px rgba(0,0,0,.25)" }}>
              <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 14, color: "#cbd5e1" }}>
                עיכוב יחסי (%) — זמן נוכחי מול היסטורי
              </div>
              <ResponsiveContainer width="100%" height={Math.max(300, sortedIrr.length * 30)}>
                <BarChart data={sortedIrr.map((d) => ({
                  name: d.name.length > 22 ? d.name.slice(0, 20) + "…" : d.name,
                  delay: pct(d.time, d.hist),
                  jam: d.jam,
                }))} layout="vertical" margin={{ right: 10 }}>
                  <XAxis type="number" tick={{ fill: "#64748b", fontSize: 11 }} unit="%" />
                  <YAxis type="category" dataKey="name" width={140} tick={{ fill: "#94a3b8", fontSize: 11 }} />
                  <Tooltip
                    formatter={(v) => `${v}%`}
                    contentStyle={{ background: "#0f172a", border: "1px solid #334155", borderRadius: 8, direction: "rtl", fontSize: 12 }}
                    itemStyle={{ color: "#e2e8f0" }}
                  />
                  <Bar dataKey="delay" radius={[0, 6, 6, 0]} barSize={18}>
                    {sortedIrr.map((d, i) => <Cell key={i} fill={JAM_COLORS[d.jam]} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* Table */}
            <div style={{ background: "#1e293b", borderRadius: 12, padding: 20, overflowX: "auto", boxShadow: "0 2px 8px rgba(0,0,0,.25)" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ borderBottom: "2px solid #334155" }}>
                    {["מיקום", "מ →‎ אל", "רמה", "נוכחי", "היסטורי", "עיכוב", "אורך"].map((h) => (
                      <th key={h} style={{ padding: "10px 8px", textAlign: "right", color: "#64748b", fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {sortedIrr.map((d, i) => {
                    const delay = pct(d.time, d.hist);
                    return (
                      <tr key={i} style={{ borderBottom: "1px solid #1e293b", background: i % 2 === 0 ? "#1e293b" : "#162033" }}>
                        <td style={{ padding: "10px 8px", fontWeight: 600 }}>{d.name}</td>
                        <td style={{ padding: "10px 8px", color: "#94a3b8" }}>
                          {d.from || "—"} → {d.to || "—"}
                        </td>
                        <td style={{ padding: "10px 8px" }}>
                          <span style={{
                            display: "inline-block", padding: "2px 10px", borderRadius: 99,
                            background: JAM_COLORS[d.jam], color: "#0f172a", fontWeight: 700, fontSize: 11,
                          }}>{d.jam}</span>
                        </td>
                        <td style={{ padding: "10px 8px", fontVariantNumeric: "tabular-nums" }}>{fmt(d.time)}</td>
                        <td style={{ padding: "10px 8px", fontVariantNumeric: "tabular-nums", color: "#64748b" }}>{fmt(d.hist)}</td>
                        <td style={{ padding: "10px 8px", fontWeight: 700, color: delay > 500 ? "#ef4444" : delay > 200 ? "#f97316" : "#facc15" }}>
                          +{delay}%
                        </td>
                        <td style={{ padding: "10px 8px", fontVariantNumeric: "tabular-nums" }}>
                          {d.len >= 1000 ? `${(d.len / 1000).toFixed(1)} ק״מ` : `${d.len} מ׳`}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* TAB 2: Route 1 */}
        {tab === 2 && (
          <div>
            {/* Route summary card */}
            <div style={{
              background: "linear-gradient(135deg, #1e3a5f 0%, #1e293b 100%)",
              borderRadius: 14, padding: 24, marginBottom: 22,
              border: "1px solid #334155",
              boxShadow: "0 4px 16px rgba(0,0,0,.3)",
            }}>
              <div style={{ fontSize: 18, fontWeight: 800, marginBottom: 4 }}>{RAW.route.name}</div>
              <div style={{ fontSize: 13, color: "#94a3b8", marginBottom: 18 }}>
                {RAW.route.from} ← → {RAW.route.to} · {(RAW.route.len / 1000).toFixed(1)} ק״מ
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16 }}>
                <div>
                  <div style={{ fontSize: 11, color: "#64748b" }}>זמן נוכחי</div>
                  <div style={{ fontSize: 30, fontWeight: 800, color: "#38bdf8" }}>{fmt(RAW.route.time)}</div>
                </div>
                <div>
                  <div style={{ fontSize: 11, color: "#64748b" }}>זמן היסטורי</div>
                  <div style={{ fontSize: 30, fontWeight: 800, color: "#4ade80" }}>{fmt(RAW.route.hist)}</div>
                </div>
                <div>
                  <div style={{ fontSize: 11, color: "#64748b" }}>עיכוב</div>
                  <div style={{ fontSize: 30, fontWeight: 800, color: "#f97316" }}>+{pct(RAW.route.time, RAW.route.hist)}%</div>
                </div>
              </div>
            </div>

            {/* Sub-route segments */}
            <div style={{ fontSize: 14, fontWeight: 700, color: "#cbd5e1", marginBottom: 14 }}>
              מקטעים
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
              {RAW.subRoutes.map((s, i) => {
                const delay = pct(s.time, s.hist);
                const pctOfTotal = Math.round((s.len / RAW.route.len) * 100);
                return (
                  <div key={i} style={{ display: "flex", alignItems: "stretch" }}>
                    {/* Timeline dot */}
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", width: 30 }}>
                      <div style={{
                        width: 14, height: 14, borderRadius: "50%", flexShrink: 0,
                        background: JAM_COLORS[s.jam], border: "2px solid #0f172a",
                        boxShadow: `0 0 8px ${JAM_COLORS[s.jam]}66`,
                      }} />
                      {i < RAW.subRoutes.length - 1 && (
                        <div style={{ flex: 1, width: 2, background: "#334155" }} />
                      )}
                    </div>
                    {/* Card */}
                    <div style={{
                      flex: 1, background: "#1e293b", borderRadius: 10, padding: "14px 18px",
                      marginBottom: 12, marginRight: 8,
                      borderRight: `3px solid ${JAM_COLORS[s.jam]}`,
                      boxShadow: "0 2px 6px rgba(0,0,0,.2)",
                    }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 6 }}>
                        <span style={{ fontWeight: 700, fontSize: 14 }}>
                          {s.from || "התחלה"} → {s.to}
                        </span>
                        <span style={{
                          padding: "2px 10px", borderRadius: 99, fontSize: 11, fontWeight: 700,
                          background: JAM_COLORS[s.jam], color: "#0f172a",
                        }}>
                          {JAM_LABELS[s.jam]}
                        </span>
                      </div>
                      <div style={{ display: "flex", gap: 20, fontSize: 12, color: "#94a3b8" }}>
                        <span>{(s.len / 1000).toFixed(1)} ק״מ ({pctOfTotal}%)</span>
                        <span>נוכחי: <b style={{ color: "#e2e8f0" }}>{fmt(s.time)}</b></span>
                        <span>היסטורי: {fmt(s.hist)}</span>
                        {delay > 0 && <span style={{ color: "#f97316", fontWeight: 600 }}>+{delay}%</span>}
                      </div>
                      {/* Progress bar */}
                      <div style={{ marginTop: 8, height: 5, borderRadius: 99, background: "#0f172a", overflow: "hidden" }}>
                        <div style={{
                          width: `${Math.min(100, (s.time / s.hist) * 100)}%`,
                          height: "100%", borderRadius: 99,
                          background: `linear-gradient(90deg, ${JAM_COLORS[s.jam]}cc, ${JAM_COLORS[s.jam]})`,
                          transition: "width .5s ease",
                        }} />
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Comparison bar chart */}
            <div style={{ background: "#1e293b", borderRadius: 12, padding: 20, marginTop: 10, boxShadow: "0 2px 8px rgba(0,0,0,.25)" }}>
              <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 14, color: "#cbd5e1" }}>
                השוואת זמנים — נוכחי מול היסטורי
              </div>
              <ResponsiveContainer width="100%" height={180}>
                <BarChart data={RAW.subRoutes.map((s) => ({
                  name: s.to,
                  נוכחי: Math.round(s.time / 60 * 10) / 10,
                  היסטורי: Math.round(s.hist / 60 * 10) / 10,
                }))} margin={{ right: 10 }}>
                  <XAxis dataKey="name" tick={{ fill: "#94a3b8", fontSize: 11 }} />
                  <YAxis tick={{ fill: "#64748b", fontSize: 11 }} unit=" דק׳" />
                  <Tooltip
                    formatter={(v) => `${v} דק׳`}
                    contentStyle={{ background: "#0f172a", border: "1px solid #334155", borderRadius: 8, direction: "rtl", fontSize: 12 }}
                    itemStyle={{ color: "#e2e8f0" }}
                  />
                  <Bar dataKey="היסטורי" fill="#334155" radius={[4, 4, 0, 0]} barSize={28} />
                  <Bar dataKey="נוכחי" fill="#38bdf8" radius={[4, 4, 0, 0]} barSize={28} />
                  <Legend formatter={(v) => <span style={{ color: "#94a3b8", fontSize: 11 }}>{v}</span>} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
