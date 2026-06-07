import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import {
  HistogramChart,
  BarChart, Bar,
  ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Cell, ReferenceLine
} from 'recharts';

// ─── Paleta & estilos globales ────────────────────────────────────────────────
const STYLES = `
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg:        #f7f5f2;
    --surface:   #ffffff;
    --surface2:  #f0ede8;
    --border:    #e0dbd3;
    --accent:    #a07830;
    --accent2:   #4a9080;
    --text:      #1a1714;
    --muted:     #8a8078;
    --danger:    #b05a38;
    --radius:    12px;
    --font-disp: 'Playfair Display', Georgia, serif;
    --font-body: 'DM Sans', system-ui, sans-serif;
  }

  body {
    background: var(--bg);
    color: var(--text);
    font-family: var(--font-body);
    font-weight: 300;
    line-height: 1.6;
    min-height: 100vh;
  }

  .app {
    max-width: 1100px;
    margin: 0 auto;
    padding: 40px 24px 80px;
  }

  /* Header */
  .header {
    display: flex;
    align-items: flex-end;
    justify-content: space-between;
    margin-bottom: 48px;
    padding-bottom: 24px;
    border-bottom: 1px solid var(--border);
  }
  .header-title {
    font-family: var(--font-disp);
    font-size: clamp(28px, 4vw, 42px);
    font-weight: 700;
    letter-spacing: -0.5px;
    line-height: 1.1;
    color: var(--text);
  }
  .header-title span { color: var(--accent); }
  .header-sub {
    font-size: 13px;
    color: var(--muted);
    letter-spacing: 0.08em;
    text-transform: uppercase;
  }

  /* Filters bar */
  .filters {
    display: flex;
    gap: 16px;
    flex-wrap: wrap;
    margin-bottom: 40px;
  }
  .filter-group {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .filter-label {
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: var(--muted);
  }
  .filter-select {
    background: var(--surface);
    border: 1px solid var(--border);
    color: var(--text);
    font-family: var(--font-body);
    font-size: 14px;
    padding: 10px 14px;
    border-radius: var(--radius);
    cursor: pointer;
    outline: none;
    transition: border-color 0.2s;
    min-width: 180px;
  }
  .filter-select:hover, .filter-select:focus { border-color: var(--accent); }

  /* Loading / error */
  .state-box {
    display: flex;
    align-items: center;
    justify-content: center;
    height: 200px;
    color: var(--muted);
    font-size: 15px;
    gap: 12px;
  }
  .spinner {
    width: 20px; height: 20px;
    border: 2px solid var(--border);
    border-top-color: var(--accent);
    border-radius: 50%;
    animation: spin 0.8s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* Stats grid */
  .stats-section { margin-bottom: 48px; }
  .section-title {
    font-family: var(--font-disp);
    font-size: 20px;
    font-weight: 600;
    color: var(--accent);
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .section-title::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--border);
  }
  .stats-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 16px;
  }
  @media (max-width: 680px) { .stats-grid { grid-template-columns: 1fr; } }

  .stat-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px;
  }
  .stat-card-title {
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: var(--muted);
    margin-bottom: 16px;
  }
  .stat-row {
    display: flex;
    justify-content: space-between;
    align-items: baseline;
    padding: 7px 0;
    border-bottom: 1px solid var(--border);
    font-size: 13.5px;
  }
  .stat-row:last-child { border-bottom: none; }
  .stat-key { color: var(--muted); }
  .stat-val { font-weight: 500; font-variant-numeric: tabular-nums; }
  .stat-val.accent { color: var(--accent); }
  .stat-val.accent2 { color: var(--accent2); }

  /* Chart section */
  .chart-section { margin-bottom: 48px; }
  .chart-controls {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
    margin-bottom: 24px;
  }
  .chart-btn {
    background: var(--surface);
    border: 1px solid var(--border);
    color: var(--muted);
    font-family: var(--font-body);
    font-size: 13px;
    padding: 8px 16px;
    border-radius: 100px;
    cursor: pointer;
    transition: all 0.2s;
  }
  .chart-btn:hover { border-color: var(--accent); color: var(--text); }
  .chart-btn.active {
    background: var(--accent);
    border-color: var(--accent);
    color: #ffffff;
    font-weight: 500;
  }
  .chart-box {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 24px;
  }
  .chart-title {
    font-size: 14px;
    color: var(--muted);
    margin-bottom: 20px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
  }
  .outlier-toggle {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 16px;
    font-size: 13px;
    color: var(--muted);
    cursor: pointer;
    user-select: none;
  }
  .outlier-toggle input { accent-color: var(--accent); }

  /* Finance modules */
  .finance-section { margin-bottom: 32px; }
  .expander {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
  }
  .expander-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 18px 24px;
    cursor: pointer;
    user-select: none;
    transition: background 0.2s;
  }
  .expander-header:hover { background: var(--surface2); }
  .expander-header-title {
    font-size: 15px;
    font-weight: 500;
    color: var(--text);
  }
  .expander-arrow {
    color: var(--accent);
    transition: transform 0.3s;
    font-size: 18px;
  }
  .expander-arrow.open { transform: rotate(180deg); }
  .expander-body {
    padding: 24px;
    border-top: 1px solid var(--border);
    display: none;
  }
  .expander-body.open { display: block; }

  .input-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 16px;
    margin-bottom: 24px;
  }
  @media (max-width: 680px) { .input-grid { grid-template-columns: 1fr; } }
  .input-grid-2 { grid-template-columns: repeat(2, 1fr); }

  .input-group {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .input-label {
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: var(--muted);
  }
  .input-field {
    background: var(--surface2);
    border: 1px solid var(--border);
    color: var(--text);
    font-family: var(--font-body);
    font-size: 14px;
    padding: 10px 14px;
    border-radius: 8px;
    outline: none;
    transition: border-color 0.2s;
    width: 100%;
  }
  .input-field:focus { border-color: var(--accent); }

  .metrics-row {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 16px;
    margin-bottom: 24px;
  }
  .metric-card {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 16px 20px;
  }
  .metric-label {
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: var(--muted);
    margin-bottom: 6px;
  }
  .metric-value {
    font-family: var(--font-disp);
    font-size: 22px;
    font-weight: 600;
    color: var(--accent);
    font-variant-numeric: tabular-nums;
  }

  .flow-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    margin-top: 12px;
  }
  .flow-table th {
    text-align: right;
    padding: 8px 12px;
    color: var(--muted);
    font-weight: 400;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    border-bottom: 1px solid var(--border);
  }
  .flow-table th:first-child { text-align: left; }
  .flow-table td {
    text-align: right;
    padding: 8px 12px;
    border-bottom: 1px solid var(--border);
    font-variant-numeric: tabular-nums;
    color: var(--text);
  }
  .flow-table td:first-child { text-align: left; color: var(--muted); }
  .flow-table tr:last-child td { border-bottom: none; }

  .caption {
    font-size: 12px;
    color: var(--muted);
    margin-bottom: 8px;
    font-style: italic;
  }
`;

// ─── Helpers ─────────────────────────────────────────────────────────────────

const URLS = {
  Departamento: '/data/departamentos_cochabamba.xlsx',
  Casa: '/data/casas_cochabamba.xlsx',
};

// Fetch + parse XLSX via SheetJS
async function fetchXLSX(url) {
  const res = await fetch(url);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws);
}

// Format numbers with Spanish locale: dots as thousands, commas as decimals
function fmtNum(value, decimals = 0) {
  if (value == null || isNaN(value)) return '—';
  return new Intl.NumberFormat('es-BO', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  }).format(value);
}

function fmtUSD(value, decimals = 0) {
  if (value == null || isNaN(value)) return '—';
  return 'USD ' + fmtNum(value, decimals);
}

function fmtBOB(value, decimals = 0) {
  if (value == null || isNaN(value)) return '—';
  return 'Bs ' + fmtNum(value, decimals);
}

// Descriptive stats
function describe(arr) {
  const valid = arr.filter(v => v != null && !isNaN(v)).sort((a, b) => a - b);
  const n = valid.length;
  if (!n) return {};
  const sum = valid.reduce((a, b) => a + b, 0);
  const mean = sum / n;
  const sorted = [...valid];
  const q = p => {
    const idx = p * (n - 1);
    const lo = Math.floor(idx), hi = Math.ceil(idx);
    return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
  };
  const variance = valid.reduce((acc, v) => acc + (v - mean) ** 2, 0) / n;
  return {
    count: n,
    mean,
    std: Math.sqrt(variance),
    min: sorted[0],
    q25: q(0.25),
    median: q(0.50),
    q75: q(0.75),
    max: sorted[n - 1],
  };
}

// Build histogram bins for recharts
function buildHistogram(arr, bins = 20) {
  const valid = arr.filter(v => v != null && !isNaN(v));
  if (!valid.length) return [];
  const min = Math.min(...valid);
  const max = Math.max(...valid);
  const width = (max - min) / bins || 1;
  const counts = Array(bins).fill(0).map((_, i) => ({
    x: min + i * width,
    midpoint: min + (i + 0.5) * width,
    count: 0,
  }));
  valid.forEach(v => {
    let idx = Math.floor((v - min) / width);
    if (idx >= bins) idx = bins - 1;
    counts[idx].count++;
  });
  return counts;
}

// IQR outlier removal
function removeOutliers(arr) {
  const sorted = [...arr].filter(v => v != null && !isNaN(v)).sort((a, b) => a - b);
  const n = sorted.length;
  const q1 = sorted[Math.floor(n * 0.25)];
  const q3 = sorted[Math.floor(n * 0.75)];
  const iqr = q3 - q1;
  return [q1 - 1.5 * iqr, q3 + 1.5 * iqr];
}


// ─── StatCard component ───────────────────────────────────────────────────────
function StatCard({ title, stats, fmt }) {
  if (!stats || !Object.keys(stats).length) return null;
  const rows = [
    { key: 'Observaciones', val: fmtNum(stats.count) },
    { key: 'Promedio', val: fmt(stats.mean), cls: 'accent' },
    { key: 'Desvío estándar', val: fmt(stats.std) },
    { key: 'Mínimo', val: fmt(stats.min) },
    { key: '25%', val: fmt(stats.q25) },
    { key: 'Mediana', val: fmt(stats.median), cls: 'accent2' },
    { key: '75%', val: fmt(stats.q75) },
    { key: 'Máximo', val: fmt(stats.max) },
  ];
  return (
    <div className="stat-card">
      <div className="stat-card-title">{title}</div>
      {rows.map(r => (
        <div className="stat-row" key={r.key}>
          <span className="stat-key">{r.key}</span>
          <span className={`stat-val${r.cls ? ' ' + r.cls : ''}`}>{r.val}</span>
        </div>
      ))}
    </div>
  );
}

// ─── Custom Tooltip ───────────────────────────────────────────────────────────
const CustomTooltip = ({ active, payload, label, xFmt, yFmt }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{
      background: '#ffffff', border: '1px solid #e0dbd3', borderRadius: 8,
      padding: '10px 14px', fontSize: 13, color: '#1a1714',
    }}>
      {xFmt && <div style={{ color: '#6b6560', marginBottom: 4 }}>{xFmt(label ?? payload[0]?.payload?.x)}</div>}
      {payload.map((p, i) => (
        <div key={i}><span style={{ color: p.color }}>{p.name}: </span>{yFmt ? yFmt(p.value) : p.value}</div>
      ))}
    </div>
  );
};

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [tipoPropiedad, setTipoPropiedad] = useState('Departamento');
  const [rawData, setRawData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const [habitacion, setHabitacion] = useState('Todos');
  const [zona, setZona] = useState('Todas');
  const [chartMode, setChartMode] = useState('Precios');
  const [eliminarOutliers, setEliminarOutliers] = useState(true);

  // Finance state
  const [openRenta, setOpenRenta] = useState(false);
  const [precioInmueble, setPrecioInmueble] = useState('');
  const [rentaPct, setRentaPct] = useState(5);

  // Load data
  useEffect(() => {
    setLoading(true);
    setError(null);
    setHabitacion('Todos');
    setZona('Todas');
    fetchXLSX(URLS[tipoPropiedad])
      .then(rows => {
        const data = rows.map(r => ({
          ...r,
          Precio_USD: parseFloat(r['Precio_USD'] ?? r['precio_usd'] ?? r['PRECIO_USD']) || null,
          Superficie_m2: parseFloat(r['Superficie_m2'] ?? r['superficie_m2'] ?? r['SUPERFICIE_M2']) || null,
          habitaciones: r['Nro de Habitaciones'] ?? r['habitaciones'] ?? r['Habitaciones'] ?? r['HABITACIONES'],
          zona: r['Zona'] ?? r['zona'] ?? r['ZONA'],
        })).map(r => ({ ...r, Precio_m2: r.Precio_USD && r.Superficie_m2 ? r.Precio_USD / r.Superficie_m2 : null }));
        setRawData(data);
        setLoading(false);
      })
      .catch(e => { setError(e.message); setLoading(false); });
  }, [tipoPropiedad]);

  // Derived data
  const habitaciones = ['Todos', '1 habitación', '2 habitaciones', '3 o más habitaciones'];
  const zonas = ['Todas', 'centro', 'norte', 'este', 'oeste', 'sur'];
  const df = rawData.filter(r =>
    (habitacion === 'Todos' || r.habitaciones === habitacion) &&
    (zona === 'Todas' || r.zona === zona)
  );

  // Scatter data (with outlier filter)
  const scatterData = (() => {
    let d = df.filter(r => r.Precio_USD && r.Superficie_m2);
    if (eliminarOutliers) {
      const [lo, hi] = removeOutliers(d.map(r => r.Superficie_m2));
      d = d.filter(r => r.Superficie_m2 >= lo && r.Superficie_m2 <= hi);
    }
    return d;
  })();

  // Stats
  const statsPrecios = describe(df.map(r => r.Precio_USD));
  const statsSuperficie = describe(df.map(r => r.Superficie_m2));
  const statsM2 = describe(df.map(r => r.Precio_m2));

  // Finance calculations
  const precioNum = parseFloat(precioInmueble) || 0;
  const alqAnualReq = precioNum * (rentaPct / 100);
  const alqMensualReq = alqAnualReq / 12;

  // Chart data
  const histPrecios = buildHistogram(df.map(r => r.Precio_USD).filter(Boolean));
  const histSuperficie = buildHistogram(df.map(r => r.Superficie_m2).filter(Boolean));
  const histM2 = buildHistogram(df.map(r => r.Precio_m2).filter(Boolean));

  const ACCENT = '#a07830';
  const ACCENT2 = '#4a9080';
  const DANGER = '#b05a38';

  return (
    <>
      <style>{STYLES}</style>
      <div className="app">
        {/* Header */}
        <div className="header">
          <div>
            <div className="header-title">Dashboard <span>Inmobiliario</span></div>
          </div>
          <div className="header-sub">Análisis del Mercado de Cochabamba</div>
        </div>

        {/* Filters */}
        <div className="filters">
          <div className="filter-group">
            <label className="filter-label">Tipo de propiedad</label>
            <select className="filter-select" value={tipoPropiedad} onChange={e => setTipoPropiedad(e.target.value)}>
              <option>Departamento</option>
              <option>Casa</option>
            </select>
          </div>
          <div className="filter-group">
            <label className="filter-label">Habitaciones</label>
            <select className="filter-select" value={habitacion} onChange={e => setHabitacion(e.target.value)}>
              {habitaciones.map(h => <option key={h}>{h}</option>)}
            </select>
          </div>
          <div className="filter-group">
            <label className="filter-label">Zona</label>
            <select className="filter-select" value={zona} onChange={e => setZona(e.target.value)}>
              {zonas.map(z => <option key={z}>{z}</option>)}
            </select>
          </div>
        </div>

        {/* States */}
        {loading && <div className="state-box"><div className="spinner" /> Cargando datos…</div>}
        {error && <div className="state-box" style={{ color: '#c97a5a' }}>Error al cargar datos: {error}</div>}

        {!loading && !error && df.length > 0 && (
          <>
            {/* Stats */}
            <div className="stats-section">
              <div className="section-title">Estadísticas descriptivas</div>
              <div className="stats-grid">
                <StatCard title="Precio (USD)" stats={statsPrecios} fmt={v => fmtUSD(v, 0)} />
                <StatCard title="Superficie (m²)" stats={statsSuperficie} fmt={v => fmtNum(v, 1) + ' m²'} />
                <StatCard title="Precio por m² (USD/m²)" stats={statsM2} fmt={v => fmtUSD(v, 0)} />
              </div>
            </div>

            {/* Chart */}
            <div className="chart-section">
              <div className="section-title">Visualización</div>
              <div className="chart-controls">
                {['Precios', 'Superficie', 'Precio por m²', 'Precios y Superficie'].map(m => (
                  <button key={m} className={`chart-btn${chartMode === m ? ' active' : ''}`} onClick={() => setChartMode(m)}>{m}</button>
                ))}
              </div>

              <div className="chart-box">
                {chartMode === 'Precios' && (
                  <>
                    <div className="chart-title">Distribución de Precios (USD)</div>
                    <ResponsiveContainer width="100%" height={300}>
                      <BarChart data={histPrecios} barCategoryGap="5%">
                        <CartesianGrid strokeDasharray="3 3" stroke="#e0dbd3" />
                        <XAxis dataKey="midpoint" tickFormatter={v => `$${(v/1000).toFixed(0)}k`} tick={{ fill: '#8a8078', fontSize: 11 }} />
                        <YAxis tick={{ fill: '#8a8078', fontSize: 11 }} />
                        <Tooltip content={<CustomTooltip xFmt={v => fmtUSD(v)} yFmt={v => fmtNum(v) + ' propiedades'} />} />
                        <Bar dataKey="count" name="Frecuencia" fill={ACCENT} radius={[3, 3, 0, 0]} fillOpacity={0.85} />
                      </BarChart>
                    </ResponsiveContainer>
                  </>
                )}

                {chartMode === 'Superficie' && (
                  <>
                    <div className="chart-title">Distribución de Superficies (m²)</div>
                    <ResponsiveContainer width="100%" height={300}>
                      <BarChart data={histSuperficie} barCategoryGap="5%">
                        <CartesianGrid strokeDasharray="3 3" stroke="#e0dbd3" />
                        <XAxis dataKey="midpoint" tickFormatter={v => `${v.toFixed(0)} m²`} tick={{ fill: '#8a8078', fontSize: 11 }} />
                        <YAxis tick={{ fill: '#8a8078', fontSize: 11 }} />
                        <Tooltip content={<CustomTooltip xFmt={v => fmtNum(v, 1) + ' m²'} yFmt={v => fmtNum(v) + ' propiedades'} />} />
                        <Bar dataKey="count" name="Frecuencia" fill={ACCENT2} radius={[3, 3, 0, 0]} fillOpacity={0.85} />
                      </BarChart>
                    </ResponsiveContainer>
                  </>
                )}

                {chartMode === 'Precio por m²' && (
                  <>
                    <div className="chart-title">Distribución de Precio por m² (USD/m²)</div>
                    <ResponsiveContainer width="100%" height={300}>
                      <BarChart data={histM2} barCategoryGap="5%">
                        <CartesianGrid strokeDasharray="3 3" stroke="#e0dbd3" />
                        <XAxis dataKey="midpoint" tickFormatter={v => `$${v.toFixed(0)}`} tick={{ fill: '#8a8078', fontSize: 11 }} />
                        <YAxis tick={{ fill: '#8a8078', fontSize: 11 }} />
                        <Tooltip content={<CustomTooltip xFmt={v => fmtUSD(v, 0)} yFmt={v => fmtNum(v) + ' propiedades'} />} />
                        <Bar dataKey="count" name="Frecuencia" fill={DANGER} radius={[3, 3, 0, 0]} fillOpacity={0.85} />
                      </BarChart>
                    </ResponsiveContainer>
                  </>
                )}

                {chartMode === 'Precios y Superficie' && (
                  <>
                    <div className="chart-title">Precio vs. Superficie</div>
                    <label className="outlier-toggle">
                      <input type="checkbox" checked={eliminarOutliers} onChange={e => setEliminarOutliers(e.target.checked)} />
                      Eliminar outliers en superficie (IQR)
                    </label>
                    <ResponsiveContainer width="100%" height={320}>
                      <ScatterChart>
                        <CartesianGrid strokeDasharray="3 3" stroke="#e0dbd3" />
                        <XAxis dataKey="Superficie_m2" name="Superficie" tickFormatter={v => `${v} m²`} tick={{ fill: '#8a8078', fontSize: 11 }} />
                        <YAxis dataKey="Precio_USD" name="Precio" tickFormatter={v => `$${(v/1000).toFixed(0)}k`} tick={{ fill: '#8a8078', fontSize: 11 }} />
                        <Tooltip
                          cursor={{ strokeDasharray: '3 3' }}
                          content={({ active, payload }) => {
                            if (!active || !payload?.length) return null;
                            const d = payload[0]?.payload;
                            return (
                              <div style={{ background: '#ffffff', border: '1px solid #e0dbd3', borderRadius: 8, padding: '10px 14px', fontSize: 13, color: '#1a1714' }}>
                                <div>Superficie: <b>{fmtNum(d?.Superficie_m2, 1)} m²</b></div>
                                <div>Precio: <b>{fmtUSD(d?.Precio_USD)}</b></div>
                              </div>
                            );
                          }}
                        />
                        <Scatter data={scatterData} fill={ACCENT} fillOpacity={0.5} />
                      </ScatterChart>
                    </ResponsiveContainer>
                    <div style={{ fontSize: 12, color: '#6b6560', marginTop: 8 }}>
                      {scatterData.length} propiedades graficadas
                    </div>
                  </>
                )}
              </div>
            </div>

            {/* Finance */}
            <div className="finance-section">
              <div className="section-title">Módulos financieros</div>

              {/* Finance: Renta */}
              <div className="expander">
                <div className="expander-header" onClick={() => setOpenRenta(o => !o)}>
                  <span className="expander-header-title">Cálculo de renta para inmueble</span>
                  <span className={`expander-arrow${openRenta ? ' open' : ''}`}>▾</span>
                </div>
                <div className={`expander-body${openRenta ? ' open' : ''}`}>
                  <div className={`input-grid input-grid-2`}>
                    <div className="input-group">
                      <label className="input-label">Precio del inmueble (USD)</label>
                      <input className="input-field" type="number" value={precioInmueble}
                        onChange={e => setPrecioInmueble(e.target.value)}
                        placeholder="0" min="0" step="1000000" />
                    </div>
                    <div className="input-group">
                      <label className="input-label">Renta pretendida (%)</label>
                      <input className="input-field" type="number" value={rentaPct}
                        onChange={e => setRentaPct(e.target.value)}
                        min="0" step="0.25" />
                    </div>
                  </div>

                  <div className="metrics-row">
                    <div className="metric-card">
                      <div className="metric-label">Alquiler anual requerido (USD)</div>
                      <div className="metric-value">{fmtUSD(alqAnualReq)}</div>
                    </div>
                    <div className="metric-card">
                      <div className="metric-label">Cuota mensual requerida (USD)</div>
                      <div className="metric-value">{fmtUSD(alqMensualReq)}</div>
                    </div>
                  </div>

                  <div className="caption">
                    Cálculo: alquiler anual requerido = precio del inmueble × renta pretendida.
                  </div>
                </div>
              </div>
            </div>
          </>
        )}

        {!loading && !error && df.length === 0 && rawData.length === 0 && (
          <div className="state-box">Sin datos disponibles.</div>
        )}
      </div>
    </>
  );
}
