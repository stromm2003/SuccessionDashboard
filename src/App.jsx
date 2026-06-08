import React, { useEffect } from 'react';
import './App.css';

// Side-effect import: runs all dashboard logic at module scope,
// registering globals on window for inline event handlers in dynamic HTML.
import './logic.js';

// logic.js assigns all public functions to window at module-load time.

export default function App() {
  useEffect(() => {

    // DOM-dependent init (must run after React renders)
    if (typeof window.renderHub === 'function') window.renderHub();

    const fileInput = document.getElementById('fileInput');
    if (fileInput) {
      fileInput.addEventListener('change', e => {
        if (e.target.files[0] && !e.target.onchange) {
          if (typeof window.ingestFile === 'function') window.ingestFile('sap', e.target.files[0]);
        }
      });
    }

    // Drag-and-drop for dual MyLearning slots
    ['transcript', 'progress'].forEach(slot => {
      const el = document.getElementById('dualSlot_' + slot);
      if (!el) return;
      el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('dragover'); });
      el.addEventListener('dragleave', () => el.classList.remove('dragover'));
      el.addEventListener('drop', e => {
        e.preventDefault(); el.classList.remove('dragover');
        if (e.dataTransfer.files[0] && typeof window._loadDualFile === 'function')
          window._loadDualFile(slot, e.dataTransfer.files[0]);
      });
    });
  }, []);

  return (
    <>
      {/* ══ UPLOAD HUB ══ */}
      <div id="upload-screen">
        <div className="wordmark">HR Talent Management</div>
        <div className="hub-shell">
          <div className="hub-head">
            <div className="hub-icon">
              <svg viewBox="0 0 24 24" fill="none" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"/>
              </svg>
            </div>
            <h1>Talent Management Dashboard</h1>
            <p>Combine data from SAP, Talent Management, PCD/Learning, and other HR systems into a unified employee profile. Upload each source independently — the system merges them automatically by NIK (with Employee Name fallback).</p>
          </div>

          <div className="hub-sources" id="hubSources">{/* populated by JS */}</div>

          <div className="hub-foot">
            <div className="hub-summary" id="hubSummary">
              <div className="hs-stat"><span className="hs-num" id="hsSources">0</span><span className="hs-lbl">Sources Loaded</span></div>
              <div className="hs-stat"><span className="hs-num" id="hsRows">0</span><span className="hs-lbl">Raw Rows</span></div>
              <div className="hs-stat"><span className="hs-num" id="hsEmployees">0</span><span className="hs-lbl">Unique Employees</span></div>
              <div className="hs-stat"><span className="hs-num" id="hsIssues">0</span><span className="hs-lbl">Validation Issues</span></div>
            </div>
            <div className="hub-actions">
              <button className="btn-ghost" onClick={() => window.openIssuesPanel()} id="btnIssues" style={{ display: 'none' }}>
                <svg viewBox="0 0 24 24" fill="none" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 9v4M12 17h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/></svg>
                Review Issues
              </button>
              <button className="btn-primary" id="btnBuild" onClick={() => window.buildAndEnter()} disabled>
                <svg viewBox="0 0 24 24" fill="none" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M5 12h14M12 5l7 7-7 7"/></svg>
                Build Dashboard
              </button>
            </div>
          </div>
        </div>
        <input type="file" id="fileInput" accept=".xlsx,.xls,.csv,.tsv" style={{ display: 'none' }} />
      </div>

      {/* ══ COLUMN MAPPER MODAL ══ */}
      <div className="modal-backdrop" id="mapperModal">
        <div className="modal-box modal-lg">
          <div className="modal-title" id="mapTitle">Map Columns</div>
          <div className="modal-sub" id="mapSub">Auto-detected mappings shown below. Adjust as needed, or leave unknowns as "Ignore".</div>
          <div className="mapper-summary" id="mapSummary"></div>
          <div className="mapper-list" id="mapList"></div>
          <div className="modal-actions">
            <button className="btn-mc" onClick={() => window.closeMapper()}>Cancel</button>
            <button className="btn-mok" onClick={() => window.confirmMapping()}>Confirm &amp; Validate</button>
          </div>
        </div>
      </div>

      {/* ══ VALIDATION ISSUES PANEL ══ */}
      <div className="modal-backdrop" id="issuesModal">
        <div className="modal-box modal-lg">
          <div className="modal-title">Validation Issues</div>
          <div className="modal-sub">Issues are non-blocking unless marked CRITICAL. The dashboard will still build with available data.</div>
          <div className="issues-tabs" id="issuesTabs"></div>
          <div className="issues-list" id="issuesList"></div>
          <div className="modal-actions">
            <button className="btn-mok" onClick={() => window.closeIssuesPanel()}>Close</button>
          </div>
        </div>
      </div>

      {/* ══ VALIDATION MODAL ══ */}
      <div className="modal-backdrop" id="validModal">
        <div className="modal-box">
          <div className="modal-title" id="vTitle">Validating File…</div>
          <div className="modal-sub" id="vSub"></div>
          <div className="modal-errors" id="vErrors"></div>
          <div className="modal-actions">
            <button className="btn-mc" onClick={() => window.closeModal()}>Close</button>
            <button className="btn-mok" id="vOk" onClick={() => window.confirmLoad()} style={{ display: 'none' }}>Load Dashboard</button>
          </div>
        </div>
      </div>

      {/* ══ DASHBOARD ══ */}
      <div id="dashboard">
        <div className="topbar">
          <div className="tb-brand">
            <div className="tb-logo">
              <svg viewBox="0 0 24 24" fill="none" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2M9 7a4 4 0 100 8 4 4 0 000-8zM23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75"/>
              </svg>
            </div>
            <div>
              <div className="tb-name">Succession Planning</div>
              <div className="tb-sub">Talent Management Dashboard</div>
            </div>
          </div>
          <div className="tb-right">
            <span className="pill" id="topFile">—</span>
            <span className="pill" id="topRec">0 records</span>
            <button className="btn-sm" onClick={() => window.resetApp()}>↑ New File</button>
          </div>
        </div>

        <div className="page" id="pdfTarget">
          {/* FILTER BAR */}
          <div className="filter-bar" id="filterBar"><span className="filter-lbl">Filter by:</span></div>

          {/* ══ PEOPLE SNAPSHOT ══ */}
          <div style={{ marginBottom: '.4rem' }}>
            <span className="sec-hdr">People Snapshot</span>
          </div>
          <div className="snap-grid mb" id="snapGrid"></div>

          {/* ══ SNAP INLINE DETAIL PANEL ══ */}
          <div id="snapInlinePanel" style={{ display: 'none', marginBottom: '1rem' }}>
            <div style={{ background: 'var(--surface)', border: '1.5px solid var(--border2)', borderRadius: 'var(--r)', boxShadow: 'var(--shadowmd)', overflow: 'hidden' }}>
              <div style={{ background: 'linear-gradient(90deg,var(--navy),var(--dark))', padding: '.9rem 1.25rem', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '12px' }}>
                <div>
                  <div id="snapModalTitle" style={{ fontSize: '14px', fontWeight: 700, color: '#fff' }}></div>
                  <div id="snapModalSub" style={{ fontSize: '11px', color: 'var(--t3)', marginTop: '2px', fontWeight: 300 }}></div>
                </div>
                <div onClick={() => window.closeSnapModal()} style={{ width: '28px', height: '28px', borderRadius: '6px', background: 'rgba(255,255,255,.1)', border: '1px solid rgba(255,255,255,.15)', display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', color: '#fff', fontSize: '16px', lineHeight: 1, flexShrink: 0 }}>&#x2715;</div>
              </div>
              <div style={{ padding: '.8rem 1.25rem .4rem', borderBottom: '1px solid var(--border)' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="var(--t3)" strokeWidth="2" strokeLinecap="round" style={{ width: '14px', height: '14px', flexShrink: 0 }}><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                  <input className="snap-modal-search" id="snapModalSearch" type="text" placeholder="Search by name, role, or job family…" onInput={() => window.filterSnapModal()} style={{ flex: 1, minWidth: '180px' }} />
                  <span style={{ fontSize: '11px', fontWeight: 600, color: 'var(--t2)', whiteSpace: 'nowrap' }}>Filter :</span>
                  <select id="snapFilterSelect" onChange={e => window.setSnapFilter(e.target.value)} style={{ padding: '4px 26px 4px 9px', background: 'var(--skyxs)', border: '1px solid var(--border2)', borderRadius: 'var(--rsm)', fontSize: '11px', fontFamily: 'inherit', color: 'var(--navy)', cursor: 'pointer', flexShrink: 0 }}>
                    <option value="unique">Unique</option>
                    <option value="nonunique">Non-Unique</option>
                    <option value="all">All</option>
                  </select>
                </div>
                <div id="snapModalCount" style={{ fontSize: '10px', color: 'var(--t2)', marginTop: '.45rem', fontWeight: 500 }}></div>
              </div>
              <div id="snapSuccFilterRow" style={{ display: 'none', padding: '.5rem 1.25rem', borderBottom: '1px solid var(--border)', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                <span style={{ fontSize: '11px', fontWeight: 600, color: 'var(--t2)', whiteSpace: 'nowrap' }}>Successor:</span>
                <button className="snap-succ-btn active" data-f="all" onClick={() => window.setSnapSuccFilter('all')} style={{ padding: '3px 10px', fontSize: '11px', fontFamily: 'inherit', borderRadius: '999px', border: '1px solid var(--border2)', background: 'var(--navy)', color: '#fff', cursor: 'pointer', fontWeight: 600 }}>All</button>
                <button className="snap-succ-btn" data-f="with" onClick={() => window.setSnapSuccFilter('with')} style={{ padding: '3px 10px', fontSize: '11px', fontFamily: 'inherit', borderRadius: '999px', border: '1px solid var(--border2)', background: 'var(--skyxs)', color: 'var(--navy)', cursor: 'pointer', fontWeight: 500 }}>With Successor</button>
                <button className="snap-succ-btn" data-f="without" onClick={() => window.setSnapSuccFilter('without')} style={{ padding: '3px 10px', fontSize: '11px', fontFamily: 'inherit', borderRadius: '999px', border: '1px solid var(--border2)', background: 'var(--skyxs)', color: 'var(--navy)', cursor: 'pointer', fontWeight: 500 }}>Without Successor</button>
              </div>
              <div style={{ padding: '.6rem 1.25rem 0', overflowX: 'auto' }}>
                <table className="snap-emp-table" id="snapModalTable">
                  <thead id="snapModalHead"></thead>
                  <tbody id="snapModalBody"></tbody>
                </table>
              </div>
              <div id="snapPgBar" className="pg-bar" style={{ padding: '.7rem 1.25rem 1rem', borderTop: '1px solid var(--border)', marginTop: '.4rem' }}></div>
            </div>
          </div>

          {/* ══ POSTURE + READINESS ══ */}
          <div className="posture-row mb" id="postureRow">
            <div className="posture-panel">
              <div className="posture-hdr">Succession Posture</div>
              <div className="posture-body"><div id="postureItems"></div></div>
            </div>
            <div className="clevel-panel">
              <div className="clevel-hdr">C-Level Succession Posture</div>
              <div className="clevel-body">
                <div id="clevelByPos"></div>
                <div id="clevelByEmp"></div>
                <div style={{ marginTop: '10px', paddingTop: '10px', borderTop: '1px solid var(--border)' }}>
                  <div style={{ fontSize: '9px', fontWeight: 700, letterSpacing: '.1em', textTransform: 'uppercase', color: 'var(--t3)', marginBottom: '6px' }}>Bench Strength</div>
                  <div style={{ background: 'var(--skyxs)', border: '1px solid var(--border)', borderRadius: '9px', padding: '.7rem 1rem', marginBottom: '8px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '16px' }}>
                      <div><div className="clevel-tile-label">Overall — Total C-Position / Total Successor Plans</div></div>
                      <div className="clevel-tile-val" id="benchStrengthValAll" style={{ flexShrink: 0, whiteSpace: 'nowrap' }}>—</div>
                    </div>
                    <div id="benchDetailAreaAll" style={{ marginTop: '8px', paddingTop: '8px', borderTop: '1px solid var(--border)' }}></div>
                  </div>
                  <div style={{ background: 'var(--skyxs)', border: '1px solid var(--border)', borderRadius: '9px', padding: '.7rem 1rem', marginBottom: '8px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '16px' }}>
                      <div><div className="clevel-tile-label">Unique — Total C-Position / Total Successor Plans</div></div>
                      <div className="clevel-tile-val" id="benchStrengthVal" style={{ flexShrink: 0, whiteSpace: 'nowrap' }}>—</div>
                    </div>
                    <div id="benchDetailArea" style={{ marginTop: '8px', paddingTop: '8px', borderTop: '1px solid var(--border)' }}></div>
                  </div>
                  <div style={{ background: 'var(--skyxs)', border: '1px solid var(--border)', borderRadius: '9px', padding: '.7rem 1rem' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '16px' }}>
                      <div><div className="clevel-tile-label">Non-Unique — Total C-Position / Total Successor Plans</div></div>
                      <div className="clevel-tile-val" id="benchStrengthValNU" style={{ flexShrink: 0, whiteSpace: 'nowrap' }}>—</div>
                    </div>
                    <div id="benchDetailAreaNU" style={{ marginTop: '8px', paddingTop: '8px', borderTop: '1px solid var(--border)' }}></div>
                  </div>
                </div>
              </div>
            </div>
            <div className="readiness-panel">
              <div className="readiness-hdr">Succession Readiness</div>
              <div className="readiness-body">
                <div id="readinessItems"></div>
                <div style={{ marginTop: '10px', paddingTop: '10px', borderTop: '1px solid var(--border)' }}>
                  <div style={{ fontSize: '9px', fontWeight: 700, letterSpacing: '.1em', textTransform: 'uppercase', color: 'var(--t3)', marginBottom: '6px' }}>Successor Availability</div>
                  <div className="readiness-card total" style={{ marginBottom: '6px', display: 'block', padding: '10px 12px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <div>
                        <div className="readiness-card-lbl">C-Position Coverage</div>
                        <div className="readiness-card-sub" id="availSub">C-Level positions covered</div>
                      </div>
                      <div style={{ textAlign: 'right', flexShrink: 0 }}>
                        <div className="readiness-card-val" style={{ color: '#7C3AED', fontSize: '32px', fontWeight: 800, letterSpacing: '-1px', lineHeight: 1 }} id="availVal">—</div>
                      </div>
                    </div>
                    <div id="clAvailBreakdown" style={{ marginTop: '6px', display: 'flex', flexDirection: 'column', gap: '3px', borderTop: '1px solid var(--border)', paddingTop: '6px' }}></div>
                    <button className="rd-view-btn" onClick={() => window.openSnapModal('avail_cl')} style={{ marginTop: '8px' }}>View Details →</button>
                  </div>
                  <div className="readiness-card" style={{ marginBottom: '4px', display: 'block', padding: '10px 12px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <div>
                        <div className="readiness-card-lbl">Non C-Position Coverage</div>
                        <div className="readiness-card-sub" id="nonClAvailSub">Non C-Level positions covered</div>
                      </div>
                      <div style={{ textAlign: 'right', flexShrink: 0 }}>
                        <div className="readiness-card-val rl" id="nonClAvailVal">—</div>
                      </div>
                    </div>
                    <div id="nonClAvailBreakdown" style={{ marginTop: '6px', display: 'flex', flexDirection: 'column', gap: '3px', borderTop: '1px solid var(--border)', paddingTop: '6px' }}></div>
                    <button className="rd-view-btn" onClick={() => window.openSnapModal('avail_noncl')} style={{ marginTop: '8px' }}>View Details →</button>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* ══ TABS ══ */}
          <div className="tabs">
            <button className="tab active" onClick={e => window.switchTab(e.currentTarget,'overview')}>CXO Distribution</button>
            <button className="tab" onClick={e => window.switchTab(e.currentTarget,'positions')}>Position Coverage</button>
            <button className="tab" onClick={e => window.switchTab(e.currentTarget,'simulation')}>Succession Planning Tools</button>
            <button className="tab" onClick={e => window.switchTab(e.currentTarget,'detail')}>Employee Detail</button>
          </div>

          {/* ── OVERVIEW TAB ── */}
          <div className="tab-panel active" id="tab-overview">
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '.8rem' }}>
              <svg viewBox="0 0 24 24" fill="none" stroke="var(--t3)" strokeWidth="2" strokeLinecap="round" style={{ width: '14px', height: '14px', flexShrink: 0 }}><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <span style={{ fontSize: '11px', fontWeight: 600, color: 'var(--t2)' }}>Filter by Business Unit:</span>
              <select id="overviewBuFilter" style={{ padding: '4px 26px 4px 9px', background: 'var(--skyxs)', border: '1px solid var(--border2)', borderRadius: 'var(--rsm)', fontSize: '11px', fontFamily: 'inherit', color: 'var(--navy)', cursor: 'pointer', minWidth: '180px' }}><option value="">All Business Units</option></select>
              <select id="overviewClFilter" style={{ display: 'none' }}><option value="Yes" defaultValue>Yes</option></select>
              <span style={{ fontSize: '11px', fontWeight: 600, color: 'var(--t2)', whiteSpace: 'nowrap' }}>Filter :</span>
              <select id="overviewUniqFilter" onChange={() => window.applyOverviewUniqFilter()} style={{ padding: '4px 26px 4px 9px', background: 'var(--skyxs)', border: '1px solid var(--border2)', borderRadius: 'var(--rsm)', fontSize: '11px', fontFamily: 'inherit', color: 'var(--navy)', cursor: 'pointer' }}>
                <option value="all">All</option><option value="unique">Unique</option><option value="nonunique">Non-Unique</option>
              </select>
            </div>
            <div className="g21 mb">
              <div className="card">
                <div className="card-title">C-Level Successor and Talent Count by Job Family</div>
                <div style={{ position: 'relative', height: '340px' }}><canvas id="buChart"></canvas></div>
              </div>
              <div>
                <div className="card mb">
                  <div className="card-title">C-Level Succession Plans by Incumbent Age Range</div>
                  <div style={{ position: 'relative', height: '200px' }}><canvas id="ageChart"></canvas></div>
                </div>
                <div className="card">
                  <div className="card-title">C-Level Succession Plans by Successor Level</div>
                  <div style={{ position: 'relative', height: '160px' }}><canvas id="layerChart"></canvas></div>
                </div>
              </div>
            </div>
            <div className="g2 mb">
              <div className="card">
                <div className="card-title">CxO Succession Plans by Job Family</div>
                <div style={{ position: 'relative' }} id="jobFamilyWrap"><canvas id="jobFamilyChart"></canvas></div>
              </div>
              <div>
                <div className="g2 mb">
                  <div className="card">
                    <div className="card-title">Readiness Distribution</div>
                    <div style={{ position: 'relative', height: '180px' }}><canvas id="readinessChart"></canvas></div>
                  </div>
                  <div className="card">
                    <div className="card-title">As Talent Calibrated</div>
                    <div style={{ position: 'relative', height: '180px' }}><canvas id="flightChart"></canvas></div>
                  </div>
                </div>
                <div className="g2">
                  <div className="card">
                    <div className="card-title">Critical Position vs Non-Critical Position</div>
                    <div style={{ position: 'relative', height: '150px' }}><canvas id="critChart"></canvas></div>
                  </div>
                  <div className="card">
                    <div className="card-title">Retirement Risk</div>
                    <div style={{ position: 'relative', height: '150px' }}><canvas id="retChart"></canvas></div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* ── POSITIONS TAB ── */}
          <div className="tab-panel" id="tab-positions">
            <div className="card mb" style={{ paddingBottom: '.6rem' }}>
              <div className="pos-toolbar">
                <div className="card-title" style={{ marginBottom: '.6rem' }}>Position Coverage</div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '.4rem' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="var(--t3)" strokeWidth="2" strokeLinecap="round" style={{ width: '14px', height: '14px', flexShrink: 0 }}><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                  <span style={{ fontSize: '11px', fontWeight: 600, color: 'var(--t2)' }}>Filter by Business Unit:</span>
                  <select id="posBuFilter" style={{ padding: '4px 26px 4px 9px', background: 'var(--skyxs)', border: '1px solid var(--border2)', borderRadius: 'var(--rsm)', fontSize: '11px', fontFamily: 'inherit', color: 'var(--navy)', cursor: 'pointer', minWidth: '180px' }}><option value="">All Business Units</option></select>
                </div>
                <div className="pos-folders" id="posFolders">
                  <button className="pf on" id="pf-all" onClick={e => window.setPF(e.currentTarget,'all')}>📋 All <span className="pf-count" id="pfct-all">—</span></button>
                  <button className="pf" id="pf-cl" onClick={e => window.setPF(e.currentTarget,'cl')}>⭐ C-Level <span className="pf-count" id="pfct-cl">—</span></button>
                  <button className="pf" id="pf-crit" onClick={e => window.setPF(e.currentTarget,'crit')}>🔴 Critical <span className="pf-count" id="pfct-crit">—</span></button>
                  <button className="pf" id="pf-risk" onClick={e => window.setPF(e.currentTarget,'risk')}>🔥 High Risk <span className="pf-count" id="pfct-risk">—</span></button>
                </div>
              </div>
              <div style={{ height: '12px' }}></div>
              <div className="pg-search-wrap">
                <svg viewBox="0 0 24 24" fill="none" strokeWidth="2" strokeLinecap="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                <input className="pg-search" id="posSearch" type="text" placeholder="Search by position name, incumbent, job family…" onInput={() => window.renderPositions()} />
                <span id="posSearchCount" style={{ fontSize: '11px', color: 'var(--t2)' }}></span>
              </div>
              <div id="posGrid"></div>
              <div className="pg-bar" id="posPgBar"></div>
            </div>
            <div className="sim-row" style={{ marginTop: '1rem' }}>
              <div className="card">
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '.6rem', flexWrap: 'wrap', gap: '8px' }}>
                  <div className="card-title" style={{ margin: 0 }}>Positions Without Any Successor (Gaps)</div>
                </div>
                <table className="btable"><thead id="gapThead"><tr><th>Position</th><th>Level</th><th>C-Level</th><th>Critical</th><th>Risk</th></tr></thead>
                <tbody id="gapBody"></tbody></table>
                <div className="pg-bar" id="pgGapBar"></div>
              </div>
              <div className="card">
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '.6rem', flexWrap: 'wrap', gap: '8px' }}>
                  <div className="card-title" style={{ margin: 0 }}>High-Risk: Critical + No Successor</div>
                </div>
                <table className="btable"><thead id="hrThead"><tr><th>Position</th><th>Level</th><th>C-Level</th><th>Incumbent</th><th>Risk</th></tr></thead>
                <tbody id="hrBody"></tbody></table>
                <div className="pg-bar" id="pgHRBar"></div>
              </div>
            </div>
          </div>

          {/* ── SIMULATION TAB ── */}
          <div className="tab-panel" id="tab-simulation">
            <div className="card mb">
              <div className="card-title">Succession Scenario Simulator</div>
              <p style={{ fontSize: '12px', color: '#5a7fa8', marginBottom: '1rem', fontWeight: 400 }}>Select a position to simulate a vacancy. The system recommends the best available successor ranked by readiness and flight risk.</p>
              <div className="sim-sel-row">
                <select id="simBuFilter" onChange={() => window._filterSimPositions()} style={{ flex: '0 0 auto', minWidth: '200px', maxWidth: '240px' }}><option value="">— All Business Units —</option></select>
                <select id="simPosSelect"><option value="">— Select a position —</option></select>
                <button className="btn-sim" onClick={() => window.runSimulation()}>▶ Run Scenario</button>
              </div>
              <div className="sim-result" id="simResult">Select a position above and click "Run Scenario".</div>
            </div>

            <div className="card mb" style={{ border: '1.5px solid var(--border2)' }}>
              <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: '.9rem', flexWrap: 'wrap', gap: '8px' }}>
                <div>
                  <div className="card-title" style={{ marginBottom: '3px' }}>📄 Generate Succession Card for Targeted Position</div>
                  <p style={{ fontSize: '12px', color: '#5a7fa8', fontWeight: 400, margin: 0 }}>Choose an employee and the position you want them to succeed to. The system will fill the card from Excel data and export a ready-to-use .pptx file.</p>
                </div>
              </div>
              <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'flex-end', marginBottom: '.85rem' }}>
                <div style={{ flex: 1, minWidth: '200px' }}>
                  <div style={{ fontSize: '10px', fontWeight: 700, letterSpacing: '.1em', textTransform: 'uppercase', color: 'var(--t3)', marginBottom: '4px' }}>Employee</div>
                  <select id="scEmp" onChange={() => window.scPreview()} style={{ width: '100%', padding: '8px 26px 8px 10px', border: '1px solid var(--border2)', borderRadius: 'var(--rsm)', fontFamily: 'inherit', fontSize: '12px', fontWeight: 500, background: 'var(--skyxs)', color: 'var(--text)', appearance: 'none', outline: 'none' }}>
                    <option value="">— Select an employee —</option>
                  </select>
                </div>
                <div style={{ flex: 1, minWidth: '200px' }}>
                  <div style={{ fontSize: '10px', fontWeight: 700, letterSpacing: '.1em', textTransform: 'uppercase', color: 'var(--t3)', marginBottom: '4px' }}>Targeted Position</div>
                  <select id="scPos" onChange={() => window.scPreview()} style={{ width: '100%', padding: '8px 26px 8px 10px', border: '1px solid var(--border2)', borderRadius: 'var(--rsm)', fontFamily: 'inherit', fontSize: '12px', fontWeight: 500, background: 'var(--skyxs)', color: 'var(--text)', appearance: 'none', outline: 'none' }}>
                    <option value="">— Select a position —</option>
                  </select>
                </div>
                <div>
                  <button className="btn-sim" id="scGenBtn" onClick={() => window.scGenerate()} style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '8px 18px' }}>
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
                    Generate Card
                  </button>
                </div>
              </div>
              <div id="scPreviewBox" style={{ display: 'none', background: 'var(--skyxs)', border: '1px solid var(--border2)', borderRadius: 'var(--rsm)', padding: '.8rem 1rem' }}>
                <div style={{ fontSize: '10px', fontWeight: 700, letterSpacing: '.1em', textTransform: 'uppercase', color: 'var(--t3)', marginBottom: '.6rem' }}>Data to be filled into the card</div>
                <div id="scPreviewGrid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill,minmax(160px,1fr))', gap: '10px', fontSize: '12px' }}></div>
                <div id="scStatusMsg" style={{ marginTop: '.6rem', fontSize: '11px', fontWeight: 600 }}></div>
              </div>
            </div>
          </div>

          {/* ── DETAIL TAB ── */}
          <div className="tab-panel" id="tab-detail">
            <div className="card">
              <div className="card-title">Full Succession Detail</div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '.8rem' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="var(--t3)" strokeWidth="2" strokeLinecap="round" style={{ width: '14px', height: '14px', flexShrink: 0 }}><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                <span style={{ fontSize: '11px', fontWeight: 600, color: 'var(--t2)' }}>Filter by Business Unit:</span>
                <select id="detailBuFilter" onChange={() => window.renderDetail()} style={{ padding: '4px 26px 4px 9px', background: 'var(--skyxs)', border: '1px solid var(--border2)', borderRadius: 'var(--rsm)', fontSize: '11px', fontFamily: 'inherit', color: 'var(--navy)', cursor: 'pointer', minWidth: '180px' }}><option value="">All Business Units</option></select>
                <span style={{ fontSize: '11px', fontWeight: 600, color: 'var(--t2)' }}>C-Level:</span>
                <select id="detailClFilter" onChange={() => window.renderDetail()} style={{ padding: '4px 26px 4px 9px', background: 'var(--skyxs)', border: '1px solid var(--border2)', borderRadius: 'var(--rsm)', fontSize: '11px', fontFamily: 'inherit', color: 'var(--navy)', cursor: 'pointer' }}>
                  <option value="">All</option><option value="Yes">Yes</option><option value="No">No</option>
                </select>
              </div>
              <div className="pg-search-wrap">
                <svg viewBox="0 0 24 24" fill="none" strokeWidth="2" strokeLinecap="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                <input className="pg-search" id="detailSearch" type="text" placeholder="Search by name, position, job family, successor…" onInput={() => window.detailSearchHandler()} />
                <span id="detailSearchCount" style={{ fontSize: '11px', color: 'var(--t2)' }}></span>
              </div>
              <div className="tbl-wrap">
                <table className="dtbl"><thead><tr id="detailHead"></tr></thead><tbody id="detailBody"></tbody></table>
              </div>
              <div className="pg-bar" id="detailPgBar"></div>
            </div>
          </div>

        </div>{/* end .page */}
      </div>{/* end #dashboard */}

      <div className="pdf-loading" id="pdfLoading"><div className="spinner"></div><p>Generating PDF…</p></div>

      {/* ══ POSITION DETAIL MODAL ══ */}
      <div className="pos-detail-backdrop" id="posDetailBackdrop" onClick={e => { if(e.target===e.currentTarget) window.closePosDetail(); }}>
        <div className="pos-detail-modal" id="posDetailModal">
          <div className="pos-detail-hdr">
            <div>
              <div id="pdTitle" style={{ fontSize: '15px', fontWeight: 700, color: '#fff', marginBottom: '3px' }}></div>
              <div id="pdIncumbent" style={{ fontSize: '12px', color: 'var(--t3)', fontWeight: 400 }}></div>
            </div>
            <div className="pos-detail-close" onClick={() => window.closePosDetail()}>&#x2715;</div>
          </div>
          <div className="pos-detail-body" id="posDetailBody"></div>
        </div>
      </div>

      {/* ══ SUCCESSOR POSITIONS POPUP ══ */}
      <div className="succ-popup-backdrop" id="succPopupBackdrop" onClick={e => { if(e.target===e.currentTarget) window.closeSuccPopup(); }}>
        <div className="succ-popup">
          <div className="succ-popup-hdr">
            <div>
              <div id="succPopupName" style={{ fontSize: '13px', fontWeight: 700, color: '#fff' }}></div>
              <div id="succPopupSub" style={{ fontSize: '10px', color: 'rgba(255,255,255,.6)', marginTop: '2px' }}></div>
            </div>
            <div className="succ-popup-close" onClick={() => window.closeSuccPopup()}>&#x2715;</div>
          </div>
          <div className="succ-popup-body" id="succPopupBody"></div>
        </div>
      </div>

      {/* ══ MYLEARNING DUAL UPLOAD MODAL ══ */}
      <div className="rd-backdrop" id="dualBackdrop" onClick={e => { if(e.target===e.currentTarget) window.closeMyLearningDualUpload(); }}>
        <div className="rd-modal" style={{ maxWidth: '780px' }}>
          <div className="rd-hdr">
            <div>
              <div className="rd-title">MyLearning Upload</div>
              <div className="rd-sub">Upload Learning Transcript and Course Completion files — they'll be merged by Employee ID or Email</div>
            </div>
            <div className="rd-close" onClick={() => window.closeMyLearningDualUpload()}>&#x2715;</div>
          </div>
          <div className="rd-body" style={{ padding: '24px' }}>
            <div className="dual-grid">
              <div className="dual-slot" id="dualSlot_transcript" data-slot="transcript">
                <div className="dual-slot-num">1</div>
                <div className="dual-slot-icn">
                  <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="13" x2="15" y2="13"/><line x1="9" y1="17" x2="15" y2="17"/></svg>
                </div>
                <div className="dual-slot-lbl">Employee Learning Transcript</div>
                <div className="dual-slot-desc">Per-employee transcript records<br/>(courses enrolled, history)</div>
                <button className="dual-slot-btn" onClick={() => window.pickDualFile('transcript')}>Choose File</button>
                <div className="dual-slot-meta" id="dualMeta_transcript"></div>
              </div>
              <div className="dual-slot" id="dualSlot_progress" data-slot="progress">
                <div className="dual-slot-num">2</div>
                <div className="dual-slot-icn">
                  <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
                </div>
                <div className="dual-slot-lbl">Course Progress / Completion</div>
                <div className="dual-slot-desc">Completion % and status<br/>per course / enrollment</div>
                <button className="dual-slot-btn" onClick={() => window.pickDualFile('progress')}>Choose File</button>
                <div className="dual-slot-meta" id="dualMeta_progress"></div>
              </div>
            </div>
            <div className="dual-info">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0, marginTop: '1px' }}><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>
              <div>
                <strong>How merge works:</strong> employees are matched by NIK/Employee ID first, then Email, then Employee Name.
                Records found in both files are combined into one row per (employee × course). Records in only one file are still kept.
                Supported: <code>.xlsx</code>, <code>.xls</code>, <code>.csv</code> (semicolon or comma).
              </div>
            </div>
            <div className="dual-progress" id="dualProgress" style={{ display: 'none' }}>
              <div className="dual-progress-lbl" id="dualProgressLbl">Processing…</div>
              <div className="dual-progress-bar"><div className="dual-progress-fill" id="dualProgressFill"></div></div>
            </div>
            <div className="dual-result" id="dualResult"></div>
            <div className="dual-footer">
              <button className="dual-cancel" onClick={() => window.closeMyLearningDualUpload()}>Cancel</button>
              <button className="dual-process" id="dualProcessBtn" onClick={() => window.processDualUpload()} disabled>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
                Process &amp; Load
              </button>
            </div>
          </div>
        </div>
      </div>
      <input type="file" id="dualFileInput" accept=".xlsx,.xls,.csv,.tsv" style={{ display: 'none' }} />

      {/* ══ READINESS DETAIL MODAL ══ */}
      <div className="rd-backdrop" id="rdBackdrop" onClick={e => { if(e.target===e.currentTarget) window.closeReadinessDetail(); }}>
        <div className="rd-modal">
          <div className="rd-hdr">
            <div>
              <div className="rd-title" id="rdTitle">Successor Details</div>
              <div className="rd-sub" id="rdSub"></div>
            </div>
            <div className="rd-actions">
              <button className="rd-export-btn" onClick={() => window.exportReadinessDetail()} title="Export to Excel">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                Export Excel
              </button>
              <div className="rd-close" onClick={() => window.closeReadinessDetail()}>&#x2715;</div>
            </div>
          </div>
          <div className="rd-filters">
            <div className="rd-filter-row">
              <div className="rd-search-wrap">
                <svg className="rd-search-icn" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                <input type="search" className="rd-search" id="rdSearch" placeholder="Search by name or employee ID…" onInput={() => window.renderReadinessDetailRows()} />
              </div>
              <select className="rd-select" id="rdDept" onChange={() => window.renderReadinessDetailRows()}><option value="">All Departments</option></select>
              <select className="rd-select" id="rdLevel" onChange={() => window.renderReadinessDetailRows()}>
                <option value="">All Levels</option>
                <option value="c">C-Level Only</option>
                <option value="non">Non C-Level Only</option>
              </select>
              <select className="rd-select" id="rdLearnStatus" onChange={() => window.renderReadinessDetailRows()}>
                <option value="">All Learning Status</option>
                <option value="Completed">Completed</option>
                <option value="In Progress">In Progress</option>
                <option value="Not Started">Not Started</option>
              </select>
              <select className="rd-select" id="rdPct" onChange={() => window.renderReadinessDetailRows()}>
                <option value="">Any Completion %</option>
                <option value="100">= 100%</option>
                <option value="75">≥ 75%</option>
                <option value="50">≥ 50%</option>
                <option value="25">≥ 25%</option>
                <option value="0">0% (Not Started)</option>
              </select>
            </div>
            <div className="rd-result-count" id="rdCount"></div>
          </div>
          <div className="rd-body">
            <table className="rd-table" id="rdTable">
              <thead>
                <tr>
                  <th>Employee</th><th>NIK</th><th>Current Position</th><th>Successor For</th>
                  <th>Department</th><th>Readiness</th><th>Courses</th>
                  <th>Overall %</th><th>Learning Status</th><th>Updated</th>
                </tr>
              </thead>
              <tbody id="rdTbody"></tbody>
            </table>
            <div className="rd-empty" id="rdEmpty" style={{ display: 'none' }}>
              <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ opacity: .4, marginBottom: '12px' }}><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <div style={{ fontSize: '14px', fontWeight: 600, marginBottom: '4px' }}>No successors match your filters</div>
              <div style={{ fontSize: '12px', color: 'var(--t3)' }}>Try clearing filters or broadening your search</div>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
