"""
build_dashboard.py — AIDA Ventures Integrated Dashboard
Generates outputs/index.html from Excel sources.
Run: python3 build_dashboard.py
"""
import os
import sys
import json
import re
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
import extract_data

OUT_DIR = os.path.join(os.path.dirname(__file__), '..', 'outputs')
GEN_DATE = datetime.now().strftime('%B %d, %Y')


# ─── CSS ─────────────────────────────────────────────────────────────────────

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,300;1,400;1,600&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --bg:         #F5F7FA;
  --surface:    #FFFFFF;
  --surface2:   #EEF2F7;
  --border:     #DDE4EE;
  --navy:       #0B1F3A;
  --navy-mid:   #1A3A5F;
  --navy-light: #2C5282;
  --accent:     #1A6FB5;
  --teal:       #0D7377;
  --teal-light: rgba(13,115,119,0.08);
  --gold:       #B8882A;
  --gold-light: rgba(184,136,42,0.10);
  --green:      #15803D;
  --green-bg:   rgba(21,128,61,0.08);
  --amber:      #D97706;
  --amber-bg:   rgba(217,119,6,0.09);
  --red:        #DC2626;
  --red-bg:     rgba(220,38,38,0.08);
  --text:       #0B1F3A;
  --text2:      #2D4A6B;
  --muted:      #5A7290;
  --dim:        #8BA0B8;
  --shadow-sm:  0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04);
  --shadow:     0 4px 12px rgba(0,0,0,0.07), 0 1px 3px rgba(0,0,0,0.05);
  --shadow-lg:  0 8px 30px rgba(0,0,0,0.09), 0 2px 8px rgba(0,0,0,0.05);
  --radius:     12px;
  --radius-sm:  8px;
  --radius-lg:  16px;
  --font-display: 'Inter', system-ui, sans-serif;
  --font-body:    'Inter', system-ui, sans-serif;
  --font-mono:    'JetBrains Mono', monospace;
}

*, *::before, *::after { margin: 0; padding: 0; box-sizing: border-box; }
html { scroll-behavior: smooth; font-size: 16px; }
body {
  background: var(--bg);
  color: var(--text);
  font-family: var(--font-body);
  line-height: 1.6;
  overflow-x: hidden;
}

/* ── SCROLLBAR ── */
::-webkit-scrollbar { width: 6px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }

/* ── NAVBAR ── */
#navbar {
  position: fixed; top: 0; left: 0; right: 0; z-index: 100;
  background: rgba(255,255,255,0.95);
  backdrop-filter: blur(10px);
  border-bottom: 1px solid var(--border);
  display: flex; align-items: center; justify-content: space-between;
  padding: 0 40px; height: 64px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.nav-brand {
  display: flex; align-items: center; gap: 10px; text-decoration: none;
}
.nav-brand-icon {
  width: 32px; height: 32px; background: var(--navy);
  border-radius: 6px; display: flex; align-items: center; justify-content: center;
}
.nav-brand-icon svg { width: 20px; height: 20px; }
.nav-brand-name {
  font-family: var(--font-display); font-size: 18px; font-weight: 600;
  color: var(--navy); letter-spacing: 0.01em;
}
.nav-brand-name span { color: var(--accent); }
.nav-links { display: flex; gap: 4px; }
.nav-link {
  padding: 6px 14px; border-radius: 6px; text-decoration: none;
  font-size: 13px; font-weight: 500; color: var(--muted);
  transition: all 0.15s ease; cursor: pointer; border: none; background: none;
  white-space: nowrap;
}
.nav-link:hover, .nav-link.active {
  background: var(--surface2); color: var(--navy);
}
.nav-right {
  font-size: 11px; color: var(--dim); font-family: var(--font-mono);
  letter-spacing: 0.02em;
}

/* ── SECTIONS ── */
.section {
  display: none; min-height: 100vh; padding: 88px 40px 48px;
  max-width: 1280px; margin: 0 auto;
}
.section.active { display: block; }

/* ── SECTION HEADERS ── */
.section-eyebrow {
  font-size: 11px; font-weight: 600; letter-spacing: 0.18em;
  text-transform: uppercase; color: var(--accent);
  margin-bottom: 8px;
}
.section-title {
  font-family: var(--font-display);
  font-size: 32px; font-weight: 700; line-height: 1.2;
  color: #0B1F3A; margin-bottom: 8px;
}
.section-title em { font-style: italic; color: var(--accent); font-weight: 300; }
.section-subtitle {
  font-size: 15px; color: var(--muted); line-height: 1.6;
  max-width: 560px; margin-bottom: 36px;
}

/* ── KPI CARDS ── */
.kpi-grid {
  display: grid; gap: 16px; margin-bottom: 36px;
}
.kpi-grid-4 { grid-template-columns: repeat(4, 1fr); }
.kpi-grid-3 { grid-template-columns: repeat(3, 1fr); }
.kpi-grid-6 { grid-template-columns: repeat(6, 1fr); }
.kpi-grid-2 { grid-template-columns: repeat(2, 1fr); }

.kpi-card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius); padding: 22px 24px;
  box-shadow: var(--shadow-sm);
  transition: box-shadow 0.2s ease, transform 0.2s ease;
}
.kpi-card:hover {
  box-shadow: var(--shadow); transform: translateY(-1px);
}
.kpi-card.accent-teal { border-left: 3px solid var(--teal); }
.kpi-card.accent-gold  { border-left: 3px solid var(--gold); }
.kpi-card.accent-navy  { border-left: 3px solid var(--navy); }
.kpi-card.accent-blue  { border-left: 3px solid var(--accent); }
.kpi-label {
  font-size: 11px; font-weight: 600; letter-spacing: 0.1em;
  text-transform: uppercase; color: var(--muted); margin-bottom: 6px;
}
.kpi-value {
  font-family: var(--font-display); font-size: 30px; font-weight: 600;
  color: var(--navy); line-height: 1.1; margin-bottom: 4px;
}
.kpi-sub { font-size: 12px; color: var(--dim); }
.kpi-badge {
  display: inline-block; margin-top: 6px;
  font-size: 11px; font-weight: 600; padding: 2px 8px;
  border-radius: 4px;
}
.badge-up   { background: var(--green-bg); color: var(--green); }
.badge-down { background: var(--red-bg); color: var(--red); }
.badge-neu  { background: var(--surface2); color: var(--muted); }

/* ── CARDS ── */
.card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius); box-shadow: var(--shadow-sm);
}
.card-lg { border-radius: var(--radius-lg); }
.card-header {
  padding: 20px 24px 16px; border-bottom: 1px solid var(--border);
  display: flex; align-items: center; justify-content: space-between;
}
.card-title {
  font-size: 14px; font-weight: 600; color: #0B1F3A;
  letter-spacing: 0.01em;
}
.card-body { padding: 20px 24px; }

/* ── INTERNAL TABS ── */
.tabs-bar {
  display: flex; gap: 2px; padding: 4px;
  background: var(--surface2); border-radius: 10px;
  margin-bottom: 28px; flex-wrap: wrap;
}
.tab-btn {
  padding: 7px 16px; border: none; background: none;
  border-radius: 7px; font-size: 13px; font-weight: 500;
  color: var(--muted); cursor: pointer;
  transition: all 0.15s ease; white-space: nowrap;
  font-family: var(--font-body);
}
.tab-btn:hover { color: var(--navy); }
.tab-btn.active {
  background: var(--surface); color: var(--navy);
  box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}
.tab-panel { display: none; }
.tab-panel.active { display: block; }

/* ── TABLES ── */
.data-table {
  width: 100%; border-collapse: collapse; font-size: 13px;
}
.data-table th {
  background: var(--surface2); padding: 10px 14px;
  text-align: left; font-size: 11px; font-weight: 600;
  letter-spacing: 0.06em; text-transform: uppercase; color: var(--muted);
  border-bottom: 1px solid var(--border);
  position: sticky; top: 0;
}
.data-table td {
  padding: 10px 14px; border-bottom: 1px solid rgba(0,0,0,0.04);
  color: var(--text2); vertical-align: middle;
}
.data-table tr:last-child td { border-bottom: none; }
.data-table tr:hover td { background: var(--surface2); }
.data-table .td-num {
  font-family: var(--font-mono); font-size: 12px; color: var(--navy);
}
.data-table .td-strong { font-weight: 600; color: var(--navy); }
.data-table .td-teal { color: var(--teal); font-weight: 500; }
.data-table .td-gold { color: var(--gold); font-weight: 500; }
.table-wrapper {
  overflow-x: auto; border-radius: var(--radius);
  border: 1px solid var(--border); background: var(--surface);
  box-shadow: var(--shadow-sm);
}

/* ── GRID LAYOUTS ── */
.grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }
.grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; }
.grid-1-2 { display: grid; grid-template-columns: 1fr 2fr; gap: 24px; }
.grid-2-1 { display: grid; grid-template-columns: 2fr 1fr; gap: 24px; }

/* ── CHART CONTAINERS ── */
.chart-card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius); padding: 20px 24px;
  box-shadow: var(--shadow-sm);
}
.chart-title {
  font-size: 13px; font-weight: 600; color: #0B1F3A;
  margin-bottom: 4px;
}
.chart-sub { font-size: 11px; color: var(--dim); margin-bottom: 16px; }
.chart-wrap { position: relative; }
.chart-wrap.h300 { height: 300px; }
.chart-wrap.h280 { height: 280px; }
.chart-wrap.h240 { height: 240px; }
.chart-wrap.h220 { height: 220px; }
.chart-wrap.h180 { height: 180px; }

/* ── HERO SECTION ── */
.hero {
  padding: 20px 0 48px;
}
.hero-eyebrow {
  font-size: 11px; font-weight: 600; letter-spacing: 0.2em;
  text-transform: uppercase; color: var(--accent); margin-bottom: 16px;
  display: flex; align-items: center; gap: 8px;
}
.hero-eyebrow::before {
  content: ''; display: inline-block;
  width: 24px; height: 1px; background: var(--accent);
}
.hero-title {
  font-family: var(--font-display);
  font-size: clamp(36px, 4.5vw, 52px);
  font-weight: 300; line-height: 1.15; color: #0B1F3A;
  margin-bottom: 20px;
}
.hero-title strong { font-weight: 800; }
.hero-title em { font-style: italic; color: var(--accent); font-weight: 300; }
.hero-desc {
  font-size: 16px; color: var(--muted); max-width: 560px;
  line-height: 1.7; margin-bottom: 40px;
}
.hero-pillars {
  display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px;
  margin-bottom: 48px;
}
.pillar-card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius); padding: 24px;
  box-shadow: var(--shadow-sm); cursor: pointer;
  transition: all 0.2s ease;
}
.pillar-card:hover {
  border-color: var(--accent); box-shadow: var(--shadow);
  transform: translateY(-2px);
}
.pillar-icon {
  width: 40px; height: 40px; border-radius: 10px;
  display: flex; align-items: center; justify-content: center;
  margin-bottom: 14px; font-size: 18px;
}
.pillar-icon.teal { background: var(--teal-light); }
.pillar-icon.gold { background: var(--gold-light); }
.pillar-icon.navy { background: rgba(11,31,58,0.06); }
.pillar-title {
  font-size: 15px; font-weight: 600; color: var(--navy);
  margin-bottom: 6px;
}
.pillar-desc { font-size: 13px; color: var(--muted); line-height: 1.5; }
.pillar-link {
  margin-top: 14px; font-size: 12px; font-weight: 500;
  color: var(--accent); display: flex; align-items: center; gap: 4px;
}

/* ── SCORECARD SPECIFIC ── */
.sc-layout {
  display: grid; grid-template-columns: 320px 1fr; gap: 28px;
  align-items: start;
}
.sc-inputs-card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius-lg); padding: 24px;
  box-shadow: var(--shadow); position: sticky; top: 80px;
}
.sc-inputs-title {
  font-size: 13px; font-weight: 600; color: var(--muted);
  letter-spacing: 0.08em; text-transform: uppercase; margin-bottom: 18px;
}
.sc-field { margin-bottom: 14px; }
.sc-label {
  display: block; font-size: 11px; font-weight: 600;
  letter-spacing: 0.06em; text-transform: uppercase;
  color: var(--muted); margin-bottom: 5px;
}
.sc-input, .sc-select {
  width: 100%; padding: 8px 12px; border: 1px solid var(--border);
  border-radius: var(--radius-sm); font-size: 13px; font-family: var(--font-body);
  background: var(--bg); color: var(--text); outline: none;
  transition: border-color 0.15s;
}
.sc-input:focus, .sc-select:focus { border-color: var(--accent); }
.sc-section-label {
  font-size: 10px; font-weight: 700; letter-spacing: 0.12em;
  text-transform: uppercase; color: var(--dim); margin: 16px 0 10px;
  padding-bottom: 6px; border-bottom: 1px solid var(--border);
}
.sc-result-panel {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius-lg); padding: 28px;
  box-shadow: var(--shadow);
}
.score-ring-wrap {
  display: flex; align-items: center; gap: 24px; margin-bottom: 28px;
  padding-bottom: 24px; border-bottom: 1px solid var(--border);
}
.score-ring {
  width: 100px; height: 100px; border-radius: 50%;
  border: 5px solid var(--border); display: flex;
  flex-direction: column; align-items: center; justify-content: center;
  flex-shrink: 0; transition: border-color 0.4s ease;
}
.score-num {
  font-family: var(--font-display); font-size: 34px; font-weight: 700;
  line-height: 1; color: var(--navy); transition: color 0.4s;
}
.score-max { font-size: 11px; color: var(--dim); margin-top: 2px; }
.score-info { flex: 1; }
.score-verdict-badge {
  display: inline-block; padding: 4px 12px; border-radius: 20px;
  font-size: 11px; font-weight: 700; letter-spacing: 0.06em;
  text-transform: uppercase; margin-bottom: 8px;
}
.verdict-strong { background: var(--green-bg); color: var(--green); }
.verdict-watch  { background: var(--amber-bg); color: var(--amber); }
.verdict-pass   { background: var(--red-bg); color: var(--red); }
.score-label { font-size: 16px; font-weight: 600; color: var(--navy); margin-bottom: 4px; }
.score-desc  { font-size: 13px; color: var(--muted); line-height: 1.5; }

/* ── METRIC ROWS (scorecard) ── */
.metrics-header {
  display: grid; grid-template-columns: 1fr 100px 100px 100px 90px;
  gap: 8px; padding: 8px 12px;
  font-size: 10px; font-weight: 700; letter-spacing: 0.08em;
  text-transform: uppercase; color: var(--dim);
  border-bottom: 1px solid var(--border); margin-bottom: 4px;
}
.metric-row {
  display: grid; grid-template-columns: 1fr 100px 100px 100px 90px;
  gap: 8px; padding: 10px 12px; border-radius: 6px;
  transition: background 0.15s;
}
.metric-row:hover { background: var(--surface2); }
.metric-name { font-size: 13px; font-weight: 500; color: var(--navy); }
.metric-sub  { font-size: 11px; color: var(--dim); margin-top: 1px; }
.metric-val  { font-family: var(--font-mono); font-size: 12px; line-height: 1.4; }
.status {
  display: inline-block; padding: 2px 8px; border-radius: 4px;
  font-size: 10px; font-weight: 700; letter-spacing: 0.04em;
}
.status-green { background: var(--green-bg); color: var(--green); }
.status-yellow { background: var(--amber-bg); color: var(--amber); }
.status-red { background: var(--red-bg); color: var(--red); }

/* ── DIM SCORE PILL ── */
.dim-pill-row {
  display: flex; align-items: center; justify-content: space-between;
  margin-bottom: 14px;
}
.dim-pill-title { font-size: 14px; font-weight: 600; color: var(--navy); }
.dim-score-pill {
  font-family: var(--font-mono); font-size: 12px; font-weight: 500;
  padding: 4px 12px; border-radius: 20px;
  background: var(--surface2); border: 1px solid var(--border);
  color: var(--muted);
}

/* ── STUDIO METRICS ── */
.studio-section {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius-lg); margin-bottom: 24px;
  box-shadow: var(--shadow-sm); overflow: hidden;
}
.studio-section-header {
  padding: 16px 24px;
  background: linear-gradient(135deg, var(--navy) 0%, var(--navy-mid) 100%);
  display: flex; align-items: center; gap: 12px;
}
.studio-section-icon {
  width: 28px; height: 28px; border-radius: 6px;
  background: rgba(255,255,255,0.12);
  display: flex; align-items: center; justify-content: center;
  font-size: 14px;
}
.studio-section-title {
  font-size: 13px; font-weight: 600; color: white;
  letter-spacing: 0.03em;
}
.studio-section-sub { font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 1px; }

/* ── METHODOLOGY ── */
.source-card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius); padding: 20px 24px;
  box-shadow: var(--shadow-sm);
}
.source-num {
  width: 28px; height: 28px; border-radius: 50%;
  background: var(--navy); color: white;
  display: flex; align-items: center; justify-content: center;
  font-size: 12px; font-weight: 700; flex-shrink: 0;
}
.source-name { font-size: 14px; font-weight: 600; color: var(--navy); }
.source-desc { font-size: 13px; color: var(--muted); margin-top: 4px; }

/* ── INSIGHT TAGS ── */
.insight-tag {
  display: inline-block; padding: 3px 10px; border-radius: 4px;
  font-size: 11px; font-weight: 600; margin: 2px;
}
.tag-teal { background: var(--teal-light); color: var(--teal); }
.tag-gold { background: var(--gold-light); color: var(--gold); }
.tag-navy { background: rgba(11,31,58,0.07); color: var(--navy-mid); }

/* ── NOTABLE DEAL CARD ── */
.deal-card {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--radius); padding: 16px 20px;
  box-shadow: var(--shadow-sm);
  display: flex; flex-direction: column; gap: 8px;
}
.deal-company { font-size: 16px; font-weight: 700; color: var(--navy); }
.deal-meta { font-size: 12px; color: var(--muted); }
.deal-amount {
  font-family: var(--font-display); font-size: 22px; font-weight: 600;
  color: var(--teal);
}
.deal-valuation { font-size: 12px; color: var(--muted); }

/* ── COMPARISON BAR ── */
.compare-bar-row { display: flex; align-items: center; gap: 12px; margin-bottom: 14px; }
.compare-bar-label { font-size: 12px; color: var(--text2); min-width: 90px; }
.compare-bar-track { flex: 1; background: var(--surface2); border-radius: 4px; height: 8px; }
.compare-bar-fill { height: 8px; border-radius: 4px; transition: width 0.6s ease; }
.compare-bar-val { font-family: var(--font-mono); font-size: 11px; color: var(--navy); min-width: 52px; text-align: right; }

/* ── MOBILE ── */
@media (max-width: 900px) {
  #navbar { padding: 0 16px; }
  .nav-links .nav-link { padding: 6px 10px; font-size: 12px; }
  .section { padding: 80px 16px 40px; }
  .kpi-grid-4, .kpi-grid-6 { grid-template-columns: repeat(2, 1fr); }
  .kpi-grid-3 { grid-template-columns: repeat(2, 1fr); }
  .grid-2, .grid-3, .grid-1-2, .grid-2-1 { grid-template-columns: 1fr; }
  .hero-pillars { grid-template-columns: 1fr; }
  .sc-layout { grid-template-columns: 1fr; }
  .sc-inputs-card { position: static; }
  .hero-title { font-size: 34px; }
}
"""

# ─── NAVBAR HTML ─────────────────────────────────────────────────────────────

def build_navbar():
    return """
<nav id="navbar">
  <a class="nav-brand" href="#" onclick="showSection('overview')">
    <img src="logo.png" alt="AIDA Ventures" style="height:36px; width:auto; display:block;">
  </a>
  <div class="nav-links">
    <button class="nav-link active" onclick="showSection('overview')">Overview</button>
    <button class="nav-link" onclick="showSection('market')">Market Intelligence</button>
    <button class="nav-link" onclick="showSection('vc')">VC Benchmarks</button>
    <button class="nav-link" onclick="showSection('scorecard')">Deal Scorecard</button>
    <button class="nav-link" onclick="showSection('studio')">Venture Studio</button>
    <button class="nav-link" onclick="showSection('methodology')">Methodology</button>
  </div>
  <div class="nav-right">Intelligence Platform v2.0</div>
</nav>
"""


# ─── OVERVIEW SECTION ────────────────────────────────────────────────────────

def build_overview(data):
    return """
<section id="overview" class="section active">
  <div class="hero">
    <div class="hero-eyebrow">AIDA Ventures — Intelligence Platform</div>
    <h1 class="hero-title">
      <strong>Investment Intelligence</strong><br>
      for <em>Early-Stage</em> Latin America
    </h1>
    <p class="hero-desc">
      A unified data platform combining fintech market intelligence, VC fund benchmarks,
      deal evaluation tools, and venture studio metrics — built for AIDA Ventures' hybrid
      VC + Studio model across LatAm.
    </p>
  </div>

  <div class="kpi-grid kpi-grid-6" style="margin-bottom:40px">
    <div class="kpi-card accent-teal">
      <div class="kpi-label">LATAM Fintech 2024</div>
      <div class="kpi-value">$3.2B</div>
      <div class="kpi-sub">Total investment</div>
      <span class="kpi-badge badge-up">+5–8% YoY</span>
    </div>
    <div class="kpi-card accent-gold">
      <div class="kpi-label">US Fintech 2024</div>
      <div class="kpi-value">$50.7B</div>
      <div class="kpi-sub">Total investment</div>
      <span class="kpi-badge badge-down">−30% YoY</span>
    </div>
    <div class="kpi-card accent-navy">
      <div class="kpi-label">US / LATAM Ratio</div>
      <div class="kpi-value">14–17×</div>
      <div class="kpi-sub">Investment gap</div>
      <span class="kpi-badge badge-neu">Converging</span>
    </div>
    <div class="kpi-card accent-blue">
      <div class="kpi-label">IRR Target Early</div>
      <div class="kpi-value">30–40%</div>
      <div class="kpi-sub">Pre-Seed / Seed</div>
      <span class="kpi-badge badge-neu">Top decil: ~28%</span>
    </div>
    <div class="kpi-card accent-teal">
      <div class="kpi-label">TVPI Benchmark</div>
      <div class="kpi-value">3.0×+</div>
      <div class="kpi-sub">Top quartile target</div>
      <span class="kpi-badge badge-neu">Median: 1.7–1.9×</span>
    </div>
    <div class="kpi-card accent-gold">
      <div class="kpi-label">Seed → Series A</div>
      <div class="kpi-value">2–5%</div>
      <div class="kpi-sub">Graduation rate (US 2024)</div>
      <span class="kpi-badge badge-down">vs 15% in 2021</span>
    </div>
  </div>

  <div class="hero-pillars">
    <div class="pillar-card" onclick="showSection('market')">
      <div class="pillar-icon teal">🗺️</div>
      <div class="pillar-title">Where to Invest</div>
      <p class="pillar-desc">
        Fintech sector mapping across 9 LATAM countries and 14 subsectors.
        Investment flows, deal activity, unicorn density and growth trajectories.
      </p>
      <div class="pillar-link">Market Intelligence →</div>
    </div>
    <div class="pillar-card" onclick="showSection('vc')">
      <div class="pillar-icon gold">📊</div>
      <div class="pillar-title">What Returns to Expect</div>
      <p class="pillar-desc">
        VC fund benchmarks for IRR, TVPI, DPI and MOIC across Pre-Seed,
        Seed and Series A. LATAM vs US comparisons.
      </p>
      <div class="pillar-link">VC Benchmarks →</div>
    </div>
    <div class="pillar-card" onclick="showSection('scorecard')">
      <div class="pillar-icon navy">🎯</div>
      <div class="pillar-title">How to Evaluate Deals</div>
      <p class="pillar-desc">
        Interactive scorecard evaluating startups across 5 dimensions:
        growth, capital efficiency, retention, valuation and LATAM context.
      </p>
      <div class="pillar-link">Deal Scorecard →</div>
    </div>
  </div>

  <div class="grid-2">
    <div class="chart-card">
      <div class="chart-title">LATAM Fintech Investment by Country</div>
      <div class="chart-sub">2024 estimated investment (USD midpoint)</div>
      <div class="chart-wrap h300">
        <canvas id="ov-country-chart"></canvas>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Stage Graduation Funnel — 1,000 Pre-Seed Startups</div>
      <div class="chart-sub">Expected survivors at each stage (US vs LATAM)</div>
      <div class="chart-wrap h300">
        <canvas id="ov-funnel-chart"></canvas>
      </div>
    </div>
  </div>
</section>
"""


# ─── MARKET INTELLIGENCE SECTION ─────────────────────────────────────────────

def build_market(data):
    ft = data['fintech']

    # Country table rows
    country_rows = ""
    for c in ft['countries']:
        growth = c['growth']
        g_class = "td-teal" if '+' in growth else ("td-gold" if '-' in growth else "")
        country_rows += f"""
        <tr>
          <td class="td-strong">{c['country']}</td>
          <td class="td-num">{c['investment']}</td>
          <td>{c['share']}</td>
          <td>{c['fintechs']}</td>
          <td>{c['deals']}</td>
          <td>{c['unicorns']}</td>
          <td class="{g_class}">{c['growth']}</td>
        </tr>"""

    # LATAM subsector rows
    latam_sub_rows = ""
    for s in ft['latam_subsectors']:
        latam_sub_rows += f"""
        <tr>
          <td class="td-strong">{s['subsector']}</td>
          <td class="td-num">{s['investment']}</td>
          <td>{s['pct']}</td>
          <td>{s['startups']}</td>
          <td class="td-teal">{s['leaders']}</td>
        </tr>"""

    # USA subsector rows
    usa_sub_rows = ""
    for s in ft['usa_subsectors']:
        trend = s['trend']
        t_cls = "td-teal" if '↑' in trend else ("td-gold" if '↓' in trend else "")
        usa_sub_rows += f"""
        <tr>
          <td class="td-strong">{s['subsector']}</td>
          <td class="td-num">{s['investment']}</td>
          <td>{s['startups']}</td>
          <td class="{t_cls}">{trend}</td>
        </tr>"""

    # Comparison rows
    cmp_rows = ""
    for c in ft['comparison']:
        cmp_rows += f"""
        <tr>
          <td class="td-strong">{c['metric']}</td>
          <td class="td-gold td-num">{c['usa']}</td>
          <td class="td-teal td-num">{c['latam']}</td>
          <td class="td-num">{c['ratio']}</td>
        </tr>"""

    return f"""
<section id="market" class="section">
  <div class="section-eyebrow">Market Intelligence</div>
  <h2 class="section-title">Fintech <em>Sector Analysis</em></h2>
  <p class="section-subtitle">
    Investment flows, subsector dynamics and competitive positioning across LATAM and the US fintech ecosystem.
  </p>

  <div class="tabs-bar" id="market-tabs">
    <button class="tab-btn active" onclick="mktTab('overview',this)">Overview</button>
    <button class="tab-btn" onclick="mktTab('countries',this)">LATAM Countries</button>
    <button class="tab-btn" onclick="mktTab('latam-sub',this)">LATAM Subsectors</button>
    <button class="tab-btn" onclick="mktTab('usa',this)">United States</button>
    <button class="tab-btn" onclick="mktTab('compare',this)">LATAM vs USA</button>
    <button class="tab-btn" onclick="mktTab('opps',this)">GP Opportunities</button>
  </div>

  <!-- OVERVIEW TAB -->
  <div class="tab-panel active" id="mkt-overview">
    <div class="kpi-grid kpi-grid-4" style="margin-bottom:28px">
      <div class="kpi-card accent-teal">
        <div class="kpi-label">LATAM Total 2024</div>
        <div class="kpi-value">$3.0–3.5B</div>
        <div class="kpi-sub">Fintech investment</div>
        <span class="kpi-badge badge-up">+5–8% YoY</span>
      </div>
      <div class="kpi-card accent-gold">
        <div class="kpi-label">USA Total 2024</div>
        <div class="kpi-value">$50.7B</div>
        <div class="kpi-sub">Fintech investment</div>
        <span class="kpi-badge badge-down">−30.3% YoY</span>
      </div>
      <div class="kpi-card accent-navy">
        <div class="kpi-label">LATAM Active Fintechs</div>
        <div class="kpi-value">3,620+</div>
        <div class="kpi-sub">Across 9 markets</div>
        <span class="kpi-badge badge-neu">~370–450 deals</span>
      </div>
      <div class="kpi-card accent-blue">
        <div class="kpi-label">LATAM Unicorns</div>
        <div class="kpi-value">15–20</div>
        <div class="kpi-sub">Fintech unicorns</div>
        <span class="kpi-badge badge-neu">vs ~165 in US</span>
      </div>
    </div>
    <div class="grid-2">
      <div class="chart-card">
        <div class="chart-title">Investment by Country — Top 6 LATAM Markets</div>
        <div class="chart-sub">Estimated midpoint 2024 (USD millions)</div>
        <div class="chart-wrap h280"><canvas id="mkt-country-bar"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Top LATAM Subsectors by Investment</div>
        <div class="chart-sub">2024 estimated investment (USD)</div>
        <div class="chart-wrap h280"><canvas id="mkt-sub-donut"></canvas></div>
      </div>
    </div>
  </div>

  <!-- COUNTRIES TAB -->
  <div class="tab-panel" id="mkt-countries">
    <div class="grid-1-2" style="align-items:start">
      <div class="chart-card">
        <div class="chart-title">Regional Share Distribution</div>
        <div class="chart-sub">% of total LATAM fintech investment</div>
        <div class="chart-wrap h280"><canvas id="mkt-share-donut"></canvas></div>
      </div>
      <div class="table-wrapper">
        <table class="data-table">
          <thead><tr>
            <th>Country</th><th>Investment</th><th>Share</th>
            <th>Fintechs</th><th>Deals</th><th>Unicorns</th><th>Growth YoY</th>
          </tr></thead>
          <tbody>{country_rows}</tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- LATAM SUBSECTORS TAB -->
  <div class="tab-panel" id="mkt-latam-sub">
    <div class="grid-1-2" style="align-items:start; margin-bottom:24px">
      <div class="chart-card">
        <div class="chart-title">LATAM Investment by Subsector</div>
        <div class="chart-sub">Top 10 subsectors (USD midpoint)</div>
        <div class="chart-wrap h300"><canvas id="mkt-latam-sub-bar"></canvas></div>
      </div>
      <div class="table-wrapper">
        <table class="data-table">
          <thead><tr>
            <th>Subsector</th><th>Investment</th><th>% Total</th>
            <th>Startups</th><th>Leading Markets</th>
          </tr></thead>
          <tbody>{latam_sub_rows}</tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- USA TAB -->
  <div class="tab-panel" id="mkt-usa">
    <div class="kpi-grid kpi-grid-3" style="margin-bottom:24px">
      <div class="kpi-card accent-gold">
        <div class="kpi-label">Payments (Largest)</div>
        <div class="kpi-value">$31.0B</div>
        <div class="kpi-sub">↑ B2B, cross-border</div>
      </div>
      <div class="kpi-card accent-gold">
        <div class="kpi-label">Crypto / Digital Assets</div>
        <div class="kpi-value">$8.5B</div>
        <div class="kpi-sub">↑ Institutional adoption</div>
      </div>
      <div class="kpi-card accent-gold">
        <div class="kpi-label">AI-Powered Finance</div>
        <div class="kpi-value">$4.7B</div>
        <div class="kpi-sub">↑↑ Explosive growth</div>
      </div>
    </div>
    <div class="grid-1-2" style="align-items:start">
      <div class="chart-card">
        <div class="chart-title">US Fintech — Top Subsectors</div>
        <div class="chart-sub">2024 investment (USD billions)</div>
        <div class="chart-wrap h300"><canvas id="mkt-usa-bar"></canvas></div>
      </div>
      <div class="table-wrapper">
        <table class="data-table">
          <thead><tr><th>Subsector</th><th>Investment</th><th>Startups</th><th>2025 Trend</th></tr></thead>
          <tbody>{usa_sub_rows}</tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- COMPARISON TAB -->
  <div class="tab-panel" id="mkt-compare">
    <div class="grid-2">
      <div>
        <div class="table-wrapper" style="margin-bottom:24px">
          <table class="data-table">
            <thead><tr><th>Metric</th><th class="td-gold">USA</th><th class="td-teal">LATAM</th><th>Ratio</th></tr></thead>
            <tbody>{cmp_rows}</tbody>
          </table>
        </div>
        <div class="card" style="padding:20px">
          <div class="card-title" style="margin-bottom:14px">Key Takeaways</div>
          <div style="font-size:13px; color:var(--text2); line-height:1.7">
            <p style="margin-bottom:8px">
              📍 <strong>14–17× investment gap</strong> — but LATAM is growing while US contracted 30% in 2024.
            </p>
            <p style="margin-bottom:8px">
              💡 <strong>AI adoption higher in LATAM</strong> (25–30% of fintechs vs 15% in US), suggesting faster iteration cycles.
            </p>
            <p style="margin-bottom:8px">
              🎯 <strong>Payments dominates both markets</strong> (28–32% LATAM / 32% US), validating the sector thesis.
            </p>
            <p>
              📈 <strong>Colombia grew +36.3% YoY</strong> — strongest momentum in the region.
            </p>
          </div>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Investment Gap — LATAM vs USA</div>
        <div class="chart-sub">Selected metrics normalized for comparison</div>
        <div class="chart-wrap h300"><canvas id="mkt-compare-radar"></canvas></div>
      </div>
    </div>
  </div>

  <!-- GP OPPORTUNITIES TAB -->
  <div class="tab-panel" id="mkt-opps">
    <div class="kpi-grid kpi-grid-3" style="margin-bottom:24px">
      <div class="kpi-card accent-teal">
        <div class="kpi-label">🥇 Priority 1 — Payments</div>
        <div class="kpi-value">$800M–1B</div>
        <div class="kpi-sub">LATAM 2024 · 28–32% share</div>
        <div style="margin-top:10px; font-size:12px; color:var(--muted)">
          Cross-border corridors Mexico-US, B2B rails, stablecoin settlement
        </div>
      </div>
      <div class="kpi-card accent-teal">
        <div class="kpi-label">🥈 Priority 2 — Open Banking / BaaS</div>
        <div class="kpi-value">$250–350M</div>
        <div class="kpi-sub">LATAM 2024 · Regulatory tailwinds</div>
        <div style="margin-top:10px; font-size:12px; color:var(--muted)">
          Brazil PIX expansion, Mexico Open Finance, Colombia BaaS licensing
        </div>
      </div>
      <div class="kpi-card accent-teal">
        <div class="kpi-label">🥉 Priority 3 — B2B Finance</div>
        <div class="kpi-value">$300–400M</div>
        <div class="kpi-sub">LATAM 2024 · High growth</div>
        <div style="margin-top:10px; font-size:12px; color:var(--muted)">
          SMB treasury, embedded payroll, working capital for e-commerce
        </div>
      </div>
    </div>
    <div class="grid-2">
      <div class="chart-card">
        <div class="chart-title">Top Opportunity Subsectors</div>
        <div class="chart-sub">LATAM investment size vs growth potential</div>
        <div class="chart-wrap h280"><canvas id="mkt-opp-bubble"></canvas></div>
      </div>
      <div class="card" style="padding:24px">
        <div class="card-title" style="margin-bottom:16px">Why Colombia + Mexico Now</div>
        <div style="font-size:13px; color:var(--text2); line-height:1.75">
          <p style="margin-bottom:10px">
            <span class="insight-tag tag-teal">Colombia +36.3%</span>
            Fastest-growing market in 2024. Sandbox regulation, growing fintech
            license approvals, and proximity to US investment appetite.
          </p>
          <p style="margin-bottom:10px">
            <span class="insight-tag tag-gold">Mexico: Plata $1.5B</span>
            Neobank ecosystem maturing rapidly. Series A–B valuations still at LATAM
            discount (35–45%) despite US-quality metrics.
          </p>
          <p style="margin-bottom:10px">
            <span class="insight-tag tag-teal">AI-First Fintechs</span>
            25–30% of LATAM fintechs are AI-enabled — higher than US (15%).
            Early-mover advantage for AI credit scoring, fraud detection.
          </p>
          <p>
            <span class="insight-tag tag-navy">Valuation Entry Point</span>
            LATAM Pre-Seed at 55–60% discount vs US. Series A at 35–45% discount.
            Rare asymmetric entry for GP with global network.
          </p>
        </div>
      </div>
    </div>
  </div>
</section>
"""


# ─── VC BENCHMARKS SECTION ───────────────────────────────────────────────────

def build_vc(data):
    vc = data['vcfunds']
    bench = data['benchmarks']

    # Early Stage Fund rows
    early_rows = ""
    for r in vc['early_stage'].get('rows', []):
        while len(r) < 4:
            r.append('')
        early_rows += f"""<tr>
          <td class="td-strong">{r[0]}</td>
          <td class="td-num td-teal">{r[1]}</td>
          <td class="td-num td-teal">{r[2]}</td>
          <td class="td-num td-gold">{r[3]}</td>
        </tr>"""

    # LATAM vs US rows
    latam_us_rows = ""
    for r in vc['latam_vs_us'].get('rows', []):
        while len(r) < 3:
            r.append('')
        latam_us_rows += f"""<tr>
          <td class="td-strong">{r[0]}</td>
          <td class="td-num td-gold">{r[1]}</td>
          <td class="td-num td-teal">{r[2]}</td>
        </tr>"""

    # Time between rounds rows
    time_rows = ""
    for t in bench['time_between_rounds']:
        time_rows += f"""<tr>
          <td class="td-strong">{t['transition']}</td>
          <td class="td-num td-gold">{t['us_2021']}</td>
          <td class="td-num td-gold">{t['us_2024']}</td>
          <td class="td-num" style="color:var(--red)">{t['change']}</td>
          <td class="td-num td-teal">{t['latam_median']}</td>
          <td class="td-num td-teal">{t['latam_top']}</td>
        </tr>"""

    # Graduation rates rows
    grad_rows = ""
    for g in bench['graduation']['geography']:
        grad_rows += f"""<tr>
          <td class="td-strong">{g['transition']}</td>
          <td class="td-num td-gold">{g['us_2024']}</td>
          <td class="td-num" style="color:var(--dim)">{g['us_2021']}</td>
          <td class="td-num td-teal">{g['latam']}</td>
          <td class="td-num td-teal">{g['latam_top']}</td>
        </tr>"""

    # Valuation rows - US SaaS
    us_saas_rows = ""
    for v in bench['us_valuations']['saas']:
        us_saas_rows += f"""<tr>
          <td class="td-strong">{v['stage']}</td>
          <td class="td-num td-gold">{v['val']}</td>
          <td class="td-num">{v['multiple']}</td>
          <td style="font-size:12px; color:var(--muted)">{v['notes']}</td>
          <td class="td-num">{v['round']}</td>
        </tr>"""

    # LATAM Discount rows
    discount_rows = ""
    for d in bench['latam_discount']:
        discount_rows += f"""<tr>
          <td class="td-strong">{d['factor']}</td>
          <td class="td-num" style="color:var(--red)">{d['discount']}</td>
          <td style="font-size:12px; color:var(--muted)">{d['source']}</td>
          <td style="font-size:12px; color:var(--muted)">{d['notes']}</td>
        </tr>"""

    # Notable deals
    notable_cards = ""
    for d in bench['latam_notable'][:6]:
        val_str = f"Val: {d['valuation']}" if d['valuation'] else ""
        notable_cards += f"""
        <div class="deal-card">
          <div>
            <div class="deal-company">{d['company']}</div>
            <div class="deal-meta">{d['country']} · {d['stage']} · {d['date']}</div>
          </div>
          <div class="deal-amount">{d['amount']}</div>
          <div class="deal-valuation">{val_str}</div>
          <div><span class="insight-tag tag-teal" style="font-size:10px">{d['sector']}</span></div>
        </div>"""

    return f"""
<section id="vc" class="section">
  <div class="section-eyebrow">Fund Benchmarks</div>
  <h2 class="section-title">VC <em>Performance Intelligence</em></h2>
  <p class="section-subtitle">
    IRR, TVPI, DPI and return benchmarks across early-stage funds. Graduation rates,
    timing between rounds, and valuation context for LatAm vs US markets.
  </p>

  <div class="tabs-bar" id="vc-tabs">
    <button class="tab-btn active" onclick="vcTab('summary',this)">Fund Summary</button>
    <button class="tab-btn" onclick="vcTab('latam-us',this)">LATAM vs US</button>
    <button class="tab-btn" onclick="vcTab('valuations',this)">Valuations</button>
    <button class="tab-btn" onclick="vcTab('timing',this)">Timing & Rounds</button>
    <button class="tab-btn" onclick="vcTab('graduation',this)">Graduation Rates</button>
    <button class="tab-btn" onclick="vcTab('notable',this)">Notable Deals</button>
  </div>

  <!-- SUMMARY TAB -->
  <div class="tab-panel active" id="vc-summary">
    <div class="kpi-grid kpi-grid-4" style="margin-bottom:28px">
      <div class="kpi-card accent-navy">
        <div class="kpi-label">IRR Target Pre-Seed</div>
        <div class="kpi-value">30–40%</div>
        <div class="kpi-sub">Observed median: ~8–12%</div>
      </div>
      <div class="kpi-card accent-navy">
        <div class="kpi-label">TVPI Top Quartile</div>
        <div class="kpi-value">3.0×+</div>
        <div class="kpi-sub">Median: 1.72–1.95×</div>
      </div>
      <div class="kpi-card accent-navy">
        <div class="kpi-label">MOIC Target (deal)</div>
        <div class="kpi-value">5–10×</div>
        <div class="kpi-sub">Pre-Seed / Seed level</div>
      </div>
      <div class="kpi-card accent-navy">
        <div class="kpi-label">Capital Deployed (24m)</div>
        <div class="kpi-value">~43%</div>
        <div class="kpi-sub">Vintage 2022 benchmark</div>
      </div>
    </div>
    <div class="grid-2">
      <div class="table-wrapper">
        <table class="data-table">
          <thead><tr><th>Metric</th><th>Pre-Seed</th><th>Seed</th><th>Series A</th></tr></thead>
          <tbody>{early_rows}</tbody>
        </table>
      </div>
      <div class="chart-card">
        <div class="chart-title">IRR Target vs Observed Median</div>
        <div class="chart-sub">By fund stage — early-stage benchmarks</div>
        <div class="chart-wrap h280"><canvas id="vc-irr-chart"></canvas></div>
      </div>
    </div>
  </div>

  <!-- LATAM vs US TAB -->
  <div class="tab-panel" id="vc-latam-us">
    <div class="grid-2">
      <div class="table-wrapper">
        <table class="data-table">
          <thead><tr><th>Metric</th><th>🇺🇸 United States</th><th>🌎 LATAM</th></tr></thead>
          <tbody>{latam_us_rows}</tbody>
        </table>
      </div>
      <div class="chart-card">
        <div class="chart-title">LATAM vs US — Key VC Metrics</div>
        <div class="chart-sub">Comparative benchmarks by market</div>
        <div class="chart-wrap h280"><canvas id="vc-compare-chart"></canvas></div>
      </div>
    </div>
    <div style="margin-top:24px" class="card" style="padding:20px">
      <div style="padding:20px">
        <div class="card-title" style="margin-bottom:12px">Fund Size Comparison</div>
        <div style="display:grid; grid-template-columns: repeat(3,1fr); gap:16px; font-size:13px">
          <div>
            <div style="font-weight:600; margin-bottom:6px; color:var(--navy)">Pre-Seed</div>
            <div style="color:var(--gold)">🇺🇸 US: $1M–$25M</div>
            <div style="color:var(--teal)">🌎 LATAM: ~$10M</div>
            <div style="margin-top:6px; color:var(--muted)">Ticket: $500K–$2M (US) / $50K–$500K (LATAM)</div>
          </div>
          <div>
            <div style="font-weight:600; margin-bottom:6px; color:var(--navy)">Seed</div>
            <div style="color:var(--gold)">🇺🇸 US: $10M–$75M</div>
            <div style="color:var(--teal)">🌎 LATAM: ~$15M</div>
            <div style="margin-top:6px; color:var(--muted)">Ticket: $3–4M (US) / $100K–$500K (LATAM)</div>
          </div>
          <div>
            <div style="font-weight:600; margin-bottom:6px; color:var(--navy)">Series A</div>
            <div style="color:var(--gold)">🇺🇸 US: $50M–$150M</div>
            <div style="color:var(--teal)">🌎 LATAM: ~$35M</div>
            <div style="margin-top:6px; color:var(--muted)">Ticket: $10–15M (US) / $250K–$1.5M (LATAM)</div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- VALUATIONS TAB -->
  <div class="tab-panel" id="vc-valuations">
    <div class="grid-2" style="margin-bottom:24px">
      <div>
        <div style="font-size:13px; font-weight:600; color:var(--navy); margin-bottom:12px">
          US Valuations — Enterprise SaaS (2025)
        </div>
        <div class="table-wrapper">
          <table class="data-table">
            <thead><tr><th>Stage</th><th>Median Val.</th><th>ARR Multiple</th><th>Notes</th><th>Round Size</th></tr></thead>
            <tbody>{us_saas_rows}</tbody>
          </table>
        </div>
      </div>
      <div>
        <div style="font-size:13px; font-weight:600; color:var(--navy); margin-bottom:12px">
          LATAM Valuation Discount vs US
        </div>
        <div class="table-wrapper">
          <table class="data-table">
            <thead><tr><th>Factor</th><th>Discount</th><th>Source</th><th>Notes</th></tr></thead>
            <tbody>{discount_rows}</tbody>
          </table>
        </div>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">LATAM Valuations vs US — By Stage & Sector</div>
      <div class="chart-sub">Estimated LATAM pre-money vs US reference (USD millions)</div>
      <div class="chart-wrap h240"><canvas id="vc-val-chart"></canvas></div>
    </div>
  </div>

  <!-- TIMING TAB -->
  <div class="tab-panel" id="vc-timing">
    <div class="kpi-grid kpi-grid-4" style="margin-bottom:24px">
      <div class="kpi-card accent-blue">
        <div class="kpi-label">Seed → Series A (US 2024)</div>
        <div class="kpi-value">~26 mo</div>
        <div class="kpi-sub">774 days median</div>
        <span class="kpi-badge badge-down">+84% vs 2021</span>
      </div>
      <div class="kpi-card accent-blue">
        <div class="kpi-label">Series A → B (US 2024)</div>
        <div class="kpi-value">~24 mo</div>
        <div class="kpi-sub">732 days median</div>
        <span class="kpi-badge badge-down">+97% vs 2021</span>
      </div>
      <div class="kpi-card accent-teal">
        <div class="kpi-label">Seed → A LATAM Est.</div>
        <div class="kpi-value">30–36 mo</div>
        <div class="kpi-sub">Median estimate</div>
        <span class="kpi-badge badge-neu">Top Q: 20–24m</span>
      </div>
      <div class="kpi-card accent-teal">
        <div class="kpi-label">Fintech Seed → A (US)</div>
        <div class="kpi-value">~32 mo</div>
        <div class="kpi-sub">971 days — longest sector</div>
        <span class="kpi-badge badge-down">High regulation</span>
      </div>
    </div>
    <div class="table-wrapper">
      <table class="data-table">
        <thead><tr>
          <th>Stage Transition</th>
          <th>US Median 2021</th><th>US Median 2024</th>
          <th>Change</th><th>LATAM Median</th><th>LATAM Top Q</th>
        </tr></thead>
        <tbody>{time_rows}</tbody>
      </table>
    </div>
    <div class="chart-card" style="margin-top:24px">
      <div class="chart-title">Time Between Rounds — Evolution (US)</div>
      <div class="chart-sub">2021 peak vs 2024 current, months</div>
      <div class="chart-wrap h240"><canvas id="vc-timing-chart"></canvas></div>
    </div>
  </div>

  <!-- GRADUATION TAB -->
  <div class="tab-panel" id="vc-graduation">
    <div class="grid-2">
      <div>
        <div class="table-wrapper" style="margin-bottom:20px">
          <table class="data-table">
            <thead><tr>
              <th>Stage Transition</th>
              <th>US 2024</th><th>US 2021 Peak</th>
              <th>LATAM Est.</th><th>LATAM Top Q</th>
            </tr></thead>
            <tbody>{grad_rows}</tbody>
          </table>
        </div>
        <div class="card" style="padding:18px 20px">
          <div style="font-size:13px; color:var(--text2); line-height:1.75">
            <p style="margin-bottom:8px">
              <strong>Series A Crunch:</strong> Only 2–5% of seed companies reach Series A in 2024
              (vs 14–15% in 2021). The market is dramatically more selective.
            </p>
            <p>
              <strong>LATAM fintech outperforms:</strong> 6–12% graduation rate Seed→A vs 3–7% US average,
              driven by unbanked population TAM and financial inclusion momentum.
            </p>
          </div>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Cohort Survival — 1,000 Pre-Seed Startups</div>
        <div class="chart-sub">US vs LATAM survivors at each stage</div>
        <div class="chart-wrap h300"><canvas id="vc-grad-chart"></canvas></div>
      </div>
    </div>
  </div>

  <!-- NOTABLE DEALS TAB -->
  <div class="tab-panel" id="vc-notable">
    <div style="margin-bottom:20px; font-size:13px; color:var(--muted)">
      Reference points for LATAM valuations at scale — 2025 notable deals.
    </div>
    <div style="display:grid; grid-template-columns: repeat(3,1fr); gap:16px">
      {notable_cards}
    </div>
  </div>
</section>
"""


# ─── DEAL SCORECARD SECTION ───────────────────────────────────────────────────

def build_scorecard():
    return """
<section id="scorecard" class="section">
  <div class="section-eyebrow">Deal Evaluation</div>
  <h2 class="section-title">Startup <em>Scorecard</em></h2>
  <p class="section-subtitle">
    Evaluate any startup against LATAM benchmark data. Enter metrics below — the model
    scores the deal from 0 to 100 across 5 weighted dimensions.
  </p>

  <div class="sc-layout">
    <!-- INPUT PANEL -->
    <div class="sc-inputs-card">
      <div class="sc-inputs-title">Deal Parameters</div>

      <div class="sc-section-label">Company Context</div>
      <div class="sc-field">
        <label class="sc-label">Sector</label>
        <select class="sc-select" id="inp-sector" onchange="scRecalc()">
          <option value="saas">SaaS / Enterprise</option>
          <option value="fintech" selected>Fintech</option>
          <option value="logtech">LogTech / Supply Chain</option>
          <option value="consumer">Consumer / B2C</option>
          <option value="other">Other</option>
        </select>
      </div>
      <div class="sc-field">
        <label class="sc-label">Stage</label>
        <select class="sc-select" id="inp-stage" onchange="scRecalc()">
          <option value="preseed">Pre-Seed</option>
          <option value="seed" selected>Seed</option>
          <option value="seriea">Series A</option>
        </select>
      </div>
      <div class="sc-field">
        <label class="sc-label">Country</label>
        <select class="sc-select" id="inp-country" onchange="scRecalc()">
          <option value="brazil">Brazil</option>
          <option value="mexico" selected>Mexico</option>
          <option value="colombia">Colombia</option>
          <option value="argentina">Argentina</option>
          <option value="chile">Chile</option>
          <option value="peru">Peru</option>
          <option value="other">Other LATAM</option>
        </select>
      </div>

      <div class="sc-section-label">Growth Metrics</div>
      <div class="sc-field">
        <label class="sc-label">ARR (USD)</label>
        <input class="sc-input" type="number" id="inp-arr" value="350000" onchange="scRecalc()">
      </div>
      <div class="sc-field">
        <label class="sc-label">MoM Growth (%)</label>
        <input class="sc-input" type="number" id="inp-mom" value="18" onchange="scRecalc()">
      </div>

      <div class="sc-section-label">Efficiency & Unit Economics</div>
      <div class="sc-field">
        <label class="sc-label">Burn Multiple (×)</label>
        <input class="sc-input" type="number" id="inp-burn" value="1.8" step="0.1" onchange="scRecalc()">
      </div>
      <div class="sc-field">
        <label class="sc-label">Gross Margin (%)</label>
        <input class="sc-input" type="number" id="inp-margin" value="68" onchange="scRecalc()">
      </div>
      <div class="sc-field">
        <label class="sc-label">Runway (months)</label>
        <input class="sc-input" type="number" id="inp-runway" value="15" onchange="scRecalc()">
      </div>

      <div class="sc-section-label">Retention & Customers</div>
      <div class="sc-field">
        <label class="sc-label">NRR (%)</label>
        <input class="sc-input" type="number" id="inp-nrr" value="108" onchange="scRecalc()">
      </div>
      <div class="sc-field">
        <label class="sc-label">LTV:CAC Ratio</label>
        <input class="sc-input" type="number" id="inp-ltvcac" value="3.2" step="0.1" onchange="scRecalc()">
      </div>

      <div class="sc-section-label">Market & Team</div>
      <div class="sc-field">
        <label class="sc-label">TAM (USD millions)</label>
        <input class="sc-input" type="number" id="inp-tam" value="2500" onchange="scRecalc()">
      </div>
      <div class="sc-field">
        <label class="sc-label">Team Score (1–10)</label>
        <input class="sc-input" type="number" id="inp-team" value="7" min="1" max="10" onchange="scRecalc()">
      </div>
    </div>

    <!-- RESULT PANEL -->
    <div>
      <div class="sc-result-panel">
        <div class="score-ring-wrap">
          <div class="score-ring" id="score-ring">
            <span class="score-num" id="score-total">72</span>
            <span class="score-max">/ 100</span>
          </div>
          <div class="score-info">
            <div class="score-verdict-badge verdict-strong" id="score-verdict">Avanzar a DD</div>
            <div class="score-label" id="score-label">Candidato Fuerte</div>
            <div class="score-desc" id="score-desc">
              El perfil supera la mediana LATAM en la mayoría de dimensiones.
              Recomendar avanzar a due diligence.
            </div>
          </div>
        </div>

        <div class="tabs-bar">
          <button class="tab-btn active" onclick="scTab('methodology',this)">Metodología</button>
          <button class="tab-btn" onclick="scTab('growth',this)">Crecimiento</button>
          <button class="tab-btn" onclick="scTab('efficiency',this)">Eficiencia</button>
          <button class="tab-btn" onclick="scTab('retention',this)">Retención</button>
          <button class="tab-btn" onclick="scTab('valuation',this)">Valuación</button>
        </div>

        <div class="tab-panel active" id="sc-tab-methodology">
          <div style="font-size:13px; color:var(--text2); line-height:1.7">
            <p style="margin-bottom:12px">
              El scorecard evalúa <strong>5 dimensiones</strong> contra benchmarks LATAM
              (mediana y top quartile) usando datos del mercado 2024–2025.
            </p>
            <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px">
              <div class="card" style="padding:14px">
                <div style="font-size:12px; font-weight:600; color:var(--navy); margin-bottom:6px">📈 Crecimiento — 25 pts</div>
                <div style="font-size:12px; color:var(--muted)">MoM growth rate + ARR para la etapa</div>
              </div>
              <div class="card" style="padding:14px">
                <div style="font-size:12px; font-weight:600; color:var(--navy); margin-bottom:6px">⚙️ Eficiencia Capital — 20 pts</div>
                <div style="font-size:12px; color:var(--muted)">Burn multiple + gross margin + runway</div>
              </div>
              <div class="card" style="padding:14px">
                <div style="font-size:12px; font-weight:600; color:var(--navy); margin-bottom:6px">🔄 Retención — 20 pts</div>
                <div style="font-size:12px; color:var(--muted)">NRR — principal driver de valuación SaaS</div>
              </div>
              <div class="card" style="padding:14px">
                <div style="font-size:12px; font-weight:600; color:var(--navy); margin-bottom:6px">🎯 Unit Econ & Mercado — 25 pts</div>
                <div style="font-size:12px; color:var(--muted)">LTV:CAC + TAM + calidad del equipo (1–10)</div>
              </div>
            </div>
            <div style="margin-top:14px; padding:12px 16px; background:var(--surface2); border-radius:8px; font-size:12px">
              <strong>Umbrales:</strong>
              <span style="color:var(--green); margin:0 12px">≥70 → Avanzar a DD</span>
              <span style="color:var(--amber); margin:0 12px">50–69 → Watchlist</span>
              <span style="color:var(--red)">0–49 → Pasar</span>
            </div>
          </div>
        </div>

        <div class="tab-panel" id="sc-tab-growth">
          <div class="dim-pill-row">
            <div class="dim-pill-title">📈 Crecimiento</div>
            <div class="dim-score-pill" id="growth-pill">— / 25 pts</div>
          </div>
          <div class="metrics-header">
            <span>Métrica</span><span>LATAM Med</span><span>LATAM Top</span>
            <span>Tu Startup</span><span>Status</span>
          </div>
          <div id="growth-metrics"></div>
        </div>

        <div class="tab-panel" id="sc-tab-efficiency">
          <div class="dim-pill-row">
            <div class="dim-pill-title">⚙️ Eficiencia de Capital</div>
            <div class="dim-score-pill" id="eff-pill">— / 20 pts</div>
          </div>
          <div class="metrics-header">
            <span>Métrica</span><span>LATAM Med</span><span>LATAM Top</span>
            <span>Tu Startup</span><span>Status</span>
          </div>
          <div id="eff-metrics"></div>
        </div>

        <div class="tab-panel" id="sc-tab-retention">
          <div class="dim-pill-row">
            <div class="dim-pill-title">🔄 Retención</div>
            <div class="dim-score-pill" id="ret-pill">— / 20 pts</div>
          </div>
          <div class="metrics-header">
            <span>Métrica</span><span>LATAM Med</span><span>LATAM Top</span>
            <span>Tu Startup</span><span>Status</span>
          </div>
          <div id="ret-metrics"></div>
        </div>

        <div class="tab-panel" id="sc-tab-valuation">
          <div class="dim-pill-row">
            <div class="dim-pill-title">🎯 Unit Econ & Mercado</div>
            <div class="dim-score-pill" id="val-pill">— / 25 pts</div>
          </div>
          <div class="metrics-header">
            <span>Métrica</span><span>LATAM Med</span><span>LATAM Top</span>
            <span>Tu Startup</span><span>Status</span>
          </div>
          <div id="val-metrics"></div>
        </div>
      </div>
    </div>
  </div>
</section>
"""


# ─── VENTURE STUDIO SECTION ───────────────────────────────────────────────────

def build_studio(data):
    vs = data['venture_studio']

    def make_rows(metrics_list):
        rows = ""
        for m in metrics_list:
            rows += f"""
            <tr>
              <td class="td-strong">{m['metric']}</td>
              <td class="td-num td-teal">{m['early']}</td>
              <td class="td-num td-gold">{m['mature']}</td>
              <td class="td-num">{m['industry']}</td>
              <td style="font-size:11px; color:var(--dim)">{m['notes']}</td>
            </tr>"""
        return rows

    a_rows = make_rows(vs['A'])
    b_rows = make_rows(vs['B'])
    c_rows = make_rows(vs['C'])

    return f"""
<section id="studio" class="section">
  <div class="section-eyebrow">Venture Studio Model</div>
  <h2 class="section-title">Studio <em>Metrics Reference</em></h2>
  <p class="section-subtitle">
    Operational, portfolio and fund metrics for AIDA's hybrid VC + Studio model.
    Benchmarks segmented by studio maturity stage.
  </p>

  <div class="kpi-grid kpi-grid-4" style="margin-bottom:28px">
    <div class="kpi-card accent-navy">
      <div class="kpi-label">Startups Built / Year</div>
      <div class="kpi-value">2–4</div>
      <div class="kpi-sub">Early-stage benchmark</div>
      <span class="kpi-badge badge-neu">Mature: 4–8</span>
    </div>
    <div class="kpi-card accent-teal">
      <div class="kpi-label">Survival Rate (Year 3)</div>
      <div class="kpi-value">30–50%</div>
      <div class="kpi-sub">Higher than trad. VC</div>
      <span class="kpi-badge badge-up">Studio support effect</span>
    </div>
    <div class="kpi-card accent-teal">
      <div class="kpi-label">Series A Graduation</div>
      <div class="kpi-value">20–35%</div>
      <div class="kpi-sub">% reaching inst. round</div>
      <span class="kpi-badge badge-up">vs 2–5% mkt avg</span>
    </div>
    <div class="kpi-card accent-gold">
      <div class="kpi-label">Target Gross MOIC</div>
      <div class="kpi-value">3–5×</div>
      <div class="kpi-sub">Early studio</div>
      <span class="kpi-badge badge-neu">Mature: 5–10×</span>
    </div>
  </div>

  <!-- SECTION A -->
  <div class="studio-section">
    <div class="studio-section-header">
      <div class="studio-section-icon">🏗️</div>
      <div>
        <div class="studio-section-title">A. Studio-Level Operational Metrics</div>
        <div class="studio-section-sub">Capacity, cost and team benchmarks</div>
      </div>
    </div>
    <div class="table-wrapper" style="border:none; border-radius:0; box-shadow:none">
      <table class="data-table">
        <thead><tr>
          <th style="min-width:220px">Metric</th>
          <th>Early Studio</th><th>Mature Studio</th>
          <th>Industry Avg</th><th style="min-width:200px">Notes</th>
        </tr></thead>
        <tbody>{a_rows}</tbody>
      </table>
    </div>
  </div>

  <!-- SECTION B -->
  <div class="studio-section">
    <div class="studio-section-header">
      <div class="studio-section-icon">🚀</div>
      <div>
        <div class="studio-section-title">B. Portfolio Company Performance Metrics</div>
        <div class="studio-section-sub">Startup-level KPIs benchmarked against studio-backed cohorts</div>
      </div>
    </div>
    <div class="table-wrapper" style="border:none; border-radius:0; box-shadow:none">
      <table class="data-table">
        <thead><tr>
          <th style="min-width:220px">Metric</th>
          <th>Early Studio</th><th>Mature Studio</th>
          <th>Industry Avg</th><th style="min-width:200px">Notes</th>
        </tr></thead>
        <tbody>{b_rows}</tbody>
      </table>
    </div>
  </div>

  <!-- SECTION C -->
  <div class="studio-section">
    <div class="studio-section-header">
      <div class="studio-section-icon">💰</div>
      <div>
        <div class="studio-section-title">C. Fund &amp; Return Metrics</div>
        <div class="studio-section-sub">LP-facing return targets and fund structure benchmarks</div>
      </div>
    </div>
    <div class="table-wrapper" style="border:none; border-radius:0; box-shadow:none">
      <table class="data-table">
        <thead><tr>
          <th style="min-width:220px">Metric</th>
          <th>Early Studio</th><th>Mature Studio</th>
          <th>Industry Avg</th><th style="min-width:200px">Notes</th>
        </tr></thead>
        <tbody>{c_rows}</tbody>
      </table>
    </div>
  </div>

  <div class="grid-2" style="margin-top:24px">
    <div class="chart-card">
      <div class="chart-title">Return Profile — Early vs Mature Studio</div>
      <div class="chart-sub">Gross MOIC, IRR and Net MOIC comparison</div>
      <div class="chart-wrap h240"><canvas id="studio-returns-chart"></canvas></div>
    </div>
    <div class="card" style="padding:24px">
      <div class="card-title" style="margin-bottom:16px">Power Law Dynamics</div>
      <div style="font-size:13px; color:var(--text2); line-height:1.75">
        <p style="margin-bottom:10px">
          <strong>Top-2 Return Concentration:</strong> 60–80% of total returns from
          just 2 companies per cohort (early studio). Mature studios: 70–90%.
        </p>
        <p style="margin-bottom:10px">
          <strong>Expected Exits:</strong> 1–2 per cohort (early) vs 2–4 per cohort
          (mature), out of 8–10 companies built.
        </p>
        <p style="margin-bottom:10px">
          <strong>Average Exit Valuation:</strong> $30M–$100M (early stage) → 
          $80M–$300M (mature). Target acquisition or strategic sale.
        </p>
        <p>
          <strong>Fund Life:</strong> 7–10 years (early) / 8–12 years (mature).
          Management fee: 2–2.5%. Carry: 20–25%.
        </p>
      </div>
    </div>
  </div>
</section>
"""


# ─── METHODOLOGY SECTION ─────────────────────────────────────────────────────

def build_methodology(gen_date):
    return f"""
<section id="methodology" class="section">
  <div class="section-eyebrow">Data Sources</div>
  <h2 class="section-title"><em>Methodology</em> &amp; Sources</h2>
  <p class="section-subtitle">
    Static dashboard generated from internal Excel data sources. All data represents
    market estimates and benchmarks for reference purposes.
  </p>

  <div style="display:grid; grid-template-columns: 1fr 1fr; gap:20px; margin-bottom:32px">
    <div style="display:flex; gap:14px; align-items:flex-start;" class="source-card">
      <div class="source-num">1</div>
      <div>
        <div class="source-name">Fintech Sectors</div>
        <div class="source-desc">
          Investment by country, LATAM and US subsector analysis, cross-market comparison.
          Covers 9 LATAM countries and 14–15 subsectors with 2024 estimates.
        </div>
        <div style="margin-top:8px">
          <span class="insight-tag tag-teal">9 countries</span>
          <span class="insight-tag tag-teal">14 LATAM subsectors</span>
          <span class="insight-tag tag-gold">15 US subsectors</span>
        </div>
      </div>
    </div>
    <div style="display:flex; gap:14px; align-items:flex-start;" class="source-card">
      <div class="source-num">2</div>
      <div>
        <div class="source-name">VC Funds Metrics</div>
        <div class="source-desc">
          Early-stage fund benchmarks including IRR, TVPI, DPI, MOIC, deal sizes,
          graduation rates and capital deployment metrics.
        </div>
        <div style="margin-top:8px">
          <span class="insight-tag tag-navy">Pre-Seed / Seed / Series A</span>
          <span class="insight-tag tag-navy">LATAM vs US</span>
        </div>
      </div>
    </div>
    <div style="display:flex; gap:14px; align-items:flex-start;" class="source-card">
      <div class="source-num">3</div>
      <div>
        <div class="source-name">Startup Benchmarks — AIDA Ventures</div>
        <div class="source-desc">
          Revenue multiples by stage and sector, US and LATAM valuations, time between
          funding rounds, graduation rates and notable LATAM deals.
          Sources: Carta Q3 2025, Finro, Capstone Partners, LAVCA 2024–2025.
        </div>
        <div style="margin-top:8px">
          <span class="insight-tag tag-teal">SaaS · Fintech · LogTech</span>
          <span class="insight-tag tag-gold">US + LATAM</span>
        </div>
      </div>
    </div>
    <div style="display:flex; gap:14px; align-items:flex-start;" class="source-card">
      <div class="source-num">4</div>
      <div>
        <div class="source-name">Startup Operational Metrics (US)</div>
        <div class="source-desc">
          Income & growth, capital efficiency, retention, unit economics, market
          and valuation benchmarks for Pre-Seed, Seed and Series A.
          Used as scoring baseline in the Deal Scorecard module.
        </div>
        <div style="margin-top:8px">
          <span class="insight-tag tag-navy">8 metric categories</span>
          <span class="insight-tag tag-navy">3 stages</span>
        </div>
      </div>
    </div>
    <div style="display:flex; gap:14px; align-items:flex-start;" class="source-card">
      <div class="source-num">5</div>
      <div>
        <div class="source-name">Venture Studio Metrics Reference</div>
        <div class="source-desc">
          Studio-level operational benchmarks, portfolio company KPIs and fund &amp;
          return metrics. Segmented by Early vs Mature studio stage.
          Private studio dataset based on industry aggregates.
        </div>
        <div style="margin-top:8px">
          <span class="insight-tag tag-teal">A / B / C metrics</span>
          <span class="insight-tag tag-teal">Early vs Mature</span>
        </div>
      </div>
    </div>
    <div style="display:flex; gap:14px; align-items:flex-start;" class="source-card">
      <div class="source-num">6</div>
      <div>
        <div class="source-name">External Data Sources Referenced</div>
        <div class="source-desc">
          Carta (2024–2025), LAVCA VC Report 2024, Crunchbase, PitchBook,
          SaaS Capital Index, Finro FCA Mid-2025, Capstone Partners LogTech Jun 2025,
          Equidam 2025, Aventis Advisors, Forum Ventures State of Seed 2024.
        </div>
        <div style="margin-top:8px">
          <span class="insight-tag tag-gold">Q4 2024 – Q3 2025</span>
        </div>
      </div>
    </div>
  </div>

  <div class="card" style="padding:24px; background:var(--navy); border-color:var(--navy)">
    <div style="display:flex; justify-content:space-between; align-items:center">
      <div>
        <div style="font-family:var(--font-display); font-size:20px; font-weight:600; color:white; margin-bottom:6px">
          AIDA Ventures — Intelligence Platform
        </div>
        <div style="font-size:13px; color:rgba(255,255,255,0.55)">
          Static dashboard generated from Excel sources. Prepared for GitHub Pages deployment.
        </div>
      </div>
      <div style="text-align:right">
        <div style="font-family:var(--font-mono); font-size:11px; color:rgba(255,255,255,0.4)">GENERATED</div>
        <div style="font-family:var(--font-mono); font-size:13px; color:rgba(255,255,255,0.7)">{gen_date}</div>
      </div>
    </div>
  </div>
</section>
"""


# ─── JAVASCRIPT ──────────────────────────────────────────────────────────────

def build_js(data):
    data_json = json.dumps(data, ensure_ascii=False)
    return f"""
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script>
// ── DASHBOARD DATA ──────────────────────────────────────────────────────────
const D = {data_json};

// ── CHART DEFAULTS ──────────────────────────────────────────────────────────
const NAVY = '#0B1F3A'; const NAVY2 = '#1A3A5F'; const ACCENT = '#1A6FB5';
const TEAL = '#0D7377'; const GOLD = '#B8882A';
const COLORS = [TEAL, NAVY, ACCENT, GOLD, '#7C3AED', '#059669', '#D97706', '#DC2626', '#0891B2', '#6D28D9'];

Chart.defaults.font.family = "'Inter', system-ui, sans-serif";
Chart.defaults.font.size = 12;
Chart.defaults.color = '#5A7290';
Chart.defaults.plugins.legend.labels.boxWidth = 12;
Chart.defaults.plugins.legend.labels.padding = 16;

function gridOpts() {{
  return {{ color: 'rgba(0,0,0,0.05)', drawBorder: false }};
}}

// Returns array: max value gets TEAL, all others get NAVY
function barColors(data, base, highlight) {{
  base      = base      || NAVY2;
  highlight = highlight || TEAL;
  const max = Math.max(...data);
  return data.map(v => v === max ? highlight : base);
}}
function barBorders(data, base, highlight) {{
  base      = base      || NAVY;
  highlight = highlight || TEAL;
  const max = Math.max(...data);
  return data.map(v => v === max ? highlight : base);
}}

// ── GLOBAL NAV ──────────────────────────────────────────────────────────────
function showSection(id) {{
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  document.querySelectorAll('.nav-link').forEach(l => {{
    if (l.getAttribute('onclick') && l.getAttribute('onclick').includes("'" + id + "'")) {{
      l.classList.add('active');
    }}
  }});
  window.scrollTo(0, 0);
  if (id === 'overview')   initOverviewCharts();
  if (id === 'market')     initMktCharts('overview');
  if (id === 'vc')         initVcCharts('summary');
  if (id === 'studio')     initStudioCharts();
}}

// ── MARKET TAB SWITCH ────────────────────────────────────────────────────────
function mktTab(id, btn) {{
  document.querySelectorAll('#market .tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('#market-tabs .tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('mkt-' + id).classList.add('active');
  btn.classList.add('active');
  initMktCharts(id);
}}

// ── VC TAB SWITCH ────────────────────────────────────────────────────────────
function vcTab(id, btn) {{
  document.querySelectorAll('#vc .tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('#vc-tabs .tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('vc-' + id).classList.add('active');
  btn.classList.add('active');
  initVcCharts(id);
}}

// ── SCORECARD TAB ────────────────────────────────────────────────────────────
function scTab(id, btn) {{
  document.querySelectorAll('#scorecard .tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('#scorecard .tabs-bar .tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('sc-tab-' + id).classList.add('active');
  btn.classList.add('active');
}}

// ── CHART REGISTRY (prevent duplicate init) ──────────────────────────────────
const _charts = {{}};
function getOrCreate(id, factory) {{
  if (_charts[id]) {{ _charts[id].destroy(); delete _charts[id]; }}
  const canvas = document.getElementById(id);
  if (!canvas) return;
  _charts[id] = factory(canvas.getContext('2d'));
}}

// ── OVERVIEW CHARTS ──────────────────────────────────────────────────────────
function initOverviewCharts() {{
  const countries = D.fintech.countries.slice(0,6);
  const labels = countries.map(c => c.country);
  const vals = [1000, 850, 330, 285, 210, 100];

  getOrCreate('ov-country-chart', ctx => new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels,
      datasets: [{{ label: 'Investment ($M est.)', data: vals,
        backgroundColor: barColors(vals),
        borderColor: barBorders(vals), borderWidth: 1.5, borderRadius: 4 }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{ legend: {{ display: false }} }},
      scales: {{
        x: {{ grid: gridOpts() }},
        y: {{ grid: gridOpts(), ticks: {{ callback: v => '$' + v + 'M' }} }}
      }}
    }}
  }}));

  const stages = ['Pre-Seed (1000)', 'Seed', 'Series A', 'Series B', 'Series C+', 'Exit'];
  const usVals  = [1000, 400, 15, 6, 3, 1];
  const laVals  = [1000, 275, 22, 8, 3, 1];

  getOrCreate('ov-funnel-chart', ctx => new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: stages,
      datasets: [
        {{ label: 'US Survivors', data: usVals, backgroundColor: GOLD + 'BB', borderColor: GOLD, borderWidth: 1.5, borderRadius: 4 }},
        {{ label: 'LATAM Survivors', data: laVals, backgroundColor: TEAL + 'BB', borderColor: TEAL, borderWidth: 1.5, borderRadius: 4 }}
      ]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{ legend: {{ position: 'bottom' }} }},
      scales: {{
        x: {{ grid: gridOpts() }},
        y: {{ grid: gridOpts(), title: {{ display: true, text: 'Survivors of 1,000' }} }}
      }}
    }}
  }}));
}}

// ── MARKET CHARTS ────────────────────────────────────────────────────────────
function initMktCharts(tab) {{
  if (tab === 'overview') {{
    getOrCreate('mkt-country-bar', ctx => new Chart(ctx, {{
      type: 'bar',
      data: {{
        labels: ['Brazil','Mexico','Colombia','Argentina','Chile','Peru'],
        datasets: [{{ label: 'Investment ($M est.)',
          data: [1000, 850, 330, 285, 210, 100],
          backgroundColor: barColors([1000,850,330,285,210,100]),
          borderColor: barBorders([1000,850,330,285,210,100]), borderWidth: 1.5, borderRadius: 5 }}]
      }},
      options: {{ responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ x: {{ grid: gridOpts() }}, y: {{ grid: gridOpts(), ticks: {{ callback: v => '$' + v + 'M' }} }} }}
      }}
    }}));

    const subLabels = D.fintech.latam_subsectors.slice(0,8).map(s => s.subsector.substring(0,20));
    const subVals = [900, 675, 475, 300, 175, 350, 240, 150];
    getOrCreate('mkt-sub-donut', ctx => new Chart(ctx, {{
      type: 'doughnut',
      data: {{ labels: subLabels, datasets: [{{ data: subVals,
        backgroundColor: COLORS.map(c => c + 'CC'), borderWidth: 1 }}] }},
      options: {{ responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'right', labels: {{ font: {{ size: 10 }} }} }} }}
      }}
    }}));
  }}
  if (tab === 'countries') {{
    getOrCreate('mkt-share-donut', ctx => new Chart(ctx, {{
      type: 'doughnut',
      data: {{
        labels: ['Brazil','Mexico','Colombia','Argentina','Chile','Peru','Uruguay','C. America','Other'],
        datasets: [{{ data: [33, 28, 11, 10, 7, 3.5, 2.5, 3, 2],
          backgroundColor: COLORS.map(c => c + 'CC'), borderWidth: 1 }}]
      }},
      options: {{ responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'right', labels: {{ font: {{ size: 11 }} }} }} }}
      }}
    }}));
  }}
  if (tab === 'latam-sub') {{
    const ls = D.fintech.latam_subsectors.slice(0,10);
    getOrCreate('mkt-latam-sub-bar', ctx => new Chart(ctx, {{
      type: 'bar',
      data: {{
        labels: ls.map(s => s.subsector.substring(0,22)),
        datasets: [{{ label: 'Investment ($M est.)',
          data: [900,675,475,300,175,350,240,150,100,185],
          backgroundColor: barColors([900,675,475,300,175,350,240,150,100,185]),
          borderColor: barBorders([900,675,475,300,175,350,240,150,100,185]), borderWidth: 1.5, borderRadius: 4 }}]
      }},
      options: {{ indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ x: {{ grid: gridOpts(), ticks: {{ callback: v => '$' + v + 'M' }} }}, y: {{ grid: gridOpts() }} }}
      }}
    }}));
  }}
  if (tab === 'usa') {{
    const us = D.fintech.usa_subsectors.slice(0,10);
    const usVals = [31.0, 8.5, 5.3, 4.9, 4.7, 4.5, 4.2, 3.8, 3.2, 2.8];
    getOrCreate('mkt-usa-bar', ctx => new Chart(ctx, {{
      type: 'bar',
      data: {{
        labels: us.map(s => s.subsector.substring(0,24)),
        datasets: [{{ label: 'Investment ($B)',
          data: usVals,
          backgroundColor: barColors(usVals),
          borderColor: barBorders(usVals), borderWidth: 1.5, borderRadius: 4 }}]
      }},
      options: {{ indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ display: false }} }},
        scales: {{ x: {{ grid: gridOpts(), ticks: {{ callback: v => '$' + v + 'B' }} }}, y: {{ grid: gridOpts() }} }}
      }}
    }}));
  }}
  if (tab === 'compare') {{
    getOrCreate('mkt-compare-radar', ctx => new Chart(ctx, {{
      type: 'radar',
      data: {{
        labels: ['Investment Volume', 'Deal Count', 'Deal Size', 'Unicorns', 'AI Adoption', 'YoY Growth'],
        datasets: [
          {{ label: 'USA', data: [100, 83, 100, 100, 50, 20], fill: true,
            backgroundColor: GOLD + '22', borderColor: GOLD, pointBackgroundColor: GOLD }},
          {{ label: 'LATAM', data: [7, 17, 37, 12, 85, 85], fill: true,
            backgroundColor: TEAL + '22', borderColor: TEAL, pointBackgroundColor: TEAL }}
        ]
      }},
      options: {{ responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom' }} }},
        scales: {{ r: {{ grid: {{ color: 'rgba(0,0,0,0.06)' }}, ticks: {{ display: false }}, max: 100 }} }}
      }}
    }}));
  }}
  if (tab === 'opps') {{
    getOrCreate('mkt-opp-bubble', ctx => new Chart(ctx, {{
      type: 'bubble',
      data: {{
        datasets: [
          {{ label: 'Payments', data: [{{ x: 900, y: 32, r: 20 }}], backgroundColor: TEAL + 'CC', borderColor: TEAL }},
          {{ label: 'Open Banking', data: [{{ x: 300, y: 45, r: 14 }}], backgroundColor: ACCENT + 'CC', borderColor: ACCENT }},
          {{ label: 'B2B Finance', data: [{{ x: 350, y: 38, r: 15 }}], backgroundColor: NAVY2 + 'CC', borderColor: NAVY2 }},
          {{ label: 'Insurtech', data: [{{ x: 175, y: 28, r: 10 }}], backgroundColor: '#7C3AED' + 'CC', borderColor: '#7C3AED' }},
          {{ label: 'Crypto/Stable', data: [{{ x: 215, y: 52, r: 12 }}], backgroundColor: GOLD + 'CC', borderColor: GOLD }},
          {{ label: 'Alt Credit Score', data: [{{ x: 130, y: 55, r: 11 }}], backgroundColor: '#059669' + 'CC', borderColor: '#059669' }},
        ]
      }},
      options: {{ responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom', labels: {{ font: {{ size: 11 }} }} }} }},
        scales: {{
          x: {{ grid: gridOpts(), title: {{ display: true, text: 'Investment Size ($M)' }} }},
          y: {{ grid: gridOpts(), title: {{ display: true, text: 'Growth Potential (Index)' }} }}
        }}
      }}
    }}));
  }}
}}

// ── VC CHARTS ────────────────────────────────────────────────────────────────
function initVcCharts(tab) {{
  if (tab === 'summary') {{
    getOrCreate('vc-irr-chart', ctx => new Chart(ctx, {{
      type: 'bar',
      data: {{
        labels: ['Pre-Seed', 'Seed', 'Series A'],
        datasets: [
          {{ label: 'IRR Target (%)', data: [35, 32.5, 25], backgroundColor: NAVY + 'CC', borderColor: NAVY, borderWidth: 1.5, borderRadius: 4 }},
          {{ label: 'IRR Median Observed (%)', data: [10, 12.5, 15], backgroundColor: ACCENT + 'CC', borderColor: ACCENT, borderWidth: 1.5, borderRadius: 4 }}
        ]
      }},
      options: {{ responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom' }} }},
        scales: {{
          x: {{ grid: gridOpts() }},
          y: {{ grid: gridOpts(), ticks: {{ callback: v => v + '%' }} }}
        }}
      }}
    }}));
  }}
  if (tab === 'latam-us') {{
    getOrCreate('vc-compare-chart', ctx => new Chart(ctx, {{
      type: 'bar',
      data: {{
        labels: ['IRR Median', 'IRR Top Decile', 'Grad Rate Seed→A'],
        datasets: [
          {{ label: 'US', data: [12, 28, 22.5], backgroundColor: GOLD + 'BB', borderColor: GOLD, borderWidth: 1.5, borderRadius: 4 }},
          {{ label: 'LATAM', data: [10, 22.5, 12.5], backgroundColor: TEAL + 'BB', borderColor: TEAL, borderWidth: 1.5, borderRadius: 4 }}
        ]
      }},
      options: {{ responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom' }} }},
        scales: {{ x: {{ grid: gridOpts() }}, y: {{ grid: gridOpts() }} }}
      }}
    }}));
  }}
  if (tab === 'valuations') {{
    getOrCreate('vc-val-chart', ctx => new Chart(ctx, {{
      type: 'bar',
      data: {{
        labels: ['Seed', 'Series A', 'Series B'],
        datasets: [
          {{ label: 'US SaaS ($M)', data: [19.8, 60, 175], backgroundColor: GOLD + 'BB', borderColor: GOLD, borderWidth: 1.5, borderRadius: 4 }},
          {{ label: 'LATAM SaaS ($M)', data: [10, 30, 87.5], backgroundColor: TEAL + 'BB', borderColor: TEAL, borderWidth: 1.5, borderRadius: 4 }},
          {{ label: 'LATAM Fintech ($M)', data: [11, 35, 105], backgroundColor: ACCENT + 'BB', borderColor: ACCENT, borderWidth: 1.5, borderRadius: 4 }}
        ]
      }},
      options: {{ responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom' }} }},
        scales: {{
          x: {{ grid: gridOpts() }},
          y: {{ grid: gridOpts(), ticks: {{ callback: v => '$' + v + 'M' }} }}
        }}
      }}
    }}));
  }}
  if (tab === 'timing') {{
    const timingData2024 = [15, 25.8, 24.4, 28];
    const timingData2021 = [9, 14, 10, 15];
    getOrCreate('vc-timing-chart', ctx => new Chart(ctx, {{
      type: 'bar',
      data: {{
        labels: ['Pre-Seed→Seed', 'Seed→Series A', 'A→Series B', 'B→Series C'],
        datasets: [
          {{ label: '2021 (months)', data: timingData2021, backgroundColor: barColors(timingData2021, '#94A3B8', TEAL), borderColor: barBorders(timingData2021, '#64748B', TEAL), borderWidth: 1.5, borderRadius: 4 }},
          {{ label: '2024 (months)', data: timingData2024, backgroundColor: barColors(timingData2024), borderColor: barBorders(timingData2024), borderWidth: 1.5, borderRadius: 4 }}
        ]
      }},
      options: {{ responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom' }} }},
        scales: {{
          x: {{ grid: gridOpts() }},
          y: {{ grid: gridOpts(), title: {{ display: true, text: 'Months' }} }}
        }}
      }}
    }}));
  }}
  if (tab === 'graduation') {{
    getOrCreate('vc-grad-chart', ctx => new Chart(ctx, {{
      type: 'bar',
      data: {{
        labels: ['Pre-Seed', 'Seed', 'Series A', 'Series B', 'Series C+', 'Exit'],
        datasets: [
          {{ label: 'US Survivors', data: [1000, 400, 15, 6, 3, 1], backgroundColor: GOLD + 'BB', borderColor: GOLD, borderWidth: 1.5, borderRadius: 4 }},
          {{ label: 'LATAM Survivors', data: [1000, 275, 22, 8, 3, 1], backgroundColor: TEAL + 'BB', borderColor: TEAL, borderWidth: 1.5, borderRadius: 4 }}
        ]
      }},
      options: {{ responsive: true, maintainAspectRatio: false,
        plugins: {{ legend: {{ position: 'bottom' }} }},
        scales: {{
          x: {{ grid: gridOpts() }},
          y: {{ grid: gridOpts(), title: {{ display: true, text: 'Survivors of 1,000' }} }}
        }}
      }}
    }}));
  }}
}}

// ── STUDIO CHARTS ─────────────────────────────────────────────────────────────
function initStudioCharts() {{
  getOrCreate('studio-returns-chart', ctx => new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: ['Target Gross MOIC', 'Target Net MOIC', 'Gross IRR (%)', 'Net IRR (%)'],
      datasets: [
        {{ label: 'Early Studio', data: [4, 3, 25, 20], backgroundColor: TEAL + 'BB', borderColor: TEAL, borderWidth: 1.5, borderRadius: 4 }},
        {{ label: 'Mature Studio', data: [7.5, 5, 32.5, 27.5], backgroundColor: NAVY + 'BB', borderColor: NAVY, borderWidth: 1.5, borderRadius: 4 }}
      ]
    }},
    options: {{ responsive: true, maintainAspectRatio: false,
      plugins: {{ legend: {{ position: 'bottom' }} }},
      scales: {{ x: {{ grid: gridOpts() }}, y: {{ grid: gridOpts() }} }}
    }}
  }}));
}}

// ── SCORECARD LOGIC (preserved) ───────────────────────────────────────────────
function scScore(val, latamMed, latamTop, max) {{
  const [lmLo, lmHi] = latamMed;
  const [ltLo, ltHi] = latamTop;
  const lmMid = (lmLo + lmHi) / 2;
  const ltMid = (ltLo + ltHi) / 2;
  if (val >= ltMid) return max;
  if (val >= lmMid) return Math.round(max * 0.7 + (val - lmMid) / (ltMid - lmMid) * max * 0.3);
  if (val >= lmLo)  return Math.round(max * 0.4 + (val - lmLo) / (lmMid - lmLo) * max * 0.3);
  return Math.max(0, Math.round(max * 0.4 * (val / lmLo)));
}}

function scScoreBurn(val, max) {{
  if (val <= 1.5) return max;
  if (val <= 2.0) return Math.round(max * 0.75);
  if (val <= 2.5) return Math.round(max * 0.5);
  if (val <= 3.5) return Math.round(max * 0.25);
  return 0;
}}

function scStatusClass(pct) {{
  if (pct >= 0.75) return 'status-green';
  if (pct >= 0.45) return 'status-yellow';
  return 'status-red';
}}
function scStatusLabel(pct) {{
  if (pct >= 0.75) return '✓ Fuerte';
  if (pct >= 0.45) return '~ Aceptable';
  return '✗ Bajo';
}}

function scRenderMetrics(id, rows) {{
  const el = document.getElementById(id);
  if (!el) return;
  el.innerHTML = rows.map(r => `
    <div class="metric-row">
      <div><div class="metric-name">${{r.name}}</div><div class="metric-sub">${{r.sub||''}}</div></div>
      <div class="metric-val" style="color:var(--muted)">${{r.latamMed}}</div>
      <div class="metric-val" style="color:var(--teal)">${{r.latamTop}}</div>
      <div class="metric-val" style="color:var(--navy); font-weight:600">${{r.yours}}</div>
      <div style="text-align:center"><span class="status ${{r.cls}}">${{r.lbl}}</span></div>
    </div>
  `).join('');
}}

function scRecalc() {{
  const arr    = parseFloat(document.getElementById('inp-arr').value) || 0;
  const mom    = parseFloat(document.getElementById('inp-mom').value) || 0;
  const nrr    = parseFloat(document.getElementById('inp-nrr').value) || 0;
  const burn   = parseFloat(document.getElementById('inp-burn').value) || 0;
  const ltvcac = parseFloat(document.getElementById('inp-ltvcac').value) || 0;
  const margin = parseFloat(document.getElementById('inp-margin').value) || 0;
  const team   = parseFloat(document.getElementById('inp-team').value) || 0;
  const tam    = parseFloat(document.getElementById('inp-tam').value) || 0;
  const runway = parseFloat(document.getElementById('inp-runway').value) || 0;
  const stage  = document.getElementById('inp-stage').value;

  const arrBench = {{ preseed: [0, 100000], seed: [100000, 1000000], seriea: [1500000, 3000000] }};
  const arrLM = arrBench[stage];

  const sGrowth   = scScore(mom, [8, 15], [20, 30], 15) + scScore(arr, arrLM, [arrLM[1], arrLM[1]*2], 10);
  const sRetention = scScore(nrr, [90, 110], [110, 130], 20);
  const sBurn     = scScoreBurn(burn, 10) + scScore(margin, [55, 65], [65, 75], 6) + scScore(runway, [12, 15], [15, 24], 4);
  const sUnitEcon = scScore(ltvcac, [2.5, 3.5], [3.0, 4.5], 15);
  const sTeam     = Math.round((team / 10) * 10);
  const sMkt      = scScore(tam, [1000, 2000], [2000, 5000], 10);
  const total     = Math.min(100, sGrowth + sRetention + sBurn + sUnitEcon + sTeam + sMkt);

  document.getElementById('score-total').textContent = total;
  const ring    = document.getElementById('score-ring');
  const verdict = document.getElementById('score-verdict');
  const label   = document.getElementById('score-label');
  const desc    = document.getElementById('score-desc');
  const num     = document.getElementById('score-total');

  if (total >= 70) {{
    label.textContent = 'Candidato Fuerte';
    desc.textContent  = 'El perfil supera la mediana LATAM en la mayoría de dimensiones. Recomendar avanzar a due diligence.';
    verdict.textContent = 'Avanzar a DD';
    verdict.className = 'score-verdict-badge verdict-strong';
    ring.style.borderColor = 'var(--green)'; num.style.color = 'var(--green)';
  }} else if (total >= 50) {{
    label.textContent = 'Watchlist — Potencial Condicional';
    desc.textContent  = 'Algunas dimensiones clave están por debajo del benchmark. Requiere más tracción.';
    verdict.textContent = 'Watchlist';
    verdict.className = 'score-verdict-badge verdict-watch';
    ring.style.borderColor = 'var(--amber)'; num.style.color = 'var(--amber)';
  }} else {{
    label.textContent = 'No califica en este ciclo';
    desc.textContent  = 'El perfil actual está por debajo de los umbrales mínimos. Re-evaluar en 6–12 meses.';
    verdict.textContent = 'Pasar';
    verdict.className = 'score-verdict-badge verdict-pass';
    ring.style.borderColor = 'var(--red)'; num.style.color = 'var(--red)';
  }}

  document.getElementById('growth-pill').textContent = sGrowth + ' / 25 pts';
  document.getElementById('eff-pill').textContent    = sBurn + ' / 20 pts';
  document.getElementById('ret-pill').textContent    = sRetention + ' / 20 pts';
  document.getElementById('val-pill').textContent    = (sUnitEcon + sTeam + sMkt) + ' / 35 pts';

  const arrFmt = arr >= 1e6 ? '$'+(arr/1e6).toFixed(1)+'M' : arr >= 1000 ? '$'+(arr/1000).toFixed(0)+'K' : '$'+arr;
  const arrLMFmt = arrLM[0] >= 1e6 ? '$'+(arrLM[0]/1e6).toFixed(1)+'–'+(arrLM[1]/1e6).toFixed(1)+'M'
                                    : '$'+(arrLM[0]/1000).toFixed(0)+'–'+(arrLM[1]/1000).toFixed(0)+'K';
  const momPct = mom >= 20 ? 1 : mom >= 8 ? 0.65 : 0.25;
  const arrPct = arr >= arrLM[1] ? 1 : arr >= arrLM[0] ? 0.65 : 0.3;

  scRenderMetrics('growth-metrics', [
    {{ name:'MoM Growth Rate', sub:'Crecimiento mensual MRR', latamMed:'8–15%', latamTop:'20–30%', yours:mom+'%', cls:scStatusClass(momPct), lbl:scStatusLabel(momPct) }},
    {{ name:'ARR actual', sub:'Benchmark para etapa '+stage, latamMed:arrLMFmt, latamTop:'>'+arrLMFmt.split('–')[1], yours:arrFmt, cls:scStatusClass(arrPct), lbl:scStatusLabel(arrPct) }},
    {{ name:'ARR YoY est.', sub:'Si MoM se mantiene', latamMed:'55–75%', latamTop:'100–150%', yours:'~'+Math.round((Math.pow(1+mom/100,12)-1)*100)+'%', cls:scStatusClass(mom>=15?0.8:mom>=8?0.55:0.25), lbl:scStatusLabel(mom>=15?0.8:mom>=8?0.55:0.25) }}
  ]);
  const burnPct = burn<=1.5?1:burn<=2?0.75:burn<=2.5?0.5:0.2;
  const mrgPct  = margin>=65?1:margin>=55?0.65:0.3;
  const rnwPct  = runway>=18?1:runway>=12?0.65:0.25;
  scRenderMetrics('eff-metrics', [
    {{ name:'Burn Multiple', sub:'Menor es mejor', latamMed:'2.0–2.5×', latamTop:'<1.5×', yours:burn.toFixed(1)+'×', cls:scStatusClass(burnPct), lbl:scStatusLabel(burnPct) }},
    {{ name:'Gross Margin', sub:'Margen bruto', latamMed:'55–65%', latamTop:'65–75%', yours:margin+'%', cls:scStatusClass(mrgPct), lbl:scStatusLabel(mrgPct) }},
    {{ name:'Runway', sub:'Meses disponibles', latamMed:'12–15m', latamTop:'15–18m', yours:runway+'m', cls:scStatusClass(rnwPct), lbl:scStatusLabel(rnwPct) }}
  ]);
  const nrrPct = nrr>=120?1:nrr>=100?0.75:nrr>=90?0.5:0.25;
  const ltvPct = ltvcac>=4?1:ltvcac>=3?0.75:ltvcac>=2.5?0.5:0.2;
  scRenderMetrics('ret-metrics', [
    {{ name:'NRR', sub:'Net Revenue Retention', latamMed:'90–110%', latamTop:'>120%', yours:nrr+'%', cls:scStatusClass(nrrPct), lbl:scStatusLabel(nrrPct) }},
    {{ name:'Logo Retention est.', sub:'Derivada de NRR', latamMed:'60–75%', latamTop:'>80%', yours:nrr>=110?'>85%':nrr>=100?'75–85%':'<75%', cls:scStatusClass(nrr>=110?0.9:nrr>=100?0.65:0.35), lbl:scStatusLabel(nrr>=110?0.9:nrr>=100?0.65:0.35) }}
  ]);
  scRenderMetrics('val-metrics', [
    {{ name:'LTV:CAC', sub:'Mínimo 3:1', latamMed:'2.5–3.5:1', latamTop:'3–4.5:1', yours:ltvcac.toFixed(1)+':1', cls:scStatusClass(ltvPct), lbl:scStatusLabel(ltvPct) }},
    {{ name:'TAM', sub:'Mercado total (USD)', latamMed:'>$1B', latamTop:'>$2B', yours:'$'+(tam>=1000?(tam/1000).toFixed(1)+'B':tam+'M'), cls:scStatusClass(tam>=2000?1:tam>=1000?0.7:0.3), lbl:scStatusLabel(tam>=2000?1:tam>=1000?0.7:0.3) }},
    {{ name:'Equipo', sub:'Score 1–10 subjetivo', latamMed:'5–7', latamTop:'8–10', yours:team+'/10', cls:scStatusClass(team>=8?1:team>=6?0.65:0.3), lbl:scStatusLabel(team>=8?1:team>=6?0.65:0.3) }}
  ]);
}}

// ── INIT ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {{
  initOverviewCharts();
  scRecalc();
}});
</script>
"""


# ─── FULL HTML ASSEMBLER ─────────────────────────────────────────────────────

def build_html(data):
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>AIDA Ventures — Intelligence Platform</title>
  <style>{CSS}</style>
</head>
<body>

{build_navbar()}

<div style="max-width:1280px; margin:0 auto">
{build_overview(data)}
{build_market(data)}
{build_vc(data)}
{build_scorecard()}
{build_studio(data)}
{build_methodology(GEN_DATE)}
</div>

{build_js(data)}
</body>
</html>"""


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    print("AIDA Ventures — Integrated Dashboard Builder")
    print("=" * 50)
    print("Step 1: Extracting Excel data...")
    data = extract_data.extract_all()

    print("Step 2: Saving processed data JSON...")
    extract_data.save_json(data)

    print("Step 3: Generating HTML dashboard...")
    html = build_html(data)

    os.makedirs(OUT_DIR, exist_ok=True)
    out_path = os.path.join(OUT_DIR, 'index.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    size_kb = len(html.encode('utf-8')) / 1024
    print(f"  → Output: {out_path}")
    print(f"  → Size: {size_kb:.1f} KB")
    print("\nDone! Open outputs/index.html in your browser.")


if __name__ == '__main__':
    main()