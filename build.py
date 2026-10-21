# -*- coding: utf-8 -*-
"""Build the self-contained index.html from Sheet1 of 词表.xlsx."""
import io
import json

try:
    import openpyxl
except ModuleNotFoundError:
    raise SystemExit(
        "缺少 openpyxl。请先安装：\n"
        "  python -m pip install openpyxl\n"
        "如果电脑上有多个 Python，请安装到实际运行脚本的那个解释器：\n"
        "  python -c \"import sys; print(sys.executable)\""
    )

WORDS_PATH = "词表.xlsx"
SHEET = "Sheet1"
OUTPUT = "index.html"

TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#f4f2ed" media="(prefers-color-scheme: light)">
<meta name="theme-color" content="#101210" media="(prefers-color-scheme: dark)">
<title>GRE 单词</title>
<style>
  :root {
    color-scheme: light dark;
    --bg: #f4f2ed;
    --card: #fffefb;
    --ink: #1c1b18;
    --muted: #77736a;
    --line: #e7e2d8;
    --soft: #f3f0e9;
    --known: #2f7d57;
    --unknown: #bf4a3c;
    --known-soft: #e7f3ec;
    --known-text: #256b49;
    --unknown-soft: #f9e9e6;
    --unknown-text: #a33b2e;
    --shadow: 0 10px 40px rgba(40, 35, 25, .08);
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #101210;
      --card: #191c19;
      --ink: #ecebe4;
      --muted: #9a988f;
      --line: #2c302c;
      --soft: #232722;
      --known-soft: #1d3227;
      --known-text: #8fd7ae;
      --unknown-soft: #38221f;
      --unknown-text: #f0a196;
      --shadow: 0 10px 40px rgba(0, 0, 0, .35);
    }
  }
  * { box-sizing: border-box; }
  [hidden] { display: none !important; }
  html { -webkit-text-size-adjust: 100%; }
  body {
    margin: 0;
    background: var(--bg);
    color: var(--ink);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
    -webkit-font-smoothing: antialiased;
    -webkit-tap-highlight-color: transparent;
  }
  .app {
    width: min(100%, 540px);
    min-height: 100vh;
    min-height: 100dvh;
    margin: 0 auto;
    padding: 18px 16px calc(16px + env(safe-area-inset-bottom));
    display: flex;
    flex-direction: column;
  }
  .topbar {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 12px;
    padding: 4px 4px 14px;
  }
  .brand {
    font-size: 13px;
    font-weight: 650;
    letter-spacing: .09em;
    color: var(--muted);
  }
  .today {
    font-size: 12px;
    color: var(--muted);
    font-variant-numeric: tabular-nums;
    text-align: right;
  }
  .card {
    margin-block: auto;
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 22px;
    box-shadow: var(--shadow);
    padding: 30px 20px 24px;
  }
  .prompt {
    margin: 0 0 18px;
    text-align: center;
    font-size: 12px;
    letter-spacing: .12em;
    color: var(--muted);
  }
  .word {
    margin: 0 0 28px;
    text-align: center;
    font-family: Georgia, "Times New Roman", "Songti SC", serif;
    font-weight: 500;
    font-size: clamp(42px, 13vw, 64px);
    line-height: 1.08;
    letter-spacing: -.01em;
    overflow-wrap: anywhere;
  }
  .word.long { font-size: clamp(24px, 7.5vw, 40px); line-height: 1.2; }
  .actions { display: flex; gap: 12px; }
  .btn {
    appearance: none;
    border: none;
    font: inherit;
    cursor: pointer;
    border-radius: 14px;
    padding: 16px 18px;
    font-size: 17px;
    font-weight: 600;
    transition: transform .08s ease, filter .15s ease;
  }
  .btn:active { transform: scale(.975); }
  .btn:focus-visible, .link:focus-visible { outline: 2px solid var(--ink); outline-offset: 2px; }
  .actions .btn { flex: 1; color: #fff; }
  .btn.known { background: var(--known); }
  .btn.unknown { background: var(--unknown); }
  @media (hover: hover) {
    .btn:hover { filter: brightness(1.06); }
  }
  .word-meta {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: center;
    gap: 7px 10px;
    margin: -10px 0 22px;
    font-size: 12.5px;
    color: var(--muted);
    font-variant-numeric: tabular-nums;
    text-align: center;
  }
  .tag {
    padding: 4px 9px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 650;
  }
  .tag.known { background: var(--known-soft); color: var(--known-text); }
  .tag.unknown { background: var(--unknown-soft); color: var(--unknown-text); }
  .details { animation: rise .22s ease both; }
  @keyframes rise {
    from { opacity: 0; transform: translateY(6px); }
    to { opacity: 1; transform: none; }
  }
  .fields { margin: 0; }
  .field { padding: 14px 0; border-top: 1px solid var(--line); }
  .field:first-child { padding-top: 2px; border-top: none; }
  .field dt {
    margin-bottom: 5px;
    font-size: 11px;
    letter-spacing: .1em;
    color: var(--muted);
  }
  .field dd {
    margin: 0;
    font-size: 16px;
    line-height: 1.65;
    white-space: pre-line;
    overflow-wrap: anywhere;
  }
  .reveal-actions {
    display: flex;
    gap: 10px;
    margin-top: 20px;
  }
  .btn.next {
    flex: 1;
    padding: 16px;
    background: var(--ink);
    color: var(--card);
    font-size: 17px;
  }
  .btn.mastered {
    flex: none;
    padding: 16px 20px;
    border: 1px solid var(--line);
    background: var(--soft);
    color: var(--ink);
  }
  .neighbors {
    margin-top: 26px;
    padding-top: 18px;
    border-top: 1px dashed var(--line);
  }
  .neighbor-group + .neighbor-group { margin-top: 15px; }
  .neighbor-label {
    margin-bottom: 8px;
    font-size: 11px;
    letter-spacing: .1em;
    color: var(--muted);
  }
  .chips {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 8px;
  }
  .chip {
    display: flex;
    flex-direction: column;
    gap: 3px;
    padding: 9px 11px;
    border: 1px solid var(--line);
    border-radius: 12px;
    background: var(--soft);
  }
  .chip-word {
    color: var(--ink);
    font-family: Georgia, "Times New Roman", "Songti SC", serif;
    font-size: 14.5px;
    font-weight: 600;
    line-height: 1.3;
    overflow-wrap: anywhere;
  }
  .chip-meaning {
    color: var(--muted);
    font-size: 12px;
    line-height: 1.45;
    white-space: pre-line;
    overflow-wrap: anywhere;
  }
  .footer {
    padding: 20px 4px 2px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    font-size: 12px;
    color: var(--muted);
    font-variant-numeric: tabular-nums;
  }
  .footer-actions { display: flex; align-items: center; gap: 2px; }
  .link {
    padding: 6px;
    border: none;
    background: none;
    color: var(--muted);
    font: inherit;
    text-decoration: underline;
    text-underline-offset: 2px;
    white-space: nowrap;
    cursor: pointer;
  }
  .icon-btn:focus-visible, .restore-btn:focus-visible { outline: 2px solid var(--ink); outline-offset: 2px; }
  .sheet {
    position: fixed;
    inset: 0;
    z-index: 50;
    display: flex;
    justify-content: center;
    padding: 16px;
    background: rgba(20, 20, 16, .42);
    -webkit-backdrop-filter: blur(3px);
    backdrop-filter: blur(3px);
  }
  .sheet-panel {
    display: flex;
    flex-direction: column;
    width: min(100%, 540px);
    max-height: calc(100vh - 32px);
    max-height: calc(100dvh - 32px);
    overflow: hidden;
    border: 1px solid var(--line);
    border-radius: 20px;
    background: var(--card);
    box-shadow: var(--shadow);
    animation: rise .2s ease both;
  }
  .sheet-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    padding: 16px 18px 14px;
    border-bottom: 1px solid var(--line);
  }
  .sheet-head h2 { margin: 0; font-size: 17px; font-weight: 650; }
  .sheet-count { margin-left: 4px; color: var(--muted); font-size: 13px; font-weight: 400; }
  .icon-btn {
    flex: none;
    width: 34px;
    height: 34px;
    padding: 0;
    border: 1px solid var(--line);
    border-radius: 999px;
    background: var(--soft);
    color: var(--ink);
    font: inherit;
    font-size: 16px;
    line-height: 1;
    cursor: pointer;
  }
  .mastered-list {
    overflow-y: auto;
    padding: 4px 18px calc(18px + env(safe-area-inset-bottom));
    -webkit-overflow-scrolling: touch;
  }
  .mastered-item {
    display: flex;
    align-items: flex-start;
    gap: 12px;
    padding: 13px 0;
    border-bottom: 1px solid var(--line);
  }
  .mastered-item:last-child { border-bottom: none; }
  .mastered-text { flex: 1; min-width: 0; }
  .mastered-word {
    display: block;
    font-family: Georgia, "Times New Roman", "Songti SC", serif;
    font-size: 17px;
    font-weight: 600;
    line-height: 1.35;
    overflow-wrap: anywhere;
  }
  .mastered-meaning {
    display: block;
    margin-top: 2px;
    color: var(--muted);
    font-size: 13px;
    line-height: 1.5;
    white-space: pre-line;
    overflow-wrap: anywhere;
  }
  .restore-btn {
    flex: none;
    padding: 6px 12px;
    border: 1px solid var(--line);
    border-radius: 999px;
    background: var(--soft);
    color: var(--muted);
    font: inherit;
    font-size: 12px;
    cursor: pointer;
  }
  .mastered-empty {
    padding: 44px 8px;
    color: var(--muted);
    font-size: 14px;
    text-align: center;
  }
  .toast {
    position: fixed;
    left: 50%;
    bottom: calc(24px + env(safe-area-inset-bottom));
    z-index: 80;
    padding: 10px 16px;
    border-radius: 999px;
    background: var(--ink);
    color: var(--card);
    box-shadow: var(--shadow);
    font-size: 13px;
    pointer-events: none;
    transform: translateX(-50%);
    animation: toast-in .18s ease both;
  }
  @keyframes toast-in {
    from { opacity: 0; transform: translate(-50%, 6px); }
    to { opacity: 1; transform: translate(-50%, 0); }
  }
  @media (prefers-reduced-motion: reduce) {
    .btn { transition: none; }
    .details, .sheet-panel, .toast { animation: none; }
  }
</style>
</head>
<body>
<main class="app">
  <header class="topbar">
    <span class="brand">GRE 单词</span>
    <span class="today" id="todayStats"></span>
  </header>

  <section class="card">
    <p class="prompt" id="prompt">认识这个单词吗？</p>
    <h1 class="word" id="word"></h1>
    <div class="word-meta" id="wordMeta" hidden>
      <span class="tag" id="resultTag"></span>
      <span id="countText"></span>
    </div>

    <div class="actions" id="actions">
      <button class="btn unknown" id="btnUnknown" type="button">不认识</button>
      <button class="btn known" id="btnKnown" type="button">认识</button>
    </div>

    <div class="details" id="details" hidden>
      <dl class="fields">
        <div class="field" id="meaningField">
          <dt>词义</dt>
          <dd id="meaning"></dd>
        </div>
        <div class="field" id="synField" hidden>
          <dt>近义词</dt>
          <dd id="synonyms"></dd>
        </div>
        <div class="field" id="relField" hidden>
          <dt>相关</dt>
          <dd id="related"></dd>
        </div>
      </dl>
      <div class="reveal-actions">
        <button class="btn mastered" id="btnMastered" type="button">标熟</button>
        <button class="btn next" id="btnNext" type="button">下一个 →</button>
      </div>
      <section class="neighbors" id="neighbors" aria-label="相邻词"></section>
    </div>
  </section>

  <footer class="footer">
    <span id="overallStats"></span>
    <div class="footer-actions">
      <button class="link" id="masteredBtn" type="button">熟词 (0)</button>
      <button class="link" id="reset" type="button">重置统计</button>
    </div>
  </footer>
</main>

<div class="sheet" id="masteredSheet" hidden>
  <div class="sheet-panel" role="dialog" aria-modal="true" aria-labelledby="masteredTitle">
    <header class="sheet-head">
      <h2 id="masteredTitle">熟词列表 <span class="sheet-count" id="masteredSheetCount"></span></h2>
      <button class="icon-btn" id="closeMastered" type="button" aria-label="关闭">✕</button>
    </header>
    <div class="mastered-list" id="masteredList"></div>
  </div>
</div>
<div class="toast" id="toast" role="status" aria-live="polite" hidden></div>

<noscript><p style="text-align:center;padding:24px">需要启用 JavaScript 才能使用。</p></noscript>

<script>
(() => {
  'use strict';

  const WORDS = @@WORDS@@;
  const WORD_KEYS = new Set(WORDS.map((w) => w.seq));

  const STORAGE_KEY = 'gre-word-stats-v1';
  const MASTERED_KEY = 'gre-mastered-v1';
  const NEIGHBOR_COUNT = 10;

  const els = {
    todayStats: document.getElementById('todayStats'),
    prompt: document.getElementById('prompt'),
    word: document.getElementById('word'),
    wordMeta: document.getElementById('wordMeta'),
    resultTag: document.getElementById('resultTag'),
    countText: document.getElementById('countText'),
    actions: document.getElementById('actions'),
    btnKnown: document.getElementById('btnKnown'),
    btnUnknown: document.getElementById('btnUnknown'),
    details: document.getElementById('details'),
    meaning: document.getElementById('meaning'),
    synonyms: document.getElementById('synonyms'),
    related: document.getElementById('related'),
    synField: document.getElementById('synField'),
    relField: document.getElementById('relField'),
    btnNext: document.getElementById('btnNext'),
    btnMastered: document.getElementById('btnMastered'),
    neighbors: document.getElementById('neighbors'),
    overallStats: document.getElementById('overallStats'),
    masteredBtn: document.getElementById('masteredBtn'),
    masteredSheet: document.getElementById('masteredSheet'),
    masteredSheetCount: document.getElementById('masteredSheetCount'),
    masteredList: document.getElementById('masteredList'),
    closeMastered: document.getElementById('closeMastered'),
    toast: document.getElementById('toast'),
    reset: document.getElementById('reset'),
  };

  let stats = loadStats();
  let mastered = loadMastered();
  let currentIndex = -1;
  let revealed = false;
  let toastTimer = 0;
  let lastFocused = null;

  function todayKey() {
    const d = new Date();
    return d.getFullYear() + '-' +
      String(d.getMonth() + 1).padStart(2, '0') + '-' +
      String(d.getDate()).padStart(2, '0');
  }

  function loadStats() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      const parsed = raw ? JSON.parse(raw) : {};
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return {};
      return parsed;
    } catch (err) {
      return {};
    }
  }

  function saveStats() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(stats));
    } catch (err) {
      // Storage unavailable (private mode etc.); the app still works this session.
    }
  }

  function loadMastered() {
    try {
      const raw = localStorage.getItem(MASTERED_KEY);
      const parsed = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(parsed)) return new Set();
      return new Set(parsed.map((key) => String(key)).filter((key) => WORD_KEYS.has(key)));
    } catch (err) {
      return new Set();
    }
  }

  function saveMastered() {
    try {
      localStorage.setItem(MASTERED_KEY, JSON.stringify([...mastered]));
    } catch (err) {
      // Same as stats: keep working for this session.
    }
  }

  function getTodayTotals() {
    const day = todayKey();
    let seen = 0;
    let unknown = 0;
    for (const key in stats) {
      if (!WORD_KEYS.has(key)) continue;
      const s = stats[key];
      if (s && s.day === day) {
        seen += s.daySeen || 0;
        unknown += s.dayUnknown || 0;
      }
    }
    return { seen, unknown };
  }

  function getOverallTotals() {
    let learned = 0;
    let seen = 0;
    for (const key in stats) {
      if (!WORD_KEYS.has(key)) continue;
      const s = stats[key];
      if (s && s.seen > 0) {
        learned += 1;
        seen += s.seen;
      }
    }
    return { learned, seen };
  }

  function updateHeader() {
    const t = getTodayTotals();
    els.todayStats.textContent = '今日 ' + t.seen + ' 次 · 不认识 ' + t.unknown;
    const o = getOverallTotals();
    els.overallStats.textContent = '已学 ' + o.learned + '/' + WORDS.length + ' · 累计 ' + o.seen + ' 次';
    els.masteredBtn.textContent = '熟词 (' + mastered.size + ')';
  }

  function weightOf(s, day) {
    if (!s) return 1;
    const cumulativeRate = s.seen ? s.unknown / s.seen : 0;
    const todayRate = (s.day === day && s.daySeen) ? s.dayUnknown / s.daySeen : 0;
    return 1 + 3 * cumulativeRate + 5 * todayRate;
  }

  function pickIndex(exclude) {
    const day = todayKey();
    const weights = new Array(WORDS.length);
    let total = 0;
    for (let i = 0; i < WORDS.length; i += 1) {
      const skipped = i === exclude || mastered.has(WORDS[i].seq);
      const weight = skipped ? 0 : weightOf(stats[WORDS[i].seq], day);
      weights[i] = weight;
      total += weight;
    }
    if (total <= 0) return -1;
    let r = Math.random() * total;
    for (let i = 0; i < weights.length; i += 1) {
      if (weights[i] <= 0) continue;
      r -= weights[i];
      if (r <= 0) return i;
    }
    for (let i = weights.length - 1; i >= 0; i -= 1) {
      if (weights[i] > 0) return i;
    }
    return -1;
  }

  function record(word, known) {
    const day = todayKey();
    let s = stats[word.seq];
    if (!s) s = { seen: 0, unknown: 0, day: '', daySeen: 0, dayUnknown: 0 };
    if (s.day !== day) {
      s.day = day;
      s.daySeen = 0;
      s.dayUnknown = 0;
    }
    s.seen += 1;
    s.daySeen += 1;
    if (!known) {
      s.unknown += 1;
      s.dayUnknown += 1;
    }
    stats[word.seq] = s;
    saveStats();
    return s;
  }

  function renderNeighbors(index) {
    els.neighbors.replaceChildren();
    const groups = [
      { label: '前', start: Math.max(0, index - NEIGHBOR_COUNT), end: index },
      { label: '后', start: index + 1, end: Math.min(WORDS.length, index + 1 + NEIGHBOR_COUNT) },
    ];
    for (const group of groups) {
      if (group.start >= group.end) continue;
      const box = document.createElement('div');
      box.className = 'neighbor-group';
      const label = document.createElement('div');
      label.className = 'neighbor-label';
      label.textContent = group.label + ' ' + (group.end - group.start) + ' 个词';
      const list = document.createElement('div');
      list.className = 'chips';
      for (let i = group.start; i < group.end; i += 1) {
        const chip = document.createElement('div');
        chip.className = 'chip';
        const chipWord = document.createElement('span');
        chipWord.className = 'chip-word';
        chipWord.textContent = WORDS[i].word;
        const chipMeaning = document.createElement('span');
        chipMeaning.className = 'chip-meaning';
        chipMeaning.textContent = WORDS[i].meaning;
        chip.append(chipWord, chipMeaning);
        list.appendChild(chip);
      }
      box.append(label, list);
      els.neighbors.appendChild(box);
    }
  }

  function renderEmpty() {
    currentIndex = -1;
    revealed = false;
    els.word.textContent = '🎉';
    els.word.classList.remove('long');
    els.prompt.textContent = '所有单词都已标熟';
    els.prompt.hidden = false;
    els.wordMeta.hidden = true;
    els.actions.hidden = true;
    els.details.hidden = true;
    els.neighbors.replaceChildren();
    window.scrollTo({ top: 0, behavior: 'auto' });
  }

  function renderWord(index) {
    if (index < 0) {
      renderEmpty();
      return;
    }
    currentIndex = index;
    revealed = false;
    const word = WORDS[index];

    els.word.textContent = word.word;
    els.word.classList.toggle('long', word.word.length > 13);
    els.prompt.textContent = '认识这个单词吗？';
    els.prompt.hidden = false;
    els.wordMeta.hidden = true;
    els.actions.hidden = false;
    els.details.hidden = true;

    els.meaning.textContent = word.meaning;
    els.synonyms.textContent = word.synonyms;
    els.related.textContent = word.related;
    els.synField.hidden = !word.synonyms;
    els.relField.hidden = !word.related;
    els.neighbors.replaceChildren();

    window.scrollTo({ top: 0, behavior: 'auto' });
  }

  function reveal(known) {
    if (revealed) return;
    revealed = true;
    const word = WORDS[currentIndex];
    const s = record(word, known);

    els.resultTag.textContent = known ? '认识' : '不认识';
    els.resultTag.className = 'tag ' + (known ? 'known' : 'unknown');
    const percent = s.seen ? Math.round((s.unknown / s.seen) * 100) : 0;
    els.countText.textContent =
      '累计 ' + s.unknown + '/' + s.seen + ' 不认识 (' + percent + '%) · 今日 ' +
      s.dayUnknown + '/' + s.daySeen;

    els.prompt.hidden = true;
    els.actions.hidden = true;
    els.wordMeta.hidden = false;
    renderNeighbors(currentIndex);
    els.details.hidden = false;
    updateHeader();
  }

  function showToast(message) {
    els.toast.textContent = message;
    els.toast.hidden = false;
    window.clearTimeout(toastTimer);
    toastTimer = window.setTimeout(() => {
      els.toast.hidden = true;
    }, 1400);
  }

  function markMastered() {
    if (currentIndex < 0) return;
    const word = WORDS[currentIndex];
    mastered.add(word.seq);
    saveMastered();
    updateHeader();
    showToast('已标熟：' + word.word);
    renderWord(pickIndex(-1));
  }

  function renderMasteredList() {
    els.masteredList.replaceChildren();
    const list = WORDS.filter((word) => mastered.has(word.seq));
    els.masteredSheetCount.textContent = '(' + list.length + ')';

    if (list.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'mastered-empty';
      empty.textContent = '还没有标熟的单词';
      els.masteredList.appendChild(empty);
      return;
    }

    for (const word of list) {
      const item = document.createElement('div');
      item.className = 'mastered-item';

      const text = document.createElement('div');
      text.className = 'mastered-text';
      const wordEl = document.createElement('span');
      wordEl.className = 'mastered-word';
      wordEl.textContent = word.word;
      const meaningEl = document.createElement('span');
      meaningEl.className = 'mastered-meaning';
      meaningEl.textContent = word.meaning;
      text.append(wordEl, meaningEl);

      const restore = document.createElement('button');
      restore.className = 'restore-btn';
      restore.type = 'button';
      restore.textContent = '恢复';
      restore.addEventListener('click', () => {
        mastered.delete(word.seq);
        saveMastered();
        updateHeader();
        renderMasteredList();
        showToast('已恢复：' + word.word);
        if (currentIndex < 0) renderWord(pickIndex(-1));
      });

      item.append(text, restore);
      els.masteredList.appendChild(item);
    }
  }

  function openMasteredList() {
    lastFocused = document.activeElement;
    renderMasteredList();
    els.masteredSheet.hidden = false;
    document.body.style.overflow = 'hidden';
    els.closeMastered.focus();
  }

  function closeMasteredList() {
    els.masteredSheet.hidden = true;
    document.body.style.overflow = '';
    if (lastFocused && typeof lastFocused.focus === 'function') lastFocused.focus();
    lastFocused = null;
  }

  els.btnKnown.addEventListener('click', () => reveal(true));
  els.btnUnknown.addEventListener('click', () => reveal(false));
  els.btnMastered.addEventListener('click', markMastered);
  els.btnNext.addEventListener('click', () => renderWord(pickIndex(currentIndex)));
  els.masteredBtn.addEventListener('click', openMasteredList);
  els.closeMastered.addEventListener('click', closeMasteredList);
  els.masteredSheet.addEventListener('click', (event) => {
    if (event.target === els.masteredSheet) closeMasteredList();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !els.masteredSheet.hidden) closeMasteredList();
  });

  els.reset.addEventListener('click', () => {
    if (!window.confirm('确定要清空所有学习统计吗？熟词列表不会被清除。')) return;
    stats = {};
    saveStats();
    window.location.reload();
  });

  renderWord(pickIndex(-1));
  updateHeader();
})();
</script>
</body>
</html>
"""


def load_words():
    wb = openpyxl.load_workbook(WORDS_PATH, data_only=True)
    ws = wb[SHEET]
    words = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        seq, word, meaning, synonyms, related = row[:5]
        if seq is None or word is None:
            continue
        words.append({
            "seq": str(seq).strip(),
            "word": str(word).strip(),
            "meaning": "" if meaning is None else str(meaning).strip(),
            "synonyms": "" if synonyms is None else str(synonyms).strip(),
            "related": "" if related is None else str(related).strip(),
        })
    return words


def main():
    words = load_words()
    data = json.dumps(words, ensure_ascii=False, separators=(",", ":"))
    data = data.replace("<", "\\u003c")
    data = data.replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    html = TEMPLATE.replace("@@WORDS@@", data)
    with io.open(OUTPUT, "w", encoding="utf-8", newline="\n") as f:
        f.write(html)
    print("wrote {} ({} words, {:,} bytes)".format(OUTPUT, len(words), len(html.encode("utf-8"))))


if __name__ == "__main__":
    main()
