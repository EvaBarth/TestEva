/* ===== NAVIGATION ===== */
function showScreen(id) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));

  const screen = document.getElementById('screen-' + id);
  if (screen) screen.classList.add('active');

  const navLink = document.getElementById('nav-' + id);
  if (navLink) navLink.classList.add('active');

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ===== TRANSLATE MODE TABS ===== */
function setMode(mode, btn) {
  document.querySelectorAll('.mode-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.mode-panel').forEach(p => p.classList.add('hidden'));
  btn.classList.add('active');
  document.getElementById('mode-' + mode).classList.remove('hidden');
}

/* ===== TRANSLATE LOGIC (simulated) ===== */
const DEMO_TRANSLATIONS = {
  de: {
    en: (t) => t
      .replace(/Guten Morgen/g, 'Good morning')
      .replace(/Hallo/g, 'Hello')
      .replace(/Danke/g, 'Thank you')
      .replace(/Bitte/g, 'Please')
      .replace(/Wie geht es Ihnen/g, 'How are you')
      .replace(/Auf Wiedersehen/g, 'Goodbye')
      || `[EN] ${t}`,
    fr: (t) => `[FR] ${t}`,
    es: (t) => `[ES] ${t}`,
  },
  en: {
    de: (t) => t
      .replace(/Good morning/g, 'Guten Morgen')
      .replace(/Hello/g, 'Hallo')
      .replace(/Thank you/g, 'Danke')
      .replace(/Please/g, 'Bitte')
      .replace(/How are you/g, 'Wie geht es Ihnen?')
      .replace(/Goodbye/g, 'Auf Wiedersehen')
      || `[DE] ${t}`,
    fr: (t) => `[FR] ${t}`,
    es: (t) => `[ES] ${t}`,
  }
};

const LANG_NAMES = {
  auto: 'Automatisch', de: 'Deutsch', en: 'Englisch', fr: 'Französisch',
  es: 'Spanisch', it: 'Italienisch', pt: 'Portugiesisch', nl: 'Niederländisch',
  pl: 'Polnisch', ru: 'Russisch', zh: 'Chinesisch', ja: 'Japanisch', ar: 'Arabisch'
};

function translateText() {
  const sourceText = document.getElementById('source-text').value.trim();
  if (!sourceText) { showToast('Bitte Text eingeben.'); return; }

  const resultEl = document.getElementById('result-text');
  resultEl.className = 'result-text loading-dots';
  resultEl.textContent = 'Übersetze';

  const srcLang = document.getElementById('src-lang').value;
  const tgtLang = document.getElementById('tgt-lang').value;

  setTimeout(() => {
    let translated;
    const effectiveSrc = srcLang === 'auto' ? detectLanguage(sourceText) : srcLang;

    if (DEMO_TRANSLATIONS[effectiveSrc] && DEMO_TRANSLATIONS[effectiveSrc][tgtLang]) {
      translated = DEMO_TRANSLATIONS[effectiveSrc][tgtLang](sourceText);
    } else {
      translated = `[${(LANG_NAMES[tgtLang] || tgtLang).toUpperCase()}] ${sourceText}`;
    }

    resultEl.className = 'result-text';
    resultEl.textContent = translated;

    if (srcLang === 'auto') {
      document.getElementById('detect-label').textContent =
        `Erkannt: ${LANG_NAMES[effectiveSrc] || effectiveSrc}`;
    }

    addToHistory({
      src: LANG_NAMES[effectiveSrc] || effectiveSrc,
      tgt: LANG_NAMES[tgtLang] || tgtLang,
      source: sourceText,
      result: translated,
      type: 'text'
    });
  }, 900);
}

function detectLanguage(text) {
  const germanWords = /\b(ich|du|er|sie|wir|und|der|die|das|ist|nicht|mit|von|zu|in|für)\b/i;
  return germanWords.test(text) ? 'de' : 'en';
}

function clearSource() {
  document.getElementById('source-text').value = '';
  document.getElementById('result-text').className = 'result-text placeholder';
  document.getElementById('result-text').textContent = 'Übersetzung erscheint hier...';
  document.getElementById('char-count').textContent = '0 / 5000 Zeichen';
  document.getElementById('detect-label').textContent = '';
}

function updateCharCount() {
  const len = document.getElementById('source-text').value.length;
  document.getElementById('char-count').textContent = `${len} / 5000 Zeichen`;
}

function swapLanguages() {
  const src = document.getElementById('src-lang');
  const tgt = document.getElementById('tgt-lang');
  if (src.value === 'auto') {
    showToast('Automatische Erkennung kann nicht getauscht werden.');
    return;
  }
  const tmp = src.value;
  src.value = tgt.value;
  tgt.value = tmp;

  const srcText = document.getElementById('source-text').value;
  const resText = document.getElementById('result-text').textContent;
  if (resText && !document.getElementById('result-text').classList.contains('placeholder')) {
    document.getElementById('source-text').value = resText;
    document.getElementById('result-text').className = 'result-text placeholder';
    document.getElementById('result-text').textContent = 'Übersetzung erscheint hier...';
    updateCharCount();
  }
}

function copyResult() {
  const text = document.getElementById('result-text').textContent;
  if (document.getElementById('result-text').classList.contains('placeholder')) {
    showToast('Noch keine Übersetzung vorhanden.');
    return;
  }
  navigator.clipboard?.writeText(text).catch(() => {});
  showToast('Übersetzung kopiert!');
}

function speakResult() {
  const text = document.getElementById('result-text').textContent;
  if (document.getElementById('result-text').classList.contains('placeholder')) {
    showToast('Noch keine Übersetzung vorhanden.');
    return;
  }
  if ('speechSynthesis' in window) {
    const utt = new SpeechSynthesisUtterance(text);
    utt.lang = document.getElementById('tgt-lang').value;
    speechSynthesis.speak(utt);
    showToast('Wird vorgelesen...');
  } else {
    showToast('Vorlesen nicht unterstützt.');
  }
}

/* ===== HISTORY ===== */
let historyItems = [
  {
    id: 1, src: 'Deutsch', tgt: 'Englisch', type: 'text', date: '31.03.2026, 09:15',
    source: 'Guten Morgen, wie geht es Ihnen?',
    result: 'Good morning, how are you?'
  },
  {
    id: 2, src: 'Englisch', tgt: 'Deutsch', type: 'file', date: '30.03.2026, 14:40',
    source: 'Vertrag_Q1_2026.pdf',
    result: 'Vertrag_Q1_2026_DE.pdf'
  },
  {
    id: 3, src: 'Deutsch', tgt: 'Französisch', type: 'human', date: '28.03.2026, 11:00',
    source: 'Bitte übersetzen Sie dieses juristische Dokument...',
    result: 'Veuillez traduire ce document juridique...'
  },
  {
    id: 4, src: 'Englisch', tgt: 'Spanisch', type: 'text', date: '27.03.2026, 16:22',
    source: 'Thank you for your order. We will process it shortly.',
    result: 'Gracias por su pedido. Lo procesaremos en breve.'
  }
];

let historyFilter = 'all';
let historySearch = '';

function renderHistory() {
  const list = document.getElementById('history-list');
  const items = historyItems.filter(item => {
    const matchFilter = historyFilter === 'all' || item.type === historyFilter;
    const search = historySearch.toLowerCase();
    const matchSearch = !search ||
      item.source.toLowerCase().includes(search) ||
      item.result.toLowerCase().includes(search) ||
      item.src.toLowerCase().includes(search) ||
      item.tgt.toLowerCase().includes(search);
    return matchFilter && matchSearch;
  });

  if (items.length === 0) {
    list.innerHTML = '<p style="color:var(--text-muted);text-align:center;padding:2rem;">Keine Einträge gefunden.</p>';
    return;
  }

  list.innerHTML = items.map(item => `
    <div class="history-item" onclick="openHistoryItem(${item.id})">
      <div class="history-item-header">
        <div class="history-langs">
          <span class="lang-tag">${item.src}</span>
          <span>→</span>
          <span class="lang-tag">${item.tgt}</span>
          <span class="type-badge type-${item.type}">${typeLabel(item.type)}</span>
        </div>
        <div class="history-meta">
          <span>${item.date}</span>
        </div>
      </div>
      <div class="history-preview">${item.source}</div>
    </div>
  `).join('');
}

function typeLabel(type) {
  return { text: 'Text', file: 'Datei', human: 'Mensch' }[type] || type;
}

function setHistoryFilter(filter, btn) {
  historyFilter = filter;
  document.querySelectorAll('.filter-chips .chip').forEach(c => c.classList.remove('active'));
  btn.classList.add('active');
  renderHistory();
}

function filterHistory(val) {
  historySearch = val;
  renderHistory();
}

function openHistoryItem(id) {
  const item = historyItems.find(i => i.id === id);
  if (!item) return;
  document.getElementById('source-text').value = item.source;
  updateCharCount();
  const resultEl = document.getElementById('result-text');
  resultEl.className = 'result-text';
  resultEl.textContent = item.result;
  showScreen('translate');
  showToast('Übersetzung wiederhergestellt');
}

function addToHistory(data) {
  const now = new Date();
  const date = now.toLocaleDateString('de-DE') + ', ' + now.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  historyItems.unshift({
    id: Date.now(),
    src: data.src,
    tgt: data.tgt,
    type: data.type,
    date,
    source: data.source,
    result: data.result
  });
}

function saveToHistory() {
  showToast('Im Verlauf gespeichert');
}

/* ===== MODALS ===== */
function showModal(id) {
  document.getElementById(id).classList.remove('hidden');
}

function closeModal(id) {
  document.getElementById(id).classList.add('hidden');
}

function closeModalOutside(event, id) {
  if (event.target === document.getElementById(id)) closeModal(id);
}

function submitOrder() {
  closeModal('modal-human');
  showToast('Auftrag erfolgreich angefragt! Sie erhalten eine E-Mail-Bestätigung.');
}

function startFileTranslation() {
  closeModal('modal-file');
  showToast('Dokument wird übersetzt... Sie werden benachrichtigt.');
}

function removeFile(btn) {
  btn.closest('.file-chip').remove();
}

/* ===== DARK MODE ===== */
function toggleDarkMode() {
  document.body.classList.toggle('dark');
  const isDark = document.body.classList.contains('dark');
  showToast(isDark ? 'Dunkelmodus aktiviert' : 'Hellmodus aktiviert');
}

/* ===== TOAST ===== */
let toastTimer = null;

function showToast(msg) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.classList.remove('hidden');
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.add('hidden'), 3000);
}

/* ===== KEYBOARD SHORTCUT ===== */
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
    if (document.getElementById('screen-translate').classList.contains('active')) {
      translateText();
    }
  }
  if (e.key === 'Escape') {
    document.querySelectorAll('.modal-overlay:not(.hidden)').forEach(m => m.classList.add('hidden'));
  }
});

/* ===== INIT ===== */
renderHistory();
