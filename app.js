// Names to pick from. Can be set via dialog and saved to localStorage.
const DEFAULT_NAMES = [
  "Alice",
  "Bob",
  "Charlie",
  "Diana",
  "Eve",
  "Frank",
  "Grace",
  "Heidi",
  "Ivan",
  "Judy",
];

const STORAGE_KEY = "flightBoardNames";
const WINNER_KEY = "flightBoardFixedWinner";
const RANKING_KEY = "flightBoardRanking";

let names = loadNames();
let fixedWinner = loadFixedWinner();
let pendingConsumeFixedWinner = false;
let ranking = loadRanking();

// Compute a fixed board width using the longest name to avoid layout shifts
let maxLen = computeMaxLen();

const board = document.getElementById("board");
const startBtn = document.getElementById("startBtn");
const setNamesBtn = document.getElementById("setNamesBtn");
const namesDialog = document.getElementById("namesDialog");
const namesTextarea = document.getElementById("namesTextarea");
const saveNamesBtn = document.getElementById("saveNamesBtn");
const cancelNamesBtn = document.getElementById("cancelNamesBtn");
const setWinnerBtn = document.getElementById("setWinnerBtn");
const winnerDialog = document.getElementById("winnerDialog");
const winnerGrid = document.getElementById("winnerGrid");
const cancelWinnerBtn = document.getElementById("cancelWinnerBtn");
const clearWinnerBtn = document.getElementById("clearWinnerBtn");
const winnerStatus = document.getElementById("winnerStatus");

function computeMaxLen() {
  return Math.max(5, ...names.map((n) => n.length));
}

function loadNames() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const arr = JSON.parse(raw);
      if (Array.isArray(arr) && arr.every((s) => typeof s === "string")) {
        return arr;
      }
    }
  } catch {}
  return DEFAULT_NAMES.slice();
}

function saveNamesToStorage() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(names));
  } catch {}
}

function loadFixedWinner() {
  try {
    const s = localStorage.getItem(WINNER_KEY);
    if (typeof s === "string" && s.trim().length) return s;
  } catch {}
  return null;
}

function saveFixedWinner() {
  try {
    if (fixedWinner && fixedWinner.trim()) {
      localStorage.setItem(WINNER_KEY, fixedWinner);
    } else {
      localStorage.removeItem(WINNER_KEY);
    }
  } catch {}
}

function loadRanking() {
  try {
    const raw = localStorage.getItem(RANKING_KEY);
    if (raw) {
      const arr = JSON.parse(raw);
      if (Array.isArray(arr)) {
        let migrated = false;
        const out = [];
        for (const item of arr) {
          if (typeof item === "string") {
            migrated = true;
            out.push({ name: item, ts: Date.now() });
          } else if (item && typeof item.name === "string") {
            out.push({ name: item.name, ts: typeof item.ts === "number" ? item.ts : Date.now() });
          }
        }
        if (migrated) {
          localStorage.setItem(RANKING_KEY, JSON.stringify(out));
        }
        return out;
      }
    }
  } catch {}
  return [];
}

function saveRanking() {
  try {
    localStorage.setItem(RANKING_KEY, JSON.stringify(ranking));
  } catch {}
}

function addRankingEntry(name) {
  ranking.push({ name, ts: Date.now() });
  saveRanking();
  renderWeeklyList();
}

// CSV download removed per request; ranking persists only in localStorage

// Audio setup (Web Audio API)
let audioCtx;
function ensureAudio() {
  if (!audioCtx) {
    const Ctx = window.AudioContext || window.webkitAudioContext;
    if (Ctx) audioCtx = new Ctx();
  }
  if (audioCtx && audioCtx.state === "suspended") {
    audioCtx.resume().catch(() => {});
  }
}

function playTone(freq, startTime, duration, type = "sine", gainDb = -6) {
  if (!audioCtx) return;
  const osc = audioCtx.createOscillator();
  const gain = audioCtx.createGain();
  const now = audioCtx.currentTime;
  const t0 = now + startTime;
  const t1 = t0 + duration;
  const linearGain = Math.pow(10, gainDb / 20);

  osc.type = type;
  osc.frequency.setValueAtTime(freq, t0);

  gain.gain.setValueAtTime(0.0001, t0);
  gain.gain.exponentialRampToValueAtTime(linearGain, t0 + 0.01);
  gain.gain.exponentialRampToValueAtTime(linearGain * 0.6, t0 + Math.max(0.02, duration * 0.3));
  gain.gain.exponentialRampToValueAtTime(0.0001, t1);

  osc.connect(gain).connect(audioCtx.destination);
  osc.start(t0);
  osc.stop(t1 + 0.02);
}

function playCongrats() {
  ensureAudio();
  if (!audioCtx) return;

  // Quick uplifting arpeggio + short chord (C major-ish)
  const base = audioCtx.currentTime + 0.05;
  const notes = [523.25, 659.25, 783.99]; // C5, E5, G5

  // Arpeggio
  notes.forEach((f, i) => {
    playTone(f, 0.18 * i, 0.16, "triangle", -8);
  });

  // Chord (stacked, slightly quieter)
  notes.forEach((f, i) => {
    playTone(f, 0.6, 0.55, i === 1 ? "square" : "sine", -12);
  });
}

let stopShuffleNoise = null;
let shuffleTickTimer = null;

function startShuffleSoundEffect() {
  ensureAudio();
  if (!audioCtx) return;

  stopShuffleSoundEffect();

  const bufferLength = Math.max(1, Math.floor(audioCtx.sampleRate * 0.4));
  const noiseBuffer = audioCtx.createBuffer(1, bufferLength, audioCtx.sampleRate);
  const data = noiseBuffer.getChannelData(0);
  for (let i = 0; i < data.length; i++) {
    data[i] = (Math.random() * 2 - 1) * (1 - i / data.length);
  }

  const noiseSource = audioCtx.createBufferSource();
  noiseSource.buffer = noiseBuffer;
  noiseSource.loop = true;

  const filter = audioCtx.createBiquadFilter();
  filter.type = "bandpass";
  filter.frequency.value = 950;
  filter.Q.value = 1.4;

  const gain = audioCtx.createGain();
  const now = audioCtx.currentTime;
  gain.gain.setValueAtTime(0.0001, now);
  gain.gain.exponentialRampToValueAtTime(0.16, now + 0.08);

  noiseSource.connect(filter).connect(gain).connect(audioCtx.destination);
  noiseSource.start(now);

  shuffleTickTimer = setInterval(() => {
    const baseFreq = 420 + Math.random() * 120;
    playTone(baseFreq, 0, 0.045, "square", -18);
  }, 95);

  const localStop = () => {
    const stopTime = audioCtx.currentTime;
    gain.gain.cancelScheduledValues(stopTime);
    const current = Math.max(0.0001, gain.gain.value);
    gain.gain.setValueAtTime(current, stopTime);
    gain.gain.exponentialRampToValueAtTime(0.0001, stopTime + 0.2);
    noiseSource.stop(stopTime + 0.24);
  };

  noiseSource.onended = () => {
    if (stopShuffleNoise === localStop) {
      stopShuffleNoise = null;
    }
  };

  stopShuffleNoise = localStop;
}

function stopShuffleSoundEffect() {
  if (shuffleTickTimer) {
    clearInterval(shuffleTickTimer);
    shuffleTickTimer = null;
  }
  if (stopShuffleNoise) {
    stopShuffleNoise();
    stopShuffleNoise = null;
  }
}

// Random character set similar to split-flap boards
const CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ";
function randChar() {
  return CHARS[(Math.random() * CHARS.length) | 0];
}

function padToMax(s) {
  // Center the string within maxLen using spaces
  const totalPad = Math.max(0, maxLen - s.length);
  const left = Math.floor(totalPad / 2);
  const right = totalPad - left;
  return " ".repeat(left) + s + " ".repeat(right);
}

function randomRow(len = maxLen) {
  let out = "";
  for (let i = 0; i < len; i++) out += randChar();
  return out;
}

let isBusy = false;

function startShuffleAndPick({ shuffleMs = 5000 } = {}) {
  if (isBusy) return;
  isBusy = true;
  startBtn.disabled = true;
  ensureAudio();
  startShuffleSoundEffect();

  // Initial board content fixed width
  board.textContent = randomRow();

  // Rapid shuffle loop
  const tickMs = 30;
  const interval = setInterval(() => {
    board.textContent = randomRow();
  }, tickMs);

  setTimeout(() => {
    clearInterval(interval);
    stopShuffleSoundEffect();
    if (!names.length) {
      board.textContent = padToMax("NO NAMES");
      isBusy = false;
      startBtn.disabled = false;
      return;
    }
    let winner;
    pendingConsumeFixedWinner = false;
    if (fixedWinner && names.some((n) => n.toUpperCase() === fixedWinner.toUpperCase())) {
      winner = fixedWinner;
      pendingConsumeFixedWinner = true;
    } else {
      winner = names[(Math.random() * names.length) | 0];
    }
    revealWinner(winner);
  }, shuffleMs);
}

function revealWinner(name) {
  const target = padToMax(name.toUpperCase());
  // Start from whatever is on screen now, or spaces
  let current = (board.textContent || "").padEnd(maxLen, " ").slice(0, maxLen).split("");

  // For each position, flip a few random chars before locking the final char
  const flipsPerChar = 7;
  const stepDelay = 18; // ms between flips within a slot
  const cascadeDelay = 55; // ms between starting each slot

  for (let i = 0; i < maxLen; i++) {
    const finalChar = target[i];
    const startAt = i * cascadeDelay;

    for (let f = 0; f < flipsPerChar; f++) {
      setTimeout(() => {
        current[i] = randChar();
        board.textContent = current.join("");
      }, startAt + f * stepDelay);
    }

    // Lock final character
    setTimeout(() => {
      current[i] = finalChar;
      board.textContent = current.join("");
      // When last char set, play sound and re-enable
      if (i === maxLen - 1) {
        playCongrats();
        // Record winner and update/download ranking CSV
        addRankingEntry(name);
        if (pendingConsumeFixedWinner) {
          fixedWinner = null;
          saveFixedWinner();
          updateWinnerStatus();
        }
        setTimeout(() => {
          isBusy = false;
          startBtn.disabled = false;
        }, 50);
      }
    }, startAt + flipsPerChar * stepDelay + 12);
  }
}

startBtn.addEventListener("click", () => {
  const ms = Math.round((shuffleSeconds || 5) * 1000);
  startShuffleAndPick({ shuffleMs: ms });
});

// Ensure stable width right away
function updateStartEnabled() {
  startBtn.disabled = isBusy || names.length === 0;
}

function openNamesDialog() {
  // Pre-fill with current names
  namesTextarea.value = names.join("\n");
  if (typeof namesDialog.showModal === "function") {
    namesDialog.showModal();
  } else {
    namesDialog.setAttribute("open", "");
  }
  setTimeout(() => namesTextarea.focus(), 0);
}

function closeNamesDialog() {
  if (typeof namesDialog.close === "function") {
    namesDialog.close();
  } else {
    namesDialog.removeAttribute("open");
  }
}

function parseNames(text) {
  const list = text
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  // Deduplicate while keeping order
  const seen = new Set();
  const out = [];
  for (const n of list) {
    const key = n.toUpperCase();
    if (!seen.has(key)) {
      seen.add(key);
      out.push(n);
    }
  }
  return out;
}

function applyNames(newNames) {
  names = newNames.slice();
  maxLen = computeMaxLen();
  saveNamesToStorage();
  board.textContent = padToMax("READY");
  updateStartEnabled();
  // If the fixed winner no longer exists in the list, clear it
  if (fixedWinner && !names.some((n) => n.toUpperCase() === fixedWinner.toUpperCase())) {
    fixedWinner = null;
    saveFixedWinner();
  }
  updateWinnerStatus();
}

setNamesBtn.addEventListener("click", openNamesDialog);
cancelNamesBtn.addEventListener("click", closeNamesDialog);
saveNamesBtn.addEventListener("click", () => {
  const parsed = parseNames(namesTextarea.value);
  if (!parsed.length) {
    alert("Please enter at least one name.");
    return;
  }
  applyNames(parsed);
  closeNamesDialog();
});

function openWinnerDialog() {
  // Render grid of name buttons
  winnerGrid.innerHTML = "";
  if (!names.length) {
    const p = document.createElement("p");
    p.textContent = "No names available. Please set names first.";
    p.style.color = "var(--subtle)";
    winnerGrid.appendChild(p);
  } else {
    names.forEach((name) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "name-btn";
      btn.textContent = name;
      btn.addEventListener("click", () => {
        fixedWinner = name;
        saveFixedWinner();
        updateWinnerStatus();
        closeWinnerDialog();
      });
      winnerGrid.appendChild(btn);
    });
  }
  if (typeof winnerDialog.showModal === "function") {
    winnerDialog.showModal();
  } else {
    winnerDialog.setAttribute("open", "");
  }
}

function closeWinnerDialog() {
  if (typeof winnerDialog.close === "function") {
    winnerDialog.close();
  } else {
    winnerDialog.removeAttribute("open");
  }
}

function updateWinnerStatus() {
  if (fixedWinner) {
    winnerStatus.innerHTML = `<span class="winner-badge">Preset winner: <strong>${escapeHtml(fixedWinner)}</strong></span>`;
    winnerStatus.style.display = "";
  } else {
    winnerStatus.textContent = "";
    winnerStatus.style.display = "none";
  }
}

function escapeHtml(s) {
  return s.replace(/[&<>"']/g, (c) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  })[c]);
}

setWinnerBtn.addEventListener("click", openWinnerDialog);
cancelWinnerBtn.addEventListener("click", closeWinnerDialog);
clearWinnerBtn.addEventListener("click", () => {
  fixedWinner = null;
  saveFixedWinner();
  updateWinnerStatus();
  closeWinnerDialog();
});

board.textContent = padToMax("READY");
updateStartEnabled();
updateWinnerStatus();
renderWeeklyList();

// Weekly winners rendering as two-column table (Week | Name)
function getISOWeekInfo(ts) {
  const d = new Date(ts);
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7; // 1..7
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
  const year = date.getUTCFullYear();
  return { year, week, label: `W${week}` };
}

function buildWeekSequence() {
  // Determine sequential week numbers (W1, W2, ...) based on first occurrence of each ISO week
  const sorted = ranking
    .filter((e) => e && typeof e.name === "string" && typeof e.ts === "number")
    .slice()
    .sort((a, b) => a.ts - b.ts);
  const weekSeq = new Map(); // key -> seq
  let seq = 0;
  for (const e of sorted) {
    const { year, week } = getISOWeekInfo(e.ts);
    const key = `${year}-W${week}`;
    if (!weekSeq.has(key)) {
      seq += 1;
      weekSeq.set(key, seq);
    }
  }
  return weekSeq;
}

function getAllRowsWithWeekLabels() {
  // Sequential label per run: W1, W2, ... regardless of calendar week
  return ranking
    .filter((e) => e && typeof e.name === "string" && typeof e.ts === "number")
    .slice()
    .sort((a, b) => a.ts - b.ts)
    .map((e, idx) => ({ label: `W${idx + 1}`, name: e.name, ts: e.ts }));
}

function renderWeeklyList() {
  const listEl = document.getElementById("weeklyList");
  if (!listEl) return;
  const rows = getAllRowsWithWeekLabels();
  listEl.innerHTML = "";
  if (!rows.length) return;
  rows.forEach((r, idx) => {
    const li = document.createElement("li");
    li.className = `weekly-item ${idx === 0 ? "p1" : idx === 1 ? "p2" : idx === 2 ? "p3" : ""}`.trim();
    li.dataset.ts = String(r.ts);
    const weekEl = document.createElement("div");
    weekEl.className = "weekly-week";
    weekEl.textContent = r.label;
    const nameEl = document.createElement("div");
    nameEl.className = "weekly-name";
    nameEl.textContent = r.name;
    li.appendChild(weekEl);
    li.appendChild(nameEl);
    listEl.appendChild(li);
  });
}

// Shuffle time slider wiring (1–10 seconds)
const shuffleRange = document.getElementById("shuffleRange");
const shuffleValue = document.getElementById("shuffleValue");
let shuffleSeconds = 5;

const weeklyListEl = document.getElementById("weeklyList");
if (weeklyListEl) {
  weeklyListEl.addEventListener("click", (event) => {
    const item = event.target.closest("li.weekly-item");
    if (!item || !weeklyListEl.contains(item)) return;
    const ts = Number(item.dataset.ts || "");
    if (!Number.isFinite(ts)) return;
    const idx = ranking.findIndex((entry) => entry && typeof entry.ts === "number" && entry.ts === ts);
    if (idx === -1) return;
    const entry = ranking[idx];
    const name = entry && typeof entry.name === "string" ? entry.name : "this winner";
    const ok = confirm(`Remove ${name} from the winners list?`);
    if (!ok) return;
    ranking.splice(idx, 1);
    saveRanking();
    renderWeeklyList();
  });
}

function updateShuffleUI(val) {
  shuffleValue.textContent = `${val}s`;
}

if (shuffleRange) {
  shuffleSeconds = Number(shuffleRange.value) || 5;
  updateShuffleUI(shuffleSeconds);
  shuffleRange.addEventListener("input", () => {
    shuffleSeconds = Math.min(10, Math.max(1, Number(shuffleRange.value) || 5));
    updateShuffleUI(shuffleSeconds);
  });
}

// Clear weekly table (clears all ranking history)
const clearWeeklyBtn = document.getElementById("clearWeeklyBtn");
if (clearWeeklyBtn) {
  clearWeeklyBtn.addEventListener("click", () => {
    const ok = confirm("Clear all saved winners from the weekly list? This cannot be undone.");
    if (!ok) return;
    // Reset winners only
    ranking = [];
    saveRanking();
    renderWeeklyList();
  });
}
