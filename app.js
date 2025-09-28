// Names to pick from. Can be set via dialog and persisted to Firestore.
const DEFAULT_NAMES = [
  "Oscar",
  "Lando",
  "Max",
  "George",
  "Charles",
  "Lewis",
  "Kimi",
  "Alexander",
  "Isack",
  "Nico",
  "Lance",
  "Carlos",
  "Liam",
  "Fernando",
  "Esteban",
  "Pierre",
  "Yuki",
  "Gabriel",
  "Oliver",
  "Franco",
];

const GLOBAL_STATS_DOC_ID = "globalUsage";

let names = DEFAULT_NAMES.slice();
let fixedWinner = null;
let defaultFixedWinner = null;
let rankingStore = { default: [] };

// Compute a fixed board width using the longest name to avoid layout shifts
let maxLen = computeMaxLen();

const board = document.getElementById("board");
const startBtn = document.getElementById("startBtn");
if (startBtn) startBtn.disabled = true;
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
const loginBtn = document.getElementById("loginBtn");
const settingsBtn = document.getElementById("settingsBtn");
const settingsDialog = document.getElementById("settingsDialog");
const classListEl = document.getElementById("classList");
const classDetailEl = document.getElementById("classDetail");
const classSelector = document.getElementById("classSelector");
const classNameInput = document.getElementById("classNameInput");
const studentListTextarea = document.getElementById("studentListTextarea");
const addClassBtn = document.getElementById("addClassBtn");
const closeSettingsBtn = document.getElementById("closeSettingsBtn");
const classUsageNotice = document.getElementById("classUsageNotice");
const qrOverlay = document.getElementById("qrOverlay");
const qrOverlayCard = qrOverlay ? qrOverlay.querySelector(".qr-overlay-card") : null;

let firebaseAuth = null;
let firebaseGoogleProvider = null;
let firebaseDb = null;
let currentUserId = null;
let userClasses = [];
let classesLoading = false;
let classesError = null;
let classesFetchToken = 0;
let selectedClassId = null;
let editingClassId = null;
let totalClassCount = null;
let classCountFetchInFlight = false;
let previousFocus = null;
let qrOverlayVisible = false;
let classWinnerUnsubscribe = null;
let globalStatsUnsubscribe = null;
const latestClassUsageMetrics = {
  percentage: null,
  mvpCount: null,
  studentCount: null,
  classCount: null,
};

function getLatestGlobalStats() {
  if (typeof window === "undefined") return {};
  return window.__latestGlobalStats || {};
}

function setLatestGlobalStats(partial) {
  if (typeof window === "undefined") return;
  const currentStats = window.__latestGlobalStats || {};
  const nextStats = { ...currentStats };
  if (partial && typeof partial === "object") {
    for (const [key, value] of Object.entries(partial)) {
      if (value !== undefined && value !== null) {
        nextStats[key] = value;
      }
    }
  }
  window.__latestGlobalStats = nextStats;
  updateClassUsageNoticeDisplay();
}

function getSelectedClass() {
  if (!selectedClassId) return null;
  return userClasses.find((entry) => entry.id === selectedClassId) || null;
}

function getSelectedClassStudents() {
  const classEntry = getSelectedClass();
  if (!classEntry || !Array.isArray(classEntry.students)) return [];
  return classEntry.students.slice();
}

function getActiveNamePool() {
  const classStudents = getSelectedClassStudents();
  const usingClass = !!selectedClassId;
  const source = usingClass ? classStudents : names;
  return source
    .map((entry) => (typeof entry === "string" ? entry.trim() : ""))
    .filter((entry) => entry.length > 0);
}

function getDisplayWidth() {
  const active = getActiveNamePool();
  const longestActive = active.length ? Math.max(...active.map((s) => s.length)) : 0;
  return Math.max(5, maxLen, longestActive);
}

function reconcileFixedWinnerWithPool() {
  if (!fixedWinner) {
    updateWinnerStatus();
    updateStartEnabled();
    return;
  }
  const pool = getActiveNamePool();
  const exists = pool.some((n) => n.toUpperCase() === fixedWinner.toUpperCase());
  if (!exists) {
    updateActiveWinner(null, { persist: true, skipReconcile: true });
    return;
  }
  updateWinnerStatus();
}

function showQrOverlay() {
  if (!qrOverlay || qrOverlayVisible) return;
  previousFocus = document.activeElement && typeof document.activeElement.focus === "function" ? document.activeElement : null;
  qrOverlay.hidden = false;
  qrOverlay.setAttribute("aria-hidden", "false");
  qrOverlayVisible = true;
  if (typeof qrOverlay.focus === "function") {
    qrOverlay.focus();
  }
}

function hideQrOverlay() {
  if (!qrOverlay || !qrOverlayVisible) return;
  qrOverlay.setAttribute("aria-hidden", "true");
  qrOverlay.hidden = true;
  qrOverlayVisible = false;
  if (previousFocus && typeof previousFocus.focus === "function") {
    try {
      previousFocus.focus();
    } catch {}
  }
  previousFocus = null;
}

function handleGlobalKeydown(event) {
  if (event.defaultPrevented || event.repeat) return;
  if (event.ctrlKey || event.metaKey || event.altKey) return;
  const key = event.key;
  if (!key) return;
  if (key === "Escape") {
    if (qrOverlayVisible) {
      hideQrOverlay();
      event.stopPropagation();
      event.preventDefault();
    }
    return;
  }

  const target = event.target;
  const tag = target && target.tagName ? target.tagName.toUpperCase() : "";
  const isEditable = target && (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable);
  if (isEditable) return;

  if (key.toLowerCase() === "q") {
    showQrOverlay();
    event.preventDefault();
  }
}

if (qrOverlay) {
  qrOverlay.addEventListener("click", (event) => {
    if (event.target === qrOverlay) {
      hideQrOverlay();
    }
  });
  if (qrOverlayCard) {
    qrOverlayCard.addEventListener("click", (event) => {
      event.stopPropagation();
    });
  }
  document.addEventListener("keydown", handleGlobalKeydown);
}

function updateClassUsageNoticeDisplay() {
  if (!classUsageNotice) return;
  const hasGlobalCount = typeof totalClassCount === "number" && !Number.isNaN(totalClassCount);
  const count = hasGlobalCount ? totalClassCount : userClasses.length;
  const statsDoc = getLatestGlobalStats();
  const totalStudents = typeof statsDoc.studentCount === "number" ? statsDoc.studentCount : null;
  const totalMvps = typeof statsDoc.mvpCount === "number" ? statsDoc.mvpCount : null;
  const classCountFromStats = typeof statsDoc.classCount === "number" ? statsDoc.classCount : null;
  const classCount = hasGlobalCount ? totalClassCount : (classCountFromStats !== null ? classCountFromStats : count);

  const canShowAggregate = classCount !== null && classCount > 0 && totalStudents !== null && totalStudents > 0 && totalMvps !== null;
  const previousMetrics = { ...latestClassUsageMetrics };

  if (canShowAggregate) {
    const percentage = Math.min(100, Math.max(0, (totalMvps / totalStudents) * 100));
    latestClassUsageMetrics.percentage = percentage.toFixed(1);
    latestClassUsageMetrics.mvpCount = totalMvps;
    latestClassUsageMetrics.studentCount = totalStudents;
    latestClassUsageMetrics.classCount = classCount;
    classUsageNotice.innerHTML = `<span class="class-usage-number" data-highlight="percentage">${latestClassUsageMetrics.percentage}%</span> of students (<span class="class-usage-number" data-highlight="mvpCount">${latestClassUsageMetrics.mvpCount}</span> out of <span class="class-usage-number" data-highlight="studentCount">${latestClassUsageMetrics.studentCount}</span>) across <span class="class-usage-number" data-highlight="classCount">${latestClassUsageMetrics.classCount}</span> classes achieved MVP status.`;
  } else if (classCountFetchInFlight) {
    classUsageNotice.textContent = "Checking how many classes are using MVP...";
    latestClassUsageMetrics.percentage = null;
    latestClassUsageMetrics.mvpCount = null;
    latestClassUsageMetrics.studentCount = null;
    latestClassUsageMetrics.classCount = null;
  } else {
    classUsageNotice.textContent = "Be the first class to use MVP to make students exciting.";
    latestClassUsageMetrics.percentage = null;
    latestClassUsageMetrics.mvpCount = null;
    latestClassUsageMetrics.studentCount = null;
    latestClassUsageMetrics.classCount = null;
  }

  classUsageNotice.hidden = false;

  const highlightDurationMs = 500;
  if (canShowAggregate) {
    classUsageNotice.querySelectorAll(".class-usage-number").forEach((element) => {
      const key = element.getAttribute("data-highlight");
      if (!key) return;
      const previousValue = previousMetrics[key];
      const currentValue = latestClassUsageMetrics[key];
      if (previousValue === null || previousValue === undefined || String(previousValue) !== String(currentValue)) {
        element.classList.add("highlight");
        setTimeout(() => element.classList.remove("highlight"), highlightDurationMs);
      }
    });
  }
}

async function fetchGlobalClassCount() {
  if (!firebaseDb || classCountFetchInFlight) return;
  classCountFetchInFlight = true;
  updateClassUsageNoticeDisplay();
  try {
    const user = firebaseAuth && firebaseAuth.currentUser;
    if (!user) {
      const doc = await firebaseDb.collection("stats").doc(GLOBAL_STATS_DOC_ID).get();
      if (doc.exists) {
        const data = doc.data() || {};
        const count = data && typeof data.classCount === "number" ? data.classCount : null;
        totalClassCount = typeof count === "number" && !Number.isNaN(count) ? count : null;
        setLatestGlobalStats({
          classCount: typeof count === "number" ? count : undefined,
          studentCount: typeof data.studentCount === "number" ? data.studentCount : undefined,
          mvpCount: typeof data.mvpCount === "number" ? data.mvpCount : undefined,
        });
      } else {
        totalClassCount = null;
        setLatestGlobalStats({ classCount: undefined });
      }
    } else {
      const snapshot = await firebaseDb.collectionGroup("classes").get();
      totalClassCount = snapshot.size;
      try {
        await firebaseDb
          .collection("stats")
          .doc(GLOBAL_STATS_DOC_ID)
          .set(
            {
              classCount: totalClassCount,
              updatedAt: firebase.firestore && firebase.firestore.FieldValue && typeof firebase.firestore.FieldValue.serverTimestamp === "function"
                ? firebase.firestore.FieldValue.serverTimestamp()
                : Date.now(),
            },
            { merge: true }
          );
      } catch (writeErr) {
        console.warn("Failed to persist global class count", writeErr);
      }
      setLatestGlobalStats({ classCount: totalClassCount });
    }
  } catch (err) {
    console.error("Failed to load global class count", err);
  } finally {
    classCountFetchInFlight = false;
    updateClassUsageNoticeDisplay();
  }
}

async function adjustGlobalUsageCount(delta) {
  if (!firebaseDb || typeof delta !== "number" || !Number.isFinite(delta) || delta === 0) return;
  const statsRef = firebaseDb.collection("stats").doc(GLOBAL_STATS_DOC_ID);
  const fieldValue = firebase && firebase.firestore && firebase.firestore.FieldValue;
  const timestamp = fieldValue && typeof fieldValue.serverTimestamp === "function"
    ? fieldValue.serverTimestamp()
    : Date.now();
  let updatedClassCount = null;

  try {
    if (fieldValue && typeof fieldValue.increment === "function") {
      await statsRef.set(
        {
          classCount: fieldValue.increment(delta),
          updatedAt: timestamp,
        },
        { merge: true }
      );
      const currentStats = getLatestGlobalStats();
      const base = typeof currentStats.classCount === "number" ? currentStats.classCount : 0;
      updatedClassCount = Math.max(0, base + delta);
      return;
    }

    await firebaseDb.runTransaction(async (tx) => {
      const snapshot = await tx.get(statsRef);
      const data = snapshot.exists ? snapshot.data() || {} : {};
      const base = typeof data.classCount === "number" && !Number.isNaN(data.classCount) ? data.classCount : 0;
      const next = Math.max(0, base + delta);
      tx.set(
        statsRef,
        {
          classCount: next,
          updatedAt: Date.now(),
        },
        { merge: true }
      );
      updatedClassCount = next;
    });
  } catch (err) {
    console.warn("Failed to adjust global usage count", err);
    return;
  }

  if (updatedClassCount !== null) {
    setLatestGlobalStats({ classCount: updatedClassCount });
  }
}

async function updateGlobalStudentTotals() {
  if (!firebaseDb || !currentUserId) return;
  const classesCollection = getClassesCollection();
  if (!classesCollection) return;
  try {
    const snapshot = await classesCollection.get();
    let totalStudents = 0;
    snapshot.forEach((doc) => {
      const data = doc.data() || {};
      const students = Array.isArray(data.students) ? data.students : [];
      totalStudents += students.length;
    });
    await firebaseDb
      .collection("stats")
      .doc(GLOBAL_STATS_DOC_ID)
      .set(
        {
          studentCount: totalStudents,
        },
        { merge: true }
      );
    setLatestGlobalStats({ studentCount: totalStudents });
  } catch (err) {
    console.error("Failed to update global student count", err);
  }
}

function hasValidFirebaseConfig(config) {
  if (!config || typeof config !== "object") return false;
  const required = ["apiKey", "authDomain", "projectId", "appId"];
  for (const key of required) {
    const value = config[key];
    if (typeof value !== "string" || !value.trim() || value.includes("YOUR_")) {
      return false;
    }
  }
  return true;
}

function ensureSettingsVisibility() {
  if (!settingsBtn) return;
  const signedIn = !!(firebaseAuth && firebaseAuth.currentUser);
  if (!signedIn) {
    settingsBtn.hidden = true;
    settingsBtn.disabled = true;
    settingsBtn.title = "Sign in to manage classes";
    if (classSelector) {
      classSelector.hidden = true;
      classSelector.classList.add("is-empty");
      classSelector.innerHTML = "";
      classSelector.disabled = true;
    }
    totalClassCount = null;
    updateClassUsageNoticeDisplay();
    if (firebaseDb) {
      fetchGlobalClassCount();
    }
    return;
  }
  settingsBtn.hidden = false;
  const enabled = !!firebaseDb;
  settingsBtn.disabled = !enabled;
  settingsBtn.title = enabled ? "Open settings" : "Firestore unavailable";
  updateClassSelector();
  if (enabled) fetchGlobalClassCount();
}

function updateLoginButton() {
  if (!loginBtn) return;
  const user = firebaseAuth && firebaseAuth.currentUser;
  if (user) {
    const display = (user.displayName || user.email || "").trim();
    loginBtn.textContent = display ? `Sign out (${display})` : "Sign out";
    loginBtn.title = display ? `Signed in as ${display}` : "Signed in";
  } else {
    loginBtn.textContent = "Login with Google";
    loginBtn.title = "Login with Google";
  }
  ensureSettingsVisibility();
}

function showFirebaseUnavailableAlert() {
  alert("Firebase scripts failed to load. Verify the CDN script tags in index.html.");
}

function showFirebaseConfigAlert() {
  alert("Firebase is not configured yet. Update window.FIREBASE_CONFIG with your project credentials.");
}

function sanitizeStudentList(list) {
  if (!Array.isArray(list)) return [];
  const out = [];
  for (const item of list) {
    if (typeof item === "string") {
      const trimmed = item.trim();
      if (trimmed.length) out.push(trimmed);
    }
  }
  return out;
}

function parseStudentList(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
}

function getClassesCollection(uid = currentUserId) {
  if (!firebaseDb || !uid) return null;
  return firebaseDb.collection("users").doc(uid).collection("classes");
}

function createSettingsMessage(text, type = "info") {
  const p = document.createElement("p");
  p.className = `settings-message${type === "error" ? " error" : ""}`;
  p.textContent = text;
  return p;
}

function renderClassList() {
  if (!classListEl) return;
  classListEl.innerHTML = "";

  const showListMessage = (text, type) => {
    classListEl.appendChild(createSettingsMessage(text, type));
    renderClassDetail();
    updateClassSelector();
    syncActiveWinnerFromSelection({ persist: false });
    renderWeeklyList();
    unsubscribeClassWinner();
  };

  if (!currentUserId) {
    selectedClassId = null;
    showListMessage("Sign in to manage classes.");
    return;
  }

  if (!firebaseDb) {
    selectedClassId = null;
    showListMessage("Firestore is unavailable. Check your Firebase configuration.");
    return;
  }

  if (classesLoading) {
    selectedClassId = null;
    showListMessage("Loading classes...");
    return;
  }

  if (classesError) {
    selectedClassId = null;
    showListMessage(classesError, "error");
    return;
  }

  if (!userClasses.length) {
    selectedClassId = null;
    showListMessage("No classes yet. Add one below.");
    return;
  }

  if (!selectedClassId || !userClasses.some((classEntry) => classEntry.id === selectedClassId)) {
    selectedClassId = userClasses[0].id;
  }

  userClasses.forEach((classEntry) => {
    const itemBtn = document.createElement("button");
    itemBtn.type = "button";
    itemBtn.className = "class-item";
    const isSelected = classEntry.id === selectedClassId;
    if (isSelected) {
      itemBtn.classList.add("selected");
      itemBtn.setAttribute("aria-current", "true");
    }
    itemBtn.setAttribute("role", "option");
    itemBtn.setAttribute("aria-selected", isSelected ? "true" : "false");
    itemBtn.setAttribute("aria-label", `${classEntry.name}, ${classEntry.students.length} student${classEntry.students.length === 1 ? "" : "s"}`);

    const nameEl = document.createElement("span");
    nameEl.className = "class-item-name";
    nameEl.textContent = classEntry.name;

    const countEl = document.createElement("span");
    countEl.className = "class-item-count";
    countEl.textContent = `${classEntry.students.length} student${classEntry.students.length === 1 ? "" : "s"}`;

    itemBtn.appendChild(nameEl);
    itemBtn.appendChild(countEl);

    itemBtn.addEventListener("click", () => {
      if (selectedClassId !== classEntry.id) {
        selectedClassId = classEntry.id;
        renderClassList();
      }
    });

    classListEl.appendChild(itemBtn);
  });

  renderClassDetail();
  updateClassSelector();
  syncActiveWinnerFromSelection({ persist: false });
  renderWeeklyList();
  refreshClassWinnerSubscription();
}

function renderClassDetail() {
  if (!classDetailEl) return;
  classDetailEl.innerHTML = "";

  const showDetailMessage = (text, type = "info") => {
    const p = document.createElement("p");
    p.className = `class-detail-message${type === "error" ? " class-detail-error" : ""}`;
    p.textContent = text;
    classDetailEl.appendChild(p);
  };

  if (!currentUserId) {
    showDetailMessage("Sign in to view students.");
    return;
  }

  if (!firebaseDb) {
    showDetailMessage("Firestore is unavailable. Check your Firebase configuration.");
    return;
  }

  if (classesLoading) {
    showDetailMessage("Loading students...");
    return;
  }

  if (classesError) {
    showDetailMessage(classesError, "error");
    return;
  }

  if (!userClasses.length) {
    showDetailMessage("Add a class to see its students here.");
    return;
  }

  if (!selectedClassId) {
    showDetailMessage("Select a class from the left.");
    return;
  }

  const classEntry = userClasses.find((entry) => entry.id === selectedClassId);
  if (!classEntry) {
    showDetailMessage("Select a class from the left.");
    return;
  }

  const title = document.createElement("h3");
  title.textContent = classEntry.name;
  classDetailEl.appendChild(title);

  const meta = document.createElement("p");
  meta.className = "class-detail-meta";
  meta.textContent = `${classEntry.students.length} student${classEntry.students.length === 1 ? "" : "s"}`;
  classDetailEl.appendChild(meta);

  const list = document.createElement("ul");
  list.className = "student-list";
  classEntry.students.forEach((student) => {
    const li = document.createElement("li");
    li.textContent = student;
    list.appendChild(li);
  });
  classDetailEl.appendChild(list);

  const actions = document.createElement("div");
  actions.className = "class-detail-actions";
  const editBtn = document.createElement("button");
  editBtn.type = "button";
  editBtn.className = "btn btn-secondary btn-sm";
  editBtn.textContent = editingClassId === classEntry.id ? "Close Edit" : "Edit Class";
  editBtn.addEventListener("click", () => {
    if (editingClassId === classEntry.id) {
      editingClassId = null;
      renderClassDetail();
    } else {
      editingClassId = classEntry.id;
      renderClassDetail();
    }
  });
  actions.appendChild(editBtn);
  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "btn btn-secondary btn-sm";
  deleteBtn.textContent = "Delete Class";
  deleteBtn.addEventListener("click", () => {
    void deleteClass(classEntry, deleteBtn);
  });
  actions.appendChild(deleteBtn);
  classDetailEl.appendChild(actions);

  if (editingClassId === classEntry.id) {
    const editPanel = document.createElement("div");
    editPanel.className = "class-edit-panel";

    const editLabel = document.createElement("label");
    editLabel.className = "settings-label";
    editLabel.textContent = "Edit Students";
    editLabel.setAttribute("for", "classEditTextarea");
    editPanel.appendChild(editLabel);

    const textarea = document.createElement("textarea");
    textarea.id = "classEditTextarea";
    textarea.className = "settings-textarea";
    textarea.rows = 8;
    textarea.value = classEntry.students.join("\n");
    editPanel.appendChild(textarea);

    const message = document.createElement("p");
    message.className = "settings-message";
    message.hidden = true;
    editPanel.appendChild(message);

    const editActions = document.createElement("div");
    editActions.className = "dialog-actions class-edit-actions";

    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "btn btn-secondary";
    cancelBtn.textContent = "Cancel";
    cancelBtn.addEventListener("click", () => {
      editingClassId = null;
      renderClassDetail();
    });

    const saveBtn = document.createElement("button");
    saveBtn.type = "button";
    saveBtn.className = "btn";
    saveBtn.textContent = "Save";
    saveBtn.addEventListener("click", async () => {
      const students = parseStudentList(textarea.value);
      if (!students.length) {
        message.textContent = "Enter at least one student.";
        message.classList.add("error");
        message.hidden = false;
        return;
      }
      saveBtn.disabled = true;
      cancelBtn.disabled = true;
      message.textContent = "Saving...";
      message.classList.remove("error");
      message.hidden = false;
      try {
        const classesCollection = getClassesCollection();
        if (!classesCollection) throw new Error("Missing Firestore reference");
        const docRef = classesCollection.doc(classEntry.id);
        const payload = { students };
        const fieldValue = firebase && firebase.firestore && firebase.firestore.FieldValue;
        if (fieldValue && typeof fieldValue.serverTimestamp === "function") {
          payload.updatedAt = fieldValue.serverTimestamp();
        } else {
          payload.updatedAt = Date.now();
        }
        await docRef.update(payload);
        const localEntry = userClasses.find((entry) => entry.id === classEntry.id);
        if (localEntry) localEntry.students = students;
        editingClassId = null;
        renderClassList();
      } catch (err) {
        console.error("Failed to update class", err);
        message.textContent = "Failed to save changes. Check the console for details.";
        message.classList.add("error");
        message.hidden = false;
        saveBtn.disabled = false;
        cancelBtn.disabled = false;
      }
    });

    editActions.appendChild(cancelBtn);
    editActions.appendChild(saveBtn);
    editPanel.appendChild(editActions);
    classDetailEl.appendChild(editPanel);
  }
}

function updateClassSelector() {
  if (!classSelector) return;
  if (!currentUserId || !firebaseDb || classesLoading) {
    classSelector.hidden = true;
    classSelector.classList.add("is-empty");
    classSelector.disabled = true;
    classSelector.innerHTML = "";
    return;
  }
  if (classesError || !userClasses.length) {
    classSelector.hidden = true;
    classSelector.classList.add("is-empty");
    classSelector.disabled = true;
    classSelector.innerHTML = "";
    return;
  }
  classSelector.innerHTML = "";
  userClasses.forEach((classEntry) => {
    const option = document.createElement("option");
    option.value = classEntry.id;
    option.textContent = classEntry.name;
    if (classEntry.id === selectedClassId) option.selected = true;
    classSelector.appendChild(option);
  });
  classSelector.hidden = false;
  classSelector.classList.remove("is-empty");
  classSelector.disabled = false;
  updateStartEnabled();
}

async function deleteClass(classEntry, button) {
  if (!firebaseDb || !currentUserId) {
    alert("Firestore is unavailable. Cannot delete class.");
    return;
  }
  const ok = confirm(`Delete class "${classEntry.name}"?`);
  if (!ok) return;
  const originalLabel = button ? button.textContent : null;
  if (button) {
    button.disabled = true;
    button.textContent = "Deleting...";
  }
  try {
    const classesCollection = getClassesCollection();
    if (!classesCollection) throw new Error("Missing Firestore reference");
    await classesCollection.doc(classEntry.id).delete();
    userClasses = userClasses.filter((entry) => entry.id !== classEntry.id);
    if (!userClasses.some((entry) => entry.id === selectedClassId)) {
      selectedClassId = userClasses.length ? userClasses[0].id : null;
    }
    classesError = null;
    const rankingKey = getRankingKeyForClass(classEntry.id);
    if (rankingStore[rankingKey]) {
      delete rankingStore[rankingKey];
    }
    renderClassList();
    fetchGlobalClassCount();
    void adjustGlobalUsageCount(-1);
    void refreshGlobalStudentCount();
    void updateGlobalMvpTotals();
  } catch (err) {
    console.error("Failed to delete class", err);
    alert("Failed to delete class. Check the console for details.");
    if (button) {
      button.disabled = false;
      button.textContent = originalLabel;
    }
  }
}

function openSettingsDialog() {
  if (!settingsDialog) return;
  if (currentUserId && firebaseDb && !classesLoading && (userClasses.length === 0 || classesError)) {
    void fetchClassesForUser(currentUserId);
  }
  renderClassList();
  if (typeof settingsDialog.showModal === "function") {
    if (!settingsDialog.open) settingsDialog.showModal();
  } else {
    settingsDialog.setAttribute("open", "");
  }
}

function closeSettingsDialog() {
  if (!settingsDialog) return;
  if (typeof settingsDialog.close === "function") {
    settingsDialog.close();
  } else {
    settingsDialog.removeAttribute("open");
  }
}

async function fetchClassesForUser(uid) {
  if (!uid || !firebaseDb) {
    userClasses = [];
    classesError = firebaseDb ? null : "Firestore unavailable. Configure Firebase Firestore.";
    classesLoading = false;
    selectedClassId = null;
    renderClassList();
    updateClassSelector();
    return;
  }
  const token = ++classesFetchToken;
  classesLoading = true;
  classesError = null;
  renderClassList();
  try {
    const classesCollection = getClassesCollection(uid);
    if (!classesCollection) throw new Error("Missing Firestore collection reference");
    const snapshot = await classesCollection.get();
    if (token !== classesFetchToken || uid !== currentUserId) return;
    const items = [];
    const seenRankingKeys = new Set();
    snapshot.forEach((doc) => {
      const data = doc.data() || {};
      const name = typeof data.name === "string" ? data.name.trim() : "";
      const students = sanitizeStudentList(data.students);
      if (!name || !students.length) return;
      const rankingKey = getRankingKeyForClass(doc.id);
      const weeklyWinners = normalizeRankingArray(data.weeklyWinners);
      rankingStore[rankingKey] = weeklyWinners;
      seenRankingKeys.add(rankingKey);
      const currentWinner = typeof data.currentWinner === "string" ? data.currentWinner.trim() : null;
      items.push({ id: doc.id, name, students, currentWinner });
    });
    for (const key of Object.keys(rankingStore)) {
      if (key.startsWith("class:") && !seenRankingKeys.has(key)) {
        delete rankingStore[key];
      }
    }
    items.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
    userClasses = items;
    if (!userClasses.some((entry) => entry.id === selectedClassId)) {
      selectedClassId = userClasses.length ? userClasses[0].id : null;
    }
    classesLoading = false;
    renderClassList();
    updateClassSelector();
    void updateGlobalMvpTotals();
  } catch (err) {
    if (token !== classesFetchToken || uid !== currentUserId) return;
    console.error("Failed to load classes", err);
    userClasses = [];
    classesError = "Failed to load classes. Try again later.";
    classesLoading = false;
    selectedClassId = null;
    renderClassList();
    updateClassSelector();
    void updateGlobalMvpTotals();
  }
}

function handleLoginClick() {
  if (!firebaseAuth || !firebaseGoogleProvider) {
    showFirebaseConfigAlert();
    return;
  }
  const user = firebaseAuth.currentUser;
  if (user) {
    firebaseAuth.signOut().catch((err) => {
      console.error("Firebase sign-out failed", err);
      alert("Sign-out failed. Check the console for details.");
    });
    closeSettingsDialog();
    return;
  }
  firebaseAuth
    .signInWithPopup(firebaseGoogleProvider)
    .catch((err) => {
      if (err && err.code === "auth/popup-closed-by-user") return;
      console.error("Google sign-in failed", err);
      alert("Google sign-in failed. Check the console for details.");
    });
}

function initFirebaseAuth() {
  if (!loginBtn) return;
  updateLoginButton();
  const config = typeof window !== "undefined" ? window.FIREBASE_CONFIG : undefined;
  if (typeof firebase === "undefined" || !firebase || typeof firebase.auth !== "function") {
    loginBtn.addEventListener("click", showFirebaseUnavailableAlert);
    loginBtn.title = "Firebase scripts missing";
    ensureSettingsVisibility();
    unsubscribeGlobalStats();
    return;
  }
  if (!hasValidFirebaseConfig(config)) {
    loginBtn.addEventListener("click", showFirebaseConfigAlert);
    loginBtn.title = "Configure Firebase to enable login";
    ensureSettingsVisibility();
    return;
  }
  try {
    if (!firebase.apps || !firebase.apps.length) {
      firebase.initializeApp(config);
    }
    firebaseAuth = firebase.auth();
    firebaseGoogleProvider = new firebase.auth.GoogleAuthProvider();
    if (typeof firebase.firestore === "function") {
      try {
        firebaseDb = firebase.firestore();
      } catch (err) {
        console.error("Failed to initialize Firestore", err);
        firebaseDb = null;
        unsubscribeGlobalStats();
      }
    } else {
      firebaseDb = null;
      unsubscribeGlobalStats();
    }
    refreshGlobalStatsSubscription();
    ensureSettingsVisibility();
    firebaseAuth.onAuthStateChanged((user) => {
      currentUserId = user && user.uid ? user.uid : null;
      if (!currentUserId) {
        classesFetchToken += 1;
        userClasses = [];
        classesError = null;
        classesLoading = false;
        selectedClassId = null;
        unsubscribeClassWinner();
        resetUserStateToDefaults();
        closeSettingsDialog();
        updateLoginButton();
        renderClassList();
        updateClassSelector();
        if (classNameInput) classNameInput.value = "";
        if (studentListTextarea) studentListTextarea.value = "";
        return;
      }
      updateLoginButton();
      if (classNameInput) classNameInput.value = "";
      if (studentListTextarea) studentListTextarea.value = "";
      void loadUserState(currentUserId);
      void fetchClassesForUser(currentUserId);
    });
    firebaseAuth.onIdTokenChanged(() => {
      ensureSettingsVisibility();
    });
    loginBtn.addEventListener("click", handleLoginClick);
    updateLoginButton();
  } catch (err) {
    console.error("Failed to initialize Firebase auth", err);
    loginBtn.disabled = true;
    loginBtn.title = "Firebase auth failed to initialize";
  }
}

function computeMaxLen() {
  return Math.max(5, ...names.map((n) => n.length));
}

function normalizeRankingArray(value) {
  const out = [];
  if (Array.isArray(value)) {
    for (const item of value) {
      if (typeof item === "string") {
        out.push({ name: item, ts: Date.now() });
      } else if (item && typeof item.name === "string") {
        const rawTs = item.ts;
        let ts;
        if (typeof rawTs === "number" && Number.isFinite(rawTs)) {
          ts = rawTs;
        } else if (rawTs && typeof rawTs.toMillis === "function") {
          ts = rawTs.toMillis();
        } else {
          ts = Date.now();
        }
        out.push({ name: item.name, ts });
      }
    }
  }
  return out;
}

function normalizeRankingStore(value) {
  const out = {};
  if (value && typeof value === "object") {
    for (const [key, entry] of Object.entries(value)) {
      out[key] = normalizeRankingArray(entry);
    }
  }
  if (!out.default) out.default = [];
  return out;
}

function sanitizeNameListForState(list) {
  if (!Array.isArray(list)) return [];
  const out = [];
  for (const item of list) {
    if (typeof item === "string") {
      const trimmed = item.trim();
      if (trimmed.length) out.push(trimmed);
    }
  }
  return out;
}

function getRankingEntriesForStorage(list) {
  const out = [];
  if (!Array.isArray(list)) return out;
  for (const entry of list) {
    if (!entry) continue;
    const name = typeof entry.name === "string" ? entry.name.trim() : "";
    if (!name) continue;
    const tsValue = entry.ts;
    let ts;
    if (typeof tsValue === "number" && Number.isFinite(tsValue)) {
      ts = tsValue;
    } else if (tsValue && typeof tsValue.toMillis === "function") {
      ts = tsValue.toMillis();
    } else {
      ts = Date.now();
    }
    out.push({ name, ts });
  }
  return out;
}

function getCurrentWinnerForStart() {
  if (selectedClassId) {
    const selectedClassEntry = getSelectedClass();
    if (selectedClassEntry && typeof selectedClassEntry.currentWinner === "string") {
      const trimmed = selectedClassEntry.currentWinner.trim();
      if (trimmed.length) return trimmed;
    }
    return null;
  }
  if (typeof fixedWinner === "string") {
    const trimmed = fixedWinner.trim();
    if (trimmed.length) return trimmed;
  }
  return null;
}

function resetUserStateToDefaults({ render = true } = {}) {
  names = DEFAULT_NAMES.slice();
  fixedWinner = null;
  defaultFixedWinner = null;
  rankingStore = { default: [] };
  maxLen = computeMaxLen();
  if (!render) return;
  board.textContent = padToMax("READY");
  updateStartEnabled();
  reconcileFixedWinnerWithPool();
  updateWinnerStatus();
  renderWeeklyList();
}

function getUserDocRef(uid = currentUserId) {
  if (!firebaseDb || !uid) return null;
  return firebaseDb.collection("users").doc(uid);
}

function persistUserState(partial) {
  if (!firebaseDb || !currentUserId) return Promise.resolve();
  const docRef = getUserDocRef();
  if (!docRef || !partial || typeof partial !== "object") return Promise.resolve();
  const payload = {};
  if (Object.prototype.hasOwnProperty.call(partial, "names")) {
    payload.names = Array.isArray(partial.names)
      ? partial.names
          .map((item) => (typeof item === "string" ? item.trim() : ""))
          .filter((item) => item.length)
      : [];
  }
  if (Object.prototype.hasOwnProperty.call(partial, "fixedWinner")) {
    if (typeof partial.fixedWinner === "string") {
      const trimmed = partial.fixedWinner.trim();
      payload.fixedWinner = trimmed.length ? trimmed : null;
    } else {
      payload.fixedWinner = null;
    }
  }
  if (Object.prototype.hasOwnProperty.call(partial, "rankingStore")) {
    try {
      payload.rankingStore = partial.rankingStore
        ? JSON.parse(JSON.stringify(partial.rankingStore))
        : { default: [] };
    } catch {
      payload.rankingStore = { default: [] };
    }
  }
  if (!Object.keys(payload).length) return Promise.resolve();
  return docRef
    .set(payload, { merge: true })
    .catch((err) => {
      console.error("Failed to persist user state", err);
    });
}

function persistNames() {
  void persistUserState({ names });
}

function persistFixedWinner() {
  defaultFixedWinner = fixedWinner;
  void persistUserState({ fixedWinner });
}

function persistRankingStore(classId = selectedClassId) {
  if (!firebaseDb || !currentUserId) return Promise.resolve();
  const key = getRankingKeyForClass(classId);
  const list = getRankingEntriesForKey(key);
  const sanitized = getRankingEntriesForStorage(list);
  list.splice(0, list.length, ...sanitized);
  if (!classId) {
    return persistUserState({ rankingStore: { default: sanitized } });
  }
  const classesCollection = getClassesCollection();
  if (!classesCollection) return Promise.resolve();
  return classesCollection
    .doc(classId)
    .set({ weeklyWinners: sanitized }, { merge: true })
    .catch((err) => {
      console.error("Failed to persist class ranking", err);
    });
}

function persistClassCurrentWinner(classId, winner) {
  if (!firebaseDb || !currentUserId || !classId) return Promise.resolve();
  const classesCollection = getClassesCollection();
  if (!classesCollection) return Promise.resolve();
  const value = typeof winner === "string" ? winner.trim() : "";
  const payload = { currentWinner: value.length ? value : null };
  return classesCollection
    .doc(classId)
    .set(payload, { merge: true })
    .catch((err) => {
      console.error("Failed to persist class winner", err);
    });
}

function updateActiveWinner(nextWinner, { persist = true, skipReconcile = false } = {}) {
  const selectedClassEntry = getSelectedClass();
  const trimmed = typeof nextWinner === "string" ? nextWinner.trim() : "";
  fixedWinner = trimmed.length ? trimmed : null;
  if (selectedClassEntry) {
    selectedClassEntry.currentWinner = fixedWinner;
    if (persist) void persistClassCurrentWinner(selectedClassEntry.id, fixedWinner);
  } else if (persist) {
    persistFixedWinner();
  } else {
    defaultFixedWinner = fixedWinner;
  }
  updateWinnerStatus();
  updateStartEnabled();
  if (!skipReconcile) reconcileFixedWinnerWithPool();
}

function syncActiveWinnerFromSelection({ persist = false } = {}) {
  if (selectedClassId) {
    const selectedClassEntry = getSelectedClass();
    const winner = selectedClassEntry && typeof selectedClassEntry.currentWinner === "string" ? selectedClassEntry.currentWinner : null;
    updateActiveWinner(winner, { persist, skipReconcile: true });
  } else {
    updateActiveWinner(defaultFixedWinner, { persist: false, skipReconcile: true });
  }
  reconcileFixedWinnerWithPool();
}

async function refreshGlobalStudentCount() {
  try {
    await updateGlobalStudentTotals();
  } catch (err) {
    console.error("Failed to refresh global student count", err);
  }
}

async function updateGlobalMvpTotals() {
  if (!firebaseDb) return;
  try {
    const totalMvps = Object.values(rankingStore || {}).reduce((sum, entries) => {
      if (Array.isArray(entries)) {
        return sum + entries.length;
      }
      return sum;
    }, 0);

    await firebaseDb
      .collection("stats")
      .doc(GLOBAL_STATS_DOC_ID)
      .set(
        {
          mvpCount: totalMvps,
        },
        { merge: true }
      );
    setLatestGlobalStats({ mvpCount: totalMvps });
  } catch (err) {
    console.error("Failed to update global MVP count", err);
  }
}

function unsubscribeClassWinner() {
  if (classWinnerUnsubscribe) {
    classWinnerUnsubscribe();
    classWinnerUnsubscribe = null;
  }
}

function refreshClassWinnerSubscription() {
  unsubscribeClassWinner();
  if (!firebaseDb || !currentUserId || !selectedClassId) return;
  const classesCollection = getClassesCollection();
  if (!classesCollection) return;
  try {
    classWinnerUnsubscribe = classesCollection.doc(selectedClassId).onSnapshot(
      (snapshot) => {
        const data = snapshot && snapshot.exists ? snapshot.data() || {} : {};
        const winner = typeof data.currentWinner === "string" ? data.currentWinner.trim() : null;
        const selectedClassEntry = userClasses.find((entry) => entry.id === selectedClassId);
        if (selectedClassEntry) selectedClassEntry.currentWinner = winner || null;
        if (Array.isArray(data.weeklyWinners)) {
          applyRemoteWeeklyWinners(selectedClassId, data.weeklyWinners, { render: true });
        }
        updateActiveWinner(winner, { persist: false, skipReconcile: true });
        reconcileFixedWinnerWithPool();
      },
      (err) => {
        console.error("Winner listener error", err);
      }
    );
  } catch (err) {
    console.error("Failed to subscribe to class winner", err);
  }
}

function unsubscribeGlobalStats() {
  if (globalStatsUnsubscribe) {
    globalStatsUnsubscribe();
    globalStatsUnsubscribe = null;
  }
}

function refreshGlobalStatsSubscription() {
  unsubscribeGlobalStats();
  if (!firebaseDb) return;
  try {
    globalStatsUnsubscribe = firebaseDb
      .collection("stats")
      .doc(GLOBAL_STATS_DOC_ID)
      .onSnapshot(
        (snapshot) => {
        const data = snapshot && snapshot.exists ? snapshot.data() || {} : {};
        if (typeof data.classCount === "number" && !Number.isNaN(data.classCount)) {
          totalClassCount = data.classCount;
        }
        setLatestGlobalStats({
          classCount: typeof data.classCount === "number" ? data.classCount : undefined,
          studentCount: typeof data.studentCount === "number" ? data.studentCount : undefined,
          mvpCount: typeof data.mvpCount === "number" ? data.mvpCount : undefined,
        });
        renderWeeklyList();
      },
        (err) => {
          console.error("Global stats listener error", err);
        }
      );
  } catch (err) {
    console.error("Failed to subscribe to global stats", err);
  }
}

async function loadUserState(uid) {
  resetUserStateToDefaults({ render: false });
  if (!firebaseDb || !uid) {
    board.textContent = padToMax("READY");
    updateStartEnabled();
    updateWinnerStatus();
    renderWeeklyList();
    return;
  }
  try {
    const docRef = getUserDocRef(uid);
    const snapshot = docRef ? await docRef.get() : null;
    const data = snapshot && snapshot.exists ? snapshot.data() || {} : {};
    const fetchedNames = sanitizeNameListForState(data.names);
    names = fetchedNames.length ? fetchedNames : DEFAULT_NAMES.slice();
    fixedWinner = typeof data.fixedWinner === "string" && data.fixedWinner.trim().length ? data.fixedWinner.trim() : null;
    defaultFixedWinner = fixedWinner;
    rankingStore = normalizeRankingStore(data.rankingStore);
    maxLen = computeMaxLen();
    reconcileFixedWinnerWithPool();
    board.textContent = padToMax("READY");
    updateStartEnabled();
    updateWinnerStatus();
    renderWeeklyList();
    void updateGlobalMvpTotals();
    if (!snapshot || !snapshot.exists) {
      await persistUserState({ names, fixedWinner, rankingStore });
    }
  } catch (err) {
    console.error("Failed to load user state", err);
    resetUserStateToDefaults();
  }
}

function getRankingKeyForClass(classId) {
  return classId ? `class:${classId}` : "default";
}

function getActiveRankingKey() {
  return getRankingKeyForClass(selectedClassId);
}

function getRankingEntriesForKey(key) {
  if (!rankingStore[key]) rankingStore[key] = [];
  return rankingStore[key];
}

function getActiveRankingEntries() {
  return getRankingEntriesForKey(getActiveRankingKey());
}

function areRankingListsEqual(listA, listB) {
  if (!Array.isArray(listA) || !Array.isArray(listB)) return false;
  if (listA.length !== listB.length) return false;
  for (let i = 0; i < listA.length; i++) {
    const a = listA[i];
    const b = listB[i];
    if (!a || !b) return false;
    if (a.name !== b.name || a.ts !== b.ts) return false;
  }
  return true;
}

function applyRemoteWeeklyWinners(classId, entries, { render = true } = {}) {
  const key = getRankingKeyForClass(classId);
  const normalized = normalizeRankingArray(entries);
  const current = rankingStore[key] || [];
  if (!areRankingListsEqual(current, normalized)) {
    rankingStore[key] = normalized;
    if (render) renderWeeklyList();
  }
}

function addRankingEntry(name) {
  const list = getActiveRankingEntries();
  list.push({ name, ts: Date.now() });
  void persistRankingStore();
  void updateGlobalMvpTotals();
  renderWeeklyList();
}

// CSV download removed per request; ranking persists in Firestore

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
  // Center the string within the current display width using spaces
  const width = getDisplayWidth();
  const totalPad = Math.max(0, width - s.length);
  const left = Math.floor(totalPad / 2);
  const right = totalPad - left;
  return " ".repeat(left) + s + " ".repeat(right);
}

let isBusy = false;

function startShuffleAndPick({ shuffleMs = 5000 } = {}) {
  if (isBusy) return;
  const pool = getActiveNamePool();
  if (!pool.length) {
    board.textContent = padToMax("NO NAMES");
    updateStartEnabled();
    return;
  }

  isBusy = true;
  startBtn.disabled = true;
  ensureAudio();
  startShuffleSoundEffect();

  // Initial board content using one of the names in the pool
  board.textContent = padToMax(pool[(Math.random() * pool.length) | 0].toUpperCase());

  // Rapid shuffle loop cycling through names
  const tickMs = 75;
  const interval = setInterval(() => {
    const next = pool[(Math.random() * pool.length) | 0] || "";
    board.textContent = padToMax(next.toUpperCase());
  }, tickMs);

  setTimeout(() => {
    clearInterval(interval);
    stopShuffleSoundEffect();
    if (!pool.length) {
      board.textContent = padToMax("NO NAMES");
      isBusy = false;
      updateStartEnabled();
      return;
    }
    let winner;
    if (fixedWinner) {
      const match = pool.find((n) => n.toUpperCase() === fixedWinner.toUpperCase());
      if (match) {
        winner = match;
      }
    }
    if (!winner) {
      winner = pool[(Math.random() * pool.length) | 0];
    }
    revealWinner(winner);
  }, shuffleMs);
}

function revealWinner(name) {
  const target = padToMax(name.toUpperCase());
  // Start from whatever is on screen now, or spaces
  const width = getDisplayWidth();
  let current = (board.textContent || "").padEnd(width, " ").slice(0, width).split("");

  // For each position, flip a few random chars before locking the final char
  const flipsPerChar = 7;
  const stepDelay = 18; // ms between flips within a slot
  const cascadeDelay = 55; // ms between starting each slot

  for (let i = 0; i < width; i++) {
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
      if (i === width - 1) {
        playCongrats();
        // Record winner and update/download ranking CSV
        addRankingEntry(name);
        updateActiveWinner(null);
        setTimeout(() => {
          isBusy = false;
          updateStartEnabled();
        }, 50);
      }
    }, startAt + flipsPerChar * stepDelay + 12);
  }
}

async function handleStartButtonClick() {
  if (isBusy) return;
  const ms = Math.round((shuffleSeconds || 5) * 1000);
  let winner = getCurrentWinnerForStart();

  if (selectedClassId) {
    const classesCollection = getClassesCollection();
    if (!classesCollection) {
      alert("Firestore is unavailable. Cannot load the current winner.");
      updateStartEnabled();
      return;
    }
    try {
      const snapshot = await classesCollection.doc(selectedClassId).get();
      const data = snapshot && snapshot.exists ? snapshot.data() || {} : {};
      const storedWinner = typeof data.currentWinner === "string" ? data.currentWinner.trim() : "";
      const selectedClassEntry = getSelectedClass();
      if (storedWinner.length) {
        winner = storedWinner;
        if (selectedClassEntry) selectedClassEntry.currentWinner = storedWinner;
      } else {
        winner = selectedClassEntry && typeof selectedClassEntry.currentWinner === "string" && selectedClassEntry.currentWinner.trim().length
          ? selectedClassEntry.currentWinner.trim()
          : null;
      }
    } catch (err) {
      console.error("Failed to read class winner", err);
      alert("Failed to load the current winner. Please try again.");
      updateStartEnabled();
      return;
    }
  }

  if (!winner) {
    alert("Select a winner before starting the shuffle.");
    updateStartEnabled();
    return;
  }

  const pool = getActiveNamePool();
  if (!pool.length) {
    alert("Add students to the class before starting.");
    updateStartEnabled();
    return;
  }
  const winnerExists = pool.some((n) => n.toUpperCase() === winner.toUpperCase());
  if (!winnerExists) {
    alert("The selected winner is not in this class. Please choose another winner.");
    updateActiveWinner(null, { persist: true });
    return;
  }

  updateActiveWinner(winner, { persist: false });
  startShuffleAndPick({ shuffleMs: ms });
}

if (startBtn) {
  startBtn.addEventListener("click", () => {
    void handleStartButtonClick();
  });
}

// Ensure stable width right away
function updateStartEnabled() {
  if (!startBtn) return;
  const hasNames = getActiveNamePool().length > 0;
  const winner = getCurrentWinnerForStart();
  startBtn.disabled = isBusy || !hasNames || !winner;
}

function openNamesDialog() {
  // Pre-fill with current names
  namesTextarea.value = getActiveNamePool().join("\n");
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

function applyNames(newNames, { persist = true } = {}) {
  names = newNames.slice();
  maxLen = computeMaxLen();
  if (persist) persistNames();
  board.textContent = padToMax("READY");
  updateStartEnabled();
  // If the fixed winner no longer exists in the list, clear it
  if (fixedWinner && !names.some((n) => n.toUpperCase() === fixedWinner.toUpperCase())) {
    updateActiveWinner(null, { persist, skipReconcile: true });
  }
  reconcileFixedWinnerWithPool();
}

if (setNamesBtn) {
  setNamesBtn.addEventListener("click", openNamesDialog);
}
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
  const classStudents = getSelectedClassStudents();
  const sourceList = classStudents.length ? classStudents : names;

  if (!sourceList.length) {
    const p = document.createElement("p");
    p.textContent = classStudents.length
      ? "This class has no students yet."
      : "No names available. Please set names first.";
    p.style.color = "var(--subtle)";
    winnerGrid.appendChild(p);
  } else {
    sourceList.forEach((name) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "name-btn";
      btn.textContent = name;
      btn.addEventListener("click", () => {
        updateActiveWinner(name);
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
  updateActiveWinner(null);
  closeWinnerDialog();
});

board.textContent = padToMax("READY");
updateStartEnabled();
updateWinnerStatus();
renderWeeklyList();
renderClassList();
updateClassUsageNoticeDisplay();
initFirebaseAuth();

if (settingsBtn) {
  settingsBtn.addEventListener("click", () => {
    if (!firebaseAuth || !firebaseAuth.currentUser) {
      alert("Please sign in to access settings.");
      return;
    }
    if (!firebaseDb) {
      alert("Firestore is unavailable. Check your Firebase configuration.");
      return;
    }
    openSettingsDialog();
  });
}

if (closeSettingsBtn) {
  closeSettingsBtn.addEventListener("click", () => {
    closeSettingsDialog();
  });
}

if (classSelector) {
  classSelector.addEventListener("change", () => {
    const nextId = classSelector.value;
    if (nextId && userClasses.some((classEntry) => classEntry.id === nextId)) {
      if (selectedClassId !== nextId) {
        selectedClassId = nextId;
        renderClassList();
      }
    } else if (!nextId) {
      selectedClassId = null;
      renderClassList();
    }
  });
}

if (addClassBtn) {
  addClassBtn.addEventListener("click", async () => {
    if (!currentUserId) {
      alert("Sign in to add a class.");
      return;
    }
    if (!firebaseDb) {
      alert("Firestore is unavailable. Check your Firebase configuration.");
      return;
    }
    const name = classNameInput ? classNameInput.value.trim() : "";
    const students = studentListTextarea ? parseStudentList(studentListTextarea.value) : [];
    if (!name) {
      alert("Enter a course name.");
      return;
    }
    if (!students.length) {
      alert("Enter at least one student name.");
      return;
    }
    const nameAlreadyExists = userClasses.some((classEntry) => classEntry.name.toLowerCase() === name.toLowerCase());
    if (nameAlreadyExists) {
      const overwrite = confirm("A class with this name already exists. Add another entry anyway?");
      if (!overwrite) return;
    }
    const originalLabel = addClassBtn.textContent;
    addClassBtn.disabled = true;
    addClassBtn.textContent = "Saving...";
    try {
      const classesCollection = getClassesCollection();
      if (!classesCollection) throw new Error("Missing Firestore reference");
      const docRef = classesCollection.doc();
      const payload = { name, students };
      const fieldValue = firebase && firebase.firestore && firebase.firestore.FieldValue;
      if (fieldValue && typeof fieldValue.serverTimestamp === "function") {
        payload.createdAt = fieldValue.serverTimestamp();
      } else {
        payload.createdAt = Date.now();
      }
      payload.weeklyWinners = [];
      await docRef.set(payload);
      const rankingKey = getRankingKeyForClass(docRef.id);
      rankingStore[rankingKey] = [];
      userClasses.push({ id: docRef.id, name, students, currentWinner: null });
      userClasses.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
      classesError = null;
      selectedClassId = docRef.id;
      renderClassList();
      renderWeeklyList();
      fetchGlobalClassCount();
      void adjustGlobalUsageCount(1);
      void refreshGlobalStudentCount();
      if (classNameInput) classNameInput.value = "";
      if (studentListTextarea) studentListTextarea.value = "";
    } catch (err) {
      console.error("Failed to add class", err);
      alert("Failed to add class. Check the console for details.");
    } finally {
      addClassBtn.disabled = false;
      addClassBtn.textContent = originalLabel;
    }
  });
}

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

function getAllRowsWithWeekLabels(entries) {
  // Sequential label per run: W1, W2, ... regardless of calendar week
  return entries
    .filter((e) => e && typeof e.name === "string" && typeof e.ts === "number")
    .slice()
    .sort((a, b) => a.ts - b.ts)
    .map((e, idx) => ({ label: `W${idx + 1}`, name: e.name, ts: e.ts }));
}

function renderWeeklyList() {
  const listEl = document.getElementById("weeklyList");
  if (!listEl) return;
  const entries = getActiveRankingEntries();
  const rows = getAllRowsWithWeekLabels(entries);
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
    const list = getActiveRankingEntries();
    const idx = list.findIndex((entry) => entry && typeof entry.ts === "number" && entry.ts === ts);
    if (idx === -1) return;
    const entry = list[idx];
    const name = entry && typeof entry.name === "string" ? entry.name : "this winner";
    const ok = confirm(`Remove ${name} from the winners list?`);
    if (!ok) return;
    list.splice(idx, 1);
    void persistRankingStore();
    void updateGlobalMvpTotals();
    renderWeeklyList();
  });
}

function updateShuffleUI(val) {
  if (shuffleValue) shuffleValue.textContent = `${val}s`;
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
const confirmResetPanel = document.getElementById("confirmResetPanel");
const confirmResetBtn = document.getElementById("confirmResetBtn");
const cancelResetBtn = document.getElementById("cancelResetBtn");

if (clearWeeklyBtn && confirmResetPanel && confirmResetBtn && cancelResetBtn) {
  clearWeeklyBtn.addEventListener("click", () => {
    confirmResetPanel.hidden = false;
  });

  cancelResetBtn.addEventListener("click", () => {
    confirmResetPanel.hidden = true;
  });

  confirmResetBtn.addEventListener("click", () => {
    const key = getActiveRankingKey();
    rankingStore[key] = [];
    void persistRankingStore();
    void updateGlobalMvpTotals();
    renderWeeklyList();
    confirmResetPanel.hidden = true;
  });
}
