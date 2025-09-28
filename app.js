// Names to pick from. Can be set via dialog and saved to localStorage.
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

const STORAGE_KEY = "flightBoardNames";
const WINNER_KEY = "flightBoardFixedWinner";
const RANKING_KEY = "flightBoardRanking";
const GLOBAL_STATS_DOC_ID = "globalUsage";

let names = loadNames();
let fixedWinner = loadFixedWinner();
let rankingStore = loadRankingStore();

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

function getSelectedClass() {
  if (!selectedClassId) return null;
  return userClasses.find((entry) => entry.id === selectedClassId) || null;
}

function getSelectedClassStudents() {
  const cls = getSelectedClass();
  if (!cls || !Array.isArray(cls.students)) return [];
  return cls.students.slice();
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
    return;
  }
  const pool = getActiveNamePool();
  const exists = pool.some((n) => n.toUpperCase() === fixedWinner.toUpperCase());
  if (!exists) {
    fixedWinner = null;
    saveFixedWinner();
  }
  updateWinnerStatus();
}

function updateClassUsageNoticeDisplay() {
  if (!classUsageNotice) return;
  const hasGlobalCount = typeof totalClassCount === "number" && !Number.isNaN(totalClassCount);
  const count = hasGlobalCount ? totalClassCount : userClasses.length;

  if (hasGlobalCount && count > 0) {
    classUsageNotice.textContent = `${count} classes already using MVP for enhance student engagement.`;
  } else if (classCountFetchInFlight) {
    classUsageNotice.textContent = "Checking how many classes are using MVP...";
  } else {
    classUsageNotice.textContent = "Be the first class to use MVP to make students exciting.";
  }

  classUsageNotice.hidden = false;
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
      } else {
        totalClassCount = null;
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

  try {
    if (fieldValue && typeof fieldValue.increment === "function") {
      await statsRef.set(
        {
          classCount: fieldValue.increment(delta),
          updatedAt: timestamp,
        },
        { merge: true }
      );
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
    });
  } catch (err) {
    console.warn("Failed to adjust global usage count", err);
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
    renderWeeklyList();
    updateStartEnabled();
    reconcileFixedWinnerWithPool();
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

  if (!selectedClassId || !userClasses.some((cls) => cls.id === selectedClassId)) {
    selectedClassId = userClasses[0].id;
  }

  userClasses.forEach((cls) => {
    const itemBtn = document.createElement("button");
    itemBtn.type = "button";
    itemBtn.className = "class-item";
    const isSelected = cls.id === selectedClassId;
    if (isSelected) {
      itemBtn.classList.add("selected");
      itemBtn.setAttribute("aria-current", "true");
    }
    itemBtn.setAttribute("role", "option");
    itemBtn.setAttribute("aria-selected", isSelected ? "true" : "false");
    itemBtn.setAttribute("aria-label", `${cls.name}, ${cls.students.length} student${cls.students.length === 1 ? "" : "s"}`);

    const nameEl = document.createElement("span");
    nameEl.className = "class-item-name";
    nameEl.textContent = cls.name;

    const countEl = document.createElement("span");
    countEl.className = "class-item-count";
    countEl.textContent = `${cls.students.length} student${cls.students.length === 1 ? "" : "s"}`;

    itemBtn.appendChild(nameEl);
    itemBtn.appendChild(countEl);

    itemBtn.addEventListener("click", () => {
      if (selectedClassId !== cls.id) {
        selectedClassId = cls.id;
        renderClassList();
      }
    });

    classListEl.appendChild(itemBtn);
  });

  renderClassDetail();
  updateClassSelector();
  updateStartEnabled();
  renderWeeklyList();
  reconcileFixedWinnerWithPool();
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

  const cls = userClasses.find((entry) => entry.id === selectedClassId);
  if (!cls) {
    showDetailMessage("Select a class from the left.");
    return;
  }

  const title = document.createElement("h3");
  title.textContent = cls.name;
  classDetailEl.appendChild(title);

  const meta = document.createElement("p");
  meta.className = "class-detail-meta";
  meta.textContent = `${cls.students.length} student${cls.students.length === 1 ? "" : "s"}`;
  classDetailEl.appendChild(meta);

  const list = document.createElement("ul");
  list.className = "student-list";
  cls.students.forEach((student) => {
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
  editBtn.textContent = editingClassId === cls.id ? "Close Edit" : "Edit Class";
  editBtn.addEventListener("click", () => {
    if (editingClassId === cls.id) {
      editingClassId = null;
      renderClassDetail();
    } else {
      editingClassId = cls.id;
      renderClassDetail();
    }
  });
  actions.appendChild(editBtn);
  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "btn btn-secondary btn-sm";
  deleteBtn.textContent = "Delete Class";
  deleteBtn.addEventListener("click", () => {
    void deleteClass(cls, deleteBtn);
  });
  actions.appendChild(deleteBtn);
  classDetailEl.appendChild(actions);

  if (editingClassId === cls.id) {
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
    textarea.value = cls.students.join("\n");
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
        const col = getClassesCollection();
        if (!col) throw new Error("Missing Firestore reference");
        const docRef = col.doc(cls.id);
        const payload = { students };
        const fieldValue = firebase && firebase.firestore && firebase.firestore.FieldValue;
        if (fieldValue && typeof fieldValue.serverTimestamp === "function") {
          payload.updatedAt = fieldValue.serverTimestamp();
        } else {
          payload.updatedAt = Date.now();
        }
        await docRef.update(payload);
        const local = userClasses.find((entry) => entry.id === cls.id);
        if (local) local.students = students;
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
  userClasses.forEach((cls) => {
    const option = document.createElement("option");
    option.value = cls.id;
    option.textContent = cls.name;
    if (cls.id === selectedClassId) option.selected = true;
    classSelector.appendChild(option);
  });
  classSelector.hidden = false;
  classSelector.classList.remove("is-empty");
  classSelector.disabled = false;
  updateStartEnabled();
}

async function deleteClass(cls, button) {
  if (!firebaseDb || !currentUserId) {
    alert("Firestore is unavailable. Cannot delete class.");
    return;
  }
  const ok = confirm(`Delete class "${cls.name}"?`);
  if (!ok) return;
  const originalLabel = button ? button.textContent : null;
  if (button) {
    button.disabled = true;
    button.textContent = "Deleting...";
  }
  try {
    const col = getClassesCollection();
    if (!col) throw new Error("Missing Firestore reference");
    await col.doc(cls.id).delete();
    userClasses = userClasses.filter((entry) => entry.id !== cls.id);
    if (!userClasses.some((entry) => entry.id === selectedClassId)) {
      selectedClassId = userClasses.length ? userClasses[0].id : null;
    }
    classesError = null;
    const rankingKey = getRankingKeyForClass(cls.id);
    if (rankingStore[rankingKey]) {
      delete rankingStore[rankingKey];
      saveRankingStore();
    }
    renderClassList();
    fetchGlobalClassCount();
    void adjustGlobalUsageCount(-1);
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
    const col = getClassesCollection(uid);
    if (!col) throw new Error("Missing Firestore collection reference");
    const snapshot = await col.get();
    if (token !== classesFetchToken || uid !== currentUserId) return;
    const items = [];
    snapshot.forEach((doc) => {
      const data = doc.data() || {};
      const name = typeof data.name === "string" ? data.name.trim() : "";
      const students = sanitizeStudentList(data.students);
      if (!name || !students.length) return;
      items.push({ id: doc.id, name, students });
    });
    items.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
    userClasses = items;
    if (!userClasses.some((entry) => entry.id === selectedClassId)) {
      selectedClassId = userClasses.length ? userClasses[0].id : null;
    }
    classesLoading = false;
    renderClassList();
    updateClassSelector();
  } catch (err) {
    if (token !== classesFetchToken || uid !== currentUserId) return;
    console.error("Failed to load classes", err);
    userClasses = [];
    classesError = "Failed to load classes. Try again later.";
    classesLoading = false;
    selectedClassId = null;
    renderClassList();
    updateClassSelector();
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
      }
    } else {
      firebaseDb = null;
    }
    ensureSettingsVisibility();
    firebaseAuth.onAuthStateChanged((user) => {
      currentUserId = user && user.uid ? user.uid : null;
      if (!currentUserId) {
        classesFetchToken += 1;
        userClasses = [];
        classesError = null;
        classesLoading = false;
        selectedClassId = null;
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

function normalizeRankingArray(value) {
  const out = [];
  if (Array.isArray(value)) {
    for (const item of value) {
      if (typeof item === "string") {
        out.push({ name: item, ts: Date.now() });
      } else if (item && typeof item.name === "string") {
        out.push({ name: item.name, ts: typeof item.ts === "number" ? item.ts : Date.now() });
      }
    }
  }
  return out;
}

function loadRankingStore() {
  try {
    const raw = localStorage.getItem(RANKING_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        const map = { default: normalizeRankingArray(parsed) };
        localStorage.setItem(RANKING_KEY, JSON.stringify(map));
        return map;
      }
      if (parsed && typeof parsed === "object") {
        const out = {};
        for (const [key, value] of Object.entries(parsed)) {
          out[key] = normalizeRankingArray(value);
        }
        if (!out.default) out.default = [];
        localStorage.setItem(RANKING_KEY, JSON.stringify(out));
        return out;
      }
    }
  } catch {}
  return { default: [] };
}

function saveRankingStore() {
  try {
    localStorage.setItem(RANKING_KEY, JSON.stringify(rankingStore));
  } catch {}
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

function addRankingEntry(name) {
  const list = getActiveRankingEntries();
  list.push({ name, ts: Date.now() });
  saveRankingStore();
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
        setTimeout(() => {
          isBusy = false;
          updateStartEnabled();
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
  startBtn.disabled = isBusy || getActiveNamePool().length === 0;
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
    if (nextId && userClasses.some((cls) => cls.id === nextId)) {
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
    const existing = userClasses.some((cls) => cls.name.toLowerCase() === name.toLowerCase());
    if (existing) {
      const overwrite = confirm("A class with this name already exists. Add another entry anyway?");
      if (!overwrite) return;
    }
    const originalLabel = addClassBtn.textContent;
    addClassBtn.disabled = true;
    addClassBtn.textContent = "Saving...";
    try {
      const col = getClassesCollection();
      if (!col) throw new Error("Missing Firestore reference");
      const docRef = col.doc();
      const payload = { name, students };
      const fieldValue = firebase && firebase.firestore && firebase.firestore.FieldValue;
      if (fieldValue && typeof fieldValue.serverTimestamp === "function") {
        payload.createdAt = fieldValue.serverTimestamp();
      } else {
        payload.createdAt = Date.now();
      }
      await docRef.set(payload);
      userClasses.push({ id: docRef.id, name, students });
      userClasses.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
      classesError = null;
      selectedClassId = docRef.id;
      renderClassList();
      fetchGlobalClassCount();
      void adjustGlobalUsageCount(1);
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
    saveRankingStore();
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
    saveRankingStore();
    renderWeeklyList();
    confirmResetPanel.hidden = true;
  });
}
