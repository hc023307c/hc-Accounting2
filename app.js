// =====================================
// 0. Supabase 設定
// =====================================
const SUPABASE_URL = "https://kjcxngzrrncotukoxbze.supabase.co";
const SUPABASE_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtqY3huZ3pycm5jb3R1a294YnplIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Njg5ODg1NjEsImV4cCI6MjA4NDU2NDU2MX0.MgYNIEhW9v5dempDoFSvoM5foom5ST8t9hkx_0_qHvo";

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

// =====================================
// 1. DOM & Debug
// =====================================
const jsStatusEl = document.getElementById("js-status");
const debugEl = document.getElementById("debug");
const authStatusEl = document.getElementById("auth-status");
const helloBarEl = document.getElementById("hello-bar");

// ledger page DOM
const ledgerInputEl = document.getElementById("ledger-input");
const ledgerTbodyEl = document.getElementById("ledger-tbody");
const editStatusEl = document.getElementById("edit-status");
const entrySubmitBtn = document.getElementById("entry-submit-btn");
const entryCancelEditBtn = document.getElementById("entry-cancel-edit-btn");

// voice DOM
const voiceSectionEl = document.getElementById("voice-section");
const voiceTextInputEl = document.getElementById("voice-text-input");
const voiceStatusEl = document.getElementById("voice-status");
const voicePreviewBodyEl = document.getElementById("voice-preview-body");
const parseTextBtn = document.getElementById("parse-text-btn");
const voiceTextClearBtn = document.getElementById("voice-text-clear-btn");
const voiceConfirmBtn = document.getElementById("voice-confirm-btn");
const voiceClearBtn = document.getElementById("voice-clear-btn");

// stats DOM
const statIncomeEl = document.getElementById("stat-income");
const statExpenseEl = document.getElementById("stat-expense");
const statNetEl = document.getElementById("stat-net");
const statsCategoryBodyEl = document.getElementById("stats-category-body");

// datalist DOM
const categoryListEl = document.getElementById("category-list");
const placeListEl = document.getElementById("place-list");

function logDebug(msg, obj) {
  if (debugEl) {
    const text = msg + (obj ? " " + JSON.stringify(obj, null, 2) : "");
    debugEl.textContent += text + "\n";
  }
  console.log("[DEBUG]", msg, obj || "");
}

if (jsStatusEl) jsStatusEl.textContent = "✅ JS 已載入，Supabase client 建立完成。";
logDebug("Supabase client created");

// =====================================
// 2. Auth
// =====================================
function accountToEmail(account) {
  return account + "@demo.local";
}

async function handleLogin() {
  const accountInput = document.getElementById("login-account");
  const passwordInput = document.getElementById("login-password");

  if (!accountInput || !passwordInput) {
    alert("這個頁面沒有登入表單。");
    return;
  }

  const account = accountInput.value.trim();
  const password = passwordInput.value;

  if (!account || !password) {
    if (authStatusEl) authStatusEl.textContent = "請輸入帳號與密碼";
    alert("請輸入帳號與密碼");
    return;
  }

  const email = accountToEmail(account);
  logDebug("嘗試登入", { email });

  const { data, error } = await supabaseClient.auth.signInWithPassword({
    email,
    password,
  });

  if (error) {
    if (authStatusEl) authStatusEl.textContent = "登入失敗：" + error.message;
    logDebug("登入失敗", error);
    alert("登入失敗：" + error.message);
    return;
  }

  if (authStatusEl) authStatusEl.textContent = "登入成功：" + (data.user?.email || "");
  logDebug("登入成功", data);

  window.location.href = "ledger.html";
}

async function handleLogout() {
  await supabaseClient.auth.signOut();
  logDebug("已登出");
  window.location.href = "index.html";
}

async function getCurrentUser() {
  const { data, error } = await supabaseClient.auth.getUser();
  if (error) {
    if (error.name !== "AuthSessionMissingError") logDebug("getUser error", error);
    return null;
  }
  return data.user || null;
}

async function loadProfile(userId) {
  const { data, error } = await supabaseClient
    .from("profiles")
    .select("username, role")
    .eq("id", userId)
    .single();

  if (error) {
    logDebug("載入 profile 失敗", error);
    return null;
  }
  return data;
}

async function loadRandomZenQuote() {
  // 這裡用簡單方式：先抓前 120 筆，前端隨機挑
  const { data, error } = await supabaseClient
    .from("zen_quotes")
    .select("content, source")
    .eq("enabled", true)
    .limit(120);

  if (error) {
    logDebug("載入 zen_quotes 失敗", error);
    return null;
  }
  if (!data || !data.length) return null;

  const q = data[Math.floor(Math.random() * data.length)];
  return q || null;
}

async function refreshHelloBar() {
  const user = await getCurrentUser();
  if (!user) return;

  const profile = await loadProfile(user.id);
  const username = profile?.username || user.email;

  const quote = await loadRandomZenQuote();
  const quoteText = quote ? `「${quote.content}」` : "";

  if (helloBarEl) {
    helloBarEl.innerHTML = `
      <strong>Hello，${escapeHtml(username)}</strong>
      <span style="margin-left:8px; color:#6b7280;">${escapeHtml(quoteText)}</span>
    `;
  }
}

// =====================================
// 3. Suggest caches（分類/地點 + 同義詞）
// =====================================
const suggestCaches = {
  categories: [],
  places: [],
  categoryTextToId: new Map(), // key: normalized text -> uuid
  placeTextToId: new Map(),
};

function normText(s) {
  return String(s ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function mapTextToId(map, text, id) {
  const k = normText(text);
  if (!k) return;
  map.set(k, id);
}

async function loadSuggestCaches() {
  const user = await getCurrentUser();
  if (!user) return;

  // categories: builtin + own
  const { data: cats, error: cErr } = await supabaseClient
    .from("categories")
    .select("id, name, type, is_builtin, user_id")
    .order("is_builtin", { ascending: false })
    .order("name", { ascending: true });

  if (cErr) {
    logDebug("load categories error", cErr);
    return;
  }

  suggestCaches.categories = cats || [];
  suggestCaches.categoryTextToId = new Map();
  (cats || []).forEach((c) => mapTextToId(suggestCaches.categoryTextToId, c.name, c.id));

  // category_synonyms: builtin (user_id null) + own (user_id=auth.uid)
  const { data: cSyn, error: csErr } = await supabaseClient
    .from("category_synonyms")
    .select("category_id, synonym, user_id");

  if (csErr) {
    logDebug("load category_synonyms error", csErr);
  } else {
    (cSyn || []).forEach((s) => mapTextToId(suggestCaches.categoryTextToId, s.synonym, s.category_id));
  }

  // places: builtin + own
  const { data: pls, error: pErr } = await supabaseClient
    .from("places")
    .select("id, name, kind, is_builtin, user_id")
    .order("is_builtin", { ascending: false })
    .order("name", { ascending: true });

  if (pErr) {
    logDebug("load places error", pErr);
    return;
  }

  suggestCaches.places = pls || [];
  suggestCaches.placeTextToId = new Map();
  (pls || []).forEach((p) => mapTextToId(suggestCaches.placeTextToId, p.name, p.id));

  // place_synonyms: builtin + own
  const { data: pSyn, error: psErr } = await supabaseClient
    .from("place_synonyms")
    .select("place_id, synonym, user_id");

  if (psErr) {
    logDebug("load place_synonyms error", psErr);
  } else {
    (pSyn || []).forEach((s) => mapTextToId(suggestCaches.placeTextToId, s.synonym, s.place_id));
  }

  logDebug("Suggest caches loaded", {
    categories: suggestCaches.categories.length,
    places: suggestCaches.places.length,
    categoryTextToId: suggestCaches.categoryTextToId.size,
    placeTextToId: suggestCaches.placeTextToId.size,
  });

  // 更新 datalist（autocomplete）
  refreshDatalists();
}

function refreshDatalists() {
  if (categoryListEl) {
    categoryListEl.innerHTML = "";
    // datalist 只塞「名稱」（避免 synonym 太雜）
    suggestCaches.categories.slice(0, 300).forEach((c) => {
      const opt = document.createElement("option");
      opt.value = c.name;
      categoryListEl.appendChild(opt);
    });
  }

  if (placeListEl) {
    placeListEl.innerHTML = "";
    suggestCaches.places.slice(0, 300).forEach((p) => {
      const opt = document.createElement("option");
      opt.value = p.name;
      placeListEl.appendChild(opt);
    });
  }
}

function findCategoryIdByText(text) {
  const k = normText(text);
  if (!k) return null;
  return suggestCaches.categoryTextToId.get(k) || null;
}

function findPlaceIdByText(text) {
  const k = normText(text);
  if (!k) return null;
  return suggestCaches.placeTextToId.get(k) || null;
}

// =====================================
// 4. 手動輸入：自動建立自訂分類/地點（你要求的）
// =====================================
async function ensureCustomCategoryId(user, name, type) {
  const raw = String(name ?? "").trim();
  if (!raw) return null;

  const hit = findCategoryIdByText(raw);
  if (hit) return hit;

  const payload = {
    user_id: user.id,
    name: raw,
    type: type,
    grp: null,
    is_builtin: false,
  };

  const { data, error } = await supabaseClient
    .from("categories")
    .insert(payload)
    .select("id, name")
    .single();

  if (error) {
    logDebug("ensureCustomCategoryId insert error", error);
    throw error;
  }

  await ensureUserCategorySynonym(user, data.id, raw);
  mapTextToId(suggestCaches.categoryTextToId, raw, data.id);

  // 也更新 categories datalist
  suggestCaches.categories.push({ id: data.id, name: raw, type, is_builtin: false, user_id: user.id });
  refreshDatalists();

  return data.id;
}

async function ensureUserCategorySynonym(user, categoryId, synonymText) {
  const syn = String(synonymText ?? "").trim();
  if (!syn) return;

  // 先查有沒有（避免一直插造成 log 很吵）
  // 如果你覺得效能要更高，可以改成直接 insert 然後 ignore duplicate
  const { data, error } = await supabaseClient
    .from("category_synonyms")
    .select("id")
    .eq("user_id", user.id)
    .eq("category_id", categoryId)
    .eq("synonym", syn)
    .limit(1);

  if (!error && data && data.length > 0) return;

  const { error: iErr } = await supabaseClient
    .from("category_synonyms")
    .insert({ user_id: user.id, category_id: categoryId, synonym: syn });

  if (iErr) {
    logDebug("ensureCategorySynonym insert error", iErr);
  }
}

async function ensureCustomPlaceId(user, name) {
  const raw = String(name ?? "").trim();
  if (!raw) return null;

  const hit = findPlaceIdByText(raw);
  if (hit) return hit;

  const payload = {
    user_id: user.id,
    name: raw,
    kind: null,
    address: null,
    city: null,
    country: null,
    lat: null,
    lng: null,
    source: "manual",
    is_builtin: false,
  };

  const { data, error } = await supabaseClient
    .from("places")
    .insert(payload)
    .select("id, name")
    .single();

  if (error) {
    logDebug("ensureCustomPlaceId insert error", error);
    throw error;
  }

  await ensureUserPlaceSynonym(user, data.id, raw);
  mapTextToId(suggestCaches.placeTextToId, raw, data.id);

  suggestCaches.places.push({ id: data.id, name: raw, is_builtin: false, user_id: user.id });
  refreshDatalists();

  return data.id;
}

async function ensureUserPlaceSynonym(user, placeId, synonymText) {
  const syn = String(synonymText ?? "").trim();
  if (!syn) return;

  const { data, error } = await supabaseClient
    .from("place_synonyms")
    .select("id")
    .eq("user_id", user.id)
    .eq("place_id", placeId)
    .eq("synonym", syn)
    .limit(1);

  if (!error && data && data.length > 0) return;

  const { error: iErr } = await supabaseClient
    .from("place_synonyms")
    .insert({ user_id: user.id, place_id: placeId, synonym: syn });

  if (iErr) {
    logDebug("ensurePlaceSynonym insert error", iErr);
  }
}

// =====================================
// 5. Ledger：載入/新增/編輯/刪除
// =====================================
let ledgerData = [];
let editingId = null;
let currentFilterStart = null;
let currentFilterEnd = null;
let isSubmittingEntry = false;
let isSubmittingVoice = false;

async function loadLedger(startDate, endDate) {
  if (!ledgerTbodyEl) return;

  const user = await getCurrentUser();
  if (!user) {
    ledgerTbodyEl.innerHTML = '<tr><td colspan="7">請先登入</td></tr>';
    window.location.href = "index.html";
    return;
  }

  ledgerTbodyEl.innerHTML = '<tr><td colspan="7">載入中...</td></tr>';

  let query = supabaseClient
    .from("ledger_entries")
    .select(`
      id,
      occurred_at,
      type,
      amount,
      item,
      category_id,
      place_id,
      categories:category_id ( id, name ),
      places:place_id ( id, name )
    `)
    .order("occurred_at", { ascending: false })
    .order("id", { ascending: false });

  if (startDate) query = query.gte("occurred_at", startDate);
  if (endDate) query = query.lte("occurred_at", endDate);

  const { data, error } = await query;

  if (error) {
    ledgerTbodyEl.innerHTML = `<tr><td colspan="7">載入失敗：${error.message}</td></tr>`;
    logDebug("loadLedger error", error);
    return;
  }

  ledgerData = data || [];

  if (!ledgerData.length) {
    ledgerTbodyEl.innerHTML = '<tr><td colspan="7">目前沒有記帳資料</td></tr>';
    updateStats();
    return;
  }

  ledgerTbodyEl.innerHTML = "";
  ledgerData.forEach((row) => {
    const tr = document.createElement("tr");

    const typeLabel =
      row.type === "income"
        ? '<span class="badge-income">收入</span>'
        : row.type === "expense"
        ? '<span class="badge-expense">支出</span>'
        : '<span class="badge-expense">轉帳</span>';

    const catName = row.categories?.name || "";
    const placeName = row.places?.name || "";

    tr.innerHTML = `
      <td>${escapeHtml(row.occurred_at)}</td>
      <td>${typeLabel}</td>
      <td>${escapeHtml(catName)}</td>
      <td>${escapeHtml(row.item || "")}</td>
      <td>${escapeHtml(row.amount)}</td>
      <td>${escapeHtml(placeName)}</td>
      <td>
        <button type="button" class="btn-secondary" onclick="startEditEntry('${row.id}')">編輯</button>
        <button type="button" class="btn-secondary" onclick="deleteEntry('${row.id}')">刪除</button>
      </td>
    `;
    ledgerTbodyEl.appendChild(tr);
  });

  updateStats();
}

function startEditEntry(id) {
  const row = ledgerData.find((r) => r.id === id);
  if (!row) return;

  const dateEl = document.getElementById("entry-date");
  const typeEl = document.getElementById("entry-type");
  const categoryEl = document.getElementById("entry-category");
  const itemEl = document.getElementById("entry-item");
  const amountEl = document.getElementById("entry-amount");
  const placeEl = document.getElementById("entry-place");

  if (!dateEl || !typeEl || !categoryEl || !itemEl || !amountEl || !placeEl) {
    alert("找不到編輯表單。");
    return;
  }

  dateEl.value = row.occurred_at;
  typeEl.value = row.type;
  categoryEl.value = row.categories?.name || "";
  itemEl.value = row.item || "";
  amountEl.value = row.amount;
  placeEl.value = row.places?.name || "";

  editingId = id;

  if (editStatusEl) editStatusEl.textContent = `正在編輯（ID: ${id}）`;
  if (entrySubmitBtn) entrySubmitBtn.textContent = "儲存修改";
  if (entryCancelEditBtn) entryCancelEditBtn.classList.remove("hidden");
}

function cancelEditEntry() {
  editingId = null;

  const dateEl = document.getElementById("entry-date");
  const typeEl = document.getElementById("entry-type");
  const categoryEl = document.getElementById("entry-category");
  const itemEl = document.getElementById("entry-item");
  const amountEl = document.getElementById("entry-amount");
  const placeEl = document.getElementById("entry-place");

  if (dateEl) dateEl.value = "";
  if (typeEl) typeEl.value = "expense";
  if (categoryEl) categoryEl.value = "";
  if (itemEl) itemEl.value = "";
  if (amountEl) amountEl.value = "";
  if (placeEl) placeEl.value = "";

  if (editStatusEl) editStatusEl.textContent = "";
  if (entrySubmitBtn) entrySubmitBtn.textContent = "新增記帳";
  if (entryCancelEditBtn) entryCancelEditBtn.classList.add("hidden");
}

async function submitEntry() {
  const user = await getCurrentUser();
  if (!user) {
    alert("請先登入");
    window.location.href = "index.html";
    return;
  }

  if (isSubmittingEntry) return;
  isSubmittingEntry = true;

  if (entrySubmitBtn) {
    entrySubmitBtn.disabled = true;
    entrySubmitBtn.textContent = editingId ? "儲存中..." : "新增中...";
  }

  try {
    const dateEl = document.getElementById("entry-date");
    const typeEl = document.getElementById("entry-type");
    const categoryEl = document.getElementById("entry-category");
    const itemEl = document.getElementById("entry-item");
    const amountEl = document.getElementById("entry-amount");
    const placeEl = document.getElementById("entry-place");

    if (!dateEl || !typeEl || !categoryEl || !itemEl || !amountEl || !placeEl) {
      alert("這個頁面沒有完整的記帳表單。");
      return;
    }

    const occurred_at = dateEl.value;
    const type = typeEl.value;
    const categoryText = categoryEl.value.trim();
    const item = itemEl.value.trim();
    const amountStr = amountEl.value;
    const placeText = placeEl.value.trim();

    if (!occurred_at || !amountStr) {
      alert("請至少填寫日期與金額");
      return;
    }

    const amount = Number(amountStr);
    if (!Number.isFinite(amount) || amount <= 0) {
      alert("金額需為正數");
      return;
    }

    // 手動輸入：分類/地點找不到就建立自訂（你要的）
    const category_id = categoryText ? await ensureCustomCategoryId(user, categoryText, type) : null;
    const place_id = placeText ? await ensureCustomPlaceId(user, placeText) : null;

    const payload = {
      user_id: user.id,     // ✅ RLS 必要
      occurred_at,
      type,
      category_id,
      place_id,
      amount,
      item: item || null,
      source: "manual",
    };

    let error = null;
    if (editingId) {
      ({ error } = await supabaseClient
        .from("ledger_entries")
        .update(payload)
        .eq("id", editingId));
    } else {
      ({ error } = await supabaseClient
        .from("ledger_entries")
        .insert(payload));
    }

    if (error) {
      logDebug("submitEntry error", error);
      alert((editingId ? "儲存修改失敗：" : "新增記帳失敗：") + error.message);
      return;
    }

    cancelEditEntry();
    await loadSuggestCaches(); // 新增分類/地點後更新快取
    await loadLedger(currentFilterStart, currentFilterEnd);
  } finally {
    isSubmittingEntry = false;
    if (entrySubmitBtn) {
      entrySubmitBtn.disabled = false;
      entrySubmitBtn.textContent = editingId ? "儲存修改" : "新增記帳";
    }
  }
}

async function deleteEntry(id) {
  const ok = confirm("確定要刪除此筆記帳嗎？");
  if (!ok) return;

  const { error } = await supabaseClient
    .from("ledger_entries")
    .delete()
    .eq("id", id);

  if (error) {
    logDebug("deleteEntry error", error);
    alert("刪除失敗：" + error.message);
    return;
  }

  await loadLedger(currentFilterStart, currentFilterEnd);
}

// =====================================
// 6. 統計
// =====================================
function updateStats() {
  let income = 0;
  let expense = 0;
  const categorySum = new Map();

  for (const row of ledgerData) {
    const amt = Number(row.amount) || 0;
    if (row.type === "income") {
      income += amt;
    } else if (row.type === "expense") {
      expense += amt;
      const key = row.categories?.name || "未分類";
      categorySum.set(key, (categorySum.get(key) || 0) + amt);
    }
  }

  const net = income - expense;

  if (statIncomeEl) statIncomeEl.textContent = income.toString();
  if (statExpenseEl) statExpenseEl.textContent = expense.toString();
  if (statNetEl) statNetEl.textContent = net.toString();

  if (!statsCategoryBodyEl) return;

  if (categorySum.size === 0) {
    statsCategoryBodyEl.innerHTML = '<tr><td colspan="2">尚無資料</td></tr>';
    return;
  }

  statsCategoryBodyEl.innerHTML = "";
  for (const [cat, amt] of categorySum.entries()) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escapeHtml(cat)}</td><td>${escapeHtml(amt)}</td>`;
    statsCategoryBodyEl.appendChild(tr);
  }
}

// =====================================
// 7. 語音文字解析（支援：在全聯 / 地點全聯）
// 規則：語音不建立分類/地點；命中就填 id，否則 null 並補到 item
// =====================================
let pendingVoiceEntries = [];

function parseTextInput() {
  if (!voiceTextInputEl || !voiceStatusEl) return;

  const text = voiceTextInputEl.value.trim();
  if (!text) {
    alert("請先輸入語音文字。");
    return;
  }

  const newEntries = parseVoiceTextToEntries(text);
  if (!newEntries.length) {
    voiceStatusEl.textContent =
      "⚠ 無法解析，建議格式：今天 支出 餐飲 晚餐便當 250 在全聯";
    return;
  }

  pendingVoiceEntries.push(...newEntries);
  renderVoicePreview();
  voiceStatusEl.textContent = `✅ 新增解析 ${newEntries.length} 筆，目前共 ${pendingVoiceEntries.length} 筆暫存。`;
}

function clearVoiceText() {
  if (!voiceTextInputEl) return;
  voiceTextInputEl.value = "";
  if (voiceStatusEl) voiceStatusEl.textContent = "已清空文字輸入區。";
}

function clearVoiceEntries() {
  pendingVoiceEntries = [];
  renderVoicePreview();
  if (voiceStatusEl) voiceStatusEl.textContent = "已清除所有暫存資料。";
}

async function confirmVoiceEntries() {
  const user = await getCurrentUser();
  if (!user) {
    alert("請先登入");
    window.location.href = "index.html";
    return;
  }

  if (!pendingVoiceEntries.length) {
    alert("目前沒有任何暫存資料。");
    return;
  }

  if (isSubmittingVoice) return;
  isSubmittingVoice = true;

  if (voiceConfirmBtn) {
    voiceConfirmBtn.disabled = true;
    voiceConfirmBtn.textContent = "送出中...";
  }

  try {
    const ok = confirm(`確定要送出 ${pendingVoiceEntries.length} 筆語音記帳資料嗎？`);
    if (!ok) return;

    const payloads = pendingVoiceEntries.map((e) => {
      const catText = (e.categoryText || "").trim();
      const plcText = (e.placeText || "").trim();

      const category_id = catText ? findCategoryIdByText(catText) : null;
      const place_id = plcText ? findPlaceIdByText(plcText) : null;

      const extra = [];
      if (catText && !category_id) extra.push(`分類:${catText}`);
      if (plcText && !place_id) extra.push(`地點:${plcText}`);

      const mergedItem = [e.item || "", ...extra].filter(Boolean).join(" / ");

      return {
        user_id: user.id, // ✅ RLS 必要
        occurred_at: e.occurred_at,
        type: e.type,
        category_id: category_id,
        place_id: place_id,
        amount: Number(e.amount),
        item: mergedItem || null,
        source: "speech",
      };
    });

    const { error } = await supabaseClient
      .from("ledger_entries")
      .insert(payloads);

    if (error) {
      logDebug("voice insert error", error);
      alert("語音記帳寫入失敗：" + error.message);
      return;
    }

    pendingVoiceEntries = [];
    renderVoicePreview();
    if (voiceStatusEl) voiceStatusEl.textContent = `✅ 已成功寫入 ${payloads.length} 筆語音記帳。`;
    await loadLedger(currentFilterStart, currentFilterEnd);
  } finally {
    isSubmittingVoice = false;
    if (voiceConfirmBtn) {
      voiceConfirmBtn.disabled = false;
      voiceConfirmBtn.textContent = "✅ 確認送出";
    }
  }
}

function renderVoicePreview() {
  if (!voicePreviewBodyEl) return;

  if (!pendingVoiceEntries.length) {
    voicePreviewBodyEl.innerHTML = '<tr><td colspan="8">尚無語音解析資料</td></tr>';
    if (voiceConfirmBtn) voiceConfirmBtn.disabled = true;
    if (voiceClearBtn) voiceClearBtn.disabled = true;
    return;
  }

  if (voiceConfirmBtn) voiceConfirmBtn.disabled = false;
  if (voiceClearBtn) voiceClearBtn.disabled = false;

  voicePreviewBodyEl.innerHTML = "";
  pendingVoiceEntries.forEach((e, idx) => {
    const typeLabel = e.type === "income" ? "收入" : "支出";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${escapeHtml(e.occurred_at)}</td>
      <td>${escapeHtml(typeLabel)}</td>
      <td>${escapeHtml(e.categoryText || "")}</td>
      <td>${escapeHtml(e.item || "")}</td>
      <td>${escapeHtml(e.amount)}</td>
      <td>${escapeHtml(e.placeText || "")}</td>
      <td>
        <button type="button" class="btn-secondary" onclick="editVoiceEntry(${idx})">編輯</button>
        <button type="button" class="btn-secondary" onclick="removeVoiceEntry(${idx})">刪除</button>
      </td>
    `;
    voicePreviewBodyEl.appendChild(tr);
  });
}

function editVoiceEntry(index) {
  const e = pendingVoiceEntries[index];
  if (!e) return;

  const newDate = prompt("日期 (YYYY-MM-DD)", e.occurred_at) || e.occurred_at;
  let newType = prompt("類型（輸入：收入 或 支出）", e.type === "income" ? "收入" : "支出");
  newType = newType && newType.includes("收") ? "income" : "expense";

  const newCategory = prompt("分類", e.categoryText || "") ?? (e.categoryText || "");
  const newItem = prompt("項目", e.item || "") ?? (e.item || "");
  const newAmountStr = prompt("金額", String(e.amount)) ?? String(e.amount);
  const newAmount = Number(newAmountStr);
  const newPlace = prompt("地點（可空）", e.placeText || "") ?? (e.placeText || "");

  if (!newDate || !Number.isFinite(newAmount) || newAmount <= 0) {
    alert("日期或金額不合法，已取消修改。");
    return;
  }

  pendingVoiceEntries[index] = {
    occurred_at: newDate,
    type: newType,
    categoryText: String(newCategory || "").trim(),
    item: String(newItem || "").trim(),
    amount: newAmount,
    placeText: String(newPlace || "").trim(),
  };

  renderVoicePreview();
  if (voiceStatusEl) voiceStatusEl.textContent = `已更新第 ${index + 1} 筆暫存資料。`;
}

function removeVoiceEntry(index) {
  if (index < 0 || index >= pendingVoiceEntries.length) return;
  const ok = confirm(`確定要刪除第 ${index + 1} 筆暫存資料嗎？`);
  if (!ok) return;

  pendingVoiceEntries.splice(index, 1);
  renderVoicePreview();
  if (voiceStatusEl) voiceStatusEl.textContent = `刪除完成，目前剩餘 ${pendingVoiceEntries.length} 筆暫存資料。`;
}

// ===== 解析核心 =====
function parseVoiceTextToEntries(rawText) {
  const entries = [];
  if (!rawText) return entries;

  let cleaned = rawText.replace(/\s+/g, " ").trim();
  cleaned = cleaned.replace(/。/g, "，");

  const segments = cleaned.split(/下一筆|下ㄧ筆|下一笔/);

  for (const segRaw of segments) {
    const seg = segRaw.trim();
    if (!seg) continue;
    const entry = parseSingleVoiceEntry(seg);
    if (entry) entries.push(entry);
  }
  return entries;
}

function extractDateFromSegment(seg) {
  let rest = seg;

  // YYYY年M月D日
  let m = seg.match(/(\d{4})年(\d{1,2})月(\d{1,2})日?/);
  if (m) {
    const year = parseInt(m[1], 10);
    const month = parseInt(m[2], 10);
    const day = parseInt(m[3], 10);
    const dateStr = [year, String(month).padStart(2, "0"), String(day).padStart(2, "0")].join("-");
    rest = seg.replace(m[0], "");
    return { dateStr, rest };
  }

  // M月D日（年取今年）
  m = seg.match(/(\d{1,2})月(\d{1,2})日?/);
  if (m) {
    const now = new Date();
    const year = now.getFullYear();
    const month = parseInt(m[1], 10);
    const day = parseInt(m[2], 10);
    const dateStr = [year, String(month).padStart(2, "0"), String(day).padStart(2, "0")].join("-");
    rest = seg.replace(m[0], "");
    return { dateStr, rest };
  }

  // 今天 / 昨天 / 前天
  if (seg.includes("前天")) return { dateStr: formatDateFromOffset(2), rest: seg.replace("前天", "") };
  if (seg.includes("昨天")) return { dateStr: formatDateFromOffset(1), rest: seg.replace("昨天", "") };
  if (seg.includes("今天")) return { dateStr: formatDateFromOffset(0), rest: seg.replace("今天", "") };

  // 預設今天
  return { dateStr: formatDateFromOffset(0), rest };
}

// 支援地點關鍵字：在XXX / 地點XXX
function extractPlaceFromText(text) {
  let t = String(text || "");

  // 地點xxx
  let m = t.match(/地點\s*([^\s，,]+)\s*/);
  if (m) {
    const place = m[1].trim();
    t = t.replace(m[0], " ");
    return { placeText: place, rest: t.trim() };
  }

  // 在xxx
  m = t.match(/在\s*([^\s，,]+)\s*/);
  if (m) {
    const place = m[1].trim();
    t = t.replace(m[0], " ");
    return { placeText: place, rest: t.trim() };
  }

  return { placeText: "", rest: t.trim() };
}

// 單句解析：日期 類型 分類 項目 金額 地點
function parseSingleVoiceEntry(seg) {
  if (!seg) return null;

  const { dateStr, rest } = extractDateFromSegment(seg);

  // 先抽地點（可有可無）
  const { placeText, rest: afterPlace } = extractPlaceFromText(rest);

  const base = afterPlace.trim();
  if (!base) return null;

  // 類型
  let type = null;
  if (base.includes("收入")) type = "income";
  else if (base.includes("支出")) type = "expense";
  else return null;

  // 把類型字拿掉，剩下 "分類 項目 金額 ..."
  const afterType = base.replace("收入", " ").replace("支出", " ").trim();

  // 金額：抓第一個數字（最後也可）
  const amountMatch = afterType.match(/(\d+)(元)?/);
  if (!amountMatch) return null;
  const amount = parseInt(amountMatch[1], 10);

  // 金額前的文字：分類 + 項目（用第一個空白切）
  const beforeAmount = afterType.slice(0, amountMatch.index).trim();
  const parts = beforeAmount.split(/\s+/).filter(Boolean);

  const categoryText = parts[0] ? parts[0].trim() : "";
  const item = parts.length >= 2 ? parts.slice(1).join(" ").trim() : "";

  // 金額後的文字若還有，補到 item（更耐用）
  const afterAmount = afterType.slice(amountMatch.index + amountMatch[0].length).trim();
  const mergedItem = [item, afterAmount].filter(Boolean).join(" ").trim();

  return {
    occurred_at: dateStr,
    type,
    categoryText,
    item: mergedItem,
    amount,
    placeText,
  };
}

function formatDateFromOffset(offset) {
  const d = new Date();
  d.setDate(d.getDate() - offset);
  return formatDate(d);
}

function formatDate(d) {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

// =====================================
// 8. 日期篩選 & CSV 匯出
// =====================================
function applyDateFilter() {
  const startEl = document.getElementById("filter-start");
  const endEl = document.getElementById("filter-end");
  currentFilterStart = startEl?.value || null;
  currentFilterEnd = endEl?.value || null;
  loadLedger(currentFilterStart, currentFilterEnd);
}

function quickRange(type) {
  const today = new Date();
  const endStr = formatDate(today);
  let start = new Date();

  if (type === "14d") start.setDate(today.getDate() - 14);
  else if (type === "30d") start.setDate(today.getDate() - 30);

  const startStr = formatDate(start);

  const startEl = document.getElementById("filter-start");
  const endEl = document.getElementById("filter-end");
  if (startEl) startEl.value = startStr;
  if (endEl) endEl.value = endStr;

  currentFilterStart = startStr;
  currentFilterEnd = endStr;

  loadLedger(startStr, endStr);
}

function exportCsv() {
  if (!ledgerData || !ledgerData.length) {
    alert("目前沒有可以匯出的資料。");
    return;
  }

  const header = ["日期", "類型", "分類", "項目", "金額", "地點"];
  const rows = [header];

  ledgerData.forEach((row) => {
    const typeLabel = row.type === "income" ? "收入" : "支出";
    rows.push([
      row.occurred_at,
      typeLabel,
      row.categories?.name || "",
      row.item || "",
      row.amount,
      row.places?.name || "",
    ]);
  });

  const csvContent = rows
    .map((cols) =>
      cols
        .map((c) => {
          const v = String(c ?? "");
          if (/[",\n]/.test(v)) return `"${v.replace(/"/g, '""')}"`;
          return v;
        })
        .join(",")
    )
    .join("\n");

  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const start = currentFilterStart || "all";
  const end = currentFilterEnd || "all";
  const fileName = `ledger_${start}_${end}.csv`;

  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.href = url;
  link.setAttribute("download", fileName);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// =====================================
// 9. 安全：escape HTML
// =====================================
function escapeHtml(str) {
  return String(str ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

// =====================================
// 10. DOMContentLoaded：依頁面初始化
// =====================================
document.addEventListener("DOMContentLoaded", async () => {
  logDebug("page loaded");

  const isIndexPage = document.getElementById("index-page") !== null;
  const isLedgerPage = document.getElementById("ledger-page") !== null;

  if (isIndexPage) {
    const user = await getCurrentUser();
    if (user) {
      logDebug("index: 已登入，跳轉 ledger", user);
      window.location.href = "ledger.html";
      return;
    }
    return;
  }

  if (isLedgerPage) {
    const user = await getCurrentUser();
    if (!user) {
      logDebug("ledger: 尚未登入，回 index");
      window.location.href = "index.html";
      return;
    }

    await refreshHelloBar();

    // 先載入分類/地點快取（語音解析 & autocomplete 會用到）
    await loadSuggestCaches();

    // 預設：最近 14 天
    const today = new Date();
    const start = new Date();
    start.setDate(today.getDate() - 14);
    currentFilterStart = formatDate(start);
    currentFilterEnd = formatDate(today);

    const startEl = document.getElementById("filter-start");
    const endEl = document.getElementById("filter-end");
    if (startEl) startEl.value = currentFilterStart;
    if (endEl) endEl.value = currentFilterEnd;

    await loadLedger(currentFilterStart, currentFilterEnd);

    // 綁定語音按鈕
    if (parseTextBtn) parseTextBtn.addEventListener("click", parseTextInput);
    if (voiceTextClearBtn) voiceTextClearBtn.addEventListener("click", clearVoiceText);
    if (voiceConfirmBtn) voiceConfirmBtn.addEventListener("click", confirmVoiceEntries);
    if (voiceClearBtn) voiceClearBtn.addEventListener("click", clearVoiceEntries);

    renderVoicePreview();
  }
});
