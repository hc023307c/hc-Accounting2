// ===============================
// 0) Supabase 設定
// ===============================
const SUPABASE_URL = "https://kjcxngzrrncotukoxbze.supabase.co";
const SUPABASE_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtqY3huZ3pycm5jb3R1a294YnplIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Njg5ODg1NjEsImV4cCI6MjA4NDU2NDU2MX0.MgYNIEhW9v5dempDoFSvoM5foom5ST8t9hkx_0_qHvo";

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

// ===============================
// 1) Debug
// ===============================
const jsStatusEl = document.getElementById("js-status");
const debugEl = document.getElementById("debug");

function logDebug(msg, obj) {
  if (!debugEl) return;
  const text = msg + (obj ? " " + JSON.stringify(obj, null, 2) : "");
  debugEl.textContent += text + "\n";
  console.log("[DEBUG]", msg, obj || "");
}

function logRlsHint(err, tableName) {
  // 給你快速定位：這不是前端壞，是 RLS / 欄位 / default / policy 的問題
  if (!err) return;
  const msg = (err.message || "").toLowerCase();
  if (msg.includes("row-level security") || err.code === "42501") {
    logDebug(`⚠ RLS 擋住：${tableName}`, {
      hint:
        "這通常代表 policy 的 WITH CHECK 不滿足（例如 insert 時 user_id 沒填、或 policy 要求 user_id=auth.uid()）。" +
        " 也可能是你表格根本沒有 user_id 欄位，導致 policy 判斷永遠不成立。" +
        " 下一步要先用 SQL 確認 ledger_entries / category_synonyms / place_synonyms 是否有 user_id 欄位、default、與正確 policy。",
    });
  }
}

if (jsStatusEl) jsStatusEl.textContent = "✅ JS 已載入，Supabase client 建立完成。";

// page markers
const isIndexPage = document.getElementById("index-page") !== null;
const isLedgerPage = document.getElementById("ledger-page") !== null;

// index DOM
const authStatusEl = document.getElementById("auth-status");

// ledger DOM
const helloLineEl = document.getElementById("hello-line");
const zenLineEl = document.getElementById("zen-line");
const editStatusEl = document.getElementById("edit-status");
const entrySubmitBtn = document.getElementById("entry-submit-btn");
const entryCancelEditBtn = document.getElementById("entry-cancel-edit-btn");
const ledgerTbodyEl = document.getElementById("ledger-tbody");
// date filter DOM
const filterStartEl = document.getElementById("filter-start");
const filterEndEl = document.getElementById("filter-end");


// manual hint DOM
const entryCategoryHitEl = document.getElementById("entry-category-hit");
const entryPlaceHitEl = document.getElementById("entry-place-hit");

// datalist DOM
const categorySuggestEl = document.getElementById("category-suggest");
const placeSuggestEl = document.getElementById("place-suggest");

// voice DOM
const voiceTextInputEl = document.getElementById("voice-text-input");
const voiceStatusEl = document.getElementById("voice-status");
const voicePreviewBodyEl = document.getElementById("voice-preview-body");
const parseTextBtn = document.getElementById("parse-text-btn");
const voiceTextClearBtn = document.getElementById("voice-text-clear-btn");
const voiceConfirmBtn = document.getElementById("voice-confirm-btn");
const voiceClearBtn = document.getElementById("voice-clear-btn");

// ===============================
// 2) Auth
// ===============================
function accountToEmail(account) {
  return account + "@demo.local";
}

async function handleLogin() {
  const accountInput = document.getElementById("login-account");
  const passwordInput = document.getElementById("login-password");
  if (!accountInput || !passwordInput) return;

  const account = accountInput.value.trim();
  const password = passwordInput.value;

  if (!account || !password) {
    if (authStatusEl) authStatusEl.textContent = "請輸入帳號與密碼";
    alert("請輸入帳號與密碼");
    return;
  }

  const email = accountToEmail(account);
  logDebug("嘗試登入", { email });

  const { data, error } = await supabaseClient.auth.signInWithPassword({ email, password });

  if (error) {
    if (authStatusEl) authStatusEl.textContent = "登入失敗：" + error.message;
    alert("登入失敗：" + error.message);
    logDebug("登入失敗", error);
    return;
  }

  if (authStatusEl) authStatusEl.textContent = "登入成功：" + (data.user?.email || "");
  location.href = "ledger.html";
}

async function handleLogout() {
  await supabaseClient.auth.signOut();
  location.href = "index.html";
}

async function getCurrentUser() {
  const { data, error } = await supabaseClient.auth.getUser();
  if (error) return null;
  return data.user || null;
}

async function loadProfile(userId) {
  const { data, error } = await supabaseClient
    .from("profiles")
    .select("display_name, account_code, role")
    .eq("id", userId)
    .maybeSingle(); // ✅ 比 single 更不容易炸

  if (error) {
    logDebug("載入 profile 失敗", error);
    return null;
  }
  return data || null;
}


// ===============================
// 3) Zen Quote（108 自在語）
// ===============================
async function loadZenQuote() {
  if (!zenLineEl) return;

  try {
    // 1) 先用「你原本的欄位」嘗試（content/source + enabled）
    let { data, error } = await supabaseClient
      .from("zen_quotes")
      .select("content, source, enabled")
      .eq("enabled", true)
      .limit(300);

    // 2) 如果是「欄位不存在」(常見：enabled/content/source 不在你現在表)
    //    就 fallback：直接 select * 不加 enabled 條件
    if (error && (String(error.message || "").includes("column") || error.code === "42703")) {
      logDebug("zen_quotes 欄位可能不同，改用 fallback select(*)", error);

      ({ data, error } = await supabaseClient
        .from("zen_quotes")
        .select("*")
        .limit(300));
    }

    // 3) 如果是 RLS/權限錯
    if (error) {
      zenLineEl.textContent = "（自在語載入失敗：請看 Debug）";
      logDebug("zen_quotes error", error);
      return;
    }

    if (!data || !data.length) {
      zenLineEl.textContent = "（目前沒有自在語）";
      logDebug("zen_quotes empty", { rows: 0 });
      return;
    }

    // 4) 兼容不同欄位命名：content/text/quote + source/author
    const pick = data[Math.floor(Math.random() * data.length)];
    const content =
      pick.content ?? pick.text ?? pick.quote ?? pick.body ?? pick.message ?? "";
    const source =
      pick.source ?? pick.author ?? pick.from ?? pick.ref ?? "";

    if (!String(content).trim()) {
      zenLineEl.textContent = "（自在語資料有，但找不到內容欄位）";
      logDebug("zen_quotes missing content field", pick);
      return;
    }

    zenLineEl.textContent = `「${content}」${source ? " — " + source : ""}`;
    logDebug("zen_quotes ok", { content, source });
  } catch (e) {
    zenLineEl.textContent = "（自在語載入失敗：JS 例外，請看 Debug）";
    logDebug("zen_quotes exception", { message: String(e?.message || e), e });
  }
}


// ===============================
// 4) 日期工具
// ===============================
function formatDate(d) {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

// ===============================
// 5) Autocomplete 快取（分類/地點 + 同義詞）
// ===============================
const categoryTextToId = new Map(); // key: anyText(name or synonym) -> category_id
const categoryIdToName = new Map(); // id -> canonical name
const placeTextToId = new Map(); // key: anyText(name or synonym) -> place_id
const placeIdToName = new Map(); // id -> canonical name

let suggestCategories = []; // datalist 顯示用（去重）
let suggestPlaces = []; // datalist 顯示用（去重）

function normKey(s) {
  return (s || "").trim().toLowerCase();
}

function setHint(el, ok, text) {
  if (!el) return;
  el.classList.remove("ok", "warn");
  if (!text) {
    el.textContent = "";
    return;
  }
  el.textContent = text;
  el.classList.add(ok ? "ok" : "warn");
}

async function loadSuggestCaches() {
  // ---- categories
  {
    const { data: cats, error } = await supabaseClient.from("categories").select("id, name").limit(1000);
    if (error) {
      logDebug("load categories error", error);
    } else {
      for (const c of cats || []) {
        categoryIdToName.set(c.id, c.name);
        categoryTextToId.set(normKey(c.name), c.id);
      }
    }
  }

  // ---- category_synonyms
  {
    const { data: syns, error } = await supabaseClient
      .from("category_synonyms")
      .select("category_id, synonym")
      .limit(3000);

    if (error) {
      logDebug("load category_synonyms error", error);
    } else {
      for (const s of syns || []) {
        if (!s.synonym || !s.category_id) continue;
        categoryTextToId.set(normKey(s.synonym), s.category_id);
      }
    }
  }

  // ---- places
  {
    const { data: places, error } = await supabaseClient.from("places").select("id, name").limit(1000);

    if (error) {
      logDebug("load places error", error);
    } else {
      for (const p of places || []) {
        placeIdToName.set(p.id, p.name);
        placeTextToId.set(normKey(p.name), p.id);
      }
    }
  }

  // ---- place_synonyms
  {
    const { data: ps, error } = await supabaseClient.from("place_synonyms").select("place_id, synonym").limit(3000);

    if (error) {
      logDebug("load place_synonyms error", error);
    } else {
      for (const s of ps || []) {
        if (!s.synonym || !s.place_id) continue;
        placeTextToId.set(normKey(s.synonym), s.place_id);
      }
    }
  }

  suggestCategories = Array.from(new Set(Array.from(categoryIdToName.values()))).sort();
  suggestPlaces = Array.from(new Set(Array.from(placeIdToName.values()))).sort();

  renderDatalist(categorySuggestEl, suggestCategories);
  renderDatalist(placeSuggestEl, suggestPlaces);

  logDebug("Suggest caches loaded", {
    categories: suggestCategories.length,
    places: suggestPlaces.length,
    categoryTextToId: categoryTextToId.size,
    placeTextToId: placeTextToId.size,
  });
}

function renderDatalist(datalistEl, items) {
  if (!datalistEl) return;
  datalistEl.innerHTML = "";
  const frag = document.createDocumentFragment();
  for (const it of items.slice(0, 1000)) {
    const opt = document.createElement("option");
    opt.value = it;
    frag.appendChild(opt);
  }
  datalistEl.appendChild(frag);
}

function resolveCategoryIdFast(text) {
  const key = normKey(text);
  if (!key) return null;
  return categoryTextToId.get(key) || null;
}

function resolvePlaceIdFast(text) {
  const key = normKey(text);
  if (!key) return null;
  return placeTextToId.get(key) || null;
}

function getCategoryNameById(id) {
  return categoryIdToName.get(id) || "";
}

function getPlaceNameById(id) {
  return placeIdToName.get(id) || "";
}

// 手動輸入即時提示（不會改資料，只顯示命中）
function wireManualHitHints() {
  const catInput = document.getElementById("entry-category");
  const placeInput = document.getElementById("entry-place");

  if (catInput) {
    catInput.addEventListener("input", () => {
      const id = resolveCategoryIdFast(catInput.value);
      if (id) setHint(entryCategoryHitEl, true, `✅ 命中：${getCategoryNameById(id)}`);
      else if (catInput.value.trim()) setHint(entryCategoryHitEl, false, "⚠ 未命中（送出時會建立自訂分類）");
      else setHint(entryCategoryHitEl, true, "");
    });
  }

  if (placeInput) {
    placeInput.addEventListener("input", () => {
      const id = resolvePlaceIdFast(placeInput.value);
      if (id) setHint(entryPlaceHitEl, true, `✅ 命中：${getPlaceNameById(id)}`);
      else if (placeInput.value.trim()) setHint(entryPlaceHitEl, false, "⚠ 未命中（送出時會建立自訂地點）");
      else setHint(entryPlaceHitEl, true, "");
    });
  }
}

// ===============================
// 6) ✅ 自訂分類/地點建立（僅限手動表單）
// ===============================
async function ensureCategorySynonym(categoryId, synonym) {
  const s = (synonym || "").trim();
  if (!categoryId || !s) return;

  const { data, error } = await supabaseClient
    .from("category_synonyms")
    .select("id")
    .eq("category_id", categoryId)
    .eq("synonym", s)
    .limit(1)
    .maybeSingle();

  if (!error && data?.id) return;

  const { error: insErr } = await supabaseClient.from("category_synonyms").insert({ category_id: categoryId, synonym: s });

  if (insErr) {
    logDebug("ensureCategorySynonym insert error", insErr);
    logRlsHint(insErr, "category_synonyms");
  }
}

async function ensurePlaceSynonym(placeId, synonym) {
  const s = (synonym || "").trim();
  if (!placeId || !s) return;

  const { data, error } = await supabaseClient
    .from("place_synonyms")
    .select("id")
    .eq("place_id", placeId)
    .eq("synonym", s)
    .limit(1)
    .maybeSingle();

  if (!error && data?.id) return;

  const { error: insErr } = await supabaseClient.from("place_synonyms").insert({ place_id: placeId, synonym: s });

  if (insErr) {
    logDebug("ensurePlaceSynonym insert error", insErr);
    logRlsHint(insErr, "place_synonyms");
  }
}

async function createCustomCategory(userId, name, type) {
  const n = (name || "").trim();
  if (!userId || !n) return null;

  const payload = {
    user_id: userId,
    name: n,
    type: type === "income" ? "income" : "expense",
    grp: null,
    is_builtin: false,
  };

  const { data, error } = await supabaseClient.from("categories").insert(payload).select("id, name").single();

  if (error) {
    logDebug("createCustomCategory error", error);
    logRlsHint(error, "categories");
    return null;
  }

  categoryIdToName.set(data.id, data.name);
  categoryTextToId.set(normKey(data.name), data.id);
  suggestCategories = Array.from(new Set(Array.from(categoryIdToName.values()))).sort();
  renderDatalist(categorySuggestEl, suggestCategories);

  return data?.id || null;
}

async function createCustomPlace(userId, name) {
  const n = (name || "").trim();
  if (!userId || !n) return null;

  const payload = {
    user_id: userId,
    name: n,
    kind: null,
    is_builtin: false,
    source: "manual",
  };

  const { data, error } = await supabaseClient.from("places").insert(payload).select("id, name").single();

  if (error) {
    logDebug("createCustomPlace error", error);
    logRlsHint(error, "places");
    return null;
  }

  placeIdToName.set(data.id, data.name);
  placeTextToId.set(normKey(data.name), data.id);
  suggestPlaces = Array.from(new Set(Array.from(placeIdToName.values()))).sort();
  renderDatalist(placeSuggestEl, suggestPlaces);

  return data?.id || null;
}

async function resolveOrCreateCategoryIdForManual(userId, catText, type) {
  const text = (catText || "").trim();
  if (!text) return null;

  let id = resolveCategoryIdFast(text);
  if (id) {
    await ensureCategorySynonym(id, text);
    categoryTextToId.set(normKey(text), id);
    return id;
  }

  id = await createCustomCategory(userId, text, type);
  if (!id) return null;

  await ensureCategorySynonym(id, text);
  categoryTextToId.set(normKey(text), id);
  return id;
}

async function resolveOrCreatePlaceIdForManual(userId, placeText) {
  const text = (placeText || "").trim();
  if (!text) return null;

  let id = resolvePlaceIdFast(text);
  if (id) {
    await ensurePlaceSynonym(id, text);
    placeTextToId.set(normKey(text), id);
    return id;
  }

  id = await createCustomPlace(userId, text);
  if (!id) return null;

  await ensurePlaceSynonym(id, text);
  placeTextToId.set(normKey(text), id);
  return id;
}

// ===============================
// 7) Ledger：列表 + 編輯/刪除
// ===============================
let ledgerRows = [];
let editingEntryId = null;
let isSubmittingEntry = false;

// ✅ 日期篩選狀態（會影響 loadLedger 與 CSV 匯出）
let currentFilterStart = null;
let currentFilterEnd = null;


async function loadLedger(startDate, endDate) {
  if (!ledgerTbodyEl) return;

  const user = await getCurrentUser();
  if (!user) {
    ledgerTbodyEl.innerHTML = '<tr><td colspan="7">請先登入</td></tr>';
    location.href = "index.html";
    return;
  }

  ledgerTbodyEl.innerHTML = '<tr><td colspan="7">載入中...</td></tr>';

  let q = supabaseClient
    .from("ledger_entries")
    .select(`
      id,
      occurred_at,
      type,
      amount,
      item,
      category_id,
      place_id,
      categories(name),
      places(name)
    `)
    .order("occurred_at", { ascending: false })
    .order("id", { ascending: false });

  // ✅ 套用日期範圍（只影響當前畫面/CSV匯出）
  if (startDate) q = q.gte("occurred_at", startDate);
  if (endDate) q = q.lte("occurred_at", endDate);

  const { data, error } = await q;

  if (error) {
    ledgerTbodyEl.innerHTML = `<tr><td colspan="7">載入失敗：${error.message}</td></tr>`;
    logDebug("loadLedger error", error);
    return;
  }

  ledgerRows = data || [];
  if (!ledgerRows.length) {
    ledgerTbodyEl.innerHTML = '<tr><td colspan="7">目前沒有記帳資料</td></tr>';
    return;
  }

  ledgerTbodyEl.innerHTML = "";
  for (const r of ledgerRows) {
    const typeLabel =
      r.type === "income"
        ? '<span class="badge-income">收入</span>'
        : '<span class="badge-expense">支出</span>';

    const catName = r.categories?.name || "";
    const placeName = r.places?.name || "";

    ledgerTbodyEl.innerHTML += `
      <tr>
        <td>${r.occurred_at}</td>
        <td>${typeLabel}</td>
        <td>${escapeHtml(catName)}</td>
        <td>${escapeHtml(r.item || "")}</td>
        <td>${r.amount}</td>
        <td>${escapeHtml(placeName)}</td>
        <td>
          <button type="button" class="btn-secondary" onclick="startEditEntry('${r.id}')">編輯</button>
          <button type="button" class="btn-secondary" onclick="deleteEntry('${r.id}')">刪除</button>
        </td>
      </tr>
    `;
  }
}
// ===============================
// D) 日期範圍篩選 + 快捷 + CSV 匯出
// ===============================

function applyDateFilter() {
  currentFilterStart = filterStartEl?.value || null;
  currentFilterEnd = filterEndEl?.value || null;
  loadLedger(currentFilterStart, currentFilterEnd);
}

// 快捷：一週 / 兩週 / 一個月 / 三個月
function quickRange(type) {
  const today = new Date();
  const endStr = formatDate(today);

  const start = new Date();
  if (type === "7d") start.setDate(today.getDate() - 7);
  else if (type === "14d") start.setDate(today.getDate() - 14);
  else if (type === "30d") start.setDate(today.getDate() - 30);
  else if (type === "90d") start.setDate(today.getDate() - 90);
  else start.setDate(today.getDate() - 7);

  const startStr = formatDate(start);

  if (filterStartEl) filterStartEl.value = startStr;
  if (filterEndEl) filterEndEl.value = endStr;

  currentFilterStart = startStr;
  currentFilterEnd = endStr;

  loadLedger(startStr, endStr);
}

// 匯出目前畫面（已套用篩選）的 CSV
function exportCsv() {
  if (!ledgerRows || !ledgerRows.length) {
    alert("目前沒有可以匯出的資料。");
    return;
  }

  const header = ["日期", "類型", "分類", "項目", "金額", "地點"];
  const rows = [header];

  ledgerRows.forEach((r) => {
    const typeLabel = r.type === "income" ? "收入" : "支出";
    rows.push([
      r.occurred_at,
      typeLabel,
      r.categories?.name || "",
      r.item || "",
      r.amount,
      r.places?.name || "",
    ]);
  });

  const csv = rows
    .map((cols) =>
      cols
        .map((v) => {
          const s = String(v ?? "");
          return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
        })
        .join(",")
    )
    .join("\n");

  // ✅ 關鍵：加上 UTF-8 BOM，Excel 才不會亂碼
  const BOM = "\uFEFF";
  const blob = new Blob([BOM + csv], { type: "text/csv;charset=utf-8;" });

  const start = currentFilterStart || "all";
  const end = currentFilterEnd || "all";
  const fileName = `ledger_${start}_${end}.csv`;

  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}



function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function fillFormForRow(row) {
  const dateEl = document.getElementById("entry-date");
  const typeEl = document.getElementById("entry-type");
  const catEl = document.getElementById("entry-category");
  const itemEl = document.getElementById("entry-item");
  const amtEl = document.getElementById("entry-amount");
  const placeEl = document.getElementById("entry-place");

  if (dateEl) dateEl.value = row.occurred_at;
  if (typeEl) typeEl.value = row.type;
  if (catEl) catEl.value = row.categories?.name || "";
  if (itemEl) itemEl.value = row.item || "";
  if (amtEl) amtEl.value = row.amount;
  if (placeEl) placeEl.value = row.places?.name || "";
}

function clearForm() {
  const dateEl = document.getElementById("entry-date");
  const typeEl = document.getElementById("entry-type");
  const catEl = document.getElementById("entry-category");
  const itemEl = document.getElementById("entry-item");
  const amtEl = document.getElementById("entry-amount");
  const placeEl = document.getElementById("entry-place");

  if (dateEl) dateEl.value = formatDate(new Date());
  if (typeEl) typeEl.value = "expense";
  if (catEl) catEl.value = "";
  if (itemEl) itemEl.value = "";
  if (amtEl) amtEl.value = "";
  if (placeEl) placeEl.value = "";

  setHint(entryCategoryHitEl, true, "");
  setHint(entryPlaceHitEl, true, "");
}

function startEditEntry(id) {
  const row = ledgerRows.find((r) => r.id === id);
  if (!row) return;

  editingEntryId = id;
  fillFormForRow(row);

  if (editStatusEl) {
    editStatusEl.textContent = `正在編輯：${row.occurred_at} / ${row.categories?.name || "未分類"} / ${row.item || ""}（ID: ${id}）`;
  }
  if (entrySubmitBtn) entrySubmitBtn.textContent = "儲存修改";
  if (entryCancelEditBtn) entryCancelEditBtn.classList.remove("hidden");
}

function cancelEditEntry() {
  editingEntryId = null;
  clearForm();

  if (editStatusEl) editStatusEl.textContent = "";
  if (entrySubmitBtn) entrySubmitBtn.textContent = "新增記帳";
  if (entryCancelEditBtn) entryCancelEditBtn.classList.add("hidden");
}

async function submitEntry() {
  const user = await getCurrentUser();
  if (!user) {
    alert("請先登入");
    location.href = "index.html";
    return;
  }

  if (isSubmittingEntry) return;
  isSubmittingEntry = true;

  try {
    if (entrySubmitBtn) {
      entrySubmitBtn.disabled = true;
      entrySubmitBtn.textContent = editingEntryId ? "儲存中..." : "新增中...";
    }

    const date = document.getElementById("entry-date")?.value;
    const type = document.getElementById("entry-type")?.value || "expense";
    const catText = document.getElementById("entry-category")?.value?.trim() || "";
    const item = document.getElementById("entry-item")?.value?.trim() || "";
    const amountStr = document.getElementById("entry-amount")?.value;
    const placeText = document.getElementById("entry-place")?.value?.trim() || "";

    if (!date || !amountStr) return alert("請至少填寫日期與金額");

    const amount = parseFloat(amountStr);
    if (!Number.isFinite(amount) || amount <= 0) return alert("金額需為正數");

    // ✅ 手動輸入：分類/地點查不到就建立自訂項 + synonym
    const category_id = await resolveOrCreateCategoryIdForManual(user.id, catText, type);
    const place_id = await resolveOrCreatePlaceIdForManual(user.id, placeText);

const payload = {
  user_id: user.id, // ✅ 新增這行（RLS 需要）
  occurred_at: date,
  type,
  amount,
  item: item || null,
  category_id: category_id || null,
  place_id: place_id || null,
  source: "manual",
  updated_at: new Date().toISOString(),
};

    let error = null;
    if (editingEntryId) {
      ({ error } = await supabaseClient.from("ledger_entries").update(payload).eq("id", editingEntryId));
    } else {
      ({ error } = await supabaseClient.from("ledger_entries").insert(payload));
    }

    if (error) {
      alert((editingEntryId ? "儲存修改失敗：" : "新增失敗：") + error.message);
      logDebug("submitEntry error", error);
      logRlsHint(error, "ledger_entries");
      return;
    }

    cancelEditEntry();
    await loadLedger();
  } finally {
    isSubmittingEntry = false;
    if (entrySubmitBtn) {
      entrySubmitBtn.disabled = false;
      entrySubmitBtn.textContent = editingEntryId ? "儲存修改" : "新增記帳";
    }
  }
}

async function deleteEntry(id) {
  const ok = confirm("確定要刪除此筆記帳嗎？");
  if (!ok) return;

  const { error } = await supabaseClient.from("ledger_entries").delete().eq("id", id);

  if (error) {
    alert("刪除失敗：" + error.message);
    logDebug("delete error", error);
    logRlsHint(error, "ledger_entries");
    return;
  }

  if (editingEntryId === id) cancelEditEntry();
  await loadLedger();
}

// ===============================
// 8) 語音：解析 + 預覽可編輯（但仍不建立）
// ===============================
let pendingVoiceEntries = [];
let isSubmittingVoice = false;

function normalizeText(s) {
  if (!s) return "";
  return s
    .replace(/[，、。]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function dateFromOffset(daysAgo) {
  const d = new Date();
  d.setDate(d.getDate() - daysAgo);
  return formatDate(d);
}

function extractDate(seg) {
  let rest = seg;

  // 0) YYYY/MM/DD or YYYY-M-D
  let m = rest.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    const y = parseInt(m[1], 10);
    const mo = String(parseInt(m[2], 10)).padStart(2, "0");
    const da = String(parseInt(m[3], 10)).padStart(2, "0");
    rest = rest.replace(m[0], " ");
    return { dateStr: `${y}-${mo}-${da}`, rest };
  }

  // 1) YYYY年M月D日
  m = rest.match(/(\d{4})年(\d{1,2})月(\d{1,2})日?/);
  if (m) {
    const y = parseInt(m[1], 10);
    const mo = String(parseInt(m[2], 10)).padStart(2, "0");
    const da = String(parseInt(m[3], 10)).padStart(2, "0");
    rest = rest.replace(m[0], " ");
    return { dateStr: `${y}-${mo}-${da}`, rest };
  }

  // 2) M/D or M-D（年取今年）✅你現在缺的就是這個
  m = rest.match(/(\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    const now = new Date();
    const y = now.getFullYear();
    const mo = String(parseInt(m[1], 10)).padStart(2, "0");
    const da = String(parseInt(m[2], 10)).padStart(2, "0");
    rest = rest.replace(m[0], " ");
    return { dateStr: `${y}-${mo}-${da}`, rest };
  }

  // 3) M月D日（年取今年）
  m = rest.match(/(\d{1,2})月(\d{1,2})日?/);
  if (m) {
    const now = new Date();
    const y = now.getFullYear();
    const mo = String(parseInt(m[1], 10)).padStart(2, "0");
    const da = String(parseInt(m[2], 10)).padStart(2, "0");
    rest = rest.replace(m[0], " ");
    return { dateStr: `${y}-${mo}-${da}`, rest };
  }

  // 4) 今天 / 昨天 / 前天
  if (rest.includes("前天")) {
    rest = rest.replace("前天", " ");
    return { dateStr: dateFromOffset(2), rest };
  }
  if (rest.includes("昨天")) {
    rest = rest.replace("昨天", " ");
    return { dateStr: dateFromOffset(1), rest };
  }
  if (rest.includes("今天")) {
    rest = rest.replace("今天", " ");
    return { dateStr: dateFromOffset(0), rest };
  }

  // 5) 都沒抓到 → 預設今天
  return { dateStr: dateFromOffset(0), rest };
}


// 支援 B：地點關鍵字：「在全聯」或「地點全聯」或「地點:全聯」
function extractPlaceKeyword(seg) {
  let placeText = null;
  let rest = seg;

  const re = /(在|地點)\s*[:：]?\s*([^\s]+)/;
  const m = rest.match(re);
  if (m) {
    placeText = m[2];
    rest = rest.replace(m[0], " ");
  }
  return { placeText, rest };
}

function extractType(seg) {
  let rest = seg;
  if (rest.includes("收入")) {
    rest = rest.replace("收入", " ");
    return { type: "income", rest };
  }
  if (rest.includes("支出")) {
    rest = rest.replace("支出", " ");
    return { type: "expense", rest };
  }
  return { type: null, rest };
}

// 取「最後出現的數字」當金額
function extractAmount(seg) {
  const matches = [...seg.matchAll(/(\d+(\.\d+)?)/g)];
  if (!matches.length) return { amount: null, rest: seg };

  const last = matches[matches.length - 1];
  const amount = parseFloat(last[1]);
  const rest = seg.replace(last[0], " ");
  return { amount, rest };
}

// 解析一段：日期 類型 分類 項目 金額 地點(可用關鍵字)
function parseOneVoiceSegment(raw) {
  const seg0 = normalizeText(raw);
  if (!seg0) return null;

  const { placeText, rest: seg1 } = extractPlaceKeyword(seg0);
  const { dateStr, rest: seg2 } = extractDate(seg1);
  const { type, rest: seg3 } = extractType(seg2);
  if (!type) return null;

  const { amount, rest: seg4 } = extractAmount(seg3);
  if (!Number.isFinite(amount) || amount <= 0) return null;

  const tokens = normalizeText(seg4).split(" ").filter(Boolean);
  const categoryText = tokens[0] || "";
  const itemText = tokens.slice(1).join(" ") || "";

  return {
    occurred_at: dateStr,
    type,
    categoryText,
    itemText,
    amount,
    placeText: placeText || "",
  };
}

function parseVoiceTextToEntries(rawText) {
  const entries = [];
  if (!rawText) return entries;

  const parts = rawText.replace(/\r/g, "\n").split(/下一筆|下ㄧ筆|下一笔/);
  for (const p of parts) {
    const e = parseOneVoiceSegment(p);
    if (e) entries.push(e);
  }
  return entries;
}

function calcHitStatus(e) {
  const cId = resolveCategoryIdFast(e.categoryText);
  const pId = resolvePlaceIdFast(e.placeText);
  return {
    categoryOk: !!cId || !e.categoryText.trim(),
    placeOk: !!pId || !e.placeText.trim(),
    categoryId: cId,
    placeId: pId,
  };
}

// ✅ 命中欄位文字/樣式
function buildHitText(hit) {
  const hitText = (hit.categoryOk ? "分類✅" : "分類⚠") + " / " + (hit.placeOk ? "地點✅" : "地點⚠");
  const hitCls = hit.categoryOk && hit.placeOk ? "hit-ok" : "hit-warn";
  return { hitText, hitCls };
}

// ✅ 只更新某一列命中欄位（不重繪整表，避免中文輸入法被打斷）
function updateHitCellByIndex(idx) {
  if (!voicePreviewBodyEl) return;
  const cell = voicePreviewBodyEl.querySelector(`[data-hit-idx="${idx}"]`);
  if (!cell) return;

  const hit = calcHitStatus(pendingVoiceEntries[idx]);
  const { hitText, hitCls } = buildHitText(hit);

  cell.textContent = hitText;
  cell.classList.remove("hit-ok", "hit-warn");
  cell.classList.add(hitCls);
}

function renderVoicePreview() {
  if (!voicePreviewBodyEl) return;

  if (!pendingVoiceEntries.length) {
    voicePreviewBodyEl.innerHTML = '<tr><td colspan="9">尚無語音解析資料</td></tr>';
    if (voiceConfirmBtn) voiceConfirmBtn.disabled = true;
    if (voiceClearBtn) voiceClearBtn.disabled = true;
    return;
  }

  if (voiceConfirmBtn) voiceConfirmBtn.disabled = false;
  if (voiceClearBtn) voiceClearBtn.disabled = false;

  voicePreviewBodyEl.innerHTML = "";

  pendingVoiceEntries.forEach((e, idx) => {
    const typeLabel = e.type === "income" ? "收入" : "支出";

    const hit = calcHitStatus(e);
    const { hitText, hitCls } = buildHitText(hit);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${e.occurred_at}</td>
      <td>${typeLabel}</td>

      <td>
        <input class="table-input" data-idx="${idx}" data-field="categoryText" list="category-suggest" value="${escapeHtml(
      e.categoryText
    )}" />
      </td>

      <td>
        <input class="table-input" data-idx="${idx}" data-field="itemText" value="${escapeHtml(e.itemText)}" />
      </td>

      <td>
        <input class="table-input" data-idx="${idx}" data-field="amount" inputmode="numeric" value="${e.amount}" />
      </td>

      <td>
        <input class="table-input" data-idx="${idx}" data-field="placeText" list="place-suggest" value="${escapeHtml(
      e.placeText
    )}" />
      </td>

      <!-- ✅ 加 data-hit-idx，讓我們可以只更新命中欄 -->
      <td class="${hitCls}" data-hit-idx="${idx}">${hitText}</td>

      <td>
        <button type="button" class="btn-secondary" onclick="removeVoiceEntry(${idx})">刪除</button>
      </td>
    `;

    voicePreviewBodyEl.appendChild(tr);
  });

  bindVoicePreviewInputs();
}

// ✅ 關鍵修正：不要在 input 每打一個字就 render 整表（會打斷中文輸入法 composition）
let voicePreviewBound = false;

function bindVoicePreviewInputs() {
  if (voicePreviewBound || !voicePreviewBodyEl) return;
  voicePreviewBound = true;

  let composing = false;

  voicePreviewBodyEl.addEventListener("compositionstart", (ev) => {
    const t = ev.target;
    if (t instanceof HTMLInputElement) composing = true;
  });

  voicePreviewBodyEl.addEventListener("compositionend", (ev) => {
    const t = ev.target;
    if (!(t instanceof HTMLInputElement)) return;
    composing = false;

    const idx = parseInt(t.getAttribute("data-idx"), 10);
    const field = t.getAttribute("data-field");
    if (!Number.isFinite(idx) || idx < 0 || idx >= pendingVoiceEntries.length) return;
    if (!field) return;

    if (field === "amount") {
      const v = parseFloat(t.value);
      if (Number.isFinite(v)) pendingVoiceEntries[idx].amount = v;
    } else {
      pendingVoiceEntries[idx][field] = t.value;
    }

    // ✅ composition 結束後再更新命中欄（不重繪整表）
    updateHitCellByIndex(idx);
  });

  voicePreviewBodyEl.addEventListener(
    "input",
    (ev) => {
      const t = ev.target;
      if (!(t instanceof HTMLInputElement)) return;

      // ✅ 組字中不要動（不然注音會被打斷）
      if (composing) return;

      const idx = parseInt(t.getAttribute("data-idx"), 10);
      const field = t.getAttribute("data-field");
      if (!Number.isFinite(idx) || idx < 0 || idx >= pendingVoiceEntries.length) return;
      if (!field) return;

      if (field === "amount") {
        const v = parseFloat(t.value);
        if (Number.isFinite(v)) pendingVoiceEntries[idx].amount = v;
      } else {
        pendingVoiceEntries[idx][field] = t.value;
      }

      // ✅ 只更新該列命中欄位
      updateHitCellByIndex(idx);
    },
    { passive: true }
  );
}

function removeVoiceEntry(index) {
  if (index < 0 || index >= pendingVoiceEntries.length) return;
  pendingVoiceEntries.splice(index, 1);
  // 注意：這裡會重繪整表（正常，因為刪列）
  renderVoicePreview();
  if (voiceStatusEl) voiceStatusEl.textContent = `已刪除一筆，目前剩 ${pendingVoiceEntries.length} 筆。`;
}

function parseTextInput() {
  if (!voiceTextInputEl) return;
  const text = voiceTextInputEl.value.trim();
  if (!text) return alert("請先輸入/貼上語音文字");

  const newEntries = parseVoiceTextToEntries(text);
  if (!newEntries.length) {
    if (voiceStatusEl)
      voiceStatusEl.textContent = "⚠ 無法解析，請用「日期 類型 分類 項目 金額（在地點/地點xxx）」格式。";
    return;
  }

  pendingVoiceEntries.push(...newEntries);
  renderVoicePreview();

  let catOk = 0,
    placeOk = 0;
  for (const e of pendingVoiceEntries) {
    const hit = calcHitStatus(e);
    if (hit.categoryOk) catOk++;
    if (hit.placeOk) placeOk++;
  }

  if (voiceStatusEl) {
    voiceStatusEl.textContent =
      `✅ 解析新增 ${newEntries.length} 筆，目前暫存 ${pendingVoiceEntries.length} 筆。` +
      `（分類命中 ${catOk}/${pendingVoiceEntries.length}，地點命中 ${placeOk}/${pendingVoiceEntries.length}）`;
  }
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
    location.href = "index.html";
    return;
  }

  if (!pendingVoiceEntries.length) return alert("目前沒有暫存資料");

  if (isSubmittingVoice) return;
  isSubmittingVoice = true;

  if (voiceConfirmBtn) {
    voiceConfirmBtn.disabled = true;
    voiceConfirmBtn.textContent = "送出中...";
  }

  try {
    const ok = confirm(`確定送出 ${pendingVoiceEntries.length} 筆語音記帳嗎？`);
    if (!ok) return;

    const payloads = [];
    for (const e of pendingVoiceEntries) {
      const category_id = resolveCategoryIdFast(e.categoryText) || null;
      const place_id = resolvePlaceIdFast(e.placeText) || null;

      let finalItem = (e.itemText || "").trim();

      if (!category_id && e.categoryText.trim()) {
        finalItem = `${finalItem} [分類:${e.categoryText.trim()}]`.trim();
      }
      if (!place_id && e.placeText.trim()) {
        finalItem = `${finalItem} [地點:${e.placeText.trim()}]`.trim();
      }

      const amount = parseFloat(e.amount);
      if (!Number.isFinite(amount) || amount <= 0) continue;

payloads.push({
  user_id: user.id, // ✅ 新增這行（RLS 需要）
  occurred_at: e.occurred_at,
  type: e.type,
  amount,
  item: finalItem || null,
  category_id,
  place_id,
  source: "speech",
});
    }

    if (!payloads.length) {
      alert("沒有可送出的有效資料（請檢查金額/格式）");
      return;
    }

    const { error } = await supabaseClient.from("ledger_entries").insert(payloads);
    if (error) {
      alert("語音寫入失敗：" + error.message);
      logDebug("voice insert error", error);
      logRlsHint(error, "ledger_entries");
      return;
    }

    pendingVoiceEntries = [];
    renderVoicePreview();
    if (voiceStatusEl) voiceStatusEl.textContent = "✅ 語音記帳已寫入成功。";
    await loadLedger();
  } finally {
    isSubmittingVoice = false;
    if (voiceConfirmBtn) {
      voiceConfirmBtn.disabled = false;
      voiceConfirmBtn.textContent = "✅ 確認送出";
    }
  }
}

// ===============================
// 9) 初始化
// ===============================
document.addEventListener("DOMContentLoaded", async () => {
  logDebug("page loaded");

if (isLedgerPage) {
  const user = await getCurrentUser();
  if (!user) {
    location.href = "index.html";
    return;
  }

  // 1) 載入 autocomplete / 快取
  await loadSuggestCaches();
  wireManualHitHints();

  // 2) 載入 profile
  const profile = await loadProfile(user.id);
  logDebug("profile loaded", { userId: user.id, profile });

  // 3) 設定 Hello（只做一次）
  const displayName = (profile?.display_name || "").trim();
  if (helloLineEl) {
    helloLineEl.textContent = displayName
      ? `Hello，${displayName}`
      : "Hello，（未設定姓名）";
  }

  // 4) 載入自在語（一定要在 Hello 後）
  await loadZenQuote();

  // 5) 預設日期
  const dateEl = document.getElementById("entry-date");
  if (dateEl && !dateEl.value) {
    dateEl.value = formatDate(new Date());
  }

  // 6) 載入記帳列表
  await loadLedger();

  // 7) 綁定事件
  if (parseTextBtn) parseTextBtn.addEventListener("click", parseTextInput);
  if (voiceTextClearBtn) voiceTextClearBtn.addEventListener("click", clearVoiceText);
  if (voiceClearBtn) voiceClearBtn.addEventListener("click", clearVoiceEntries);
  if (voiceConfirmBtn) voiceConfirmBtn.addEventListener("click", confirmVoiceEntries);

  renderVoicePreview();
}
});
