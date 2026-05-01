import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js";
import {
  getFirestore,
  collection,
  addDoc,
  getDocs,
  doc,
  updateDoc,
  runTransaction,
  serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js";

const env = window.__ENV || {};
const firebaseConfig = env.firebase || {};
const stockAlertEmailConfig = env.stockAlertEmail || {};

const statusEl = document.getElementById("firebaseStatus");
const toastEl = document.getElementById("toast");
const notificationToggle = document.getElementById("notificationToggle");
const notificationPanel = document.getElementById("notificationPanel");
const notificationList = document.getElementById("notificationList");
const notificationCount = document.getElementById("notificationCount");
const profileSelect = document.getElementById("profileSelect");
const profileNameInput = document.getElementById("profileNameInput");
const addProfileBtn = document.getElementById("addProfileBtn");

const PROFILE_LIST_KEY = "ims.profile.list";
const PROFILE_ACTIVE_KEY = "ims.profile.active";
const STOCK_ALERT_SIGNATURE_KEY = "ims.stockAlert.lastSignature";

const STOCK_ALERT_DEFAULT_RECIPIENT = "phoenixpathlabs@gmail.com";

const isConfigured =
  firebaseConfig.apiKey &&
  !firebaseConfig.apiKey.includes("YOUR_") &&
  firebaseConfig.projectId &&
  !firebaseConfig.projectId.includes("YOUR_");

let db = null;
let firebaseReady = false;

if (isConfigured) {
  const app = initializeApp(firebaseConfig);
  db = getFirestore(app);
  firebaseReady = true;
  statusEl.textContent = "Firebase: connected";
  refreshAll();
} else {
  statusEl.textContent = "Firebase: add config in env.js";
}

const state = {
  profiles: [],
  activeProfile: "",
  categories: [],
  units: [],
  vendors: [],
  items: [],
  purchases: [],
  consumptions: [],
  returns: [],
  adjustments: [],
  auditLogs: []
};

const navButtons = Array.from(document.querySelectorAll(".nav-btn"));
const featureButtons = Array.from(document.querySelectorAll(".feature-btn"));
const pages = Array.from(document.querySelectorAll(".page"));
const featurePanels = Array.from(document.querySelectorAll(".feature-panel"));

const categoryForm = document.getElementById("categoryForm");
const unitForm = document.getElementById("unitForm");
const vendorForm = document.getElementById("vendorForm");
const itemForm = document.getElementById("itemForm");
const excelImportForm = document.getElementById("excelImportForm");
const excelFileInput = document.getElementById("excelFile");

const purchaseForm = document.getElementById("purchaseForm");
const returnForm = document.getElementById("returnForm");
const adjustmentForm = document.getElementById("adjustmentForm");

const purchaseItem = document.getElementById("purchaseItem");
const purchaseQty = document.getElementById("purchaseQty");
const purchaseTotal = document.getElementById("purchaseTotal");

const itemCategory = document.getElementById("itemCategory");
const itemUnit = document.getElementById("itemUnit");
const itemVendor = document.getElementById("itemVendor");
const itemHasExpiry = document.getElementById("itemHasExpiry");
const expiryRow = document.getElementById("expiryRow");
const categorySlideBased = document.getElementById("categorySlideBased");
const categorySlideConfig = document.getElementById("categorySlideConfig");
const categorySlideUnit = document.getElementById("categorySlideUnit");
const categorySlideValue = document.getElementById("categorySlideValue");

const returnItem = document.getElementById("returnItem");
const adjustItem = document.getElementById("adjustItem");

const consumptionList = document.getElementById("consumptionList");
const consumptionSearch = document.getElementById("consumptionSearch");
const configItemsList = document.getElementById("configItemsList");

const statTotalItems = document.getElementById("statTotalItems");
const statLowStock = document.getElementById("statLowStock");
const statOutStock = document.getElementById("statOutStock");
const statSpend = document.getElementById("statSpend");
const dashTotalItems = document.getElementById("dashTotalItems");
const dashLowStock = document.getElementById("dashLowStock");
const dashOutStock = document.getElementById("dashOutStock");
const dashSpend = document.getElementById("dashSpend");
const lowStockList = document.getElementById("lowStockList");
const outStockList = document.getElementById("outStockList");
const expiryList = document.getElementById("expiryList");
const activityList = document.getElementById("activityList");
const statsDatePreset = document.getElementById("statsDatePreset");
const statsDateFrom = document.getElementById("statsDateFrom");
const statsDateTo = document.getElementById("statsDateTo");
const statsProfileFilter = document.getElementById("statsProfileFilter");
const statsActionFilter = document.getElementById("statsActionFilter");
const statsCategoryFilter = document.getElementById("statsCategoryFilter");
const statsItemSearch = document.getElementById("statsItemSearch");
const statsResetFilters = document.getElementById("statsResetFilters");
const statsExportCsv = document.getElementById("statsExportCsv");
const analyticsEventCount = document.getElementById("analyticsEventCount");
const analyticsPurchaseQty = document.getElementById("analyticsPurchaseQty");
const analyticsConsumptionQty = document.getElementById("analyticsConsumptionQty");
const analyticsNetStock = document.getElementById("analyticsNetStock");
const analyticsSlidesUsed = document.getElementById("analyticsSlidesUsed");
const analyticsTopConsumed = document.getElementById("analyticsTopConsumed");
const analyticsConsumedAll = document.getElementById("analyticsConsumedAll");
const statsAuditList = document.getElementById("statsAuditList");
const logsActionFilter = document.getElementById("logsActionFilter");
const logsProfileFilter = document.getElementById("logsProfileFilter");
const logsDateFrom = document.getElementById("logsDateFrom");
const logsDateTo = document.getElementById("logsDateTo");
const logsItemSearch = document.getElementById("logsItemSearch");
const logsCount = document.getElementById("logsCount");
const logsList = document.getElementById("logsList");
const auditActionFilter = document.getElementById("auditActionFilter");
const auditEntityFilter = document.getElementById("auditEntityFilter");
const auditProfileFilter = document.getElementById("auditProfileFilter");
const auditDateFrom = document.getElementById("auditDateFrom");
const auditDateTo = document.getElementById("auditDateTo");
const auditSearch = document.getElementById("auditSearch");
const auditCount = document.getElementById("auditCount");
const auditLogsList = document.getElementById("auditLogsList");
const repeatMonthFilter = document.getElementById("repeatMonthFilter");
const repeatSpecificDate = document.getElementById("repeatSpecificDate");
const repeatItemSearch = document.getElementById("repeatItemSearch");
const repeatNetSlides = document.getElementById("repeatNetSlides");
const repeatLogCount = document.getElementById("repeatLogCount");
const repeatByItemList = document.getElementById("repeatByItemList");
const repeatAuditList = document.getElementById("repeatAuditList");

const currencyFormatter = new Intl.NumberFormat("en-IN", {
  style: "currency",
  currency: "INR"
});

const dateFormatter = new Intl.DateTimeFormat("en-US", {
  dateStyle: "medium",
  timeStyle: "short"
});

let purchaseTotalManual = false;
let activeConfigAction = { itemId: null, mode: null };
let activeLogAction = { eventId: null, mode: null };
let consumptionSearchQuery = "";
const searchableSelects = new Map();
let statsFiltersInitialized = false;
let auditLogsInitialized = false;

function showToast(message, isError = false) {
  toastEl.textContent = message;
  toastEl.style.background = isError ? "#b4442f" : "#1f4d4c";
  toastEl.classList.add("show");
  setTimeout(() => toastEl.classList.remove("show"), 2500);
}

function friendlyFirestoreError(error, fallback = "Operation failed") {
  const code = String(error?.code || "").toLowerCase();
  const message = String(error?.message || "").toLowerCase();

  if (code.includes("resource-exhausted") || message.includes("quota exceeded")) {
    return "Firestore quota exceeded. Delete cannot run until quota resets or plan is upgraded.";
  }
  if (code.includes("permission-denied")) {
    return "Permission denied by Firestore rules.";
  }
  return error?.message || fallback;
}

function escapeHTML(value) {
  if (!value) return "";
  return String(value).replace(/[&<>"']/g, (match) => {
    const map = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    };
    return map[match] || match;
  });
}

function toNumber(value) {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function toDate(value) {
  if (!value) return null;
  if (value.toDate) return value.toDate();
  if (value instanceof Date) return value;
  return new Date(value);
}

function timeValue(value) {
  if (!value) return 0;
  if (value.toMillis) return value.toMillis();
  const date = toDate(value);
  return date ? date.getTime() : 0;
}

function formatDate(value) {
  const date = toDate(value);
  return date ? dateFormatter.format(date) : "";
}

function isoDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function monthKey(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function monthLabelFromKey(value) {
  if (!/^\d{4}-\d{2}$/.test(String(value || ""))) return value || "";
  const [year, month] = String(value).split("-");
  const monthIndex = Number.parseInt(month, 10) - 1;
  const date = new Date(Number.parseInt(year, 10), monthIndex, 1);
  return Number.isNaN(date.getTime())
    ? value
    : new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric" }).format(date);
}

function escapeHtmlForEmail(value) {
  return escapeHTML(value);
}

function isStockEmailConfigured() {
  return Boolean(
    stockAlertEmailConfig.apiUrl
    && !String(stockAlertEmailConfig.apiUrl).includes("YOUR_")
  );
}

function stockAlertSignature(lowStockItems, outStockItems) {
  const low = lowStockItems
    .map((item) => `${item.id}:${toNumber(item.stockQty)}`)
    .sort()
    .join("|");
  const out = outStockItems
    .map((item) => `${item.id}:${toNumber(item.stockQty)}`)
    .sort()
    .join("|");
  return `LOW[${low}]__OUT[${out}]`;
}

function buildStockRows(items, modeLabel) {
  if (!items.length) {
    return `
      <tr>
        <td colspan="4" style="padding:10px;border:1px solid #e6e0d7;color:#5a6b66;text-align:center;">No ${modeLabel.toLowerCase()} items</td>
      </tr>
    `;
  }

  return items
    .map((item) => {
      const unit = getUnitName(item.unitId) || "Unit";
      return `
        <tr>
          <td style="padding:10px;border:1px solid #e6e0d7;">${escapeHtmlForEmail(item.name)}</td>
          <td style="padding:10px;border:1px solid #e6e0d7;">${escapeHtmlForEmail(getCategoryName(item.categoryId))}</td>
          <td style="padding:10px;border:1px solid #e6e0d7;">${toNumber(item.stockQty)} ${escapeHtmlForEmail(unit)}</td>
          <td style="padding:10px;border:1px solid #e6e0d7;">${modeLabel}</td>
        </tr>
      `;
    })
    .join("");
}

function buildStockAlertEmailHtml(lowStockItems, outStockItems) {
  const now = new Date();
  const generatedOn = dateFormatter.format(now);
  const totalAlerts = lowStockItems.length + outStockItems.length;

  return `
    <div style="font-family:Arial,Helvetica,sans-serif;background:#f7f1e8;padding:24px;color:#1f2a2a;">
      <div style="max-width:820px;margin:0 auto;background:#ffffff;border:1px solid #e6e0d7;border-radius:14px;overflow:hidden;">
        <div style="background:#1f4d4c;color:#fff;padding:18px 22px;">
          <div style="font-size:20px;font-weight:700;">Penguin Biotechnologies</div>
          <div style="font-size:13px;opacity:.9;">Inventory Stock Alert Summary</div>
        </div>
        <div style="padding:20px 22px;">
          <p style="margin:0 0 10px;font-size:14px;">Hello Team,</p>
          <p style="margin:0 0 12px;font-size:14px;">This is a consolidated inventory alert email from IMS. Please review the low and out-of-stock items listed below.</p>
          <div style="margin:0 0 14px;padding:10px 12px;background:#f0e8dc;border:1px solid #e6e0d7;border-radius:10px;font-size:13px;">
            <strong>Total Alerts:</strong> ${totalAlerts} &nbsp; | &nbsp;
            <strong>Low Stock:</strong> ${lowStockItems.length} &nbsp; | &nbsp;
            <strong>Out of Stock:</strong> ${outStockItems.length}<br/>
            <strong>Generated:</strong> ${escapeHtmlForEmail(generatedOn)}
          </div>
          <table style="width:100%;border-collapse:collapse;font-size:13px;">
            <thead>
              <tr style="background:#f7f1e8;">
                <th style="padding:10px;border:1px solid #e6e0d7;text-align:left;">Item</th>
                <th style="padding:10px;border:1px solid #e6e0d7;text-align:left;">Category</th>
                <th style="padding:10px;border:1px solid #e6e0d7;text-align:left;">Current Stock</th>
                <th style="padding:10px;border:1px solid #e6e0d7;text-align:left;">Status</th>
              </tr>
            </thead>
            <tbody>
              ${buildStockRows(lowStockItems, "Low Stock")}
              ${buildStockRows(outStockItems, "Out of Stock")}
            </tbody>
          </table>
          <p style="margin:14px 0 0;font-size:12px;color:#5a6b66;">This is an automated alert from IMS.</p>
        </div>
      </div>
    </div>
  `;
}

async function maybeSendStockAlertEmail(lowStockItems, outStockItems) {
  if (!isStockEmailConfigured()) return;

  const total = lowStockItems.length + outStockItems.length;
  if (!total) return;

  const signature = stockAlertSignature(lowStockItems, outStockItems);
  const previousSignature = localStorage.getItem(STOCK_ALERT_SIGNATURE_KEY);
  if (signature === previousSignature) return;

  const recipient = stockAlertEmailConfig.toEmail || STOCK_ALERT_DEFAULT_RECIPIENT;
  const emailHtml = buildStockAlertEmailHtml(lowStockItems, outStockItems);

  const payload = {
    toEmail: recipient,
    subject: `IMS Stock Alert Summary (${total} items)`,
    html: emailHtml,
    totalAlerts: total,
    lowStockCount: lowStockItems.length,
    outStockCount: outStockItems.length,
    generatedAt: dateFormatter.format(new Date())
  };

  try {
    const headers = {
      "Content-Type": "application/json"
    };

    if (stockAlertEmailConfig.apiKey && !String(stockAlertEmailConfig.apiKey).includes("YOUR_")) {
      headers["x-alert-key"] = stockAlertEmailConfig.apiKey;
    }

    const response = await fetch(stockAlertEmailConfig.apiUrl, {
      method: "POST",
      headers,
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      throw new Error("Email send failed");
    }

    localStorage.setItem(STOCK_ALERT_SIGNATURE_KEY, signature);
  } catch (error) {
    console.error("Stock alert email error", error);
  }
}

function ensureFirebase() {
  if (!firebaseReady) {
    showToast("Add Firebase config in env.js", true);
    return false;
  }
  return true;
}

function ensureProfile() {
  if (!state.activeProfile) {
    showToast("Select profile to continue", true);
    return false;
  }
  return true;
}

function ensureReady() {
  return ensureFirebase() && ensureProfile();
}

function getAuditProfile() {
  return state.activeProfile || "Unknown";
}

function normalizeProfileName(value) {
  return String(value || "").trim();
}

function loadProfiles() {
  try {
    const raw = localStorage.getItem(PROFILE_LIST_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeProfileName).filter(Boolean);
  } catch {
    return [];
  }
}

function saveProfiles() {
  localStorage.setItem(PROFILE_LIST_KEY, JSON.stringify(state.profiles));
}

function ensureSearchableSelect(selectEl) {
  if (!selectEl || !window.TomSelect) return;

  const existing = searchableSelects.get(selectEl);
  if (existing) {
    existing.sync();
    existing.refreshOptions(false);
    if (selectEl.disabled) {
      existing.disable();
    } else {
      existing.enable();
    }
    return;
  }

  const instance = new window.TomSelect(selectEl, {
    create: false,
    allowEmptyOption: true,
    searchField: ["text"],
    placeholder: "Type to search..."
  });
  searchableSelects.set(selectEl, instance);

  if (selectEl.disabled) {
    instance.disable();
  }
}

function initSearchableSelects() {
  const selects = Array.from(document.querySelectorAll("select"));
  selects.forEach((selectEl) => ensureSearchableSelect(selectEl));
}

function renderProfileSelect() {
  if (!profileSelect) return;
  const current = state.activeProfile;
  profileSelect.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select Profile";
  profileSelect.appendChild(placeholder);

  state.profiles.forEach((profile) => {
    const option = document.createElement("option");
    option.value = profile;
    option.textContent = profile;
    profileSelect.appendChild(option);
  });

  profileSelect.value = current || "";
  ensureSearchableSelect(profileSelect);
}

function setAppLocked(locked) {
  const interactiveEls = Array.from(document.querySelectorAll("button, input, select, textarea"));
  interactiveEls.forEach((element) => {
    if (element.closest(".profile-picker")) return;
    element.disabled = locked;
  });

  Array.from(document.querySelectorAll("select")).forEach((selectEl) => ensureSearchableSelect(selectEl));
}

function setActiveProfile(profileName) {
  const normalized = normalizeProfileName(profileName);
  state.activeProfile = normalized;
  localStorage.setItem(PROFILE_ACTIVE_KEY, normalized);
  renderProfileSelect();
  setAppLocked(!normalized);
}

function addProfileFromInput() {
  const profileName = normalizeProfileName(profileNameInput?.value);
  if (!profileName) {
    showToast("Enter profile name", true);
    return;
  }

  const exists = state.profiles.some((profile) => profile.toLowerCase() === profileName.toLowerCase());
  if (!exists) {
    state.profiles.push(profileName);
    state.profiles.sort((a, b) => a.localeCompare(b));
    saveProfiles();
  }

  if (profileNameInput) profileNameInput.value = "";
  setActiveProfile(profileName);
  showToast(`Profile set: ${profileName}`);
}

function initProfiles() {
  state.profiles = loadProfiles();
  const storedActive = normalizeProfileName(localStorage.getItem(PROFILE_ACTIVE_KEY));

  if (storedActive && !state.profiles.some((profile) => profile.toLowerCase() === storedActive.toLowerCase())) {
    state.profiles.push(storedActive);
  }

  state.profiles = [...new Set(state.profiles)].sort((a, b) => a.localeCompare(b));
  saveProfiles();
  setActiveProfile(storedActive);

  profileSelect?.addEventListener("change", () => {
    setActiveProfile(profileSelect.value);
  });

  addProfileBtn?.addEventListener("click", addProfileFromInput);
  profileNameInput?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    addProfileFromInput();
  });
}

async function fetchCollection(name) {
  const snapshot = await getDocs(collection(db, name));
  return snapshot.docs.map((docSnap) => ({
    id: docSnap.id,
    ...docSnap.data()
  }));
}

async function refreshAll() {
  if (!firebaseReady) return;
  try {
    const [
      categories,
      units,
      vendors,
      items,
      purchases,
      consumptions,
      returns,
      adjustments,
      auditLogs
    ] = await Promise.all([
      fetchCollection("categories"),
      fetchCollection("units"),
      fetchCollection("vendors"),
      fetchCollection("items"),
      fetchCollection("purchases"),
      fetchCollection("consumptions"),
      fetchCollection("returns"),
      fetchCollection("adjustments"),
      fetchCollection("auditLogs")
    ]);

    state.categories = categories;
    state.units = units;
    state.vendors = vendors;
    state.items = items;
    state.purchases = purchases;
    state.consumptions = consumptions;
    state.returns = returns;
    state.adjustments = adjustments;
    state.auditLogs = auditLogs;

    renderSelects();
    renderConsumption();
    renderConfigItems();
    renderStats();
    updatePurchaseTotal();
  } catch (error) {
    if (error?.code === "permission-denied") {
      showToast("Permission denied. Set Firestore rules to allow public access.", true);
    } else {
      showToast("Failed to load data", true);
    }
    console.error(error);
  }
}

function renderSelect(selectEl, items, labelFn, placeholder) {
  const current = selectEl.value;
  selectEl.innerHTML = "";

  if (!items.length) {
    const option = document.createElement("option");
    option.textContent = placeholder || "No options";
    option.value = "";
    option.disabled = true;
    option.selected = true;
    selectEl.appendChild(option);
    selectEl.disabled = true;
    ensureSearchableSelect(selectEl);
    return;
  }

  selectEl.disabled = false;
  items.forEach((item) => {
    const option = document.createElement("option");
    option.value = item.id;
    option.textContent = labelFn(item);
    selectEl.appendChild(option);
  });

  if (items.some((item) => item.id === current)) {
    selectEl.value = current;
  }

  ensureSearchableSelect(selectEl);
}

function getVendorName(id) {
  return state.vendors.find((vendor) => vendor.id === id)?.name || "Unknown";
}

function getUnitName(id) {
  return state.units.find((unit) => unit.id === id)?.displayName || "";
}

function getCategoryName(id) {
  return state.categories.find((category) => category.id === id)?.name || "Uncategorized";
}

function getCategory(id) {
  return state.categories.find((category) => category.id === id);
}

function findByName(items, value) {
  const target = String(value || "").trim().toLowerCase();
  return items.find((item) => String(item.name || item.displayName || "").trim().toLowerCase() === target);
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

function toSafeStock(value) {
  return Math.max(0, toNumber(value));
}

async function ensureExcelImportSetup() {
  const profileName = getAuditProfile();

  let unit = state.units.find((entry) => String(entry.displayName || "").toLowerCase() === "ul");
  if (!unit) {
    const unitRef = await addDoc(collection(db, "units"), {
      displayName: "uL",
      unitName: "MICROLITER",
      profileName,
      createdAt: serverTimestamp()
    });
    unit = { id: unitRef.id, displayName: "uL", unitName: "MICROLITER" };
    state.units.push(unit);
  }

  let vendor = findByName(state.vendors, "Epitope");
  if (!vendor) {
    const vendorRef = await addDoc(collection(db, "vendors"), {
      name: "Epitope",
      address: "",
      mobile: "",
      email: "",
      openingBalance: 0,
      profileName,
      createdAt: serverTimestamp()
    });
    vendor = { id: vendorRef.id, name: "Epitope" };
    state.vendors.push(vendor);
  }

  let category = findByName(state.categories, "IHC marker");
  const unitsPerSlide = 50;

  if (!category) {
    const categoryRef = await addDoc(collection(db, "categories"), {
      name: "IHC marker",
      description: "Imported via Excel",
      slideBasedConsumption: true,
      slideUnitId: unit.id,
      unitsPerSlide,
      profileName,
      createdAt: serverTimestamp()
    });
    category = {
      id: categoryRef.id,
      name: "IHC marker",
      slideBasedConsumption: true,
      slideUnitId: unit.id,
      unitsPerSlide
    };
    state.categories.push(category);
  } else {
    await updateDoc(doc(db, "categories", category.id), {
      slideBasedConsumption: true,
      slideUnitId: unit.id,
      unitsPerSlide,
      lastUpdatedBy: profileName,
      updatedAt: serverTimestamp()
    });
  }

  return {
    unitId: unit.id,
    vendorId: vendor.id,
    categoryId: category.id
  };
}

async function parseExcelFile(file) {
  if (!window.XLSX) {
    throw new Error("Excel parser failed to load. Please refresh and try again.");
  }

  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) return [];

  const sheet = workbook.Sheets[firstSheetName];
  const rows = window.XLSX.utils.sheet_to_json(sheet, { defval: "" });

  return rows.map((row) => {
    const normalized = {};
    Object.entries(row).forEach(([key, value]) => {
      normalized[normalizeHeader(key)] = value;
    });
    return normalized;
  });
}

function getItem(id) {
  return state.items.find((item) => item.id === id);
}

function validateSlideCategoryUnit(categoryId, unitId) {
  const category = getCategory(categoryId);
  if (!category?.slideBasedConsumption) return null;
  if (!category.slideUnitId) return null;
  if (category.slideUnitId === unitId) return null;
  const requiredUnit = getUnitName(category.slideUnitId) || "configured category unit";
  return `Slide-based category requires item unit: ${requiredUnit}`;
}

function renderConfigItems() {
  if (!configItemsList) return;

  if (!state.items.length) {
    configItemsList.innerHTML = "<div class=\"list-empty\">Create items to manage them here.</div>";
    return;
  }

  const categories = [...state.categories].sort((a, b) => a.name.localeCompare(b.name));
  const itemsByCategory = new Map();
  categories.forEach((category) => itemsByCategory.set(category.id, []));
  itemsByCategory.set("uncategorized", []);

  state.items.forEach((item) => {
    const key = item.categoryId && itemsByCategory.has(item.categoryId) ? item.categoryId : "uncategorized";
    itemsByCategory.get(key).push(item);
  });

  const blocks = [];
  categories.forEach((category) => {
    const items = itemsByCategory.get(category.id) || [];
    blocks.push(renderConfigCategoryBlock(category, items));
  });

  const uncategorizedItems = itemsByCategory.get("uncategorized") || [];
  if (uncategorizedItems.length) {
    blocks.push(renderConfigCategoryBlock({ id: "", name: "Uncategorized" }, uncategorizedItems));
  }

  configItemsList.innerHTML = blocks.join("");
}

function renderConfigCategoryBlock(category, items) {
  const title = escapeHTML(category?.name || "Uncategorized");
  const categoryId = category?.id || "";
  const bulkControls = categoryId && items.length
    ? `
      <form class="bulk-reorder" data-form="bulkReorder" data-category-id="${categoryId}">
        <label>Set Re-order level for all</label>
        <input type="number" name="reorderLevel" min="0" step="0.01" value="0" required />
        <button type="submit" class="secondary-btn">Apply</button>
      </form>
    `
    : "";
  const itemRows = items.length
    ? items
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((item) => renderConfigItemRow(item))
      .join("")
    : "<div class=\"list-empty\">No items in this category.</div>";

  return `
    <div class="config-category">
      <div class="config-category-head">
        <h4>${title}</h4>
        ${bulkControls}
      </div>
      ${itemRows}
    </div>
  `;
}

function renderConfigItemRow(item) {
  const unitName = getUnitName(item.unitId);
  const vendorName = getVendorName(item.vendorId);
  const price = toNumber(item.price);
  const stockQty = toNumber(item.stockQty);
  const reorderLevel = toNumber(item.reorderLevel);
  const hasExpiry = Boolean(item.hasExpiry);
  const expiryDateValue = item.expiryDate ? toDate(item.expiryDate)?.toISOString().slice(0, 10) || "" : "";
  const actionOpen = activeConfigAction.itemId === item.id;
  const mode = actionOpen ? activeConfigAction.mode : "";

  const categoryOptions = state.categories
    .map((category) => `<option value="${category.id}" ${category.id === item.categoryId ? "selected" : ""}>${escapeHTML(category.name)}</option>`)
    .join("");
  const unitOptions = state.units
    .map((unit) => `<option value="${unit.id}" ${unit.id === item.unitId ? "selected" : ""}>${escapeHTML(unit.displayName)} (${escapeHTML(unit.unitName)})</option>`)
    .join("");
  const vendorOptions = state.vendors
    .map((vendor) => `<option value="${vendor.id}" ${vendor.id === item.vendorId ? "selected" : ""}>${escapeHTML(vendor.name)}</option>`)
    .join("");

  let panel = "";
  if (mode === "edit") {
    panel = `
      <form class="inline-form" data-form="edit" data-item-id="${item.id}">
        <div class="form-row">
          <label>Item Name</label>
          <input type="text" name="name" value="${escapeHTML(item.name)}" required />
        </div>
        <div class="form-row">
          <label>Price</label>
          <input type="number" name="price" min="0" step="0.01" value="${price}" required />
        </div>
        <div class="form-row">
          <label>Re-order Quantity Level</label>
          <input type="number" name="reorderLevel" min="0" step="0.01" value="${reorderLevel}" required />
        </div>
        <div class="form-row toggle-row">
          <label><input type="checkbox" name="hasExpiry" ${hasExpiry ? "checked" : ""} /> Checklist for Expiry</label>
        </div>
        <div class="form-row">
          <label>Expiry Date (optional)</label>
          <input type="date" name="expiryDate" value="${expiryDateValue}" />
        </div>
        <div class="inline-actions">
          <button type="submit" class="primary">Save Edit</button>
          <button type="button" class="secondary-btn" data-config-action="close" data-item-id="${item.id}">Cancel</button>
        </div>
      </form>
    `;
  }

  if (mode === "config") {
    panel = `
      <form class="inline-form" data-form="config" data-item-id="${item.id}">
        <div class="form-row">
          <label>Category</label>
          <select name="categoryId" required>${categoryOptions}</select>
        </div>
        <div class="form-row">
          <label>Unit</label>
          <select name="unitId" required>${unitOptions}</select>
        </div>
        <div class="form-row">
          <label>Vendor</label>
          <select name="vendorId" required>${vendorOptions}</select>
        </div>
        <div class="inline-actions">
          <button type="submit" class="primary">Save Config</button>
          <button type="button" class="secondary-btn" data-config-action="close" data-item-id="${item.id}">Cancel</button>
        </div>
      </form>
    `;
  }

  if (mode === "manage") {
    panel = `
      <form class="inline-form" data-form="manage" data-item-id="${item.id}">
        <div class="form-row">
          <label>Current Stock Quantity</label>
          <input type="number" name="stockQty" min="0" step="0.01" value="${stockQty}" required />
        </div>
        <div class="inline-actions">
          <button type="submit" class="primary">Update Stock</button>
          <button type="button" class="secondary-btn" data-config-action="close" data-item-id="${item.id}">Cancel</button>
        </div>
      </form>
    `;
  }

  return `
    <div class="config-item">
      <div class="config-item-head">
        <div>
          <strong>${escapeHTML(item.name)}</strong>
          <div class="muted">Stock: ${stockQty} ${escapeHTML(unitName)} · Vendor: ${escapeHTML(vendorName)} · Price: ${currencyFormatter.format(price)}</div>
        </div>
        <div class="config-item-actions">
          <button class="secondary-btn ${mode === "edit" ? "active" : ""}" data-config-action="edit" data-item-id="${item.id}" type="button">Edit</button>
          <button class="secondary-btn ${mode === "config" ? "active" : ""}" data-config-action="config" data-item-id="${item.id}" type="button">Config</button>
          <button class="secondary-btn ${mode === "manage" ? "active" : ""}" data-config-action="manage" data-item-id="${item.id}" type="button">Manage</button>
        </div>
      </div>
      ${panel}
    </div>
  `;
}

function renderSelects() {
  renderSelect(itemCategory, state.categories, (category) => category.name, "Create category first");
  renderSelect(itemUnit, state.units, (unit) => `${unit.displayName} (${unit.unitName})`, "Create unit first");
  renderSelect(itemVendor, state.vendors, (vendor) => vendor.name, "Create vendor first");
  renderSelect(categorySlideUnit, state.units, (unit) => `${unit.displayName} (${unit.unitName})`, "Create unit first");

  renderSelect(purchaseItem, state.items, (item) => `${item.name} — ${getVendorName(item.vendorId)}`, "Create item first");
  renderSelect(returnItem, state.items, (item) => item.name, "Create item first");
  renderSelect(adjustItem, state.items, (item) => item.name, "Create item first");
}

function renderConsumption() {
  if (!state.items.length) {
    consumptionList.innerHTML = "<div class=\"list-empty\">Create items to start logging consumption.</div>";
    return;
  }

  const searchQuery = consumptionSearchQuery.trim().toLowerCase();

  const categories = [...state.categories].sort((a, b) => a.name.localeCompare(b.name));
  const itemsByCategory = new Map();

  categories.forEach((category) => itemsByCategory.set(category.id, []));
  itemsByCategory.set("uncategorized", []);

  state.items.forEach((item) => {
    if (searchQuery && !String(item.name || "").toLowerCase().includes(searchQuery)) {
      return;
    }
    const key = item.categoryId && itemsByCategory.has(item.categoryId) ? item.categoryId : "uncategorized";
    itemsByCategory.get(key).push(item);
  });

  const blocks = [];

  categories.forEach((category) => {
    const items = itemsByCategory.get(category.id) || [];
    blocks.push(renderCategoryBlock(category, items));
  });

  const uncategorizedItems = itemsByCategory.get("uncategorized") || [];
  if (uncategorizedItems.length) {
    blocks.push(renderCategoryBlock({ name: "Uncategorized" }, uncategorizedItems));
  }

  if (!blocks.length) {
    consumptionList.innerHTML = "<div class=\"list-empty\">No matching items found.</div>";
    return;
  }

  consumptionList.innerHTML = blocks.join("");
}

function renderCategoryBlock(category, items) {
  const title = category?.name || "Uncategorized";

  if (!items.length) {
    return `
      <div class="category-block">
        <h4>${escapeHTML(title)}</h4>
        <div class="list-empty">No items yet.</div>
      </div>
    `;
  }

  const rows = items
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((item) => {
      const stockQty = toNumber(item.stockQty);
      const unitName = getUnitName(item.unitId);
      const vendorName = getVendorName(item.vendorId);
      const itemCategory = getCategory(item.categoryId);
      const itemSlideBased = Boolean(itemCategory?.slideBasedConsumption && toNumber(itemCategory?.unitsPerSlide) > 0);
      const itemUnitsPerSlide = toNumber(itemCategory?.unitsPerSlide);
      const itemSlideUnitName = getUnitName(itemCategory?.slideUnitId) || unitName;
      const consumeModeLabel = itemSlideBased ? "Slides" : "Qty";
      const slideHint = itemSlideBased
        ? `<div class="slide-meta">1 slide = ${itemUnitsPerSlide} ${escapeHTML(itemSlideUnitName)}</div>`
        : "";

      return `
        <div class="item-row" data-slide-based="${itemSlideBased ? "1" : "0"}" data-units-per-slide="${itemUnitsPerSlide}" data-slide-unit-id="${itemCategory?.slideUnitId || ""}">
          <div class="item-meta">
            <div class="item-name">${escapeHTML(item.name)}</div>
            <div class="item-sub">Stock: ${stockQty} ${escapeHTML(unitName)} · Vendor: ${escapeHTML(vendorName)}</div>
            ${slideHint}
          </div>
          <div class="item-actions">
            <label class="action-input-wrap">
              <span>Normal</span>
              <input type="number" class="qty-input" min="1" step="1" value="1" aria-label="${consumeModeLabel}" title="${consumeModeLabel}" />
            </label>
            <label class="action-input-wrap">
              <span>Repeat</span>
              <input type="number" class="repeat-input" min="0" step="1" value="0" aria-label="Repeat" title="Repeat" />
            </label>
            <button class="action-btn consume" data-action="consume" data-item-id="${item.id}" title="Consume">+</button>
            <button class="action-btn undo" data-action="undo" data-item-id="${item.id}" title="Undo">-</button>
          </div>
        </div>
      `;
    })
    .join("");

  return `
    <div class="category-block">
      <h4>${escapeHTML(title)}</h4>
      ${rows}
    </div>
  `;
}

function renderStats() {
  const totalItems = state.items.length;
  const lowStockItems = state.items.filter((item) => {
    const qty = toNumber(item.stockQty);
    const level = toNumber(item.reorderLevel);
    return qty > 0 && qty <= level;
  });
  const outStockItems = state.items.filter((item) => toNumber(item.stockQty) <= 0);

  statTotalItems.textContent = totalItems;
  statLowStock.textContent = lowStockItems.length;
  statOutStock.textContent = outStockItems.length;
  if (dashTotalItems) dashTotalItems.textContent = totalItems;
  if (dashLowStock) dashLowStock.textContent = lowStockItems.length;
  if (dashOutStock) dashOutStock.textContent = outStockItems.length;

  const totalSpend = state.purchases.reduce((sum, purchase) => sum + toNumber(purchase.totalAmount), 0);
  statSpend.textContent = currencyFormatter.format(totalSpend);
  if (dashSpend) dashSpend.textContent = currencyFormatter.format(totalSpend);

  renderList(lowStockList, lowStockItems, (item) => {
    const unitName = getUnitName(item.unitId);
    return `${escapeHTML(item.name)} (${toNumber(item.stockQty)} ${escapeHTML(unitName)})`;
  });

  renderList(outStockList, outStockItems, (item) => `${escapeHTML(item.name)} (0)`);

  renderExpiryAlerts();
  renderActivity();
  renderNotifications();
  renderAnalyticsAndAudits();
  renderLogs();
  renderAuditLogs();
}

function getAuditEvents() {
  const events = [];

  state.purchases.forEach((entry) => {
    const item = getItem(entry.itemId);
    events.push({
      id: `p-${entry.id}`,
      sourceId: entry.id,
      sourceCollection: "purchases",
      baseType: "Purchase",
      action: "Purchase",
      profileName: entry.profileName || "Unknown",
      itemId: entry.itemId || null,
      itemName: item?.name || "Item",
      categoryId: item?.categoryId || "uncategorized",
      categoryName: getCategoryName(item?.categoryId),
      quantity: toNumber(entry.quantity),
      stockDelta: toNumber(entry.quantity),
      eventDate: toDate(entry.createdAt),
      rawTime: entry.createdAt,
      note: `Amount ${currencyFormatter.format(toNumber(entry.totalAmount))}`,
      rawDoc: entry
    });
  });

  state.consumptions.forEach((entry) => {
    const item = getItem(entry.itemId);
    const qtyDelta = toNumber(entry.quantity);
    const mode = entry.inputMode === "slides" ? "slides" : "units";
    const count = Math.abs(toNumber(entry.totalInputCount) || toNumber(entry.inputCount) || 0);
    const repeatCount = Math.max(0, Math.abs(toNumber(entry.repeatCount) || 0));
    const unitsPerSlide = toNumber(entry.unitsPerSlide);
    const slideUnitName = getUnitName(entry.slideUnitId);
    const note = mode === "slides"
      ? `${count || Math.abs(qtyDelta / (unitsPerSlide || 1))} slide(s)${repeatCount ? ` (Repeat: ${repeatCount})` : ""}${unitsPerSlide > 0 ? ` (${Math.abs(qtyDelta)} ${slideUnitName})` : ""}`
      : `Qty ${Math.abs(qtyDelta)}${repeatCount ? ` (Repeat: ${repeatCount})` : ""}`;

    events.push({
      id: `c-${entry.id}`,
      sourceId: entry.id,
      sourceCollection: "consumptions",
      baseType: "Consumption",
      action: qtyDelta >= 0 ? "Consumption" : "Consumption Undo",
      inputMode: mode,
      repeatCount: Math.max(0, toNumber(entry.repeatCount)),
      slidesDelta: mode === "slides"
        ? (qtyDelta >= 0 ? 1 : -1) * (count || Math.abs(qtyDelta / (unitsPerSlide || 1)))
        : 0,
      profileName: entry.profileName || "Unknown",
      itemId: entry.itemId || null,
      itemName: item?.name || "Item",
      categoryId: item?.categoryId || "uncategorized",
      categoryName: getCategoryName(item?.categoryId),
      quantity: qtyDelta,
      stockDelta: -qtyDelta,
      eventDate: toDate(entry.createdAt),
      rawTime: entry.createdAt,
      note,
      rawDoc: entry
    });
  });

  state.returns.forEach((entry) => {
    const item = getItem(entry.itemId);
    const qty = toNumber(entry.quantity);
    events.push({
      id: `r-${entry.id}`,
      sourceId: entry.id,
      sourceCollection: "returns",
      baseType: "Return",
      action: "Return",
      profileName: entry.profileName || "Unknown",
      itemId: entry.itemId || null,
      itemName: item?.name || "Item",
      categoryId: item?.categoryId || "uncategorized",
      categoryName: getCategoryName(item?.categoryId),
      quantity: qty,
      stockDelta: -qty,
      eventDate: toDate(entry.createdAt),
      rawTime: entry.createdAt,
      note: entry.reason || "",
      rawDoc: entry
    });
  });

  state.adjustments.forEach((entry) => {
    const item = getItem(entry.itemId);
    const qty = toNumber(entry.quantity);
    const stockDelta = entry.direction === "increase" ? qty : -qty;
    events.push({
      id: `a-${entry.id}`,
      sourceId: entry.id,
      sourceCollection: "adjustments",
      baseType: "Adjustment",
      action: `Adjustment (${entry.direction === "increase" ? "Increase" : "Decrease"})`,
      profileName: entry.profileName || "Unknown",
      itemId: entry.itemId || null,
      itemName: item?.name || "Item",
      categoryId: item?.categoryId || "uncategorized",
      categoryName: getCategoryName(item?.categoryId),
      quantity: qty,
      stockDelta,
      eventDate: toDate(entry.createdAt),
      rawTime: entry.createdAt,
      note: entry.reason || "",
      rawDoc: entry
    });
  });

  return events.sort((a, b) => (b.eventDate?.getTime() || 0) - (a.eventDate?.getTime() || 0));
}

async function appendAuditLog(actionType, entityType, entityId, details = {}) {
  if (!firebaseReady) return;
  try {
    await addDoc(collection(db, "auditLogs"), {
      actionType,
      entityType,
      entityId: entityId || null,
      profileName: getAuditProfile(),
      details,
      createdAt: serverTimestamp()
    });
  } catch (error) {
    console.error("Audit log write failed", error);
  }
}

function getFilteredAuditLogs() {
  const action = auditActionFilter?.value || "all";
  const entity = auditEntityFilter?.value || "all";
  const profile = auditProfileFilter?.value || "all";
  const query = (auditSearch?.value || "").trim().toLowerCase();
  const from = toStartOfDay(auditDateFrom?.value || "");
  const to = toEndOfDay(auditDateTo?.value || "");

  return [...state.auditLogs].filter((entry) => {
    if (action !== "all" && entry.actionType !== action) return false;
    if (entity !== "all" && entry.entityType !== entity) return false;
    if (profile !== "all" && entry.profileName !== profile) return false;

    const createdAt = toDate(entry.createdAt);
    if (from && (!createdAt || createdAt < from)) return false;
    if (to && (!createdAt || createdAt > to)) return false;

    if (query) {
      const haystack = `${entry.entityId || ""} ${entry.entityType || ""} ${entry.actionType || ""} ${JSON.stringify(entry.details || {})}`.toLowerCase();
      if (!haystack.includes(query)) return false;
    }
    return true;
  }).sort((a, b) => timeValue(b.createdAt) - timeValue(a.createdAt));
}

function renderAuditLogs() {
  if (!auditLogsList) return;

  const revertEntries = state.auditLogs.filter((entry) => entry.actionType === "revert" && entry.entityType === "audit-log");
  const revertedMap = new Map();
  revertEntries.forEach((entry) => {
    const targetId = entry?.details?.targetAuditId;
    if (!targetId) return;
    revertedMap.set(targetId, entry);
  });

  if (auditActionFilter) {
    const current = auditActionFilter.value || "all";
    const actions = [...new Set(state.auditLogs.map((entry) => entry.actionType).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    auditActionFilter.innerHTML = '<option value="all">All Actions</option>';
    actions.forEach((action) => {
      const option = document.createElement("option");
      option.value = action;
      option.textContent = action;
      auditActionFilter.appendChild(option);
    });
    auditActionFilter.value = actions.includes(current) ? current : "all";
    ensureSearchableSelect(auditActionFilter);
  }

  if (auditEntityFilter) {
    const current = auditEntityFilter.value || "all";
    const entities = [...new Set(state.auditLogs.map((entry) => entry.entityType).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    auditEntityFilter.innerHTML = '<option value="all">All Entities</option>';
    entities.forEach((entity) => {
      const option = document.createElement("option");
      option.value = entity;
      option.textContent = entity;
      auditEntityFilter.appendChild(option);
    });
    auditEntityFilter.value = entities.includes(current) ? current : "all";
    ensureSearchableSelect(auditEntityFilter);
  }

  if (auditProfileFilter) {
    const current = auditProfileFilter.value || "all";
    const profiles = [...new Set(state.auditLogs.map((entry) => entry.profileName).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    auditProfileFilter.innerHTML = '<option value="all">All Profiles</option>';
    profiles.forEach((profile) => {
      const option = document.createElement("option");
      option.value = profile;
      option.textContent = profile;
      auditProfileFilter.appendChild(option);
    });
    auditProfileFilter.value = profiles.includes(current) ? current : "all";
    ensureSearchableSelect(auditProfileFilter);
  }

  const rows = getFilteredAuditLogs();
  if (auditCount) auditCount.textContent = String(rows.length);

  if (!rows.length) {
    auditLogsList.innerHTML = '<div class="list-empty">No audit logs for selected filters.</div>';
    return;
  }

  auditLogsList.innerHTML = rows
    .slice(0, 500)
    .map((entry) => {
      const isReverted = revertedMap.has(entry.id);
      const revertEntry = revertedMap.get(entry.id);
      const canRevert = !isReverted && entry.actionType !== "revert";

      return `
        <div class="audit-row">
          <div class="audit-row-head">
            <strong>${escapeHTML(entry.actionType || "Action")}</strong>
            <span>${formatDate(entry.createdAt)}</span>
          </div>
          <div>${escapeHTML(entry.entityType || "Unknown")} · ${escapeHTML(entry.entityId || "N/A")}</div>
          <div class="audit-row-meta">By: ${escapeHTML(entry.profileName || "Unknown")}</div>
          ${isReverted ? `<div class="audit-row-meta">Status: Reverted · Reason: ${escapeHTML(revertEntry?.details?.reason || "N/A")}</div>` : ""}
          <div class="log-meta">${escapeHTML(JSON.stringify(entry.details || {}, null, 2))}</div>
          ${canRevert ? `
            <div class="log-actions">
              <button class="secondary-btn" type="button" data-audit-action="revert" data-audit-id="${entry.id}">Revert</button>
            </div>
          ` : ""}
        </div>
      `;
    })
    .join("");
}

function initAuditLogs() {
  if (auditLogsInitialized) return;
  if (!auditLogsList) return;
  auditLogsInitialized = true;

  const rerender = () => renderAuditLogs();
  auditActionFilter?.addEventListener("change", rerender);
  auditEntityFilter?.addEventListener("change", rerender);
  auditProfileFilter?.addEventListener("change", rerender);
  auditDateFrom?.addEventListener("change", rerender);
  auditDateTo?.addEventListener("change", rerender);
  auditSearch?.addEventListener("input", rerender);

  auditLogsList?.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-audit-action]");
    if (!button) return;
    if (!ensureReady()) return;

    const action = button.dataset.auditAction;
    const auditId = button.dataset.auditId;
    if (action !== "revert" || !auditId) return;

    const reason = String(window.prompt("Enter reason for revert:", "") || "").trim();
    if (!reason) {
      showToast("Revert reason is required", true);
      return;
    }

    const target = state.auditLogs.find((entry) => entry.id === auditId);
    if (!target) {
      showToast("Audit entry not found", true);
      return;
    }

    const alreadyReverted = state.auditLogs.some((entry) =>
      entry.actionType === "revert"
      && entry.entityType === "audit-log"
      && entry?.details?.targetAuditId === auditId
    );
    if (alreadyReverted) {
      showToast("Audit already reverted", true);
      return;
    }

    await appendAuditLog("revert", "audit-log", auditId, {
      targetAuditId: auditId,
      targetActionType: target.actionType || "",
      targetEntityType: target.entityType || "",
      targetEntityId: target.entityId || null,
      reason
    });

    await refreshAll();
    showToast("Audit reverted");
  });
}

function getCollectionForBaseType(baseType) {
  if (baseType === "Purchase") return "purchases";
  if (baseType === "Consumption") return "consumptions";
  if (baseType === "Return") return "returns";
  if (baseType === "Adjustment") return "adjustments";
  return "";
}

function computeStockDeltaForLog(baseType, data) {
  const quantity = toNumber(data?.quantity);
  if (baseType === "Purchase") return quantity;
  if (baseType === "Consumption") return -quantity;
  if (baseType === "Return") return -quantity;
  if (baseType === "Adjustment") {
    return data?.direction === "increase" ? quantity : -quantity;
  }
  return 0;
}

function findAuditEventById(eventId) {
  return getAuditEvents().find((event) => event.id === eventId) || null;
}

function getLogEventsFiltered() {
  const action = logsActionFilter?.value || "all";
  const profile = logsProfileFilter?.value || "all";
  const from = toStartOfDay(logsDateFrom?.value || "");
  const to = toEndOfDay(logsDateTo?.value || "");
  const query = (logsItemSearch?.value || "").trim().toLowerCase();

  return getAuditEvents().filter((event) => {
    if (action !== "all" && event.baseType !== action) return false;
    if (profile !== "all" && event.profileName !== profile) return false;
    if (query && !String(event.itemName || "").toLowerCase().includes(query)) return false;
    if (from && (!event.eventDate || event.eventDate < from)) return false;
    if (to && (!event.eventDate || event.eventDate > to)) return false;
    return true;
  });
}

function renderLogActionPanel(event) {
  const isActive = activeLogAction.eventId === event.id;
  if (!isActive) return "";
  const mode = activeLogAction.mode;
  const data = event.rawDoc || {};

  if (mode === "edit") {
    const qty = toNumber(data.quantity);
    const profileName = data.profileName || "";
    const reason = data.reason || "";
    const totalAmount = toNumber(data.totalAmount);
    const repeatCount = toNumber(data.repeatCount);
    const direction = data.direction || "decrease";

    return `
      <form class="inline-form" data-form="logEdit" data-event-id="${event.id}">
        <div class="form-row">
          <label>Quantity</label>
          <input type="number" name="quantity" step="0.01" value="${qty}" required />
        </div>
        <div class="form-row">
          <label>Profile</label>
          <input type="text" name="profileName" value="${escapeHTML(profileName)}" required />
        </div>
        ${event.baseType === "Purchase" ? `
          <div class="form-row">
            <label>Total Amount</label>
            <input type="number" name="totalAmount" min="0" step="0.01" value="${totalAmount}" required />
          </div>
        ` : ""}
        ${event.baseType === "Adjustment" ? `
          <div class="form-row">
            <label>Direction</label>
            <select name="direction">
              <option value="decrease" ${direction === "decrease" ? "selected" : ""}>Decrease</option>
              <option value="increase" ${direction === "increase" ? "selected" : ""}>Increase</option>
            </select>
          </div>
        ` : ""}
        ${event.baseType === "Consumption" ? `
          <div class="form-row">
            <label>Repeat Count</label>
            <input type="number" name="repeatCount" min="0" step="1" value="${repeatCount}" />
          </div>
        ` : ""}
        ${event.baseType === "Return" || event.baseType === "Adjustment" ? `
          <div class="form-row">
            <label>Reason</label>
            <input type="text" name="reason" value="${escapeHTML(reason)}" />
          </div>
        ` : ""}
        <div class="inline-actions">
          <button type="submit" class="primary">Save Edit</button>
          <button type="button" class="secondary-btn" data-log-action="close" data-event-id="${event.id}">Cancel</button>
        </div>
      </form>
    `;
  }

  if (mode === "manage") {
    const item = getItem(event.itemId);
    const stockQty = toNumber(item?.stockQty);
    return `
      <form class="inline-form" data-form="logManage" data-event-id="${event.id}">
        <div class="form-row">
          <label>Set Linked Item Stock</label>
          <input type="number" name="stockQty" min="0" step="0.01" value="${stockQty}" required />
        </div>
        <div class="inline-actions">
          <button type="submit" class="primary">Update Stock</button>
          <button type="button" class="secondary-btn" data-log-action="close" data-event-id="${event.id}">Cancel</button>
        </div>
      </form>
    `;
  }

  if (mode === "configure") {
    return `
      <form class="inline-form" data-form="logConfigure" data-event-id="${event.id}">
        <div class="form-row">
          <label>Metadata JSON (merge patch)</label>
          <textarea name="metadata" rows="8">${escapeHTML(JSON.stringify(data, null, 2))}</textarea>
        </div>
        <div class="inline-actions">
          <button type="submit" class="primary">Save Config</button>
          <button type="button" class="secondary-btn" data-log-action="close" data-event-id="${event.id}">Cancel</button>
        </div>
      </form>
    `;
  }

  return "";
}

function renderLogs() {
  if (!logsList) return;

  const allEvents = getAuditEvents();
  if (logsProfileFilter) {
    const current = logsProfileFilter.value || "all";
    const profiles = [...new Set(allEvents.map((event) => event.profileName))].sort((a, b) => a.localeCompare(b));
    logsProfileFilter.innerHTML = '<option value="all">All Profiles</option>';
    profiles.forEach((profile) => {
      const option = document.createElement("option");
      option.value = profile;
      option.textContent = profile;
      logsProfileFilter.appendChild(option);
    });
    logsProfileFilter.value = profiles.includes(current) ? current : "all";
    ensureSearchableSelect(logsProfileFilter);
  }

  const events = getLogEventsFiltered();
  if (logsCount) logsCount.textContent = String(events.length);

  if (!events.length) {
    logsList.innerHTML = '<div class="list-empty">No logs found for selected filters.</div>';
    return;
  }

  logsList.innerHTML = events
    .slice(0, 300)
    .map((event) => {
      const metadata = {
        id: event.id,
        type: event.baseType,
        sourceCollection: event.sourceCollection,
        sourceId: event.sourceId,
        itemId: event.itemId,
        inputMode: event.inputMode || null,
        repeatCount: toNumber(event.repeatCount || 0),
        quantity: toNumber(event.quantity),
        stockDelta: toNumber(event.stockDelta)
      };

      return `
        <div class="audit-row">
          <div class="audit-row-head">
            <strong>${escapeHTML(event.action)}</strong>
            <span>${formatDate(event.rawTime)}</span>
          </div>
          <div>${escapeHTML(event.itemName)} · ${escapeHTML(event.categoryName)}</div>
          <div class="audit-row-meta">Qty: ${toNumber(event.quantity)} · Stock Δ: ${toNumber(event.stockDelta)} · By: ${escapeHTML(event.profileName)}</div>
          ${event.note ? `<div class="audit-row-meta">${escapeHTML(event.note)}</div>` : ""}
          <div class="log-meta">${escapeHTML(JSON.stringify(metadata, null, 2))}</div>
          <div class="log-actions">
            <button class="secondary-btn ${activeLogAction.eventId === event.id && activeLogAction.mode === "edit" ? "active" : ""}" data-log-action="edit" data-event-id="${event.id}" type="button">Edit</button>
            <button class="secondary-btn ${activeLogAction.eventId === event.id && activeLogAction.mode === "manage" ? "active" : ""}" data-log-action="manage" data-event-id="${event.id}" type="button">Manage</button>
            <button class="secondary-btn ${activeLogAction.eventId === event.id && activeLogAction.mode === "configure" ? "active" : ""}" data-log-action="configure" data-event-id="${event.id}" type="button">Configure</button>
            <button class="secondary-btn" data-log-action="delete" data-event-id="${event.id}" type="button">Delete</button>
          </div>
          ${renderLogActionPanel(event)}
        </div>
      `;
    })
    .join("");
}

async function withLogTransaction(eventId, buildUpdates) {
  const event = findAuditEventById(eventId);
  if (!event) throw new Error("Log entry not found");

  const collectionName = getCollectionForBaseType(event.baseType);
  if (!collectionName) throw new Error("Unsupported log type");

  await runTransaction(db, async (tx) => {
    const logRef = doc(db, collectionName, event.sourceId);
    const logSnap = await tx.get(logRef);
    if (!logSnap.exists()) throw new Error("Log entry no longer exists");
    const oldData = logSnap.data();
    const updates = buildUpdates(oldData, event.baseType) || {};

    if (Object.prototype.hasOwnProperty.call(updates, "itemId") && updates.itemId !== oldData.itemId) {
      throw new Error("Changing itemId is not allowed");
    }

    const newData = { ...oldData, ...updates };
    const oldDelta = computeStockDeltaForLog(event.baseType, oldData);
    const newDelta = computeStockDeltaForLog(event.baseType, newData);
    const deltaDiff = newDelta - oldDelta;

    if (oldData.itemId && deltaDiff !== 0) {
      const itemRef = doc(db, "items", oldData.itemId);
      const itemSnap = await tx.get(itemRef);
      if (!itemSnap.exists()) throw new Error("Linked item not found");
      const currentStock = toNumber(itemSnap.data().stockQty);
      const nextStock = currentStock + deltaDiff;
      if (nextStock < 0) throw new Error("Operation would make stock negative");
      tx.update(itemRef, {
        stockQty: nextStock,
        lastUpdatedBy: getAuditProfile(),
        updatedAt: serverTimestamp()
      });
    }

    tx.update(logRef, {
      ...updates,
      lastUpdatedBy: getAuditProfile(),
      updatedAt: serverTimestamp()
    });
  });
}

async function deleteLogEvent(eventId, options = {}) {
  const shouldRevertStock = options.revertStock !== false;
  const event = findAuditEventById(eventId);
  if (!event) throw new Error("Log entry not found");
  const collectionName = getCollectionForBaseType(event.baseType);
  if (!collectionName) throw new Error("Unsupported log type");

  const result = {
    revertedStock: false,
    skippedRevert: !shouldRevertStock
  };

  await runTransaction(db, async (tx) => {
    const logRef = doc(db, collectionName, event.sourceId);
    const logSnap = await tx.get(logRef);
    if (!logSnap.exists()) throw new Error("Log entry no longer exists");
    const logData = logSnap.data();
    const stockDelta = computeStockDeltaForLog(event.baseType, logData);

    if (logData.itemId && shouldRevertStock) {
      const itemRef = doc(db, "items", logData.itemId);
      const itemSnap = await tx.get(itemRef);
      if (!itemSnap.exists()) throw new Error("Linked item not found");
      const currentStock = toNumber(itemSnap.data().stockQty);
      const nextStock = currentStock - stockDelta;
      if (nextStock < 0) throw new Error("Delete would make stock negative");
      tx.update(itemRef, {
        stockQty: nextStock,
        lastUpdatedBy: getAuditProfile(),
        updatedAt: serverTimestamp()
      });
      result.revertedStock = true;
    }

    tx.delete(logRef);
  });

  return result;
}

function setStatsDateRangeFromPreset() {
  if (!statsDatePreset || !statsDateFrom || !statsDateTo) return;
  const preset = statsDatePreset.value;
  const today = new Date();
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  if (preset === "custom") return;

  if (preset === "all") {
    statsDateFrom.value = "";
    statsDateTo.value = "";
    return;
  }

  let start = new Date(end);
  if (preset === "today") {
    start = new Date(end);
  } else if (preset === "yesterday") {
    start.setDate(start.getDate() - 1);
    end.setDate(end.getDate() - 1);
  } else if (preset === "7d") {
    start.setDate(start.getDate() - 6);
  } else if (preset === "30d") {
    start.setDate(start.getDate() - 29);
  }

  statsDateFrom.value = isoDate(start);
  statsDateTo.value = isoDate(end);
}

function toStartOfDay(value) {
  if (!value) return null;
  return new Date(`${value}T00:00:00`);
}

function toEndOfDay(value) {
  if (!value) return null;
  return new Date(`${value}T23:59:59.999`);
}

function getFilteredAuditEvents() {
  const allEvents = getAuditEvents();
  const profile = statsProfileFilter?.value || "all";
  const action = statsActionFilter?.value || "all";
  const category = statsCategoryFilter?.value || "all";
  const query = (statsItemSearch?.value || "").trim().toLowerCase();
  const from = toStartOfDay(statsDateFrom?.value || "");
  const to = toEndOfDay(statsDateTo?.value || "");

  return allEvents.filter((event) => {
    if (profile !== "all" && event.profileName !== profile) return false;
    if (action !== "all" && event.baseType !== action) return false;
    if (category !== "all" && event.categoryId !== category) return false;
    if (query && !event.itemName.toLowerCase().includes(query)) return false;
    if (from && (!event.eventDate || event.eventDate < from)) return false;
    if (to && (!event.eventDate || event.eventDate > to)) return false;
    return true;
  });
}

function updateFilterOptions() {
  if (!statsProfileFilter || !statsCategoryFilter) return;

  const currentProfile = statsProfileFilter.value || "all";
  const currentCategory = statsCategoryFilter.value || "all";

  const profiles = [...new Set(getAuditEvents().map((event) => event.profileName))].sort((a, b) => a.localeCompare(b));
  statsProfileFilter.innerHTML = '<option value="all">All Profiles</option>';
  profiles.forEach((profile) => {
    const option = document.createElement("option");
    option.value = profile;
    option.textContent = profile;
    statsProfileFilter.appendChild(option);
  });
  statsProfileFilter.value = profiles.includes(currentProfile) ? currentProfile : "all";
  ensureSearchableSelect(statsProfileFilter);

  statsCategoryFilter.innerHTML = '<option value="all">All Categories</option>';
  const categories = [...state.categories].sort((a, b) => a.name.localeCompare(b.name));
  categories.forEach((categoryEntry) => {
    const option = document.createElement("option");
    option.value = categoryEntry.id;
    option.textContent = categoryEntry.name;
    statsCategoryFilter.appendChild(option);
  });
  statsCategoryFilter.value = categories.some((entry) => entry.id === currentCategory) ? currentCategory : "all";
  ensureSearchableSelect(statsCategoryFilter);
}

function renderAnalyticsAndAudits() {
  if (!analyticsEventCount || !statsAuditList || !analyticsTopConsumed || !analyticsConsumedAll || !analyticsSlidesUsed) return;

  updateFilterOptions();
  const events = getFilteredAuditEvents();

  const purchaseQty = events
    .filter((event) => event.baseType === "Purchase")
    .reduce((sum, event) => sum + toNumber(event.quantity), 0);
  const netConsumedQty = events
    .filter((event) => event.baseType === "Consumption")
    .reduce((sum, event) => sum + toNumber(event.quantity), 0);
  const netStockMovement = events.reduce((sum, event) => sum + toNumber(event.stockDelta), 0);

  analyticsEventCount.textContent = String(events.length);
  analyticsPurchaseQty.textContent = purchaseQty.toFixed(2);
  analyticsConsumptionQty.textContent = netConsumedQty.toFixed(2);
  analyticsNetStock.textContent = netStockMovement.toFixed(2);

  const consumedMap = new Map();
  events
    .filter((event) => event.baseType === "Consumption")
    .forEach((event) => {
      consumedMap.set(event.itemName, (consumedMap.get(event.itemName) || 0) + toNumber(event.quantity));
    });

  const topConsumed = [...consumedMap.entries()]
    .filter(([, qty]) => qty > 0)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  const ihcSlidesUsed = events
    .filter((event) => event.baseType === "Consumption"
      && event.inputMode === "slides"
      && String(event.categoryName || "").toLowerCase() === "ihc marker")
    .reduce((sum, event) => sum + toNumber(event.slidesDelta), 0);
  analyticsSlidesUsed.textContent = ihcSlidesUsed.toFixed(2);

  if (!topConsumed.length) {
    analyticsTopConsumed.innerHTML = '<div class="list-empty">No consumption data for selected filters.</div>';
  } else {
    analyticsTopConsumed.innerHTML = topConsumed
      .map(([name, qty]) => `
        <div class="list-item">
          <strong>${escapeHTML(name)}</strong>
          <span>${qty.toFixed(2)}</span>
        </div>
      `)
      .join("");
  }

  const consumedByItem = new Map();
  events
    .filter((event) => event.baseType === "Consumption")
    .forEach((event) => {
      consumedByItem.set(event.itemName, (consumedByItem.get(event.itemName) || 0) + toNumber(event.quantity));
    });

  const allItemsView = [...state.items]
    .sort((a, b) => a.name.localeCompare(b.name))
    .filter((item) => {
      const selectedCategory = statsCategoryFilter?.value || "all";
      const search = (statsItemSearch?.value || "").trim().toLowerCase();
      if (selectedCategory !== "all" && item.categoryId !== selectedCategory) return false;
      if (search && !String(item.name || "").toLowerCase().includes(search)) return false;
      return true;
    })
    .map((item) => {
      const consumedQty = toNumber(consumedByItem.get(item.name) || 0);
      const categoryName = getCategoryName(item.categoryId);
      return {
        name: item.name,
        categoryName,
        consumedQty
      };
    });

  if (!allItemsView.length) {
    analyticsConsumedAll.innerHTML = '<div class="list-empty">No items match selected filters.</div>';
  } else {
    analyticsConsumedAll.innerHTML = allItemsView
      .map((entry) => `
        <div class="audit-row">
          <div class="audit-row-head">
            <strong>${escapeHTML(entry.name)}</strong>
            <span>${entry.consumedQty.toFixed(2)}</span>
          </div>
          <div class="audit-row-meta">Category: ${escapeHTML(entry.categoryName)} · Net consumed qty (filtered)</div>
        </div>
      `)
      .join("");
  }

  const repeatEventsAll = events
    .filter((event) => event.baseType === "Consumption" && toNumber(event.repeatCount) > 0)
    .map((event) => ({
      ...event,
      repeatMonth: monthKey(event.eventDate),
      signedRepeatSlides: (toNumber(event.quantity) >= 0 ? 1 : -1) * Math.max(0, toNumber(event.repeatCount))
    }));

  if (repeatMonthFilter) {
    const currentMonth = repeatMonthFilter.value || "all";
    const months = [...new Set(repeatEventsAll.map((event) => event.repeatMonth).filter(Boolean))]
      .sort((a, b) => b.localeCompare(a));

    repeatMonthFilter.innerHTML = '<option value="all">All Months</option>';
    months.forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = monthLabelFromKey(value);
      repeatMonthFilter.appendChild(option);
    });

    repeatMonthFilter.value = (currentMonth !== "all" && months.includes(currentMonth)) ? currentMonth : "all";
    ensureSearchableSelect(repeatMonthFilter);
  }

  const selectedRepeatMonth = repeatMonthFilter?.value || "all";
  const selectedRepeatDate = repeatSpecificDate?.value || "";
  const repeatSearchQuery = (repeatItemSearch?.value || "").trim().toLowerCase();

  let repeatEvents = selectedRepeatMonth === "all"
    ? repeatEventsAll
    : repeatEventsAll.filter((event) => event.repeatMonth === selectedRepeatMonth);

  if (selectedRepeatDate) {
    repeatEvents = repeatEvents.filter((event) => event.eventDate && isoDate(event.eventDate) === selectedRepeatDate);
  }

  if (repeatSearchQuery) {
    repeatEvents = repeatEvents.filter((event) => String(event.itemName || "").toLowerCase().includes(repeatSearchQuery));
  }

  const totalRepeatLogs = repeatEvents.length;
  const totalRepeatSlidesNet = Math.max(0, repeatEvents.reduce((sum, event) => sum + toNumber(event.signedRepeatSlides), 0));

  if (repeatNetSlides) repeatNetSlides.textContent = totalRepeatSlidesNet.toFixed(2);
  if (repeatLogCount) repeatLogCount.textContent = String(totalRepeatLogs);

  const repeatByItem = new Map();
  repeatEvents.forEach((event) => {
    const existing = repeatByItem.get(event.itemName) || { itemName: event.itemName, categoryName: event.categoryName, netRepeatSlides: 0, logs: 0 };
    existing.netRepeatSlides += toNumber(event.signedRepeatSlides);
    existing.logs += 1;
    repeatByItem.set(event.itemName, existing);
  });

  const repeatByItemRows = [...repeatByItem.values()]
    .filter((entry) => entry.netRepeatSlides > 0)
    .sort((a, b) => b.netRepeatSlides - a.netRepeatSlides);

  if (repeatByItemList) {
    if (!repeatByItemRows.length) {
      repeatByItemList.innerHTML = '<div class="list-empty">No repeat entries for selected filters.</div>';
    } else {
      repeatByItemList.innerHTML = repeatByItemRows
        .map((entry) => `
          <div class="audit-row">
            <div class="audit-row-head">
              <strong>${escapeHTML(entry.itemName)}</strong>
              <span>${entry.netRepeatSlides.toFixed(2)} slide(s)</span>
            </div>
            <div class="audit-row-meta">Category: ${escapeHTML(entry.categoryName)} · Logs: ${entry.logs}</div>
          </div>
        `)
        .join("");
    }
  }

  if (repeatAuditList) {
    if (!repeatEvents.length) {
      repeatAuditList.innerHTML = '<div class="list-empty">No repeat audit entries for selected filters.</div>';
    } else {
      repeatAuditList.innerHTML = repeatEvents
        .slice(0, 200)
        .map((event) => `
          <div class="audit-row">
            <div class="audit-row-head">
              <strong>${escapeHTML(event.itemName)}</strong>
              <span>${formatDate(event.rawTime)}</span>
            </div>
            <div class="audit-row-meta">Repeat: ${toNumber(event.repeatCount)} slide(s) · Net effect: ${toNumber(event.signedRepeatSlides).toFixed(2)} · Qty: ${toNumber(event.quantity)} · By: ${escapeHTML(event.profileName)}</div>
            <div class="audit-row-meta">${escapeHTML(event.categoryName)}${event.note ? ` · ${escapeHTML(event.note)}` : ""}</div>
          </div>
        `)
        .join("");
    }
  }

  if (!events.length) {
    statsAuditList.innerHTML = '<div class="list-empty">No audit records for selected filters.</div>';
    return;
  }

  statsAuditList.innerHTML = events
    .slice(0, 200)
    .map((event) => `
      <div class="audit-row">
        <div class="audit-row-head">
          <strong>${escapeHTML(event.action)}</strong>
          <span>${formatDate(event.rawTime)}</span>
        </div>
        <div>${escapeHTML(event.itemName)} · ${escapeHTML(event.categoryName)}</div>
        <div class="audit-row-meta">Qty: ${toNumber(event.quantity)} · Stock Δ: ${toNumber(event.stockDelta)} · By: ${escapeHTML(event.profileName)}</div>
        ${event.note ? `<div class="audit-row-meta">${escapeHTML(event.note)}</div>` : ""}
      </div>
    `)
    .join("");
}

function csvCell(value) {
  const text = String(value ?? "").replace(/"/g, '""');
  return `"${text}"`;
}

function exportAuditsCsv() {
  const events = getFilteredAuditEvents();
  if (!events.length) {
    showToast("No records to export", true);
    return;
  }

  const headers = ["Date", "Action", "Profile", "Item", "Category", "Quantity", "StockDelta", "Note"];
  const rows = events.map((event) => [
    event.eventDate ? isoDate(event.eventDate) : "",
    event.action,
    event.profileName,
    event.itemName,
    event.categoryName,
    toNumber(event.quantity),
    toNumber(event.stockDelta),
    event.note || ""
  ]);

  const csv = [headers, ...rows].map((row) => row.map(csvCell).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `ims-audits-${isoDate(new Date())}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function initStatsFilters() {
  if (statsFiltersInitialized) return;
  if (!statsDatePreset) return;
  statsFiltersInitialized = true;

  setStatsDateRangeFromPreset();

  const rerender = () => renderAnalyticsAndAudits();

  statsDatePreset?.addEventListener("change", () => {
    setStatsDateRangeFromPreset();
    rerender();
  });
  statsDateFrom?.addEventListener("change", rerender);
  statsDateTo?.addEventListener("change", rerender);
  statsProfileFilter?.addEventListener("change", rerender);
  statsActionFilter?.addEventListener("change", rerender);
  statsCategoryFilter?.addEventListener("change", rerender);
  statsItemSearch?.addEventListener("input", rerender);
  repeatMonthFilter?.addEventListener("change", rerender);
  repeatSpecificDate?.addEventListener("change", rerender);
  repeatItemSearch?.addEventListener("input", rerender);

  statsResetFilters?.addEventListener("click", () => {
    if (statsDatePreset) statsDatePreset.value = "30d";
    if (statsActionFilter) statsActionFilter.value = "all";
    if (statsProfileFilter) statsProfileFilter.value = "all";
    if (statsCategoryFilter) statsCategoryFilter.value = "all";
    if (statsItemSearch) statsItemSearch.value = "";
    if (repeatMonthFilter) repeatMonthFilter.value = "all";
    if (repeatSpecificDate) repeatSpecificDate.value = "";
    if (repeatItemSearch) repeatItemSearch.value = "";
    setStatsDateRangeFromPreset();
    ensureSearchableSelect(statsDatePreset);
    ensureSearchableSelect(statsActionFilter);
    ensureSearchableSelect(statsProfileFilter);
    ensureSearchableSelect(statsCategoryFilter);
    ensureSearchableSelect(repeatMonthFilter);
    rerender();
  });

  statsExportCsv?.addEventListener("click", exportAuditsCsv);
}

function initLogs() {
  if (!logsList) return;

  const rerender = () => renderLogs();

  logsActionFilter?.addEventListener("change", rerender);
  logsProfileFilter?.addEventListener("change", rerender);
  logsDateFrom?.addEventListener("change", rerender);
  logsDateTo?.addEventListener("change", rerender);
  logsItemSearch?.addEventListener("input", rerender);

  logsList.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-log-action]");
    if (!button) return;
    if (!ensureReady()) return;

    const action = button.dataset.logAction;
    const eventId = button.dataset.eventId;
    if (!eventId) return;

    if (action === "close") {
      activeLogAction = { eventId: null, mode: null };
      renderLogs();
      return;
    }

    if (action === "delete") {
      const confirmed = window.confirm("Delete this log entry and revert linked stock?");
      if (!confirmed) return;
      const deleteReason = String(window.prompt("Enter delete reason:", "") || "").trim();
      if (!deleteReason) {
        showToast("Delete reason is required", true);
        return;
      }
      try {
        const beforeDelete = findAuditEventById(eventId);
        let deleteResult;
        try {
          deleteResult = await deleteLogEvent(eventId, { revertStock: true });
        } catch (error) {
          const message = String(error?.message || "").toLowerCase();
          if (!message.includes("stock negative")) {
            throw error;
          }
          deleteResult = await deleteLogEvent(eventId, { revertStock: false });
        }
        await appendAuditLog("delete", "action-log", eventId, {
          deletedEvent: beforeDelete ? {
            action: beforeDelete.action,
            baseType: beforeDelete.baseType,
            itemId: beforeDelete.itemId,
            itemName: beforeDelete.itemName,
            quantity: beforeDelete.quantity,
            stockDelta: beforeDelete.stockDelta
          } : null,
          reason: deleteReason,
          revertedStock: Boolean(deleteResult?.revertedStock),
          skippedRevert: Boolean(deleteResult?.skippedRevert)
        });
        activeLogAction = { eventId: null, mode: null };
        await refreshAll();
        showToast(deleteResult?.revertedStock ? "Log deleted" : "Log deleted (stock revert skipped)");
      } catch (error) {
        showToast(friendlyFirestoreError(error, "Delete failed"), true);
      }
      return;
    }

    if (activeLogAction.eventId === eventId && activeLogAction.mode === action) {
      activeLogAction = { eventId: null, mode: null };
    } else {
      activeLogAction = { eventId, mode: action };
    }
    renderLogs();
  });

  logsList.addEventListener("submit", async (event) => {
    const form = event.target.closest("form[data-form]");
    if (!form) return;
    event.preventDefault();
    if (!ensureReady()) return;

    const formType = form.dataset.form;
    const eventId = form.dataset.eventId;
    if (!eventId) return;

    try {
      if (formType === "logEdit") {
        await withLogTransaction(eventId, (_oldData, baseType) => {
          const quantity = toNumber(form.elements.quantity.value);
          const profileName = String(form.elements.profileName.value || "").trim();

          if (!profileName) throw new Error("Profile is required");
          if (baseType !== "Consumption" && quantity <= 0) throw new Error("Quantity must be greater than 0");
          if (baseType === "Consumption" && quantity === 0) throw new Error("Consumption quantity cannot be 0");

          const updates = {
            quantity,
            profileName
          };

          if (baseType === "Purchase") {
            updates.totalAmount = toNumber(form.elements.totalAmount.value);
          }
          if (baseType === "Return") {
            updates.reason = String(form.elements.reason.value || "").trim();
          }
          if (baseType === "Adjustment") {
            updates.direction = form.elements.direction.value === "increase" ? "increase" : "decrease";
            updates.reason = String(form.elements.reason.value || "").trim();
          }
          if (baseType === "Consumption") {
            updates.repeatCount = Math.max(0, Math.trunc(toNumber(form.elements.repeatCount.value)));
          }

          return updates;
        });
        await appendAuditLog("update", "action-log", eventId, {
          mode: "edit"
        });
      }

      if (formType === "logConfigure") {
        await withLogTransaction(eventId, (oldData) => {
          const raw = String(form.elements.metadata.value || "{}");
          let parsed;
          try {
            parsed = JSON.parse(raw);
          } catch {
            throw new Error("Invalid JSON");
          }

          if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
            throw new Error("Metadata must be a JSON object");
          }
          if (Object.prototype.hasOwnProperty.call(parsed, "itemId") && parsed.itemId !== oldData.itemId) {
            throw new Error("Changing itemId is not allowed");
          }
          return parsed;
        });
        await appendAuditLog("update", "action-log", eventId, {
          mode: "configure"
        });
      }

      if (formType === "logManage") {
        const entry = findAuditEventById(eventId);
        if (!entry?.itemId) throw new Error("Linked item not found for this log");
        const stockQty = toNumber(form.elements.stockQty.value);
        if (stockQty < 0) throw new Error("Stock cannot be negative");

        await updateDoc(doc(db, "items", entry.itemId), {
          stockQty,
          lastUpdatedBy: getAuditProfile(),
          updatedAt: serverTimestamp()
        });
        await appendAuditLog("update", "item", entry.itemId, {
          mode: "manage-from-log",
          sourceEventId: eventId,
          stockQty
        });
      }

      activeLogAction = { eventId: null, mode: null };
      await refreshAll();
      showToast("Log updated");
    } catch (error) {
      showToast(friendlyFirestoreError(error, "Failed to update log"), true);
    }
  });
}

function renderNotifications() {
  if (!notificationList || !notificationCount) return;

  const lowStockItems = state.items.filter((item) => {
    const qty = toNumber(item.stockQty);
    const level = toNumber(item.reorderLevel);
    return qty > 0 && qty <= level;
  });
  const outStockItems = state.items.filter((item) => toNumber(item.stockQty) <= 0);

  const alerts = [];

  lowStockItems.forEach((item) => {
    alerts.push({
      type: "Low Stock",
      text: `${item.name} is low (${toNumber(item.stockQty)} ${getUnitName(item.unitId)})`
    });
  });

  outStockItems.forEach((item) => {
    alerts.push({
      type: "Out of Stock",
      text: `${item.name} is out of stock`
    });
  });

  void maybeSendStockAlertEmail(lowStockItems, outStockItems);

  notificationCount.textContent = String(alerts.length);

  if (!alerts.length) {
    notificationList.innerHTML = '<div class="list-empty">No alerts right now.</div>';
    return;
  }

  notificationList.innerHTML = alerts
    .map((alert) => `
      <div class="notification-item">
        <strong>${escapeHTML(alert.type)}</strong>
        <div class="muted">${escapeHTML(alert.text)}</div>
      </div>
    `)
    .join("");
}

function renderList(container, items, labelFn) {
  if (!items.length) {
    container.innerHTML = "<div class=\"list-empty\">Nothing to show.</div>";
    return;
  }

  container.innerHTML = items
    .map((item) => `
      <div class="list-item">
        <strong>${labelFn(item)}</strong>
      </div>
    `)
    .join("");
}

function renderExpiryAlerts() {
  const today = new Date();
  const warningDays = 14;

  const expiring = state.items
    .filter((item) => item.hasExpiry && item.expiryDate)
    .map((item) => {
      const expiry = toDate(item.expiryDate);
      const diffDays = expiry ? Math.ceil((expiry - today) / (1000 * 60 * 60 * 24)) : null;
      return { item, expiry, diffDays };
    })
    .filter(({ diffDays }) => diffDays !== null && diffDays <= warningDays);

  if (!expiring.length) {
    expiryList.innerHTML = "<div class=\"list-empty\">No expiry alerts in the next 14 days.</div>";
    return;
  }

  expiryList.innerHTML = expiring
    .sort((a, b) => a.diffDays - b.diffDays)
    .map(({ item, expiry, diffDays }) => {
      const label = `${escapeHTML(item.name)} · ${diffDays} day(s)`;
      return `
        <div class="list-item">
          <strong>${label}</strong>
          <span>${formatDate(expiry)}</span>
        </div>
      `;
    })
    .join("");
}

function renderActivity() {
  const activities = [];

  state.purchases.forEach((purchase) => {
    const profileName = purchase.profileName || "Unknown";
    activities.push({
      type: "Purchase",
      label: `${getItem(purchase.itemId)?.name || "Item"} · Qty ${toNumber(purchase.quantity)} · By ${profileName}`,
      time: purchase.createdAt
    });
  });

  state.consumptions.forEach((consumption) => {
    const qty = toNumber(consumption.quantity);
    const mode = consumption.inputMode === "slides" ? "slides" : "units";
    const count = Math.abs(toNumber(consumption.totalInputCount) || toNumber(consumption.inputCount) || 0);
    const repeatCount = Math.max(0, Math.abs(toNumber(consumption.repeatCount) || 0));
    const unitsPerSlide = toNumber(consumption.unitsPerSlide);
    const slideUnitName = getUnitName(consumption.slideUnitId);
    const amountLabel = mode === "slides"
      ? `${count || Math.abs(qty / (unitsPerSlide || 1))} slide(s)${repeatCount ? ` (Repeat: ${repeatCount})` : ""}${unitsPerSlide > 0 ? ` (${Math.abs(qty)} ${slideUnitName})` : ""}`
      : `${Math.abs(qty)}${repeatCount ? ` (Repeat: ${repeatCount})` : ""}`;
    const profileName = consumption.profileName || "Unknown";
    const label = `${getItem(consumption.itemId)?.name || "Item"} · ${qty >= 0 ? "Consumed" : "Undo"} ${amountLabel} · By ${profileName}`;
    activities.push({
      type: "Consumption",
      label,
      time: consumption.createdAt
    });
  });

  state.returns.forEach((entry) => {
    const profileName = entry.profileName || "Unknown";
    activities.push({
      type: "Return",
      label: `${getItem(entry.itemId)?.name || "Item"} · Qty ${toNumber(entry.quantity)} · By ${profileName}`,
      time: entry.createdAt
    });
  });

  state.adjustments.forEach((entry) => {
    const direction = entry.direction === "increase" ? "Increase" : "Decrease";
    const profileName = entry.profileName || "Unknown";
    activities.push({
      type: "Adjustment",
      label: `${getItem(entry.itemId)?.name || "Item"} · ${direction} ${toNumber(entry.quantity)} · By ${profileName}`,
      time: entry.createdAt
    });
  });

  if (!activities.length) {
    activityList.innerHTML = "<div class=\"list-empty\">No activity yet.</div>";
    return;
  }

  const sorted = activities
    .sort((a, b) => timeValue(b.time) - timeValue(a.time))
    .slice(0, 10);

  activityList.innerHTML = sorted
    .map((activity) => `
      <div class="list-item">
        <div>
          <strong>${escapeHTML(activity.type)}</strong>
          <div class="muted">${escapeHTML(activity.label)}</div>
        </div>
        <span>${formatDate(activity.time)}</span>
      </div>
    `)
    .join("");
}

function updatePurchaseTotal() {
  if (!state.items.length) return;
  if (purchaseTotalManual) return;
  const item = getItem(purchaseItem.value);
  const qty = toNumber(purchaseQty.value) || 1;
  const total = item ? qty * toNumber(item.price) : 0;
  purchaseTotal.value = total.toFixed(2);
}

navButtons.forEach((btn) => {
  btn.addEventListener("click", () => {
    navButtons.forEach((item) => item.classList.remove("active"));
    btn.classList.add("active");
    const target = btn.dataset.page;
    pages.forEach((page) => page.classList.toggle("active", page.id === target));
  });
});

featureButtons.forEach((btn) => {
  btn.addEventListener("click", () => {
    const target = btn.dataset.feature;
    featureButtons.forEach((item) => item.classList.remove("active"));
    btn.classList.add("active");
    featurePanels.forEach((panel) => panel.classList.toggle("active", panel.dataset.featurePanel === target));
  });
});

notificationToggle?.addEventListener("click", () => {
  notificationPanel?.classList.toggle("open");
});

document.addEventListener("click", (event) => {
  if (!notificationPanel || !notificationToggle) return;
  const target = event.target;
  if (!(target instanceof Element)) return;
  if (target.closest(".notification-center")) return;
  notificationPanel.classList.remove("open");
});

itemHasExpiry.addEventListener("change", () => {
  expiryRow.style.display = itemHasExpiry.checked ? "flex" : "none";
});

function toggleCategorySlideConfig() {
  if (!categorySlideConfig || !categorySlideBased) return;
  categorySlideConfig.style.display = categorySlideBased.checked ? "block" : "none";
}

categorySlideBased?.addEventListener("change", toggleCategorySlideConfig);

expiryRow.style.display = "none";
toggleCategorySlideConfig();

purchaseItem.addEventListener("change", () => {
  purchaseTotalManual = false;
  updatePurchaseTotal();
});

purchaseQty.addEventListener("input", updatePurchaseTotal);

purchaseTotal.addEventListener("input", () => {
  purchaseTotalManual = true;
});

consumptionSearch?.addEventListener("input", () => {
  consumptionSearchQuery = consumptionSearch.value || "";
  renderConsumption();
});

consumptionList.addEventListener("click", async (event) => {
  const button = event.target.closest("button[data-action]");
  if (!button) return;
  if (!ensureReady()) return;

  const action = button.dataset.action;
  const itemId = button.dataset.itemId;
  const row = button.closest(".item-row");
  const qtyInput = row?.querySelector(".qty-input");
  const repeatInput = row?.querySelector(".repeat-input");
  const requestedInput = toNumber(qtyInput?.value || 1);
  const requestedRepeat = toNumber(repeatInput?.value || 0);
  const slideBased = row?.dataset.slideBased === "1";
  const unitsPerSlide = toNumber(row?.dataset.unitsPerSlide);
  const slideUnitId = row?.dataset.slideUnitId || null;
  const inputCount = slideBased
    ? Math.max(1, Math.trunc(requestedInput || 1))
    : Math.max(1, requestedInput || 1);
  const repeatCount = Math.max(0, Math.trunc(requestedRepeat || 0));
  const totalInputCount = inputCount;

  if (slideBased && unitsPerSlide <= 0) {
    showToast("Invalid slide configuration for category", true);
    return;
  }

  const stockChange = slideBased ? totalInputCount * unitsPerSlide : totalInputCount;
  const delta = action === "consume" ? stockChange : -stockChange;

  try {
    await runTransaction(db, async (tx) => {
      const itemRef = doc(db, "items", itemId);
      const itemSnap = await tx.get(itemRef);
      if (!itemSnap.exists()) throw new Error("Item not found");
      const itemData = itemSnap.data();
      const currentStock = toNumber(itemData.stockQty);
      const newStock = currentStock - delta;

      if (delta > 0 && newStock < 0) {
        throw new Error("Insufficient stock");
      }

      tx.update(itemRef, {
        stockQty: newStock,
        lastUpdatedBy: getAuditProfile(),
        updatedAt: serverTimestamp()
      });
      const logRef = doc(collection(db, "consumptions"));
      tx.set(logRef, {
        itemId,
        quantity: delta,
        inputMode: slideBased ? "slides" : "units",
        inputCount,
        repeatCount,
        totalInputCount,
        unitsPerSlide: slideBased ? unitsPerSlide : null,
        slideUnitId: slideBased ? slideUnitId : null,
        profileName: getAuditProfile(),
        createdAt: serverTimestamp()
      });
    });
    await appendAuditLog("create", "consumption", itemId, {
      action,
      inputMode: slideBased ? "slides" : "units",
      inputCount,
      repeatCount,
      quantity: delta
    });
    await refreshAll();
    showToast("Consumption updated");
  } catch (error) {
    showToast(error.message || "Consumption failed", true);
  }
});

configItemsList?.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-config-action]");
  if (!button) return;
  const action = button.dataset.configAction;
  const itemId = button.dataset.itemId;
  if (!itemId) return;

  if (action === "close") {
    activeConfigAction = { itemId: null, mode: null };
  } else if (activeConfigAction.itemId === itemId && activeConfigAction.mode === action) {
    activeConfigAction = { itemId: null, mode: null };
  } else {
    activeConfigAction = { itemId, mode: action };
  }

  renderConfigItems();
});

configItemsList?.addEventListener("submit", async (event) => {
  const form = event.target.closest("form[data-form]");
  if (!form) return;
  event.preventDefault();
  if (!ensureReady()) return;

  const formType = form.dataset.form;
  if (!formType) return;

  try {
    if (formType === "bulkReorder") {
      const categoryId = form.dataset.categoryId;
      const reorderLevel = toNumber(form.elements.reorderLevel.value);
      if (!categoryId) {
        showToast("Category missing for bulk update", true);
        return;
      }
      if (reorderLevel < 0) {
        showToast("Re-order level cannot be negative", true);
        return;
      }

      const itemsInCategory = state.items.filter((item) => item.categoryId === categoryId);
      if (!itemsInCategory.length) {
        showToast("No items found in this category", true);
        return;
      }

      await Promise.all(itemsInCategory.map((item) => updateDoc(doc(db, "items", item.id), {
        reorderLevel,
        lastUpdatedBy: getAuditProfile(),
        updatedAt: serverTimestamp()
      })));

      await appendAuditLog("update", "item", categoryId, {
        mode: "bulkReorder",
        reorderLevel,
        itemCount: itemsInCategory.length
      });

      await refreshAll();
      showToast(`Re-order level updated for ${itemsInCategory.length} item(s)`);
      return;
    }

    const itemId = form.dataset.itemId;
    if (!itemId) return;
    const itemRef = doc(db, "items", itemId);

    if (formType === "edit") {
      const name = form.elements.name.value.trim();
      const price = toNumber(form.elements.price.value);
      const reorderLevel = toNumber(form.elements.reorderLevel.value);
      const hasExpiry = Boolean(form.elements.hasExpiry.checked);
      const expiryDateRaw = form.elements.expiryDate.value;

      if (!name) {
        showToast("Item name required", true);
        return;
      }

      await updateDoc(itemRef, {
        name,
        price,
        reorderLevel,
        hasExpiry,
        expiryDate: hasExpiry && expiryDateRaw ? new Date(expiryDateRaw) : null,
        lastUpdatedBy: getAuditProfile(),
        updatedAt: serverTimestamp()
      });
      await appendAuditLog("update", "item", itemId, {
        mode: "edit",
        name,
        price,
        reorderLevel,
        hasExpiry
      });
      showToast("Item updated");
    }

    if (formType === "config") {
      const categoryId = form.elements.categoryId.value;
      const unitId = form.elements.unitId.value;
      const vendorId = form.elements.vendorId.value;

      if (!categoryId || !unitId || !vendorId) {
        showToast("Category, unit, and vendor are required", true);
        return;
      }

      const validationError = validateSlideCategoryUnit(categoryId, unitId);
      if (validationError) {
        showToast(validationError, true);
        return;
      }

      await updateDoc(itemRef, {
        categoryId,
        unitId,
        vendorId,
        lastUpdatedBy: getAuditProfile(),
        updatedAt: serverTimestamp()
      });
      await appendAuditLog("update", "item", itemId, {
        mode: "config",
        categoryId,
        unitId,
        vendorId
      });
      showToast("Item config updated");
    }

    if (formType === "manage") {
      const stockQty = toNumber(form.elements.stockQty.value);
      if (stockQty < 0) {
        showToast("Stock cannot be negative", true);
        return;
      }

      await updateDoc(itemRef, {
        stockQty,
        lastUpdatedBy: getAuditProfile(),
        updatedAt: serverTimestamp()
      });
      await appendAuditLog("update", "item", itemId, {
        mode: "manage",
        stockQty
      });
      showToast("Stock updated");
    }

    activeConfigAction = { itemId: null, mode: null };
    await refreshAll();
  } catch (error) {
    showToast("Failed to update item", true);
  }
});

categoryForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!ensureReady()) return;

  const name = document.getElementById("categoryName").value.trim();
  const description = document.getElementById("categoryDescription").value.trim();
  const slideBasedConsumption = Boolean(categorySlideBased?.checked);
  const slideUnitId = categorySlideUnit?.value || "";
  const unitsPerSlide = toNumber(categorySlideValue?.value);

  if (!name) {
    showToast("Category name required", true);
    return;
  }

  if (slideBasedConsumption) {
    if (!slideUnitId) {
      showToast("Select slide unit for slide-based category", true);
      return;
    }
    if (unitsPerSlide <= 0) {
      showToast("Units per slide must be greater than 0", true);
      return;
    }
  }

  try {
    const ref = await addDoc(collection(db, "categories"), {
      name,
      description,
      slideBasedConsumption,
      slideUnitId: slideBasedConsumption ? slideUnitId : null,
      unitsPerSlide: slideBasedConsumption ? unitsPerSlide : null,
      profileName: getAuditProfile(),
      createdAt: serverTimestamp()
    });
    await appendAuditLog("create", "category", ref.id, {
      name,
      slideBasedConsumption,
      slideUnitId: slideBasedConsumption ? slideUnitId : null,
      unitsPerSlide: slideBasedConsumption ? unitsPerSlide : null
    });
    categoryForm.reset();
    toggleCategorySlideConfig();
    await refreshAll();
    showToast("Category created");
  } catch (error) {
    showToast("Failed to create category", true);
  }
});

unitForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!ensureReady()) return;

  const displayName = document.getElementById("unitDisplay").value.trim();
  const unitName = document.getElementById("unitName").value.trim();

  if (!displayName || !unitName) {
    showToast("Unit fields required", true);
    return;
  }

  try {
    const ref = await addDoc(collection(db, "units"), {
      displayName,
      unitName,
      profileName: getAuditProfile(),
      createdAt: serverTimestamp()
    });
    await appendAuditLog("create", "unit", ref.id, {
      displayName,
      unitName
    });
    unitForm.reset();
    await refreshAll();
    showToast("Unit created");
  } catch (error) {
    showToast("Failed to create unit", true);
  }
});

vendorForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!ensureReady()) return;

  const payload = {
    name: document.getElementById("vendorName").value.trim(),
    address: document.getElementById("vendorAddress").value.trim(),
    mobile: document.getElementById("vendorMobile").value.trim(),
    email: document.getElementById("vendorEmail").value.trim(),
    openingBalance: toNumber(document.getElementById("vendorOpening").value),
    profileName: getAuditProfile(),
    createdAt: serverTimestamp()
  };

  if (!payload.name) {
    showToast("Vendor name required", true);
    return;
  }

  try {
    const ref = await addDoc(collection(db, "vendors"), payload);
    await appendAuditLog("create", "vendor", ref.id, {
      name: payload.name,
      email: payload.email,
      mobile: payload.mobile
    });
    vendorForm.reset();
    await refreshAll();
    showToast("Vendor created");
  } catch (error) {
    showToast("Failed to create vendor", true);
  }
});

itemForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!ensureReady()) return;

  const categoryId = itemCategory.value;
  const unitId = itemUnit.value;
  const vendorId = itemVendor.value;

  if (!categoryId || !unitId || !vendorId) {
    showToast("Select category, unit, and vendor", true);
    return;
  }

  const validationError = validateSlideCategoryUnit(categoryId, unitId);
  if (validationError) {
    showToast(validationError, true);
    return;
  }

  const payload = {
    categoryId,
    unitId,
    vendorId,
    name: document.getElementById("itemName").value.trim(),
    hasExpiry: itemHasExpiry.checked,
    expiryDate: itemHasExpiry.checked && document.getElementById("itemExpiryDate").value
      ? new Date(document.getElementById("itemExpiryDate").value)
      : null,
    openingQty: toNumber(document.getElementById("itemOpeningQty").value),
    price: toNumber(document.getElementById("itemPrice").value),
    reorderLevel: toNumber(document.getElementById("itemReorder").value),
    stockQty: toNumber(document.getElementById("itemOpeningQty").value),
    profileName: getAuditProfile(),
    createdAt: serverTimestamp()
  };

  if (!payload.name) {
    showToast("Item name required", true);
    return;
  }

  try {
    const ref = await addDoc(collection(db, "items"), payload);
    await appendAuditLog("create", "item", ref.id, {
      name: payload.name,
      categoryId,
      unitId,
      vendorId,
      openingQty: payload.openingQty,
      price: payload.price
    });
    itemForm.reset();
    itemHasExpiry.checked = false;
    expiryRow.style.display = "none";
    await refreshAll();
    showToast("Item created");
  } catch (error) {
    showToast("Failed to create item", true);
  }
});

purchaseForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!ensureReady()) return;

  const itemId = purchaseItem.value;
  const quantity = toNumber(purchaseQty.value);
  const totalAmount = toNumber(purchaseTotal.value);

  if (!itemId || quantity <= 0) {
    showToast("Select item and quantity", true);
    return;
  }

  try {
    await runTransaction(db, async (tx) => {
      const itemRef = doc(db, "items", itemId);
      const itemSnap = await tx.get(itemRef);
      if (!itemSnap.exists()) throw new Error("Item not found");
      const itemData = itemSnap.data();
      const currentStock = toNumber(itemData.stockQty);
      const newStock = currentStock + quantity;

      tx.update(itemRef, {
        stockQty: newStock,
        lastUpdatedBy: getAuditProfile(),
        updatedAt: serverTimestamp()
      });
      const purchaseRef = doc(collection(db, "purchases"));
      tx.set(purchaseRef, {
        itemId,
        vendorId: itemData.vendorId || null,
        quantity,
        totalAmount,
        profileName: getAuditProfile(),
        createdAt: serverTimestamp()
      });
    });

    await appendAuditLog("create", "purchase", itemId, {
      quantity,
      totalAmount
    });

    purchaseForm.reset();
    purchaseTotalManual = false;
    await refreshAll();
    showToast("Purchase saved");
  } catch (error) {
    showToast("Purchase failed", true);
  }
});

returnForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!ensureReady()) return;

  const itemId = returnItem.value;
  const quantity = toNumber(document.getElementById("returnQty").value);
  const reason = document.getElementById("returnReason").value.trim();

  if (!itemId || quantity <= 0 || !reason) {
    showToast("Return details required", true);
    return;
  }

  try {
    await runTransaction(db, async (tx) => {
      const itemRef = doc(db, "items", itemId);
      const itemSnap = await tx.get(itemRef);
      if (!itemSnap.exists()) throw new Error("Item not found");
      const itemData = itemSnap.data();
      const currentStock = toNumber(itemData.stockQty);
      const newStock = currentStock - quantity;

      if (newStock < 0) {
        throw new Error("Insufficient stock for return");
      }

      tx.update(itemRef, {
        stockQty: newStock,
        lastUpdatedBy: getAuditProfile(),
        updatedAt: serverTimestamp()
      });
      const returnRef = doc(collection(db, "returns"));
      tx.set(returnRef, {
        itemId,
        quantity,
        reason,
        profileName: getAuditProfile(),
        createdAt: serverTimestamp()
      });
    });

    await appendAuditLog("create", "return", itemId, {
      quantity,
      reason
    });

    returnForm.reset();
    await refreshAll();
    showToast("Return logged");
  } catch (error) {
    showToast(error.message || "Return failed", true);
  }
});

adjustmentForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!ensureReady()) return;

  const itemId = adjustItem.value;
  const direction = document.getElementById("adjustType").value;
  const quantity = toNumber(document.getElementById("adjustQty").value);
  const reason = document.getElementById("adjustReason").value.trim();

  if (!itemId || quantity <= 0 || !reason) {
    showToast("Adjustment details required", true);
    return;
  }

  const delta = direction === "increase" ? quantity : -quantity;

  try {
    await runTransaction(db, async (tx) => {
      const itemRef = doc(db, "items", itemId);
      const itemSnap = await tx.get(itemRef);
      if (!itemSnap.exists()) throw new Error("Item not found");
      const itemData = itemSnap.data();
      const currentStock = toNumber(itemData.stockQty);
      const newStock = currentStock + delta;

      if (newStock < 0) {
        throw new Error("Insufficient stock for adjustment");
      }

      tx.update(itemRef, {
        stockQty: newStock,
        lastUpdatedBy: getAuditProfile(),
        updatedAt: serverTimestamp()
      });
      const adjustRef = doc(collection(db, "adjustments"));
      tx.set(adjustRef, {
        itemId,
        quantity,
        direction,
        reason,
        profileName: getAuditProfile(),
        createdAt: serverTimestamp()
      });
    });

    await appendAuditLog("create", "adjustment", itemId, {
      quantity,
      direction,
      reason
    });

    adjustmentForm.reset();
    await refreshAll();
    showToast("Adjustment saved");
  } catch (error) {
    showToast(error.message || "Adjustment failed", true);
  }
});

excelImportForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!ensureReady()) return;

  const file = excelFileInput?.files?.[0];
  if (!file) {
    showToast("Select an Excel file", true);
    return;
  }

  try {
    const rows = await parseExcelFile(file);
    if (!rows.length) {
      showToast("No rows found in file", true);
      return;
    }

    const parsedRows = [];
    let invalidRows = 0;

    rows.forEach((row) => {
      const itemName = String(row["ITEM NAME"] || "").trim();
      const packedQtyRaw = row["PACKED QTY"];
      const packedQty = String(packedQtyRaw || "").trim();
      const stockAvailable = toSafeStock(row["STOCK AVAILABLE"]);
      const rate = toNumber(row["RATE"]);

      if (!itemName || !packedQty) {
        invalidRows += 1;
        return;
      }

      parsedRows.push({
        name: `${itemName} - ${packedQty}`,
        packedQty,
        packedQtyNumber: toNumber(packedQtyRaw),
        stockAvailable,
        rate: rate > 0 ? rate : 0
      });
    });

    if (!parsedRows.length) {
      showToast("No valid rows found. Check column names", true);
      return;
    }

    const setup = await ensureExcelImportSetup();
    const existingNames = new Set(state.items.map((item) => String(item.name || "").trim().toLowerCase()));

    let createdCount = 0;
    let skippedCount = 0;

    for (const row of parsedRows) {
      const key = row.name.toLowerCase();
      if (existingNames.has(key)) {
        skippedCount += 1;
        continue;
      }

      await addDoc(collection(db, "items"), {
        categoryId: setup.categoryId,
        unitId: setup.unitId,
        vendorId: setup.vendorId,
        name: row.name,
        hasExpiry: false,
        expiryDate: null,
        openingQty: row.stockAvailable,
        stockQty: row.stockAvailable,
        price: row.rate,
        reorderLevel: 0,
        packedQty: row.packedQty,
        importedFromExcel: true,
        profileName: getAuditProfile(),
        createdAt: serverTimestamp()
      });

      existingNames.add(key);
      createdCount += 1;
    }

    await appendAuditLog("create", "excel-import", "ihc-marker", {
      createdCount,
      skippedCount,
      invalidRows
    });

    excelImportForm.reset();
    await refreshAll();
    showToast(`Import done: ${createdCount} created, ${skippedCount} skipped, ${invalidRows} invalid`);
  } catch (error) {
    showToast(error.message || "Excel import failed", true);
  }
});

initProfiles();
initSearchableSelects();
initStatsFilters();
initLogs();
initAuditLogs();
refreshAll();



