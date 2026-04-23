const state = {
  records: [],
  filtered: [],
  contentCache: new Map(),
  contentTextIndex: new Map(),
  selectedKey: null,
  page: 1,
  pageSize: 50,
  indexedCount: 0,
};

const dom = {
  syncSummary: document.getElementById("syncSummary"),
  searchInput: document.getElementById("searchInput"),
  driveFilter: document.getElementById("driveFilter"),
  typeFilter: document.getElementById("typeFilter"),
  sortBy: document.getElementById("sortBy"),
  indexingState: document.getElementById("indexingState"),
  listStatus: document.getElementById("listStatus"),
  resultList: document.getElementById("resultList"),
  prevPageBtn: document.getElementById("prevPageBtn"),
  nextPageBtn: document.getElementById("nextPageBtn"),
  pageInfo: document.getElementById("pageInfo"),
  viewerEmpty: document.getElementById("viewerEmpty"),
  viewer: document.getElementById("viewer"),
  docTitle: document.getElementById("docTitle"),
  docPath: document.getElementById("docPath"),
  docDrive: document.getElementById("docDrive"),
  docType: document.getElementById("docType"),
  docModified: document.getElementById("docModified"),
  docSourceLink: document.getElementById("docSourceLink"),
  docContent: document.getElementById("docContent"),
  copyLinkBtn: document.getElementById("copyLinkBtn"),
  resultItemTemplate: document.getElementById("resultItemTemplate"),
};

const sorters = {
  name_asc: (a, b) => a.name.localeCompare(b.name),
  name_desc: (a, b) => b.name.localeCompare(a.name),
  modified_asc: (a, b) => new Date(a.last_modified || 0) - new Date(b.last_modified || 0),
  modified_desc: (a, b) => new Date(b.last_modified || 0) - new Date(a.last_modified || 0),
};

init().catch((error) => {
  dom.listStatus.textContent = "Could not load document index.";
  dom.syncSummary.textContent = "Failed to load sync metadata.";
  console.error(error);
});

async function init() {
  wireEvents();

  const index = await fetchJson("/sharepoint_sync/index.json");
  state.records = Array.isArray(index)
    ? index.map((r) => ({ ...r, _id: r.item_id || r.full_path || r.output_file }))
    : [];

  fillFilterOptions();
  renderSyncSummary();
  applyFilters();
  startBackgroundIndexing();

  const preselected = readSelectedFromUrl();
  if (preselected) selectRecordById(preselected);
}

function wireEvents() {
  [dom.searchInput, dom.driveFilter, dom.typeFilter, dom.sortBy].forEach((el) => {
    el.addEventListener("input", () => {
      state.page = 1;
      applyFilters();
    });
    el.addEventListener("change", () => {
      state.page = 1;
      applyFilters();
    });
  });

  dom.prevPageBtn.addEventListener("click", () => {
    state.page -= 1;
    renderList();
  });

  dom.nextPageBtn.addEventListener("click", () => {
    state.page += 1;
    renderList();
  });

  dom.copyLinkBtn.addEventListener("click", async () => {
    if (!state.selectedKey) return;
    const url = new URL(window.location.href);
    url.pathname = `/doc/${encodeURIComponent(state.selectedKey)}`;
    await navigator.clipboard.writeText(url.toString());
    dom.copyLinkBtn.textContent = "Copied";
    setTimeout(() => {
      dom.copyLinkBtn.textContent = "Copy link";
    }, 1200);
  });

  window.addEventListener("popstate", () => {
    const key = readSelectedFromUrl();
    selectRecordById(key, false);
  });
}

function fillFilterOptions() {
  for (const drive of unique(state.records.map((x) => x.drive_name).filter(Boolean))) {
    dom.driveFilter.add(new Option(drive, drive));
  }
  for (const ext of unique(state.records.map((x) => x.extension || "(none)"))) {
    dom.typeFilter.add(new Option(ext, ext));
  }
}

function renderSyncSummary() {
  if (!state.records.length) {
    dom.syncSummary.textContent = "No synced documents found.";
    return;
  }
  const newest = new Date(
    Math.max(...state.records.map((r) => new Date(r.last_modified || 0).getTime()))
  );
  dom.syncSummary.textContent = `${state.records.length} documents • latest update ${newest.toLocaleString()}`;
}

function applyFilters() {
  const q = dom.searchInput.value.trim().toLowerCase();
  const drive = dom.driveFilter.value;
  const ext = dom.typeFilter.value;
  const sorter = sorters[dom.sortBy.value] || sorters.modified_desc;

  state.filtered = state.records
    .filter((r) => !drive || r.drive_name === drive)
    .filter((r) => !ext || (r.extension || "(none)") === ext)
    .filter((r) => {
      if (!q) return true;
      const metaHit = [r.name, r.path, r.full_path].filter(Boolean).join(" ").toLowerCase().includes(q);
      if (metaHit) return true;
      const cachedText = state.contentTextIndex.get(r._id);
      return cachedText ? cachedText.includes(q) : false;
    })
    .sort(sorter);

  renderList(q);
}

function renderList(query = "") {
  const totalPages = Math.max(1, Math.ceil(state.filtered.length / state.pageSize));
  state.page = Math.min(state.page, totalPages);

  dom.resultList.innerHTML = "";

  if (!state.filtered.length) {
    dom.listStatus.textContent = "No results match your current filters.";
    dom.pageInfo.textContent = "Page 1 of 1";
    dom.prevPageBtn.disabled = true;
    dom.nextPageBtn.disabled = true;
    return;
  }

  const start = (state.page - 1) * state.pageSize;
  const pageRows = state.filtered.slice(start, start + state.pageSize);

  dom.listStatus.textContent = `${state.filtered.length} matching documents`;

  for (const record of pageRows) {
    const node = dom.resultItemTemplate.content.firstElementChild.cloneNode(true);
    const button = node.querySelector(".result-button");
    const titleEl = node.querySelector(".result-title");
    const pathEl = node.querySelector(".result-path");
    const metaEl = node.querySelector(".result-meta");
    const snippetEl = node.querySelector(".result-snippet");

    titleEl.innerHTML = highlight(record.name || "Untitled", query);
    pathEl.innerHTML = highlight(record.full_path || record.path || "(no path)", query);
    metaEl.textContent = `${record.drive_name || "Unknown drive"} • ${record.extension || "no extension"} • ${formatDate(record.last_modified)}`;

    const snippet = buildSnippet(record, query);
    snippetEl.innerHTML = snippet ? highlight(snippet, query) : "";

    if (record._id === state.selectedKey) node.classList.add("active");

    button.addEventListener("click", () => selectRecord(record));
    dom.resultList.appendChild(node);
  }

  dom.pageInfo.textContent = `Page ${state.page} of ${totalPages}`;
  dom.prevPageBtn.disabled = state.page <= 1;
  dom.nextPageBtn.disabled = state.page >= totalPages;
}

async function selectRecord(record, pushState = true) {
  state.selectedKey = record?._id || null;
  renderList(dom.searchInput.value.trim().toLowerCase());

  if (!record) {
    dom.viewer.classList.add("hidden");
    dom.viewerEmpty.classList.remove("hidden");
    return;
  }

  dom.viewerEmpty.classList.add("hidden");
  dom.viewer.classList.remove("hidden");

  dom.docTitle.textContent = record.name || "Untitled";
  dom.docPath.textContent = record.full_path || record.path || "—";
  dom.docDrive.textContent = record.drive_name || "—";
  dom.docType.textContent = record.mime_type || record.extension || "—";
  dom.docModified.textContent = formatDate(record.last_modified);
  dom.docSourceLink.href = record.web_url || "#";

  dom.docContent.classList.remove("unsupported");
  dom.docContent.textContent = "Loading content…";

  const contentDoc = await fetchContent(record);
  renderContent(contentDoc);

  if (pushState) {
    const url = `/doc/${encodeURIComponent(record._id)}`;
    history.pushState({}, "", url);
  }
}

function selectRecordById(id, pushState = false) {
  if (!id) {
    selectRecord(null, pushState);
    return;
  }
  const record = state.records.find((r) => r._id === id);
  if (record) selectRecord(record, pushState);
}

function renderContent(contentDoc) {
  if (!contentDoc) {
    dom.docContent.classList.add("unsupported");
    dom.docContent.textContent = "Content could not be loaded for this document.";
    return;
  }

  const type = contentDoc.content_type || "unknown";
  const content = contentDoc.content;

  if (type === "unsupported" || content == null || content === "") {
    dom.docContent.classList.add("unsupported");
    dom.docContent.textContent = "This file type is unsupported or has no extracted content available.";
    return;
  }

  if (type === "html") {
    dom.docContent.innerHTML = sanitizeHtml(content);
    return;
  }

  dom.docContent.textContent = typeof content === "string" ? content : JSON.stringify(content, null, 2);
}

async function fetchContent(record) {
  if (!record?.output_file) return null;

  const cacheKey = record._id;
  if (state.contentCache.has(cacheKey)) return state.contentCache.get(cacheKey);

  const path = normalizeOutputPath(record.output_file);
  try {
    const content = await fetchJson(path);
    state.contentCache.set(cacheKey, content);
    return content;
  } catch {
    return null;
  }
}

function startBackgroundIndexing() {
  const queue = [...state.records];
  const workers = Array.from({ length: 5 }, () => indexWorker(queue));

  Promise.all(workers).then(() => {
    dom.indexingState.textContent = "Content index ready";
    applyFilters();
  });
}

async function indexWorker(queue) {
  while (queue.length) {
    const rec = queue.shift();
    if (!rec) return;
    const doc = await fetchContent(rec);
    const text = extractSearchableText(doc);
    if (text) state.contentTextIndex.set(rec._id, text.toLowerCase());
    state.indexedCount += 1;
    dom.indexingState.textContent = `Indexed ${state.indexedCount}/${state.records.length} docs`;
  }
}

function buildSnippet(record, query) {
  if (!query) return "";
  const text = state.contentTextIndex.get(record._id);
  if (!text) return "";
  const idx = text.indexOf(query.toLowerCase());
  if (idx < 0) return "";
  const start = Math.max(0, idx - 50);
  const end = Math.min(text.length, idx + query.length + 70);
  return text.slice(start, end).replace(/\s+/g, " ");
}

function readSelectedFromUrl() {
  const match = window.location.pathname.match(/^\/doc\/(.+)$/);
  return match ? decodeURIComponent(match[1]) : null;
}

function extractSearchableText(doc) {
  if (!doc || doc.content == null || doc.content_type === "unsupported") return "";
  if (doc.content_type === "html") {
    const el = document.createElement("div");
    el.innerHTML = doc.content;
    return (el.textContent || "").trim();
  }
  if (typeof doc.content === "string") return doc.content;
  return JSON.stringify(doc.content);
}

function normalizeOutputPath(path) {
  return `/${path.replace(/^\/+/, "")}`;
}

function highlight(text, rawQuery) {
  if (!rawQuery) return escapeHtml(text);
  const query = rawQuery.trim();
  if (!query) return escapeHtml(text);

  const escaped = escapeRegExp(query);
  const regex = new RegExp(`(${escaped})`, "ig");
  return escapeHtml(text).replace(regex, "<mark>$1</mark>");
}

function sanitizeHtml(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const blockedTags = new Set(["script", "style", "iframe", "object", "embed", "link", "meta"]);

  for (const el of [...doc.body.querySelectorAll("*")]) {
    const name = el.tagName.toLowerCase();
    if (blockedTags.has(name)) {
      el.remove();
      continue;
    }
    for (const attr of [...el.attributes]) {
      const attrName = attr.name.toLowerCase();
      const value = attr.value.trim().toLowerCase();
      if (attrName.startsWith("on") || attrName === "srcdoc") {
        el.removeAttribute(attr.name);
      }
      if ((attrName === "href" || attrName === "src") && value.startsWith("javascript:")) {
        el.removeAttribute(attr.name);
      }
    }
  }
  return doc.body.innerHTML;
}

function formatDate(value) {
  if (!value) return "Unknown";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "Unknown";
  return d.toLocaleString();
}

async function fetchJson(path) {
  const res = await fetch(path, { cache: "no-store" });
  if (!res.ok) throw new Error(`Failed to fetch ${path}: ${res.status}`);
  return res.json();
}

function unique(items) {
  return [...new Set(items)].sort((a, b) => String(a).localeCompare(String(b)));
}

function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
