"use strict";

/* ---------- Einstellungen ---------- */
const CLIENT_ID = "02882712-3952-4323-a968-d11d9b045236";
// Muss exakt der in Entra eingetragenen Umleitungs-URI entsprechen (mit Schrägstrich am Ende).
const REDIRECT_URI = "https://andreasoelenberg-svg.github.io/video-kacheln/";
// Startordner ab dem OneDrive-Stamm. Von hier aus kommst du in alle Unterordner, aber nicht darüber hinaus.
const ROOT_PATH = "KI pic/Self KI";
const SCOPES = ["Files.Read"];
const GRAPH = "https://graph.microsoft.com/v1.0";

// Bildformate, die gängige Browser direkt anzeigen können (HEIC u. Ä. bewusst nicht dabei).
const IMAGE_RE = /\.(jpe?g|png|gif|webp|bmp|avif)$/i;
const VIDEO_RE = /\.mp4$/i;
const isImage = (it) => !!(it && it.file && IMAGE_RE.test(it.name || ""));

const FOLDER_SVG =
  '<svg viewBox="0 0 64 64" width="64" height="64" fill="none" stroke="currentColor" stroke-width="3" ' +
  'stroke-linejoin="round" aria-hidden="true"><path d="M6 18a4 4 0 0 1 4-4h14l6 7h24a4 4 0 0 1 4 4v23a4 4 0 0 1-4 4H10a4 4 0 0 1-4-4z"/></svg>';

/* ---------- Elemente ---------- */
const $ = (id) => document.getElementById(id);
const els = {
  grid: $("grid"), status: $("status"), count: $("count"), hint: $("hint"),
  sort: $("sort"), reload: $("reload"), logout: $("logout"),
  pathbar: $("pathbar"), back: $("back"), crumb: $("crumb"),
  player: $("player"), pvideo: $("pvideo"), pimg: $("pimg"), ptitle: $("ptitle"), pmsg: $("pmsg"), pclose: $("pclose"), pprev: $("pprev"), pnext: $("pnext"),
};

/* ---------- Zustand ---------- */
let pca = null;
let items = [];
let currentPath = ROOT_PATH;
let sortMode = "name";
let selectedId = null;
let playToken = 0;
let loadToken = 0;

const collator = new Intl.Collator("de", { numeric: true, sensitivity: "base" });

/* ---------- Anmeldung ---------- */
async function start() {
  if (typeof msal === "undefined") {
    showError("Die Anmelde-Bibliothek konnte nicht geladen werden.");
    return;
  }
  pca = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/consumers",
      redirectUri: REDIRECT_URI,
      postLogoutRedirectUri: REDIRECT_URI,
      navigateToLoginRequestUrl: false,
    },
    // Die Anmelde-Daten liegen nur im sessionStorage dieses Tabs und verschwinden, wenn der Tab geschlossen wird.
    // (memoryStorage geht mit der Weiterleitungs-Anmeldung nicht, MSAL meldet dann in_mem_redirect_unavailable.)
    cache: { cacheLocation: "sessionStorage" },
  });

  try {
    await pca.initialize();
    const result = await pca.handleRedirectPromise();
    if (result && result.account) pca.setActiveAccount(result.account);
  } catch (e) {
    showError("Anmeldung fehlgeschlagen: " + (e && e.message ? e.message : e), true);
    return;
  }

  const account = pca.getActiveAccount() || pca.getAllAccounts()[0];
  if (!account) {
    showLogin();
    return;
  }
  pca.setActiveAccount(account);
  try { history.replaceState({ path: ROOT_PATH }, ""); } catch (_) { /* nicht kritisch */ }
  await loadFolder(ROOT_PATH, false);
}

async function login() {
  try {
    await pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
  } catch (e) {
    // Fehler sichtbar machen, statt sie still zu verschlucken.
    const code = e && e.errorCode ? e.errorCode + ": " : "";
    showError("Die Anmeldung konnte nicht gestartet werden. " + code + ((e && e.message) || e), true);
  }
}

function showLogin() {
  setStatus();
  const p = document.createElement("p");
  p.textContent = "Melde dich mit deinem Microsoft-Konto an, um deine Videos zu sehen.";
  const b = document.createElement("button");
  b.className = "btn primary";
  b.type = "button";
  b.textContent = "Mit Microsoft anmelden";
  b.addEventListener("click", login);
  els.status.append(p, b);
}

async function getToken() {
  const account = pca.getActiveAccount();
  try {
    const r = await pca.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch (e) {
    await pca.acquireTokenRedirect({ scopes: SCOPES, account });
    return null; // Seite wird zur Anmeldung weitergeleitet
  }
}

async function graph(pathOrUrl) {
  const token = await getToken();
  if (!token) return null;
  const url = pathOrUrl.startsWith("http") ? pathOrUrl : GRAPH + pathOrUrl;
  const res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  if (!res.ok) {
    const err = new Error("Microsoft Graph antwortet mit Status " + res.status);
    err.status = res.status;
    throw err;
  }
  return res.json();
}

/* ---------- Ordner laden ---------- */
function insideRoot(path) {
  return path === ROOT_PATH || path.startsWith(ROOT_PATH + "/");
}

function folderUrl(path) {
  const encoded = path.split("/").map(encodeURIComponent).join("/");
  return "/me/drive/root:/" + encoded + ":/children?$expand=thumbnails&$top=200";
}

function parentOf(path) {
  const i = path.lastIndexOf("/");
  return i < 0 ? path : path.slice(0, i);
}

function updatePathbar() {
  const rootName = ROOT_PATH.split("/").pop();
  const rest = currentPath === ROOT_PATH ? "" : currentPath.slice(ROOT_PATH.length + 1);
  els.crumb.textContent = [rootName].concat(rest ? rest.split("/") : []).join(" / ");
  els.back.hidden = currentPath === ROOT_PATH;
  els.pathbar.hidden = false;
}

async function loadFolder(path, push) {
  if (!insideRoot(path)) return;
  const token = ++loadToken;
  currentPath = path;
  if (push) {
    try { history.pushState({ path }, ""); } catch (_) { /* nicht kritisch */ }
  }
  updatePathbar();
  els.hint.hidden = true;
  setStatus("Lade …");
  els.grid.replaceChildren();
  els.count.textContent = "";
  selectedId = null;
  try {
    const all = [];
    let next = folderUrl(path);
    while (next) {
      const page = await graph(next);
      if (!page) return; // Weiterleitung zur Anmeldung läuft
      all.push(...(page.value || []));
      next = page["@odata.nextLink"] || null;
    }
    if (token !== loadToken) return; // inzwischen wurde ein anderer Ordner geöffnet
    items = all.filter((it) => it.folder || (it.file && (VIDEO_RE.test(it.name || "") || IMAGE_RE.test(it.name || ""))));
    els.sort.hidden = false;
    els.reload.hidden = false;
    els.logout.hidden = false;
    setStatus();
    if (items.length === 0) {
      showMessage("Hier gibt es keine Unterordner, MP4-Videos oder Bilder.");
      return;
    }
    els.hint.hidden = false;
    render();
  } catch (e) {
    if (token !== loadToken) return;
    els.reload.hidden = false;
    els.logout.hidden = false;
    if (e.status === 404) {
      showError("Ordner nicht gefunden: " + path, true);
    } else if (e.status === 401 || e.status === 403) {
      showError("Kein Zugriff auf den Ordner (Status " + e.status + "). Melde dich bitte neu an.", true);
    } else {
      showError("Der Ordner konnte nicht geladen werden: " + (e.message || e), true);
    }
  }
}

/* ---------- Anzeige ---------- */
function setStatus(text) {
  els.status.replaceChildren();
  if (text) {
    const p = document.createElement("p");
    p.textContent = text;
    els.status.append(p);
  }
}

function showMessage(text) {
  const p = document.createElement("p");
  p.textContent = text;
  els.status.append(p);
}

function showError(text, withRetry) {
  setStatus();
  const p = document.createElement("p");
  p.className = "err";
  p.textContent = text;
  els.status.append(p);
  if (withRetry && pca) {
    const b = document.createElement("button");
    b.className = "btn primary";
    b.type = "button";
    b.textContent = "Neu anmelden";
    b.addEventListener("click", login);
    els.status.append(b);
  }
}

function sizeText(bytes) {
  if (!bytes && bytes !== 0) return "";
  const mb = bytes / (1024 * 1024);
  return mb >= 100 ? Math.round(mb) + " MB" : mb.toFixed(1).replace(".", ",") + " MB";
}

function dateText(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  return isNaN(d) ? "" : d.toLocaleDateString("de-DE");
}

function sorted() {
  // Ordner stehen immer oben und sind nach Namen sortiert, danach die Videos.
  const folders = items.filter((it) => it.folder).sort((a, b) => collator.compare(a.name, b.name));
  const videos = items.filter((it) => !it.folder); // Videos und Bilder
  if (sortMode === "date") {
    videos.sort((a, b) => String(b.lastModifiedDateTime).localeCompare(String(a.lastModifiedDateTime)));
  } else if (sortMode === "size") {
    // Größte Datei zuerst; bei gleicher Größe nach Namen.
    videos.sort((a, b) => (b.size || 0) - (a.size || 0) || collator.compare(a.name, b.name));
  } else {
    videos.sort((a, b) => collator.compare(a.name, b.name));
  }
  return folders.concat(videos);
}

function thumbUrl(item) {
  const t = item.thumbnails && item.thumbnails[0];
  if (!t) return null;
  // Größte Vorschau zuerst (large = bis ca. 800 px), kleinere nur als Rückfall.
  return (t.large && t.large.url) || (t.medium && t.medium.url) || (t.small && t.small.url) || null;
}

function countText(list) {
  const nf = list.filter((it) => it.folder).length;
  const ni = list.filter(isImage).length;
  const nv = list.length - nf - ni;
  const parts = [];
  if (nf) parts.push(nf + " Ordner");
  if (nv) parts.push(nv + (nv === 1 ? " Video" : " Videos"));
  if (ni) parts.push(ni + (ni === 1 ? " Bild" : " Bilder"));
  return parts.join(" · ");
}

function render() {
  const list = sorted();
  els.count.textContent = countText(list);
  const frag = document.createDocumentFragment();
  for (const item of list) {
    const isFolder = !!item.folder;
    const isImg = isImage(item);
    const tile = document.createElement("div");
    tile.className = "tile" + (isFolder ? " folder" : "");
    tile.setAttribute("role", "option");
    tile.tabIndex = 0;
    tile.dataset.id = item.id;
    tile.setAttribute("aria-selected", item.id === selectedId ? "true" : "false");
    tile.setAttribute("aria-label", (isFolder ? "Ordner " : isImg ? "Bild " : "Video ") + item.name);

    const ph = document.createElement("div");
    ph.className = "ph";
    ph.setAttribute("aria-hidden", "true");
    if (isFolder) {
      ph.innerHTML = FOLDER_SVG; // feste, eigene Grafik ohne Fremddaten
    } else {
      ph.textContent = isImg ? "🖼" : "▶";
    }
    tile.append(ph);

    if (!isFolder) {
      const url = thumbUrl(item);
      if (url) {
        const img = document.createElement("img");
        img.alt = "";
        img.loading = "lazy";
        img.decoding = "async";
        img.referrerPolicy = "no-referrer";
        img.src = url;
        img.addEventListener("error", () => img.remove());
        tile.append(img);
      }
    }

    const cap = document.createElement("div");
    cap.className = "cap";
    const name = document.createElement("div");
    name.className = "name";
    name.textContent = isFolder ? item.name : item.name.replace(VIDEO_RE, "").replace(IMAGE_RE, "");
    const meta = document.createElement("div");
    meta.className = "meta";
    if (isFolder) {
      const n = item.folder.childCount;
      meta.textContent = typeof n === "number" ? n + (n === 1 ? " Element" : " Elemente") : "Ordner";
    } else {
      meta.textContent = [sizeText(item.size), dateText(item.lastModifiedDateTime)].filter(Boolean).join(" · ");
    }
    cap.append(name, meta);

    const pb = document.createElement("button");
    pb.type = "button";
    pb.className = "playbtn";
    pb.setAttribute("aria-label", (isFolder ? "Öffnen: " : isImg ? "Anzeigen: " : "Abspielen: ") + item.name);
    pb.textContent = isFolder ? "›" : isImg ? "⤢" : "▶";

    tile.append(cap, pb);
    frag.append(tile);
  }
  els.grid.replaceChildren(frag);
}

function select(id) {
  selectedId = id;
  for (const t of els.grid.children) {
    t.setAttribute("aria-selected", t.dataset.id === id ? "true" : "false");
  }
}

/* Ordner öffnen, Video abspielen oder Bild anzeigen */
function activate(id) {
  const item = items.find((x) => x.id === id);
  if (!item) return;
  if (item.folder) {
    loadFolder(currentPath + "/" + item.name, true);
  } else {
    play(id);
  }
}

/* ---------- Bedienung der Kacheln ---------- */
els.grid.addEventListener("pointerup", (e) => {
  if (e.pointerType === "mouse" && e.button !== 0) return;
  const tile = e.target.closest(".tile");
  if (!tile) return;
  const id = tile.dataset.id;
  if (e.target.closest(".playbtn")) {
    select(id);
    activate(id);
    return;
  }
  // Ein Tipp genügt: Ordner öffnen, Video abspielen oder Bild anzeigen.
  select(id);
  activate(id);
});

els.grid.addEventListener("keydown", (e) => {
  const tile = e.target.closest(".tile");
  if (!tile) return;
  if (e.key === "Enter") {
    e.preventDefault();
    select(tile.dataset.id);
    activate(tile.dataset.id);
  } else if (e.key === " ") {
    e.preventDefault();
    select(tile.dataset.id);
  }
});

els.sort.addEventListener("click", (e) => {
  const b = e.target.closest("button[data-sort]");
  if (!b) return;
  sortMode = b.dataset.sort;
  for (const x of els.sort.querySelectorAll("button")) {
    x.setAttribute("aria-pressed", x === b ? "true" : "false");
  }
  render();
});

els.back.addEventListener("click", () => {
  if (currentPath !== ROOT_PATH) loadFolder(parentOf(currentPath), true);
});

// Zurück-Taste des Browsers oder Handys führt eine Ordnerebene nach oben (oder zum zuletzt besuchten Ordner).
window.addEventListener("popstate", (e) => {
  if (!pca || !pca.getActiveAccount()) return;
  const p = e.state && e.state.path;
  if (typeof p === "string" && insideRoot(p)) {
    if (!els.player.hidden) closePlayer();
    loadFolder(p, false);
  }
});

els.reload.addEventListener("click", () => loadFolder(currentPath, false));
els.logout.addEventListener("click", () => {
  try { sessionStorage.removeItem(GATE_KEY); } catch (_) { /* nicht kritisch */ }
  pca.logoutRedirect({ account: pca.getActiveAccount(), postLogoutRedirectUri: REDIRECT_URI });
});

/* ---------- Abspielen ---------- */
let fsRequested = false;
let viewList = [];
let viewIndex = -1;

// Blättern per Pfeil-Knöpfen unten im Player (und Pfeiltasten am PC).
function updateNav() {
  els.pprev.disabled = viewIndex <= 0;
  els.pnext.disabled = viewIndex < 0 || viewIndex >= viewList.length - 1;
}

function step(dir) {
  const next = viewIndex + dir;
  if (viewIndex < 0 || next < 0 || next >= viewList.length) return; // Anfang bzw. Ende der Liste
  play(viewList[next].id);
}

// Muss direkt im Tipp/Klick aufgerufen werden (vor dem ersten await), sonst verweigert der Browser den Vollbildmodus.
function enterFullscreen() {
  const el = els.player;
  const req = el.requestFullscreen || el.webkitRequestFullscreen;
  if (!req) return false;
  try {
    const p = req.call(el);
    fsRequested = true;
    if (p && p.catch) p.catch(() => { fsRequested = false; });
  } catch (_) {
    fsRequested = false;
    return false;
  }
  return true;
}

function exitFullscreen() {
  const fsEl = document.fullscreenElement || document.webkitFullscreenElement;
  if (!fsEl) return;
  const ex = document.exitFullscreen || document.webkitExitFullscreen;
  try { const p = ex.call(document); if (p && p.catch) p.catch(() => {}); } catch (_) { /* nicht kritisch */ }
}

function onFullscreenChange() {
  const active = !!(document.fullscreenElement || document.webkitFullscreenElement);
  // Wer den Vollbildmodus verlässt, kommt zurück zu den Kacheln.
  if (!active && fsRequested && !els.player.hidden) closePlayer();
  if (!active) fsRequested = false;
}
document.addEventListener("fullscreenchange", onFullscreenChange);
document.addEventListener("webkitfullscreenchange", onFullscreenChange);

async function play(id) {
  const item = items.find((x) => x.id === id);
  if (!item || item.folder) return;
  const token = ++playToken;
  blobTried = false;
  probeInfo = "";
  lastDownloadUrl = "";
  revokeBlob();
  // Blätterliste: alle Videos und Bilder des Ordners in der aktuell gewählten Sortierung.
  viewList = sorted().filter((it) => !it.folder);
  viewIndex = viewList.findIndex((it) => it.id === id);
  const isImg = isImage(item);
  // Videos starten im Vollbild; Bilder werden im Anzeigefenster gezeigt.
  // Videos starten im Vollbild, wo der Browser das für die ganze Seite erlaubt. Auf dem iPhone gibt es das nicht:
  // Dort füllt der Player das Fenster und behält seine Pfeil-Knöpfe (Safaris eigener Videovollbild hätte keine).
  if (!isImg) enterFullscreen();
  updateNav();
  els.ptitle.textContent = (viewIndex >= 0 ? (viewIndex + 1) + " / " + viewList.length + " · " : "") + item.name;
  els.pmsg.hidden = false;
  els.pmsg.className = "msg";
  els.pmsg.textContent = isImg ? "Lade Bild …" : "Lade Video …";
  els.player.hidden = false;
  els.pvideo.hidden = isImg;
  els.pimg.hidden = !isImg;
  els.pvideo.removeAttribute("src");
  els.pvideo.load();
  els.pimg.removeAttribute("src");
  try {
    // Bei jedem Start eine frische, kurzlebige Download-URL holen. Sie wird nirgends gespeichert.
    const data = await graph("/me/drive/items/" + encodeURIComponent(id));
    if (token !== playToken || !data) return;
    const url = data["@microsoft.graph.downloadUrl"];
    if (!url) throw new Error("Microsoft hat keine Download-Adresse geliefert.");
    if (isImg) {
      els.pimg.alt = item.name;
      els.pimg.src = url; // Meldung verschwindet, sobald das Bild geladen ist
      return;
    }
    lastDownloadUrl = url;
    els.pvideo.src = url;
    els.pmsg.hidden = true;
    try { await els.pvideo.play(); } catch (_) { /* Browser verlangt ggf. Tippen auf Play */ }
  } catch (e) {
    if (token !== playToken) return;
    els.pmsg.hidden = false;
    els.pmsg.className = "msg err";
    els.pmsg.textContent = (isImg ? "Bild" : "Video") + " konnte nicht geladen werden: " + (e.message || e);
  }
}

const MEDIA_ERRORS = {
  1: "Wiedergabe abgebrochen",
  2: "Netzwerkfehler beim Laden",
  3: "Dekodierfehler (Datei beschädigt oder Videoprofil nicht unterstützt)",
  4: "Format/Codec nicht unterstützt oder Adresse nicht abrufbar",
};

// Rückfall: Lehnt der Browser die Adresse direkt ab (z. B. Safari bei falschem Content-Type), wird die Datei
// komplett geladen und als Video/mp4 abgespielt.
let blobTried = false;
let blobUrl = null;
let probeInfo = "";
let lastDownloadUrl = "";

function revokeBlob() {
  if (blobUrl) { try { URL.revokeObjectURL(blobUrl); } catch (_) { /* nicht kritisch */ } blobUrl = null; }
}

// Lädt die Antwort stückweise und meldet den Fortschritt. Gibt null zurück, wenn inzwischen ein anderes Video läuft.
async function readWithProgress(r, token, onProgress) {
  const total = Number(r.headers.get("content-length")) || 0;
  if (!r.body || !r.body.getReader) {
    const b = await r.blob(); // Browser ohne Streaming: ohne Fortschritt
    return new Blob([b], { type: "video/mp4" });
  }
  const reader = r.body.getReader();
  const chunks = [];
  let got = 0;
  onProgress(0, total);
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (token !== playToken) { try { await reader.cancel(); } catch (_) { /* nicht kritisch */ } return null; }
    chunks.push(value);
    got += value.length;
    onProgress(got, total);
  }
  return new Blob(chunks, { type: "video/mp4" });
}

function showVideoError() {
  const err = els.pvideo.error;
  const code = err && err.code ? err.code : 0;
  const can = (t) => els.pvideo.canPlayType(t) || "nein";
  els.pmsg.hidden = false;
  els.pmsg.className = "msg err";
  els.pmsg.replaceChildren();
  const lines = [
    "Das Video lässt sich in diesem Browser nicht abspielen.",
    "Browser-Meldung: " + (code ? "Code " + code + " – " + (MEDIA_ERRORS[code] || "unbekannt") : "keine") +
      (err && err.message ? " (" + err.message + ")" : ""),
    "Dieser Browser kann laut eigener Angabe: H.264 " + can('video/mp4; codecs="avc1.640028"') +
      ", H.265 " + can('video/mp4; codecs="hvc1.1.6.L93.B0"') + ", AV1 " + can('video/mp4; codecs="av01.0.05M.08"'),
  ];
  if (probeInfo) lines.push(probeInfo);
  for (const t of lines) {
    const d = document.createElement("div");
    d.textContent = t;
    els.pmsg.append(d);
  }
  if (lastDownloadUrl) {
    const link = document.createElement("a");
    link.href = lastDownloadUrl; // frische, kurzlebige Adresse; gilt nur ca. 1 Stunde
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    link.textContent = "Datei direkt öffnen (neuer Tab)";
    link.style.color = "inherit";
    els.pmsg.append(link);
  }
}

els.pvideo.addEventListener("error", async () => {
  const src = els.pvideo.getAttribute("src");
  if (els.player.hidden || !src) return;
  const token = playToken;
  if (!blobTried && !src.startsWith("blob:")) {
    blobTried = true;
    els.pmsg.hidden = false;
    els.pmsg.className = "msg";
    els.pmsg.replaceChildren();
    const label = document.createElement("div");
    label.textContent = "Direktwiedergabe nicht möglich, lade die Datei komplett …";
    const bar = document.createElement("progress");
    bar.max = 100;
    bar.style.cssText = "width:100%;height:14px;margin-top:8px;accent-color:var(--accent)";
    els.pmsg.append(label, bar);
    try {
      const r = await fetch(src);
      const len = r.headers.get("content-length");
      probeInfo = "Abruf: Status " + r.status + ", Content-Type " + (r.headers.get("content-type") || "unbekannt") +
        (len ? ", " + sizeText(Number(len)) : "");
      if (!r.ok) throw new Error("Status " + r.status);
      const whole = await readWithProgress(r, token, (got, total) => {
        if (total) {
          const pct = Math.min(100, Math.floor((got / total) * 100));
          bar.value = pct;
          label.textContent = "Lade Datei … " + pct + " % (" + sizeText(got) + " von " + sizeText(total) + ")";
        } else {
          bar.removeAttribute("value"); // unbekannte Gesamtgröße: laufender Balken
          label.textContent = "Lade Datei … " + sizeText(got);
        }
      });
      if (!whole || token !== playToken) return;
      revokeBlob();
      blobUrl = URL.createObjectURL(whole);
      els.pvideo.src = blobUrl;
      els.pmsg.hidden = true;
      try { await els.pvideo.play(); } catch (_) { /* ggf. Tippen auf Play */ }
      return;
    } catch (e) {
      if (token !== playToken) return;
      probeInfo = probeInfo || "Abruf nicht möglich: " + ((e && e.message) || e);
    }
  }
  showVideoError();
});

els.pimg.addEventListener("load", () => { els.pmsg.hidden = true; });
els.pimg.addEventListener("error", () => {
  if (els.player.hidden || !els.pimg.getAttribute("src")) return;
  els.pmsg.hidden = false;
  els.pmsg.className = "msg err";
  els.pmsg.textContent = "Das Bild lässt sich in diesem Browser nicht anzeigen.";
});

function closePlayer() {
  playToken++;
  revokeBlob();
  els.pimg.removeAttribute("src");
  fsRequested = false;
  exitFullscreen();
  els.pvideo.pause();
  els.pvideo.removeAttribute("src");
  els.pvideo.load();
  els.player.hidden = true;
}

els.pclose.addEventListener("click", closePlayer);
els.pprev.addEventListener("click", () => step(-1));
els.pnext.addEventListener("click", () => step(1));
document.addEventListener("keydown", (e) => {
  if (els.player.hidden) return;
  if (e.key === "Escape") closePlayer();
  else if (e.key === "ArrowLeft") step(-1);
  else if (e.key === "ArrowRight") step(1);
});


/* ---------- Version ---------- */
// Bei jeder Änderung: Nummer in index.html (<meta name="app-version">) und version.json erhöhen, APP_DATE und APP_CHANGES neu schreiben.
const APP_VERSION = (document.querySelector('meta[name="app-version"]') || {}).content || "?";
const APP_DATE = "06.11.2026";
const APP_CHANGES = [
  "Fortschrittsbalken (Prozent und MB), wenn das Video komplett geladen werden muss",
];

/* ---------- Code-Sperre ---------- */
// Bequemlichkeits-Sperre vor dem Start der App. Kein Ersatz für die Microsoft-Anmeldung:
// Der Code steht nur als Hash im Quelltext, ein 4-stelliger Code ist aber technisch leicht zu knacken.
const PIN_HASH = "07e6809a7d52265fdbb3d6b9b825c439d6cfa620747a5aade457a7a3bf43f12a";
const GATE_KEY = "vk_gate_ok";

function gateOpen() {
  try { return sessionStorage.getItem(GATE_KEY) === "1"; } catch (_) { return false; }
}

async function sha256Hex(text) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

function runGate() {
  const gate = $("gate"), form = $("gateform"), pin = $("pin"), err = $("gateerr");
  const info = $("gateinfo");
  info.replaceChildren();
  const head = document.createElement("div");
  head.textContent = "Version " + APP_VERSION + " · " + APP_DATE;
  const ul = document.createElement("ul");
  for (const c of APP_CHANGES) {
    const li = document.createElement("li");
    li.textContent = c;
    ul.append(li);
  }
  info.append(head, ul);
  gate.hidden = false;
  pin.focus();
  let fails = 0;
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    err.textContent = "";
    let ok = false;
    try {
      ok = (await sha256Hex("video-kacheln|" + pin.value.trim())) === PIN_HASH;
    } catch (_) {
      err.textContent = "Prüfung hier nicht möglich (HTTPS erforderlich).";
      return;
    }
    if (ok) {
      try { sessionStorage.setItem(GATE_KEY, "1"); } catch (_) { /* dann fragt die App nach dem Microsoft-Login erneut */ }
      gate.hidden = true;
      start();
    } else {
      fails += 1;
      pin.value = "";
      err.textContent = "Falscher Code.";
      form.querySelector("button").disabled = true;
      setTimeout(() => { form.querySelector("button").disabled = false; pin.focus(); }, Math.min(1000 * fails, 5000));
    }
  });
}

if (gateOpen()) start(); else runGate();
