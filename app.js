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
  player: $("player"), pvideo: $("pvideo"), pimg: $("pimg"), ptitle: $("ptitle"), pmsg: $("pmsg"), pclose: $("pclose"), pprev: $("pprev"), pnext: $("pnext"), pthumb: $("pthumb"), pfolder: $("pfolder"),
  pfine: $("pfine"), ptime: $("ptime"), pslider: $("pslider"), pm1s: $("pm1s"), pmf: $("pmf"), ppf: $("ppf"), pp1s: $("pp1s"),
};

/* ---------- Zustand ---------- */
let pca = null;
let items = [];
let currentPath = ROOT_PATH;
let sortMode = "name";
let sortDesc = false; // Start: Name aufsteigend
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
    // Nach der Microsoft-Anmeldung landet man auf der Adresse ohne ?v=… und damit evtl. auf einer veralteten Startseite.
    // Dann hier einmalig auf die neue Version wechseln (die Anmeldung bleibt im sessionStorage erhalten).
    if (result && (await checkUpdate(true))) return;
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
  loadCustomThumbs(); // im Hintergrund
}

async function login() {
  try {
    await pca.loginRedirect({ scopes: SCOPES.concat([APP_SCOPE]), prompt: "select_account" });
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
    items = all.filter((it) => (it.folder && !HIDDEN_FOLDER_RE.test(it.name || "")) || (it.file && (VIDEO_RE.test(it.name || "") || IMAGE_RE.test(it.name || ""))));
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
  // Ordner stehen immer oben (nach Namen; bei der Namenssortierung in der gewählten Richtung), danach die Dateien.
  const dir = sortDesc ? -1 : 1;
  const nameDir = sortMode === "name" ? dir : 1;
  const folders = items.filter((it) => it.folder).sort((a, b) => nameDir * collator.compare(a.name, b.name));
  const videos = items.filter((it) => !it.folder); // Videos und Bilder
  if (sortMode === "date") {
    videos.sort((a, b) => dir * String(a.lastModifiedDateTime).localeCompare(String(b.lastModifiedDateTime)) || collator.compare(a.name, b.name));
  } else if (sortMode === "size") {
    videos.sort((a, b) => dir * ((a.size || 0) - (b.size || 0)) || collator.compare(a.name, b.name));
  } else {
    videos.sort((a, b) => dir * collator.compare(a.name, b.name));
  }
  return folders.concat(videos);
}

/* ---------- Eigene Vorschaubilder (App-Ordner in OneDrive) ---------- */
// Liegen als JPEG direkt im App-Ordner von OneDrive (anfangs Apps/<App-Name>/; der Ordner darf nach
// Self KI/Thumbnails verschoben werden). Dafür genügt die eng begrenzte Berechtigung Files.ReadWrite.AppFolder:
// Die App darf nur in diesem einen Ordner schreiben, sonst nirgends.
const APP_SCOPE = "Files.ReadWrite.AppFolder";
const HIDDEN_FOLDER_RE = /^thumbnails$/i; // dieser Ordner enthält nur die Vorschaubilder und wird nicht als Kachel gezeigt
let customThumbs = {}; // Dateiname -> Adresse des Bildes

const thumbName = (item) => String(item.id).replace(/[^A-Za-z0-9_-]/g, "_") + ".jpg";

// interactive = false: nie weiterleiten (still prüfen). true: bei fehlender Berechtigung zur Microsoft-Seite weiterleiten.
async function appFolderToken(interactive) {
  const account = pca.getActiveAccount();
  try {
    const r = await pca.acquireTokenSilent({ scopes: [APP_SCOPE], account });
    return r.accessToken;
  } catch (e) {
    if (!interactive) return null;
    await pca.acquireTokenRedirect({ scopes: [APP_SCOPE], account });
    return null; // Seite wird weitergeleitet
  }
}

// Zeigt in der Fußzeile, wie es mit den eigenen Vorschaubildern steht (und bei Bedarf einen Knopf zum Berechtigen).
function setThumbStatus(text, actionLabel, actionFn) {
  const el = $("ver");
  if (!el) return;
  el.textContent = "Version " + APP_VERSION + " · " + text;
  if (actionLabel) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "btn";
    b.textContent = actionLabel;
    b.addEventListener("click", actionFn);
    el.append(" ", b);
  }
}

async function loadCustomThumbs() {
  try {
    const token = await appFolderToken(false);
    if (!token) {
      setThumbStatus("Vorschaubilder: Berechtigung fehlt", "Berechtigen", () => appFolderToken(true));
      return;
    }
    const map = {};
    let next = GRAPH + "/me/drive/special/approot/children?$select=name,@microsoft.graph.downloadUrl&$top=1000";
    while (next) {
      const res = await fetch(next, { headers: { Authorization: "Bearer " + token } });
      if (res.status === 404) { setThumbStatus("Vorschaubilder: App-Ordner noch leer"); return; }
      if (!res.ok) {
        let code = "";
        try { const j = await res.json(); code = (j && j.error && j.error.code) || ""; } catch (_) { /* nicht kritisch */ }
        setThumbStatus("Vorschaubilder: OneDrive-Fehler " + res.status + (code ? " (" + code + ")" : ""));
        return;
      }
      const page = await res.json();
      for (const f of page.value || []) {
        const u = f["@microsoft.graph.downloadUrl"];
        if (u && f.name) map[f.name] = u;
      }
      next = page["@odata.nextLink"] || null;
    }
    Object.assign(customThumbs, map);
    setThumbStatus("Vorschaubilder: " + Object.keys(map).length + " geladen");
    if (items.length) render();
  } catch (e) {
    setThumbStatus("Vorschaubilder: Fehler (" + ((e && (e.errorCode || e.message)) || e) + ")");
  }
}

function thumbUrl(item) {
  if (!isImage(item) && customThumbs[thumbName(item)]) return customThumbs[thumbName(item)];
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

    let hasCover = false;
    {
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
        hasCover = true;
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
      meta.textContent = [isImg ? "BILD" : "VIDEO", sizeText(item.size), dateText(item.lastModifiedDateTime)].filter(Boolean).join(" · ");
    }
    cap.append(name, meta);

    const pb = document.createElement("button");
    pb.type = "button";
    pb.className = "playbtn";
    pb.setAttribute("aria-label", (isFolder ? "Öffnen: " : isImg ? "Anzeigen: " : "Abspielen: ") + item.name);
    pb.textContent = isFolder ? "›" : isImg ? "⤢" : "▶";

    tile.append(cap, pb);
    if (isFolder && hasCover) {
      // Ordner mit eigenem Bild: grüner Rahmen und Reiter mit Ordner-Symbol, damit er nicht mit einem Video verwechselt wird.
      tile.classList.add("hascover");
      const ft = document.createElement("div");
      ft.className = "foldertag";
      ft.setAttribute("aria-hidden", "true");
      ft.innerHTML = FOLDER_SVG; // feste, eigene Grafik
      ft.append("ORDNER");
      tile.append(ft);
    }
    if (!isFolder) {
      const badge = document.createElement("div");
      badge.className = "kindtag " + (isImg ? "image" : "video");
      badge.setAttribute("aria-hidden", "true");
      badge.textContent = isImg ? "BILD" : "VIDEO";
      tile.append(badge);
    }
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
  if (sortMode === b.dataset.sort) {
    sortDesc = !sortDesc; // erneut gedrückt: Richtung umkehren
  } else {
    sortMode = b.dataset.sort;
    sortDesc = true; // erster Druck: absteigend
  }
  updateSortButtons();
  render();
});

const SORT_LABELS = { name: "Name", date: "Datum", size: "Größe" };
function updateSortButtons() {
  for (const x of els.sort.querySelectorAll("button")) {
    const on = x.dataset.sort === sortMode;
    x.setAttribute("aria-pressed", on ? "true" : "false");
    x.textContent = SORT_LABELS[x.dataset.sort] + (on ? (sortDesc ? " ▼" : " ▲") : "");
  }
}
updateSortButtons();

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
function flash(text, isError, ms) {
  els.pmsg.hidden = false;
  els.pmsg.className = isError ? "msg err" : "msg";
  els.pmsg.textContent = text;
  const mine = text;
  setTimeout(() => { if (els.pmsg.textContent === mine) els.pmsg.hidden = true; }, ms || 3000);
}

// Nimmt ein Standbild des Videos an der Stelle t (Sekunden). Eigenes verstecktes Video, damit die Wiedergabe unberührt bleibt.
// Safari auf dem iPhone liefert für ein nicht angezeigtes, nicht abgespieltes Video oft ein schwarzes Bild. Darum hängt das
// Video unsichtbar im Dokument, und ein rein schwarzes Ergebnis wird nie gespeichert: erst ein zweiter Versuch mit kurzem Abspielen,
// danach eine klare Fehlermeldung.
function captureFrame(src, t) {
  return new Promise((resolve, reject) => {
    const v = document.createElement("video");
    v.crossOrigin = "anonymous";
    v.muted = true;
    v.setAttribute("muted", "");
    v.playsInline = true;
    v.setAttribute("playsinline", "");
    v.preload = "auto";
    v.style.cssText = "position:fixed;left:0;top:0;width:2px;height:2px;opacity:0;pointer-events:none;";
    document.body.append(v);
    let finished = false;
    const done = (err, blob) => {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      try { v.pause(); } catch (_) { /* nicht kritisch */ }
      v.removeAttribute("src");
      try { v.load(); } catch (_) { /* nicht kritisch */ }
      v.remove();
      if (err) reject(err); else resolve(blob);
    };
    const timer = setTimeout(() => done(new Error("Zeitüberschreitung beim Laden des Videos")), 25000);
    // Zeichnet das aktuelle Bild; liefert { canvas, black }.
    const grab = () => {
      if (!v.videoWidth) throw new Error("Videogröße unbekannt");
      const w = 480, h = Math.round((w * v.videoHeight) / v.videoWidth);
      const c = document.createElement("canvas");
      c.width = w;
      c.height = h;
      const ctx = c.getContext("2d");
      ctx.drawImage(v, 0, 0, w, h);
      const px = ctx.getImageData(0, 0, w, h).data;
      let black = true;
      for (let i = 0; i < px.length; i += 16) { // jedes 4. Pixel genügt
        if (px[i] > 12 || px[i + 1] > 12 || px[i + 2] > 12) { black = false; break; }
      }
      return { canvas: c, black };
    };
    const finish = (g) => g.canvas.toBlob((b) => (b ? done(null, b) : done(new Error("Bild konnte nicht erzeugt werden"))), "image/jpeg", 0.82);
    v.addEventListener("error", () => done(new Error("Video konnte für das Standbild nicht geladen werden")));
    v.addEventListener("loadedmetadata", () => { try { v.currentTime = t; } catch (e) { done(e); } });
    v.addEventListener("seeked", () => {
      setTimeout(() => {
        try {
          const first = grab();
          if (!first.black) { finish(first); return; }
          // Zweiter Versuch: kurz abspielen, damit der Browser ein Bild entschlüsselt.
          const p = v.play();
          if (p && p.catch) p.catch(() => {});
          let tried = false;
          const second = () => {
            if (tried || finished) return;
            tried = true;
            try {
              v.pause();
              const g = grab();
              if (g.black) done(new Error("Das Standbild ist komplett schwarz. Dieses Video lässt sich in diesem Browser nicht auslesen."));
              else finish(g);
            } catch (e) { done(e); }
          };
          if (v.requestVideoFrameCallback) v.requestVideoFrameCallback(() => setTimeout(second, 100));
          setTimeout(second, 1500);
        } catch (e) { done(e); }
      }, 200);
    }, { once: true });
    v.src = src;
  });
}

let savingThumb = false;

// Legt das Bild (JPEG) im App-Ordner ab; key = OneDrive-ID der Datei bzw. des Ordners, zu dem das Bild gehört.
async function putThumbBlob(key, blob, token) {
  const name = thumbName({ id: key });
  const res = await fetch(GRAPH + "/me/drive/special/approot:/" + name + ":/content", {
    method: "PUT",
    headers: { Authorization: "Bearer " + token, "Content-Type": "image/jpeg" },
    body: blob,
  });
  if (!res.ok) throw new Error("OneDrive antwortet mit Status " + res.status);
  const old = customThumbs[name];
  if (old && old.startsWith("blob:")) URL.revokeObjectURL(old);
  customThumbs[name] = URL.createObjectURL(blob);
}

// Verkleinert ein Bild aus OneDrive auf ein JPEG (480 px breit).
async function imageToJpeg(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error("Bild nicht ladbar (Status " + r.status + ")");
  const objUrl = URL.createObjectURL(await r.blob());
  try {
    const img = await new Promise((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = () => reject(new Error("Bild nicht lesbar"));
      i.src = objUrl;
    });
    const w = 480, h = Math.max(1, Math.round((w * img.naturalHeight) / img.naturalWidth));
    const c = document.createElement("canvas");
    c.width = w;
    c.height = h;
    c.getContext("2d").drawImage(img, 0, 0, w, h);
    return await new Promise((resolve, reject) => c.toBlob((b) => (b ? resolve(b) : reject(new Error("Bild konnte nicht erzeugt werden"))), "image/jpeg", 0.82));
  } finally {
    URL.revokeObjectURL(objUrl);
  }
}

// kind "video": Standbild als Kachelbild des Videos. kind "folder": Standbild bzw. Bild als Kachelbild des aktuellen Ordners.
async function saveThumbnail(kind) {
  const item = viewList[viewIndex];
  if (savingThumb || !item) return;
  if (kind === "video" && isImage(item)) return;
  if (kind === "folder" && currentPath === ROOT_PATH) { flash("Der Startordner hat keine Kachel.", true, 5000); return; }
  savingThumb = true;
  els.pthumb.disabled = true;
  els.pfolder.disabled = true;
  const what = kind === "folder" ? "Ordnerbild" : "Vorschaubild";
  try {
    const img = isImage(item);
    const t = img ? 0 : els.pvideo.currentTime || 0;
    const token = await appFolderToken(true);
    if (!token) { flash("Weiterleitung zur Berechtigung …", false, 8000); return; }
    flash("Speichere " + what + " …", false, 30000);
    const enc = currentPath.split("/").map(encodeURIComponent).join("/");
    const [data, folder] = await Promise.all([
      graph("/me/drive/items/" + encodeURIComponent(item.id)),
      kind === "folder" ? graph("/me/drive/root:/" + enc + "?$select=id") : Promise.resolve(null),
    ]);
    const url = data && data["@microsoft.graph.downloadUrl"];
    if (!url) throw new Error("keine Download-Adresse");
    const key = kind === "folder" ? folder && folder.id : item.id;
    if (!key) throw new Error("Ordner nicht gefunden");
    const blob = img ? await imageToJpeg(url) : await captureFrame(url, t);
    await putThumbBlob(key, blob, token);
    render();
    flash(what + " gespeichert.");
  } catch (e) {
    flash(what + " nicht möglich: " + ((e && e.message) || e), true, 8000);
  } finally {
    savingThumb = false;
    els.pthumb.disabled = false;
    els.pfolder.disabled = false;
  }
}

/* ---------- Feinwahl im Pausenmodus ---------- */
// Zeigt bei angehaltenem Video einen Feinregler (±5 s um die Pausenstelle), Sprünge um 1 s und Bild-für-Bild-Knöpfe.
let videoFps = 30;
let fpsKnown = false;
let anchor = 0; // Mitte des Feinreglers (Sekunden)
let sliding = false;
let slideTarget = 0;
const FPS_STEPS = [23.976, 24, 25, 29.97, 30, 48, 50, 59.94, 60, 120];

function snapFps(f) {
  for (const s of FPS_STEPS) if (Math.abs(f - s) / s < 0.02) return s;
  return Math.round(f * 10) / 10;
}

// Schätzt die Bildrate aus der Wiedergabe (requestVideoFrameCallback). Ohne Messung gelten 30 Bilder/s.
function startFpsProbe(token) {
  const v = els.pvideo;
  videoFps = 30;
  fpsKnown = false;
  if (!v.requestVideoFrameCallback) return;
  let last = null;
  const samples = [];
  const cb = (now, md) => {
    if (token !== playToken) return;
    if (last && md.presentedFrames - last.pf === 1) {
      const d = md.mediaTime - last.mt;
      if (d > 0.004 && d < 0.2) samples.push(d);
    }
    last = { pf: md.presentedFrames, mt: md.mediaTime };
    if (samples.length >= 15) {
      samples.sort((x, y) => x - y);
      const f = 1 / samples[samples.length >> 1];
      if (f >= 10 && f <= 125) { videoFps = snapFps(f); fpsKnown = true; updateFine(); }
      return;
    }
    v.requestVideoFrameCallback(cb);
  };
  v.requestVideoFrameCallback(cb);
}

function fmtTime(t) {
  if (!isFinite(t) || t < 0) t = 0;
  const m = Math.floor(t / 60);
  const s = (t - m * 60).toFixed(2).replace(".", ",");
  return m + ":" + (s.length < 5 ? "0" : "") + s;
}

function fineVisible() {
  const v = els.pvideo;
  return !els.player.hidden && !v.hidden && v.paused && v.readyState >= 1 && isFinite(v.duration) && v.duration > 0;
}

function updateFine() {
  const show = fineVisible();
  els.pfine.hidden = !show;
  if (!show) return;
  const v = els.pvideo;
  const cur = sliding ? slideTarget : v.currentTime;
  els.ptime.textContent = fmtTime(cur) + " / " + fmtTime(v.duration) + " · " +
    (fpsKnown ? videoFps + " Bilder/s" : "Bildrate unbekannt, 30 angenommen");
}

function seekTo(t) {
  const v = els.pvideo;
  const max = Math.max(0, v.duration - 0.5 / videoFps);
  v.currentTime = Math.min(Math.max(0, t), max);
}

// Ein Bild weiter/zurück: auf die Mitte des Nachbarbildes springen, damit Rundungen nicht auf dem alten Bild landen.
function stepFrame(dir) {
  const v = els.pvideo;
  if (!fineVisible()) return;
  seekTo((Math.floor(v.currentTime * videoFps + 1e-3) + dir + 0.5) / videoFps);
  updateFine();
}

function stepSeconds(dir) {
  if (!fineVisible()) return;
  seekTo(els.pvideo.currentTime + dir);
  updateFine();
}

// Knopf: sofort ein Schritt, bei gedrückt halten schnelle Wiederholung.
function holdButton(btn, fn) {
  let t1 = null, t2 = null;
  const stop = () => { clearTimeout(t1); clearInterval(t2); t1 = t2 = null; };
  btn.addEventListener("pointerdown", (e) => {
    if (e.pointerType === "mouse" && e.button !== 0) return;
    e.preventDefault();
    fn();
    stop();
    t1 = setTimeout(() => { t2 = setInterval(fn, 90); }, 450);
  });
  for (const ev of ["pointerup", "pointercancel", "pointerleave", "lostpointercapture"]) btn.addEventListener(ev, stop);
  btn.addEventListener("click", (e) => { if (e.detail === 0) fn(); }); // Tastatur
}

els.pslider.addEventListener("input", () => {
  if (!fineVisible()) return;
  sliding = true;
  slideTarget = anchor + parseFloat(els.pslider.value);
  seekTo(slideTarget);
  updateFine();
});
els.pslider.addEventListener("change", () => {
  sliding = false;
  anchor = Math.min(Math.max(0, slideTarget), els.pvideo.duration || 0);
  els.pslider.value = "0";
  updateFine();
});
holdButton(els.pm1s, () => stepSeconds(-1));
holdButton(els.pp1s, () => stepSeconds(1));
holdButton(els.pmf, () => stepFrame(-1));
holdButton(els.ppf, () => stepFrame(1));

els.pvideo.addEventListener("pause", () => { anchor = els.pvideo.currentTime; els.pslider.value = "0"; updateFine(); });
els.pvideo.addEventListener("play", updateFine);
els.pvideo.addEventListener("loadedmetadata", updateFine);
els.pvideo.addEventListener("emptied", updateFine);
els.pvideo.addEventListener("seeked", () => {
  if (!sliding) { anchor = els.pvideo.currentTime; els.pslider.value = "0"; }
  updateFine();
});

function updateNav() {
  const cur = viewList[viewIndex];
  els.pthumb.hidden = !cur || isImage(cur);
  els.pfolder.hidden = !cur;
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
  fileInfo = "";
  lastDownloadUrl = "";
  lastItemName = item.name;
  vlcTried = false;
  if (blobCtl) { blobCtl.abort(); blobCtl = null; }
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
    startFpsProbe(token);
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
let fileInfo = "";
let lastDownloadUrl = "";
let lastItemName = "";
let blobCtl = null;

const IS_IOS = /iPad|iPhone|iPod/.test(navigator.userAgent) ||
  (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1); // iPadOS meldet sich als Mac
let vlcTried = false;

function vlcStreamUrl() {
  return "vlc-x-callback://x-callback-url/stream?url=" + encodeURIComponent(lastDownloadUrl);
}

function vlcLaunch(url) {
  window.location.href = url;
}

// Auf iPhone/iPad: Kann der Browser das Video nicht abspielen, startet es automatisch in der VLC-App.
function autoVlc() {
  vlcTried = true;
  els.pmsg.hidden = false;
  els.pmsg.className = "msg";
  els.pmsg.replaceChildren();
  const l1 = document.createElement("div");
  l1.textContent = "Nicht abspielbar im Browser – wird in VLC geöffnet.";
  els.pmsg.append(l1);
  const vb = vlcBox();
  if (vb) els.pmsg.append(vb);
  vlcLaunch(vlcStreamUrl());
}

// Knöpfe für die VLC-App (iPhone/iPad). Adressen laut VLC-Dokumentation: vlc-x-callback://x-callback-url/stream bzw. /download.
function vlcBox() {
  if (!lastDownloadUrl) return null;
  const box = document.createElement("div");
  box.style.cssText = "display:flex;flex-wrap:wrap;gap:10px;margin:10px 0 4px";
  const mk = (text, href) => {
    const l = document.createElement("a");
    l.className = "btn";
    l.href = href;
    l.textContent = text;
    l.style.cssText = "text-decoration:none;display:inline-flex;align-items:center;color:var(--text)";
    l.addEventListener("click", () => { if (blobCtl) blobCtl.abort(); }); // laufendes Laden hier beenden
    return l;
  };
  const u = encodeURIComponent(lastDownloadUrl);
  box.append(
    mk("In VLC abspielen", vlcStreamUrl()),
    mk("In VLC laden", "vlc-x-callback://x-callback-url/download?url=" + u + "&filename=" + encodeURIComponent(lastItemName || "video.mp4"))
  );
  return box;
}

function revokeBlob() {
  if (blobUrl) { try { URL.revokeObjectURL(blobUrl); } catch (_) { /* nicht kritisch */ } blobUrl = null; }
}

/* ---------- MP4-Analyse (Diagnose und kleine Korrektur) ---------- */
async function readBytes(blob, start, end) {
  const part = blob.slice(start, end);
  if (part.arrayBuffer) return new Uint8Array(await part.arrayBuffer());
  return new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = () => res(new Uint8Array(fr.result));
    fr.onerror = () => rej(fr.error);
    fr.readAsArrayBuffer(part);
  });
}
const u32 = (b, o) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
const cc = (b, o) => String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]);

// Sucht in buf[start,end) entlang einer Kette von Box-Typen; liefert {s: Beginn des Inhalts, e: Ende} oder null.
function findBox(buf, start, end, path) {
  let o = start;
  while (o + 8 <= end) {
    const size = u32(buf, o);
    if (size < 8 || o + size > end) return null;
    if (cc(buf, o + 4) === path[0]) {
      if (path.length === 1) return { s: o + 8, e: o + size, box: o };
      return findBox(buf, o + 8, o + size, path.slice(1));
    }
    o += size;
  }
  return null;
}

const AVC_PROFILES = { 66: "Baseline", 77: "Main", 88: "Extended", 100: "High", 110: "High 10", 122: "High 4:2:2", 244: "High 4:4:4" };

// Liest Codec, Profil, Auflösung und Aufbau der Datei. Korrigiert, wenn nötig, Codec-Kennungen, die Safari ablehnt.
async function analyzeMp4(blob) {
  let pos = 0, moov = null, mdatPos = -1, brand = "";
  while (pos + 8 <= blob.size) {
    const h = await readBytes(blob, pos, pos + 16);
    let size = u32(h, 0);
    const type = cc(h, 4);
    let hdr = 8;
    if (size === 1) { size = u32(h, 8) * 4294967296 + u32(h, 12); hdr = 16; } else if (size === 0) size = blob.size - pos;
    if (size < hdr) break;
    if (type === "ftyp") brand = cc(h, 8);
    if (type === "mdat" && mdatPos < 0) mdatPos = pos;
    if (type === "moov") moov = { pos, size };
    pos += size;
  }
  if (!moov) return { text: "Datei-Analyse: kein moov-Block gefunden (Datei unvollständig oder kein gewöhnliches MP4)" };

  const buf = await readBytes(blob, moov.pos, moov.pos + moov.size);
  const parts = [], hints = [];
  const patches = [];
  let o = 8;
  while (o + 8 <= buf.length) {
    const size = u32(buf, o);
    if (size < 8) break;
    if (cc(buf, o + 4) === "trak") {
      const hdlr = findBox(buf, o + 8, o + size, ["mdia", "hdlr"]);
      const stsd = findBox(buf, o + 8, o + size, ["mdia", "minf", "stbl", "stsd"]);
      const tkhd = findBox(buf, o + 8, o + size, ["tkhd"]);
      const handler = hdlr ? cc(buf, hdlr.s + 8) : "";
      if (stsd) {
        const e = stsd.s + 8; // erster Eintrag
        const esize = u32(buf, e);
        const fmt = cc(buf, e + 4);
        if (handler === "vide") {
          let t = fmt;
          const kids = [e + 86, e + esize];
          const avcC = findBox(buf, kids[0], kids[1], ["avcC"]);
          const hvcC = findBox(buf, kids[0], kids[1], ["hvcC"]);
          if (avcC) {
            const p = buf[avcC.s + 1], lvl = buf[avcC.s + 3];
            t += " " + (AVC_PROFILES[p] || "Profil " + p) + " Level " + (lvl / 10).toFixed(1);
            if (p >= 110) hints.push("H.264 mit 10 Bit bzw. 4:2:2/4:4:4 (Profil " + (AVC_PROFILES[p] || p) + "): Safari unterstützt H.264 nur mit 8 Bit 4:2:0");
          }
          if (hvcC) {
            const chroma = buf[hvcC.s + 16] & 3, depth = (buf[hvcC.s + 17] & 7) + 8, prof = buf[hvcC.s + 1] & 31;
            t += " Profil " + prof + ", " + depth + " Bit, Chroma " + ({ 0: "4:0:0", 1: "4:2:0", 2: "4:2:2", 3: "4:4:4" }[chroma]);
            if (chroma !== 1 || depth > 10) hints.push("HEVC mit " + depth + " Bit / Chroma " + chroma + ": Safari unterstützt nur 4:2:0 mit 8 oder 10 Bit");
          }
          if (tkhd) {
            const v1 = buf[tkhd.s] === 1;
            const w = u32(buf, tkhd.s + (v1 ? 88 : 76)) >>> 16, hh = u32(buf, tkhd.s + (v1 ? 92 : 80)) >>> 16;
            if (w && hh) t += " " + w + "×" + hh;
          }
          parts.push("Video " + t);
          if (fmt === "hev1") { patches.push([moov.pos + e + 4, "hvc1"]); hints.push("Kennung hev1 auf hvc1 gesetzt (Safari spielt HEVC nur mit hvc1)"); }
          if (fmt === "avc3") { patches.push([moov.pos + e + 4, "avc1"]); hints.push("Kennung avc3 auf avc1 gesetzt"); }
          if (fmt === "av01") hints.push("AV1 wird von diesem Browser laut eigener Angabe nicht unterstützt");
          if (fmt === "vp09") hints.push("VP9 in MP4 unterstützt Safari nur eingeschränkt");
        } else if (handler === "soun") {
          parts.push("Audio " + fmt);
        }
      }
    }
    o += size;
  }
  let text = "Datei-Analyse: " + (parts.join(", ") || "keine Spuren gefunden") +
    "; Metadaten " + (mdatPos >= 0 && mdatPos < moov.pos ? "am Ende" : "am Anfang") + (brand ? "; Marke " + brand : "");
  if (hints.length) text += "\nHinweis (aus der Dateistruktur abgeleitet): " + hints.join("; ");
  let fixed = null;
  if (patches.length) {
    const copy = buf.slice();
    for (const [abs, name] of patches) for (let i = 0; i < 4; i++) copy[abs - moov.pos + i] = name.charCodeAt(i);
    fixed = new Blob([blob.slice(0, moov.pos), copy, blob.slice(moov.pos + moov.size)], { type: "video/mp4" });
  }
  return { text, blob: fixed };
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
  // Einzelheiten nur in der Konsole, auf dem Bildschirm steht ein einzelner Satz.
  try { console.info("Videofehler:", code, err && err.message, probeInfo, fileInfo); } catch (_) { /* nicht kritisch */ }
  els.pmsg.hidden = false;
  els.pmsg.className = "msg err";
  els.pmsg.textContent = "Das Video lässt sich in diesem Browser nicht abspielen.";
}

els.pvideo.addEventListener("error", async () => {
  const src = els.pvideo.getAttribute("src");
  if (els.player.hidden || !src) return;
  const token = playToken;
  if (IS_IOS && lastDownloadUrl && !src.startsWith("blob:")) {
    if (!vlcTried) autoVlc(); // nur einmal pro Video; kein Komplett-Download mehr auf iPhone/iPad
    return;
  }
  if (!blobTried && !src.startsWith("blob:")) {
    blobTried = true;
    els.pmsg.hidden = false;
    els.pmsg.className = "msg";
    els.pmsg.replaceChildren();
    const label = document.createElement("div");
    label.textContent = "Direktwiedergabe nicht möglich, lade die Datei komplett …";
    const track = document.createElement("div");
    track.className = "pbar";
    const bar = document.createElement("div");
    track.append(bar);
    els.pmsg.append(label, track);
    const vb = vlcBox();
    if (vb) els.pmsg.append(vb);
    blobCtl = typeof AbortController === "function" ? new AbortController() : null;
    try {
      const r = await fetch(src, blobCtl ? { signal: blobCtl.signal } : undefined);
      const len = r.headers.get("content-length");
      probeInfo = "Abruf: Status " + r.status + ", Content-Type " + (r.headers.get("content-type") || "unbekannt") +
        (len ? ", " + sizeText(Number(len)) : "");
      if (!r.ok) throw new Error("Status " + r.status);
      const whole = await readWithProgress(r, token, (got, total) => {
        if (total) {
          const pct = Math.min(100, Math.floor((got / total) * 100));
          track.classList.remove("indet");
          bar.style.width = pct + "%";
          label.textContent = "Lade Datei … " + pct + " % (" + sizeText(got) + " von " + sizeText(total) + ")";
        } else {
          track.classList.add("indet"); // unbekannte Gesamtgröße: laufender Balken
          label.textContent = "Lade Datei … " + sizeText(got);
        }
      });
      if (!whole || token !== playToken) return;
      let playable = whole;
      try {
        const res = await analyzeMp4(whole);
        fileInfo = res.text;
        if (res.blob) playable = res.blob;
      } catch (_) {
        fileInfo = "Datei-Analyse nicht möglich";
      }
      if (token !== playToken) return;
      revokeBlob();
      blobUrl = URL.createObjectURL(playable);
      els.pvideo.src = blobUrl;
      els.pmsg.hidden = true;
      try { await els.pvideo.play(); } catch (_) { /* ggf. Tippen auf Play */ }
      return;
    } catch (e) {
      if (token !== playToken) return;
      if (e && e.name === "AbortError") { // über VLC-Knopf beendet
        track.remove();
        label.textContent = "Laden hier beendet. Das Video wird in VLC geöffnet.";
        return;
      }
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
  els.pfine.hidden = true;
}

els.pclose.addEventListener("click", closePlayer);
els.pprev.addEventListener("click", () => step(-1));
els.pthumb.addEventListener("click", () => saveThumbnail("video"));
els.pfolder.addEventListener("click", () => saveThumbnail("folder"));
els.pnext.addEventListener("click", () => step(1));
document.addEventListener("keydown", (e) => {
  if (els.player.hidden) return;
  if (e.key === "Escape") closePlayer();
  else if (e.key === "ArrowLeft") step(-1);
  else if (e.key === "ArrowRight") step(1);
  else if (e.key === "," && fineVisible()) stepFrame(-1);
  else if (e.key === "." && fineVisible()) stepFrame(1);
});


/* ---------- Version ---------- */
// Bei jeder Änderung: Nummer in index.html (<meta name="app-version">) und version.json erhöhen, APP_DATE und APP_CHANGES neu schreiben.
const APP_VERSION = (document.querySelector('meta[name="app-version"]') || {}).content || "?";
const APP_DATE = "07.11.2026";
const APP_CHANGES = [
  "Fußzeile zeigt, wie viele Vorschaubilder geladen wurden (oder den Fehlergrund); Anmeldung fragt die Berechtigung dafür gleich mit",
];

/* ---------- Selbstaktualisierung ---------- */
function goTo(url) {
  window.location.replace(url);
}

// Prüft, ob es eine neuere Version gibt (z. B. wenn der Tab lange offen war) und lädt dann neu.
// afterLogin = true: direkt nach der Anmeldung (Adressteil ist verarbeitet), mit eigenem Einmal-Schutz.
// Gibt true zurück, wenn zur neuen Version gewechselt wird.
async function checkUpdate(afterLogin) {
  try {
    if (!els.player.hidden) return false; // nie mitten in einem Video
    if (!afterLogin && /[#&](code|state|error)=/.test(location.hash || "")) return false; // Anmeldung läuft
    const r = await fetch("./version.json?t=" + Date.now(), { cache: "no-store" });
    const j = await r.json();
    const latest = String(j.version);
    if (!/^\d+$/.test(latest) || latest === String(APP_VERSION)) return false;
    const key = afterLogin === true ? "vk_reload_login" : "vk_reload";
    let tried = null;
    try { tried = sessionStorage.getItem(key); } catch (_) { /* nicht kritisch */ }
    if (tried === latest) return false; // nur ein Versuch pro Version
    try { sessionStorage.setItem(key, latest); } catch (_) { /* nicht kritisch */ }
    goTo(location.pathname + "?v=" + latest + (afterLogin === true ? "" : location.hash || ""));
    return true;
  } catch (_) { return false; /* offline o. Ä.: nichts tun */ }
}
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") checkUpdate(); });
window.addEventListener("pageshow", (e) => { if (e.persisted) checkUpdate(); });
els.reload.addEventListener("click", () => checkUpdate());

const verEl = $("ver");
if (verEl && !verEl.textContent) verEl.textContent = "Version " + APP_VERSION;

// App-Name immer mit Versionsnummer: „Video-Kacheln V.24“ (Tab-Titel, Überschrift, Code-Sperre).
const APP_TITLE = "Video-Kacheln V." + APP_VERSION;
document.title = APP_TITLE;
for (const el of document.querySelectorAll("h1, #gate h2")) el.textContent = APP_TITLE;

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
