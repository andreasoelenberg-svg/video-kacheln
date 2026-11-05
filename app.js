"use strict";

/* ---------- Einstellungen ---------- */
const CLIENT_ID = "02882712-3952-4323-a968-d11d9b045236";
// Muss exakt der in Entra eingetragenen Umleitungs-URI entsprechen (mit Schrägstrich am Ende).
const REDIRECT_URI = "https://andreasoelenberg-svg.github.io/video-kacheln/";
// Ordnerpfad ab dem OneDrive-Stamm.
const FOLDER_PATH = "KI pic/Self KI/0. Harmless Videos";
const SCOPES = ["Files.Read"];
const GRAPH = "https://graph.microsoft.com/v1.0";
const DOUBLE_TAP_MS = 400;

/* ---------- Elemente ---------- */
const $ = (id) => document.getElementById(id);
const els = {
  grid: $("grid"), status: $("status"), count: $("count"), hint: $("hint"),
  sort: $("sort"), reload: $("reload"), logout: $("logout"),
  player: $("player"), pvideo: $("pvideo"), ptitle: $("ptitle"), pmsg: $("pmsg"), pclose: $("pclose"),
};

/* ---------- Zustand (nur im Arbeitsspeicher) ---------- */
let pca = null;
let items = [];
let sortMode = "name";
let selectedId = null;
let lastTap = { id: null, t: 0 };
let playToken = 0;

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
  await loadFolder();
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
function folderUrl() {
  const encoded = FOLDER_PATH.split("/").map(encodeURIComponent).join("/");
  return "/me/drive/root:/" + encoded + ":/children?$expand=thumbnails&$top=200";
}

async function loadFolder() {
  setStatus("Lade Videos …");
  els.grid.replaceChildren();
  els.count.textContent = "";
  try {
    const all = [];
    let next = folderUrl();
    while (next) {
      const page = await graph(next);
      if (!page) return; // Weiterleitung zur Anmeldung läuft
      all.push(...(page.value || []));
      next = page["@odata.nextLink"] || null;
    }
    items = all.filter((it) => it.file && /\.mp4$/i.test(it.name || ""));
    selectedId = null;
    els.sort.hidden = false;
    els.reload.hidden = false;
    els.logout.hidden = false;
    if (items.length === 0) {
      setStatus();
      showMessage("Im Ordner liegen keine MP4-Videos.");
      return;
    }
    setStatus();
    els.hint.hidden = false;
    render();
  } catch (e) {
    els.reload.hidden = false;
    els.logout.hidden = false;
    if (e.status === 404) {
      showError("Ordner nicht gefunden: " + FOLDER_PATH, true);
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
  const list = items.slice();
  if (sortMode === "date") {
    list.sort((a, b) => String(b.lastModifiedDateTime).localeCompare(String(a.lastModifiedDateTime)));
  } else {
    list.sort((a, b) => collator.compare(a.name, b.name));
  }
  return list;
}

function thumbUrl(item) {
  const t = item.thumbnails && item.thumbnails[0];
  if (!t) return null;
  return (t.medium && t.medium.url) || (t.large && t.large.url) || (t.small && t.small.url) || null;
}

function render() {
  const list = sorted();
  els.count.textContent = list.length + (list.length === 1 ? " Video" : " Videos");
  const frag = document.createDocumentFragment();
  for (const item of list) {
    const tile = document.createElement("div");
    tile.className = "tile";
    tile.setAttribute("role", "option");
    tile.tabIndex = 0;
    tile.dataset.id = item.id;
    tile.setAttribute("aria-selected", item.id === selectedId ? "true" : "false");
    tile.setAttribute("aria-label", item.name);

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
    const ph = document.createElement("div");
    ph.className = "ph";
    ph.textContent = "▶";
    ph.setAttribute("aria-hidden", "true");
    tile.prepend(ph);

    const cap = document.createElement("div");
    cap.className = "cap";
    const name = document.createElement("div");
    name.className = "name";
    name.textContent = item.name.replace(/\.mp4$/i, "");
    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = [sizeText(item.size), dateText(item.lastModifiedDateTime)].filter(Boolean).join(" · ");
    cap.append(name, meta);

    const pb = document.createElement("button");
    pb.type = "button";
    pb.className = "playbtn";
    pb.setAttribute("aria-label", "Abspielen: " + item.name);
    pb.textContent = "▶";

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

/* ---------- Bedienung der Kacheln ---------- */
els.grid.addEventListener("pointerup", (e) => {
  if (e.pointerType === "mouse" && e.button !== 0) return;
  const tile = e.target.closest(".tile");
  if (!tile) return;
  const id = tile.dataset.id;
  if (e.target.closest(".playbtn")) {
    select(id);
    lastTap = { id: null, t: 0 };
    play(id);
    return;
  }
  const now = performance.now();
  if (lastTap.id === id && now - lastTap.t < DOUBLE_TAP_MS) {
    lastTap = { id: null, t: 0 };
    select(id);
    play(id);
  } else {
    select(id);
    lastTap = { id, t: now };
  }
});

els.grid.addEventListener("keydown", (e) => {
  const tile = e.target.closest(".tile");
  if (!tile) return;
  if (e.key === "Enter") {
    e.preventDefault();
    select(tile.dataset.id);
    play(tile.dataset.id);
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

els.reload.addEventListener("click", () => loadFolder());
els.logout.addEventListener("click", () => {
  pca.logoutRedirect({ account: pca.getActiveAccount(), postLogoutRedirectUri: REDIRECT_URI });
});

/* ---------- Abspielen ---------- */
async function play(id) {
  const item = items.find((x) => x.id === id);
  if (!item) return;
  const token = ++playToken;
  els.ptitle.textContent = item.name;
  els.pmsg.hidden = false;
  els.pmsg.className = "msg";
  els.pmsg.textContent = "Lade Video …";
  els.player.hidden = false;
  els.pvideo.removeAttribute("src");
  els.pvideo.load();
  try {
    // Bei jedem Start eine frische, kurzlebige Download-URL holen. Sie wird nirgends gespeichert.
    const data = await graph("/me/drive/items/" + encodeURIComponent(id));
    if (token !== playToken || !data) return;
    const url = data["@microsoft.graph.downloadUrl"];
    if (!url) throw new Error("Microsoft hat keine Download-Adresse geliefert.");
    els.pvideo.src = url;
    els.pmsg.hidden = true;
    try { await els.pvideo.play(); } catch (_) { /* Browser verlangt ggf. Tippen auf Play */ }
  } catch (e) {
    if (token !== playToken) return;
    els.pmsg.hidden = false;
    els.pmsg.className = "msg err";
    els.pmsg.textContent = "Video konnte nicht geladen werden: " + (e.message || e);
  }
}

els.pvideo.addEventListener("error", () => {
  if (els.player.hidden || !els.pvideo.getAttribute("src")) return;
  els.pmsg.hidden = false;
  els.pmsg.className = "msg err";
  els.pmsg.textContent = "Das Video lässt sich in diesem Browser nicht abspielen.";
});

function closePlayer() {
  playToken++;
  els.pvideo.pause();
  els.pvideo.removeAttribute("src");
  els.pvideo.load();
  els.player.hidden = true;
}

els.pclose.addEventListener("click", closePlayer);
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && !els.player.hidden) closePlayer();
});

start();
