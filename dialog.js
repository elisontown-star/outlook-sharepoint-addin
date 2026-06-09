// ============================================================
//  Vshare — dialog.js
//  Lógica do pop-up de decisão no envio.
//  Autentica, carrega pastas e arquiva. Comunica a decisão
//  de volta ao commands.js via Office.context.ui.messageParent.
// ============================================================

let msalInstance = null;
let accessToken  = null;
let siteId       = null;
let driveId      = null;
let allFolders   = [];

Office.onReady(async () => {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:    CONFIG.CLIENT_ID,
      authority:   `https://login.microsoftonline.com/${CONFIG.TENANT_ID}`,
      redirectUri: CONFIG.REDIRECT_URI
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
  });

  // Tenta autenticar silenciosamente com a conta em cache
  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      const resp = await msalInstance.acquireTokenSilent({
        scopes: CONFIG.SCOPES, account: accounts[0]
      });
      accessToken = resp.accessToken;
      await prepareFolders();
    } else {
      // Sem sessão: o botão "Arquivar e Enviar" fará login via popup
      showStatus("Faça login ao escolher Arquivar e Enviar.", "loading");
    }
  } catch {
    showStatus("Faça login ao escolher Arquivar e Enviar.", "loading");
  }
});

// ── Prepara a lista de pastas ────────────────────────────────
async function prepareFolders() {
  document.getElementById("folderArea").style.display = "block";
  hideStatus();
  try {
    await loadSiteAndDrive();
    await loadRootFolders();
  } catch (e) {
    showStatus("Erro ao carregar pastas: " + e.message, "error");
  }
}

// ── Garante token (faz login se necessário) ──────────────────
async function ensureToken() {
  if (accessToken) return accessToken;
  const resp = await msalInstance.loginPopup({ scopes: CONFIG.SCOPES });
  const tok  = await msalInstance.acquireTokenSilent({
    scopes: CONFIG.SCOPES, account: resp.account
  });
  accessToken = tok.accessToken;
  return accessToken;
}

// ── Graph helpers ────────────────────────────────────────────
async function graphGet(endpoint) {
  const url = endpoint.startsWith("https://") ? endpoint
    : "https://graph.microsoft.com/v1.0" + endpoint;
  const res = await fetch(url, { headers: { Authorization: "Bearer " + accessToken } });
  if (!res.ok) throw new Error(`Graph ${res.status}`);
  return res.json();
}
async function graphGetAll(endpoint) {
  let items = [], next = endpoint;
  while (next) {
    const data = await graphGet(next);
    if (Array.isArray(data.value)) items = items.concat(data.value);
    next = data["@odata.nextLink"] || null;
  }
  return items;
}
async function graphPut(endpoint, body) {
  const res = await fetch("https://graph.microsoft.com/v1.0" + endpoint, {
    method: "PUT",
    headers: { Authorization: "Bearer " + accessToken, "Content-Type": "application/octet-stream" },
    body
  });
  if (!res.ok) throw new Error(`Graph PUT ${res.status}`);
  return res.json();
}
function isFolder(item) { return item && item.folder != null; }

async function loadSiteAndDrive() {
  const host = new URL(CONFIG.SITE_URL).hostname;
  const path = new URL(CONFIG.SITE_URL).pathname;
  const site = await graphGet(`/sites/${host}:${path}`);
  siteId = site.id;
  const drives = await graphGet(`/sites/${siteId}/drives`);
  const drive = drives.value.find(d => d.name.toLowerCase() === CONFIG.LIBRARY_NAME.toLowerCase()) || drives.value[0];
  driveId = drive.id;
}

async function loadRootFolders() {
  const items = await graphGetAll(`/drives/${driveId}/root/children?$select=id,name,folder&$top=200`);
  allFolders = items.filter(isFolder).sort((a, b) => a.name.localeCompare(b.name, "pt-BR"));
  const sel = document.getElementById("rootFolder");
  sel.innerHTML = '<option value="">— selecione a pasta —</option>';
  allFolders.forEach(f => {
    const o = document.createElement("option");
    o.value = f.id; o.textContent = f.name; sel.appendChild(o);
  });
  sel.addEventListener("change", () => loadSubfolders(sel.value));
}

async function loadSubfolders(folderId) {
  const wrap = document.getElementById("subWrap");
  const sel  = document.getElementById("subFolder");
  if (!folderId) { wrap.style.display = "none"; return; }
  try {
    const items = await graphGetAll(`/drives/${driveId}/items/${folderId}/children?$select=id,name,folder&$top=200`);
    const subs  = items.filter(isFolder).sort((a, b) => a.name.localeCompare(b.name, "pt-BR"));
    sel.innerHTML = '<option value="">— raiz da pasta acima —</option>';
    if (subs.length > 0) {
      subs.forEach(f => {
        const o = document.createElement("option");
        o.value = f.id; o.textContent = f.name; sel.appendChild(o);
      });
      wrap.style.display = "block";
    } else {
      wrap.style.display = "none";
    }
  } catch { wrap.style.display = "none"; }
}

// ── Busca simples (filtra pastas raiz carregadas) ────────────
document.addEventListener("DOMContentLoaded", () => {
  const si = document.getElementById("searchInput");
  if (si) si.addEventListener("input", () => {
    const q = si.value.toLowerCase();
    const sel = document.getElementById("rootFolder");
    sel.innerHTML = '<option value="">— selecione a pasta —</option>';
    allFolders.filter(f => f.name.toLowerCase().includes(q)).forEach(f => {
      const o = document.createElement("option");
      o.value = f.id; o.textContent = f.name; sel.appendChild(o);
    });
  });
});

// ── Ações dos botões ─────────────────────────────────────────
async function archiveAndSend() {
  const btn = document.getElementById("archiveSendBtn");
  btn.innerHTML = '<span class="loader"></span> Arquivando...';
  btn.disabled = true;

  try {
    await ensureToken();

    // Se as pastas ainda não carregaram (login feito agora), carrega
    if (!driveId) await prepareFolders();

    const rootId = document.getElementById("rootFolder").value;
    const subId  = document.getElementById("subFolder").value;
    if (!rootId) {
      showStatus("Selecione uma pasta antes de arquivar.", "error");
      btn.innerHTML = "Arquivar e Enviar"; btn.disabled = false;
      return;
    }
    const targetId = subId || rootId;

    showStatus("Arquivando email...", "loading");
    const { blob, ext, subject } = await getEmailBlob();
    const fileName = `${getLocalDate()}_${sanitize(subject)}.${ext}`;
    await graphPut(`/drives/${driveId}/items/${targetId}:/${encodeURIComponent(fileName)}:/content`, blob);

    // Avisa o commands.js que arquivou com sucesso → libera o envio
    sendDecision({ action: "archive_and_send", success: true });
  } catch (e) {
    showStatus("Erro: " + e.message, "error");
    sendDecision({ action: "archive_and_send", success: false, error: e.message });
  }
}

function sendOnly() { sendDecision({ action: "send_only" }); }
function cancelSend() { sendDecision({ action: "cancel" }); }

// ── Comunica a decisão de volta ao commands.js ───────────────
function sendDecision(obj) {
  Office.context.ui.messageParent(JSON.stringify(obj));
}

// ── Obtém o conteúdo do email sendo composto ─────────────────
// Em modo de composição usamos a Office.js para montar um EML.
async function getEmailBlob() {
  const item = Office.context.mailbox.item;

  const body = await new Promise(resolve => {
    item.body.getAsync(Office.CoercionType.Html, r =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "(corpo indisponível)")
    );
  });
  const subject = await new Promise(resolve => {
    item.subject.getAsync(r =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "(sem assunto)")
    );
  });
  const to = await new Promise(resolve => {
    item.to.getAsync(r =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded
        ? r.value.map(x => `${x.displayName} <${x.emailAddress}>`).join(", ") : "")
    );
  });

  const eml = [
    `To: ${to}`, `Subject: ${subject}`, `Date: ${new Date().toUTCString()}`,
    `MIME-Version: 1.0`, `Content-Type: text/html; charset=utf-8`, ``, body
  ].join("\r\n");

  return { blob: new Blob([eml], { type: "message/rfc822" }), ext: "eml", subject };
}

function getLocalDate() {
  const n = new Date();
  return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}-${String(n.getDate()).padStart(2,"0")}`;
}
function sanitize(s) { return (s || "email").replace(/[\\/:*?"<>|]/g, "_").substring(0, 80); }

function showStatus(msg, type) {
  const el = document.getElementById("status");
  el.textContent = msg; el.className = type; el.style.display = "block";
}
function hideStatus() { document.getElementById("status").style.display = "none"; }

// Expor para onclick inline
window.archiveAndSend = archiveAndSend;
window.sendOnly = sendOnly;
window.cancelSend = cancelSend;
