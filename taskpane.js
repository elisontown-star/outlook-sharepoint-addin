// ============================================================
//  Outlook Add-in — Arquivar Email no SharePoint
//  Autor: VtechIT  |  Versão: 1.0
// ============================================================

let msalInstance = null;
let accessToken  = null;
let currentItem  = null;
let siteId       = null;
let driveId      = null;

// ── Inicialização ────────────────────────────────────────────
Office.onReady(async () => {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:   CONFIG.CLIENT_ID,
      authority:  `https://login.microsoftonline.com/${CONFIG.TENANT_ID}`,
      redirectUri: "https://elisontown-star.github.io/outlook-sharepoint-addin/taskpane.html"
    },
    cache: { cacheLocation: "sessionStorage" }
  });

  // Tenta login silencioso com conta já em cache
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    try {
      const resp = await msalInstance.acquireTokenSilent({
        scopes:  CONFIG.SCOPES,
        account: accounts[0]
      });
      accessToken = resp.accessToken;
      await onLoggedIn();
    } catch {
      showLoginSection();
    }
  } else {
    showLoginSection();
  }
});

function showLoginSection() {
  document.getElementById("loginSection").style.display = "block";
  document.getElementById("folderSection").style.display = "none";
}

// ── Login ────────────────────────────────────────────────────
async function login() {
  try {
    showStatus("Abrindo janela de login...", "loading");
    const resp = await msalInstance.loginPopup({ scopes: CONFIG.SCOPES });
    const tokenResp = await msalInstance.acquireTokenSilent({
      scopes:  CONFIG.SCOPES,
      account: resp.account
    });
    accessToken = tokenResp.accessToken;
    await onLoggedIn();
  } catch (e) {
    showStatus("Erro no login: " + e.message, "error");
  }
}

async function onLoggedIn() {
  document.getElementById("loginSection").style.display = "none";
  document.getElementById("folderSection").style.display = "block";
  hideStatus();

  // Dados do email atual
  currentItem = Office.context.mailbox.item;
  document.getElementById("previewSubject").textContent =
    currentItem.subject || "(sem assunto)";
  document.getElementById("previewFrom").textContent =
    "De: " + (currentItem.from?.emailAddress || "—");

  await loadSiteAndDrive();
  await loadRootFolders();
}

// ── Graph API helpers ────────────────────────────────────────
async function graphGet(endpoint) {
  const res = await fetch("https://graph.microsoft.com/v1.0" + endpoint, {
    headers: { Authorization: "Bearer " + accessToken }
  });
  if (!res.ok) throw new Error(`Graph API ${res.status}: ${await res.text()}`);
  return res.json();
}

async function graphPut(endpoint, body, contentType = "application/octet-stream") {
  const res = await fetch("https://graph.microsoft.com/v1.0" + endpoint, {
    method:  "PUT",
    headers: {
      Authorization:  "Bearer " + accessToken,
      "Content-Type": contentType
    },
    body
  });
  if (!res.ok) throw new Error(`Graph PUT ${res.status}: ${await res.text()}`);
  return res.json();
}

// ── Resolução do site e drive ────────────────────────────────
async function loadSiteAndDrive() {
  showStatus("Conectando ao SharePoint...", "loading");
  try {
    // Resolve site pelo URL configurado
    const siteHost = new URL(CONFIG.SITE_URL).hostname;
    const sitePath = new URL(CONFIG.SITE_URL).pathname;
    const siteData = await graphGet(`/sites/${siteHost}:${sitePath}`);
    siteId = siteData.id;

    // Resolve o drive (biblioteca) pelo nome
    const drives = await graphGet(`/sites/${siteId}/drives`);
    const drive  = drives.value.find(
      d => d.name.toLowerCase() === CONFIG.LIBRARY_NAME.toLowerCase()
    ) || drives.value[0];
    driveId = drive.id;
    hideStatus();
  } catch (e) {
    showStatus("Erro ao conectar ao SharePoint: " + e.message, "error");
    throw e;
  }
}

// ── Pastas raiz ───────────────────────────────────────────────
async function loadRootFolders() {
  showStatus("Carregando pastas...", "loading");
  try {
    const data = await graphGet(
      `/drives/${driveId}/root/children?$filter=folder ne null&$select=id,name`
    );
    const select = document.getElementById("rootFolder");
    select.innerHTML = '<option value="">— selecione a pasta —</option>';
    data.value.forEach(f => {
      const opt = document.createElement("option");
      opt.value = f.id;
      opt.textContent = f.name;
      select.appendChild(opt);
    });
    hideStatus();
    document.getElementById("archiveBtn").style.display = "block";
  } catch (e) {
    showStatus("Erro ao carregar pastas: " + e.message, "error");
  }
}

// ── Subpastas ────────────────────────────────────────────────
async function loadSubfolders(folderId) {
  const subSection = document.getElementById("subfolderSection");
  const subSelect  = document.getElementById("subFolder");

  if (!folderId) {
    subSection.style.display = "none";
    return;
  }

  try {
    const data = await graphGet(
      `/drives/${driveId}/items/${folderId}/children?$filter=folder ne null&$select=id,name`
    );
    subSelect.innerHTML = '<option value="">— raiz da pasta acima —</option>';
    if (data.value.length > 0) {
      data.value.forEach(f => {
        const opt = document.createElement("option");
        opt.value = f.id;
        opt.textContent = f.name;
        subSelect.appendChild(opt);
      });
      subSection.style.display = "block";
    } else {
      subSection.style.display = "none";
    }
  } catch {
    subSection.style.display = "none";
  }
}

// ── Arquivamento ─────────────────────────────────────────────
async function archiveEmail() {
  const rootId = document.getElementById("rootFolder").value;
  const subId  = document.getElementById("subFolder").value;

  if (!rootId) {
    showStatus("Selecione uma pasta antes de arquivar.", "error");
    return;
  }

  const targetFolderId = subId || rootId;
  const btn = document.getElementById("archiveBtn");
  btn.innerHTML = '<span class="loader"></span>Arquivando...';
  btn.disabled  = true;
  showStatus("Obtendo conteúdo do email...", "loading");

  try {
    // 1. Obtém o EML via REST do Outlook (requer permissão Mail.Read)
    const itemId    = currentItem.itemId;
    const emlData   = await fetchEml(itemId);

    // 2. Nome do arquivo: data + assunto sanitizado
    const date      = new Date().toISOString().slice(0, 10);
    const subject   = (currentItem.subject || "sem-assunto")
      .replace(/[\\/:*?"<>|]/g, "_").substring(0, 80);
    const fileName  = `${date}_${subject}.eml`;

    // 3. Upload para o SharePoint
    showStatus("Enviando para o SharePoint...", "loading");
    await graphPut(
      `/drives/${driveId}/items/${targetFolderId}:/${encodeURIComponent(fileName)}:/content`,
      emlData
    );

    showStatus(`✅ Email arquivado como "${fileName}"`, "success");
    btn.innerHTML = "💾 Arquivar Email";
    btn.disabled  = false;

  } catch (e) {
    showStatus("❌ Erro: " + e.message, "error");
    btn.innerHTML = "💾 Arquivar Email";
    btn.disabled  = false;
  }
}

// ── Obtém o conteúdo do email via Office.js e Graph API ──────
async function fetchEml(itemId) {
  // Tenta primeiro via Graph API (mais confiável)
  try {
    // Converte o itemId REST para formato EWS se necessário
    const restId = Office.context.mailbox.convertToRestId
      ? Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0)
      : itemId;

    const res = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${restId}/$value`,
      { headers: { Authorization: "Bearer " + accessToken } }
    );
    if (res.ok) return await res.blob();
  } catch {}

  // Fallback: constrói EML a partir dos dados disponíveis via Office.js
  return await buildEmlFromOffice();
}

async function buildEmlFromOffice() {
  const item = Office.context.mailbox.item;

  const body = await new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded)
        resolve(result.value);
      else
        resolve("(corpo não disponível)");
    });
  });

  const from    = item.from ? item.from.emailAddress : "";
  const subject = item.subject || "(sem assunto)";
  const date    = new Date().toUTCString();

  const toList = (item.to || []).map(r => `${r.displayName} <${r.emailAddress}>`).join(", ");

  const eml = [
    `From: ${from}`,
    `To: ${toList}`,
    `Subject: ${subject}`,
    `Date: ${date}`,
    `MIME-Version: 1.0`,
    `Content-Type: text/html; charset=utf-8`,
    ``,
    body
  ].join("\r\n");

  return new Blob([eml], { type: "message/rfc822" });
}

async function getRestToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync(
      { isRest: true },
      result => {
        if (result.status === Office.AsyncResultStatus.Succeeded)
          resolve(result.value);
        else
          reject(new Error(result.error.message));
      }
    );
  });
}

// ── Helpers de UI ────────────────────────────────────────────
function showStatus(msg, type) {
  const el = document.getElementById("status");
  el.textContent  = msg;
  el.className    = type;
  el.style.display = "block";
}

function hideStatus() {
  const el = document.getElementById("status");
  el.style.display = "none";
}
