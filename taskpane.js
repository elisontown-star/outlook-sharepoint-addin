// ============================================================
//  Outlook Add-in — Arquivar Email no SharePoint
//  Autor: VtechIT  |  Versão: 1.1
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
      clientId:    CONFIG.CLIENT_ID,
      authority:   `https://login.microsoftonline.com/${CONFIG.TENANT_ID}`,
      redirectUri: CONFIG.REDIRECT_URI   // ← usa config.js, não hardcoded
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
    const siteHost = new URL(CONFIG.SITE_URL).hostname;
    const sitePath = new URL(CONFIG.SITE_URL).pathname;
    const siteData = await graphGet(`/sites/${siteHost}:${sitePath}`);
    siteId = siteData.id;

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
// FIX: $filter=folder ne null não é suportado pela Graph API em drives.
//      Buscamos todos os filhos e filtramos no cliente.
async function loadRootFolders() {
  showStatus("Carregando pastas...", "loading");
  try {
    const data = await graphGet(
      `/drives/${driveId}/root/children?$select=id,name,folder`
    );

    // Filtra apenas itens que são pastas (possuem a propriedade "folder")
    const folders = data.value.filter(item => item.folder !== undefined);

    const select = document.getElementById("rootFolder");
    select.innerHTML = '<option value="">— selecione a pasta —</option>';

    if (folders.length === 0) {
      select.innerHTML = '<option value="">Nenhuma pasta encontrada</option>';
      showStatus("Nenhuma pasta encontrada na biblioteca.", "error");
      return;
    }

    folders.forEach(f => {
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
      `/drives/${driveId}/items/${folderId}/children?$select=id,name,folder`
    );

    // Filtra apenas pastas no cliente
    const subFolders = data.value.filter(item => item.folder !== undefined);

    subSelect.innerHTML = '<option value="">— raiz da pasta acima —</option>';
    if (subFolders.length > 0) {
      subFolders.forEach(f => {
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
    const { blob, extension } = await fetchEmailBlob();

    // Nome: data local (BR) + assunto sanitizado + extensão correta
    const date     = getLocalDateString();
    const subject  = (currentItem.subject || "sem-assunto")
      .replace(/[\\/:*?"<>|]/g, "_")
      .substring(0, 80);
    const fileName = `${date}_${subject}.${extension}`;

    showStatus("Enviando para o SharePoint...", "loading");
    await graphPut(
      `/drives/${driveId}/items/${targetFolderId}:/${encodeURIComponent(fileName)}:/content`,
      blob
    );

    showStatus(`✅ Email arquivado como "${fileName}"`, "success");
  } catch (e) {
    showStatus("❌ Erro: " + e.message, "error");
  } finally {
    btn.innerHTML = "💾 Arquivar Email";
    btn.disabled  = false;
  }
}

// ── Data local no fuso do Brasil ─────────────────────────────
// FIX: toISOString() retorna UTC; usamos toLocaleDateString para UTC-3
function getLocalDateString() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm   = String(now.getMonth() + 1).padStart(2, "0");
  const dd   = String(now.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

// ── Obtém o conteúdo do email ─────────────────────────────────
// Retorna { blob, extension } para que o nome do arquivo seja consistente
async function fetchEmailBlob() {
  // Tenta primeiro via Graph API ($value retorna o EML bruto)
  try {
    const restId = Office.context.mailbox.convertToRestId
      ? Office.context.mailbox.convertToRestId(
          currentItem.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        )
      : currentItem.itemId;

    const res = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${restId}/$value`,
      { headers: { Authorization: "Bearer " + accessToken } }
    );

    if (res.ok) {
      // Graph API retorna EML (formato RFC-822) → extensão .eml
      return { blob: await res.blob(), extension: "eml" };
    }
  } catch (err) {
    console.warn("Graph $value falhou, usando fallback:", err);
  }

  // Fallback: constrói o arquivo a partir dos dados do Office.js
  return await buildFallbackEmail();
}

// ── Fallback: constrói arquivo de email via Office.js ─────────
async function buildFallbackEmail() {
  const item = Office.context.mailbox.item;

  const body = await new Promise(resolve => {
    item.body.getAsync(Office.CoercionType.Html, result => {
      resolve(
        result.status === Office.AsyncResultStatus.Succeeded
          ? result.value
          : "(corpo não disponível)"
      );
    });
  });

  const from    = item.from ? `${item.from.displayName} <${item.from.emailAddress}>` : "";
  const subject = item.subject || "(sem assunto)";
  const date    = new Date().toUTCString();
  const toList  = (item.to || [])
    .map(r => `${r.displayName} <${r.emailAddress}>`)
    .join(", ");

  // Tenta criar .msg usando CFB se disponível
  if (typeof CFB !== "undefined") {
    try {
      const blob = buildMsgFile({ from, subject, date, toList, body });
      return { blob, extension: "msg" };
    } catch (e) {
      console.warn("CFB falhou, usando EML simples:", e);
    }
  }

  // Último recurso: EML puro (RFC-822)
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

  return {
    blob: new Blob([eml], { type: "message/rfc822" }),
    extension: "eml"   // FIX: era ".msg" mesmo sendo EML
  };
}

// ── Construtor de .msg via CFB ────────────────────────────────
// FIX: adicionado stream __nameid_version1.0 obrigatório no formato .msg
function buildMsgFile({ from, subject, date, toList, body }) {
  function encodeUTF16LE(str) {
    const buf  = new ArrayBuffer((str.length + 1) * 2);
    const view = new Uint16Array(buf);
    for (let i = 0; i < str.length; i++) view[i] = str.charCodeAt(i);
    view[str.length] = 0;
    return new Uint8Array(buf);
  }

  const cfb = CFB.utils.cfb_new();

  // Stream obrigatório — sem ele leitores MAPI recusam o arquivo
  CFB.utils.cfb_add(cfb, "__nameid_version1.0", new Uint8Array(0));

  const stringProps = [
    { tag: 0x0037, type: 0x001F, value: subject },    // PR_SUBJECT
    { tag: 0x1013, type: 0x001F, value: body },        // PR_BODY_HTML
    { tag: 0x0C1A, type: 0x001F, value: from },        // PR_SENDER_NAME
    { tag: 0x0E04, type: 0x001F, value: toList },      // PR_DISPLAY_TO
    { tag: 0x0070, type: 0x001F, value: subject },     // PR_CONVERSATION_TOPIC
    { tag: 0x0039, type: 0x001F, value: date },        // PR_CLIENT_SUBMIT_TIME (texto)
  ];

  const fixedProps = [
    { tag: 0x0017, type: 0x0003, value: 0x00000001 }, // PR_IMPORTANCE = high
    { tag: 0x0026, type: 0x0003, value: 0x00000000 }, // PR_SENSITIVITY = normal
  ];

  const headerSize = 32 + fixedProps.length * 16;
  const headerBuf  = new ArrayBuffer(headerSize);
  const headerView = new DataView(headerBuf);

  let offset = 32;
  fixedProps.forEach(p => {
    headerView.setUint16(offset,      p.type,  true);
    headerView.setUint16(offset + 2,  p.tag,   true);
    headerView.setUint32(offset + 4,  0,       true);
    headerView.setUint32(offset + 8,  p.value, true);
    headerView.setUint32(offset + 12, 0,       true);
    offset += 16;
  });

  CFB.utils.cfb_add(cfb, "__properties_version1.0", new Uint8Array(headerBuf));

  stringProps.forEach(p => {
    const encoded = encodeUTF16LE(p.value || "");
    const name = `__substg1.0_${p.tag.toString(16).toUpperCase().padStart(4, "0")}${p.type.toString(16).toUpperCase().padStart(4, "0")}`;
    CFB.utils.cfb_add(cfb, name, encoded);
  });

  const out = CFB.write(cfb, { type: "array" });
  return new Blob([new Uint8Array(out)], { type: "application/vnd.ms-outlook" });
}

// ── Helpers de UI ────────────────────────────────────────────
function showStatus(msg, type) {
  const el = document.getElementById("status");
  el.textContent   = msg;
  el.className     = type;
  el.style.display = "block";
}

function hideStatus() {
  document.getElementById("status").style.display = "none";
}
