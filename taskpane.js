// ============================================================
//  Outlook Add-in — Vshare
//  Autor: VtechIT  |  Versão: 1.2
//  Fix: SSO nativo do Office + localStorage para sessão persistente
// ============================================================

let msalInstance  = null;
let accessToken   = null;
let currentItem   = null;
let siteId        = null;
let driveId       = null;
let loggedAccount = null;

// ── Inicialização ────────────────────────────────────────────
Office.onReady(async () => {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:    CONFIG.CLIENT_ID,
      authority:   `https://login.microsoftonline.com/${CONFIG.TENANT_ID}`,
      redirectUri: CONFIG.REDIRECT_URI
    },
    // FIX: localStorage persiste entre sessões do painel.
    // sessionStorage é limpo toda vez que o painel fecha — causa o login repetido.
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
  });

  await tryAutoLogin();
});

// ── Estratégia de autenticação em 3 camadas ──────────────────
// 1ª: SSO nativo do Office (usa a sessão já logada no Outlook — sem popup)
// 2ª: Token silencioso via MSAL (conta em cache no localStorage)
// 3ª: Popup de login (único fallback com interação do usuário)

async function tryAutoLogin() {
  // Camada 1: SSO nativo do Office
  try {
    const ssoToken = await getOfficeSSOToken();
    if (ssoToken) {
      // Troca o token do Office por um token da Graph API via OBO (On-Behalf-Of)
      const graphToken = await exchangeTokenViaOBO(ssoToken);
      if (graphToken) {
        accessToken   = graphToken;
        loggedAccount = msalInstance.getAllAccounts()[0] || null;
        await onLoggedIn();
        return;
      }
    }
  } catch (e) {
    console.log("SSO Office não disponível, tentando MSAL cache:", e.message);
  }

  // Camada 2: Token silencioso via MSAL (localStorage)
  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      const resp = await msalInstance.acquireTokenSilent({
        scopes:  CONFIG.SCOPES,
        account: accounts[0]
      });
      accessToken   = resp.accessToken;
      loggedAccount = accounts[0];
      await onLoggedIn();
      return;
    }
  } catch (e) {
    console.log("Token silencioso falhou, exibindo botão de login:", e.message);
  }

  // Camada 3: Exibe botão de login (popup manual)
  showLoginSection();
}

// ── SSO nativo do Office ─────────────────────────────────────
function getOfficeSSOToken() {
  return new Promise((resolve, reject) => {
    if (!Office.context.auth || !Office.context.auth.getAccessTokenAsync) {
      return resolve(null); // API não disponível nesta versão do Office
    }
    Office.context.auth.getAccessTokenAsync(
      { allowSignInPrompt: false, allowConsentPrompt: false, forMSGraphAccess: true },
      result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          // Códigos 13000-13010 = SSO não configurado ou não suportado — não é erro crítico
          resolve(null);
        }
      }
    );
  });
}

// ── OBO: troca token do Office por token da Graph API ────────
// Necessário porque o token SSO do Office tem audience diferente da Graph API
async function exchangeTokenViaOBO(officeToken) {
  try {
    const res = await fetch(
      `https://login.microsoftonline.com/${CONFIG.TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type:         "urn:ietf:params:oauth:grant-type:jwt-bearer",
          client_id:          CONFIG.CLIENT_ID,
          assertion:          officeToken,
          scope:              CONFIG.SCOPES.join(" "),
          requested_token_use: "on_behalf_of"
        })
      }
    );
    const data = await res.json();
    return data.access_token || null;
  } catch {
    return null;
  }
}

function showLoginSection() {
  document.getElementById("loginSection").style.display = "block";
  document.getElementById("folderSection").style.display = "none";
}

// ── Login manual (popup) — último recurso ────────────────────
async function login() {
  try {
    showStatus("Abrindo janela de login...", "loading");
    const resp = await msalInstance.loginPopup({ scopes: CONFIG.SCOPES });
    const tokenResp = await msalInstance.acquireTokenSilent({
      scopes:  CONFIG.SCOPES,
      account: resp.account
    });
    accessToken   = tokenResp.accessToken;
    loggedAccount = resp.account;
    await onLoggedIn();
  } catch (e) {
    showStatus("Erro no login: " + e.message, "error");
  }
}

async function onLoggedIn() {
  document.getElementById("loginSection").style.display = "none";
  document.getElementById("folderSection").style.display = "block";
  hideStatus();

  // Mostrar user bar e botão logout
  const email = loggedAccount?.username || loggedAccount?.name || "—";
  document.getElementById("userEmail").textContent = email;
  const initials = email.substring(0, 2).toUpperCase();
  document.getElementById("userAvatar").textContent = initials;
  document.getElementById("userBar").style.display = "flex";
  document.getElementById("logoutBtn").style.display = "flex";

  currentItem = Office.context.mailbox.item;
  document.getElementById("previewSubject").textContent =
    currentItem.subject || "(sem assunto)";
  document.getElementById("previewFrom").textContent =
    "De: " + (currentItem.from?.emailAddress || "—");

  await loadSiteAndDrive();
  await loadRootFolders();
}

// ── Logout ───────────────────────────────────────────────────
function logout() {
  if (msalInstance) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalInstance.logoutPopup({ account: accounts[0] }).catch(() => {});
    }
  }
  accessToken   = null;
  loggedAccount = null;
  document.getElementById("userBar").style.display        = "none";
  document.getElementById("logoutBtn").style.display      = "none";
  document.getElementById("folderSection").style.display  = "none";
  document.getElementById("folderIndicator").style.display = "none";
  document.getElementById("loginSection").style.display   = "block";
  hideStatus();
}

// ── Renovação automática do token ────────────────────────────
// Tokens expiram em ~1h. Antes de cada chamada, verifica e renova silenciosamente.
async function getValidToken() {
  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      const resp = await msalInstance.acquireTokenSilent({
        scopes:  CONFIG.SCOPES,
        account: accounts[0]
      });
      accessToken = resp.accessToken;
    }
  } catch {
    // Token expirado e renovação silenciosa falhou — continua com o atual
    // (vai falhar na chamada e mostrar erro descritivo)
  }
  return accessToken;
}

// ── Graph API helpers ────────────────────────────────────────
async function graphGet(endpoint) {
  const token = await getValidToken();
  const url = endpoint.startsWith("https://")
    ? endpoint
    : "https://graph.microsoft.com/v1.0" + endpoint;
  const res = await fetch(url, {
    headers: { Authorization: "Bearer " + token }
  });
  if (!res.ok) throw new Error(`Graph API ${res.status}: ${await res.text()}`);
  return res.json();
}

// FIX: busca TODAS as páginas de resultados automaticamente.
// A Graph API retorna no máximo 200 itens por vez; pastas com muitos
// itens exigem seguir @odata.nextLink até acabar.
async function graphGetAll(endpoint) {
  let items = [];
  let nextUrl = endpoint;
  while (nextUrl) {
    const data = await graphGet(nextUrl);
    if (Array.isArray(data.value)) items = items.concat(data.value);
    nextUrl = data["@odata.nextLink"] || null;
  }
  return items;
}

async function graphPut(endpoint, body, contentType = "application/octet-stream") {
  const token = await getValidToken();
  const res = await fetch("https://graph.microsoft.com/v1.0" + endpoint, {
    method:  "PUT",
    headers: {
      Authorization:  "Bearer " + token,
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
async function loadRootFolders() {
  showStatus("Carregando pastas...", "loading");
  try {
    // graphGetAll segue @odata.nextLink automaticamente — lista TODAS as pastas
    const allItems = await graphGetAll(
      `/drives/${driveId}/root/children?$select=id,name,folder&$top=200`
    );
    const folders = allItems.filter(item => item.folder !== undefined)
      .sort((a, b) => a.name.localeCompare(b.name, "pt-BR"));

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
    document.getElementById("archiveBtn").style.display = "flex";
  } catch (e) {
    showStatus("Erro ao carregar pastas: " + e.message, "error");
  }
}

// ── Subpastas + indicador de pasta ──────────────────────────
function updateFolderIndicator() {
  const rootSel = document.getElementById("rootFolder");
  const subSel  = document.getElementById("subFolder");
  const rootName = rootSel.options[rootSel.selectedIndex]?.text || "";
  const subName  = subSel.options[subSel.selectedIndex]?.text || "";
  const indicator = document.getElementById("folderIndicator");

  if (!rootSel.value) {
    indicator.style.display = "none";
    return;
  }

  const path = subSel.value ? `${rootName} / ${subName}` : rootName;
  document.getElementById("folderPathText").textContent = path;
  indicator.style.display = "flex";
}

async function onRootFolderChange(folderId) {
  updateFolderIndicator();
  await loadSubfolders(folderId);
}

async function loadSubfolders(folderId) {
  const subSection = document.getElementById("subfolderSection");
  const subSelect  = document.getElementById("subFolder");

  if (!folderId) { subSection.style.display = "none"; return; }

  try {
    // graphGetAll segue @odata.nextLink — lista TODAS as subpastas
    const allItems = await graphGetAll(
      `/drives/${driveId}/items/${folderId}/children?$select=id,name,folder&$top=200`
    );
    const subFolders = allItems.filter(item => item.folder !== undefined)
      .sort((a, b) => a.name.localeCompare(b.name, "pt-BR"));

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
    updateFolderIndicator();
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

    const date     = getLocalDateString();
    const subject  = (currentItem.subject || "sem-assunto")
      .replace(/[\\/:*?"<>|]/g, "_").substring(0, 80);
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

// ── Data local (fuso Brasil) ─────────────────────────────────
function getLocalDateString() {
  const now  = new Date();
  const yyyy = now.getFullYear();
  const mm   = String(now.getMonth() + 1).padStart(2, "0");
  const dd   = String(now.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

// ── Obtém o blob do email ─────────────────────────────────────
async function fetchEmailBlob() {
  try {
    const restId = Office.context.mailbox.convertToRestId
      ? Office.context.mailbox.convertToRestId(
          currentItem.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        )
      : currentItem.itemId;

    const token = await getValidToken();
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${restId}/$value`,
      { headers: { Authorization: "Bearer " + token } }
    );
    if (res.ok) return { blob: await res.blob(), extension: "eml" };
  } catch (err) {
    console.warn("Graph $value falhou, usando fallback:", err);
  }
  return await buildFallbackEmail();
}

// ── Fallback EML via Office.js ───────────────────────────────
async function buildFallbackEmail() {
  const item = Office.context.mailbox.item;
  const body = await new Promise(resolve => {
    item.body.getAsync(Office.CoercionType.Html, result => {
      resolve(
        result.status === Office.AsyncResultStatus.Succeeded
          ? result.value : "(corpo não disponível)"
      );
    });
  });

  const from    = item.from ? `${item.from.displayName} <${item.from.emailAddress}>` : "";
  const subject = item.subject || "(sem assunto)";
  const date    = new Date().toUTCString();
  const toList  = (item.to || []).map(r => `${r.displayName} <${r.emailAddress}>`).join(", ");

  if (typeof CFB !== "undefined") {
    try {
      return { blob: buildMsgFile({ from, subject, date, toList, body }), extension: "msg" };
    } catch (e) { console.warn("CFB falhou:", e); }
  }

  const eml = [`From: ${from}`, `To: ${toList}`, `Subject: ${subject}`,
    `Date: ${date}`, `MIME-Version: 1.0`, `Content-Type: text/html; charset=utf-8`, ``, body
  ].join("\r\n");

  return { blob: new Blob([eml], { type: "message/rfc822" }), extension: "eml" };
}

// ── Construtor .msg via CFB ───────────────────────────────────
function buildMsgFile({ from, subject, date, toList, body }) {
  function encodeUTF16LE(str) {
    const buf  = new ArrayBuffer((str.length + 1) * 2);
    const view = new Uint16Array(buf);
    for (let i = 0; i < str.length; i++) view[i] = str.charCodeAt(i);
    view[str.length] = 0;
    return new Uint8Array(buf);
  }

  const cfb = CFB.utils.cfb_new();
  CFB.utils.cfb_add(cfb, "__nameid_version1.0", new Uint8Array(0));

  const stringProps = [
    { tag: 0x0037, type: 0x001F, value: subject },
    { tag: 0x1013, type: 0x001F, value: body },
    { tag: 0x0C1A, type: 0x001F, value: from },
    { tag: 0x0E04, type: 0x001F, value: toList },
    { tag: 0x0070, type: 0x001F, value: subject },
    { tag: 0x0039, type: 0x001F, value: date },
  ];

  const fixedProps = [
    { tag: 0x0017, type: 0x0003, value: 0x00000001 },
    { tag: 0x0026, type: 0x0003, value: 0x00000000 },
  ];

  const headerBuf  = new ArrayBuffer(32 + fixedProps.length * 16);
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
    const name = `__substg1.0_${p.tag.toString(16).toUpperCase().padStart(4,"0")}${p.type.toString(16).toUpperCase().padStart(4,"0")}`;
    CFB.utils.cfb_add(cfb, name, encodeUTF16LE(p.value || ""));
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
