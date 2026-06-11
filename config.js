// ============================================================
//  CONFIGURAÇÃO — preencha com os dados do seu ambiente
// ============================================================

const CONFIG = {
  // Azure AD → App registrations → seu app
  CLIENT_ID: "e5a8fd64-798c-4040-9caf-8797976955f4",
  TENANT_ID: "51873eff-cfd8-4715-839f-10a25cdcbec9",

  // URI de redirecionamento registrada no Azure AD
  REDIRECT_URI: "https://elisontown-star.github.io/outlook-sharepoint-addin/taskpane.html",

  // URL do site SharePoint (sem barra no final) — mantido para compatibilidade
  SITE_URL: "https://vtecit.sharepoint.com/sites/armazemjuridico",

  // Nome da biblioteca de documentos onde os emails serão salvos — mantido para compatibilidade
  LIBRARY_NAME: "Documentos",

  // ── Múltiplas bibliotecas ────────────────────────────────────
  // Adicione ou remova entradas conforme necessário.
  // "label" aparece no dropdown; "siteUrl" e "libraryName" definem o destino.
  SITES: [
    {
      label:       "Armazém Jurídico",
      siteUrl:     "https://vtecit.sharepoint.com/sites/armazemjuridico",
      libraryName: "Documentos"
    },
    {
      label:       "Group Financeiro",
      siteUrl:     "https://vtecit.sharepoint.com/sites/groupfinanceiro",
      libraryName: "Documentos"
    }
  ],

  // Escopos necessários para a Graph API
  // "openid" e "profile" são obrigatórios para o SSO nativo do Office funcionar
  // "offline_access" NÃO deve ter o prefixo de URL
  SCOPES: [
    "openid",
    "profile",
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Files.ReadWrite.All",
    "https://graph.microsoft.com/Sites.ReadWrite.All",
    "offline_access"
  ]
};
