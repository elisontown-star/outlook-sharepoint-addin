// ============================================================
//  CONFIGURAÇÃO — preencha com os dados do seu ambiente
// ============================================================

const CONFIG = {
  // Azure AD → App registrations → seu app
  CLIENT_ID: "e5a8fd64-798c-4040-9caf-8797976955f4",
  TENANT_ID: "51873eff-cfd8-4715-839f-10a25cdcbec9",

  // URL do site SharePoint (sem barra no final)
  SITE_URL: "https://vtecit.sharepoint.com/sites/groupfinanceiro",

  // Nome da biblioteca de documentos onde os emails serão salvos
  // (geralmente "Documentos" ou "Documents" ou nome personalizado)
  LIBRARY_NAME: "Documentos",

  // Escopos necessários para a Graph API
  SCOPES: [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Files.ReadWrite.All",
    "https://graph.microsoft.com/Sites.ReadWrite.All",
    "https://graph.microsoft.com/offline_access"
  ]
};

