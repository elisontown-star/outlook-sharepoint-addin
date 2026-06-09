// ============================================================
//  Vshare — commands.js
//  Intercepta o envio de emails (OnMessageSend) e abre o diálogo
//  de decisão: arquivar e enviar, apenas enviar, ou cancelar.
// ============================================================

let sendEvent = null;   // referência ao evento de envio para liberar/bloquear depois
let dialog    = null;   // referência ao diálogo aberto

Office.onReady(() => {
  // Runtime de eventos pronto. Nada a fazer aqui — a função
  // validateBeforeSend é chamada pelo Office quando o usuário envia.
});

// ── Função chamada pelo Office ao clicar em "Enviar" ─────────
// Declarada no manifest como o handler do evento OnMessageSend.
function validateBeforeSend(event) {
  sendEvent = event;

  // Abre o diálogo de decisão (Office Dialog API).
  // O envio fica BLOQUEADO até o diálogo retornar uma decisão.
  const url = "https://elisontown-star.github.io/outlook-sharepoint-addin/dialog.html";

  Office.context.ui.displayDialogAsync(
    url,
    { height: 60, width: 32, displayInIframe: false },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Se o diálogo não abrir, libera o envio para não travar o usuário
        console.error("[Vshare] Falha ao abrir diálogo:", result.error.message);
        event.completed({ allowEvent: true });
        return;
      }

      dialog = result.value;

      // Recebe mensagens do diálogo (a decisão do usuário)
      dialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        onDialogMessage
      );

      // Se o usuário fechar o diálogo no X, trata como cancelamento do envio
      dialog.addEventHandler(
        Office.EventType.DialogEventReceived,
        () => {
          // Diálogo fechado sem decisão → cancela o envio
          if (sendEvent) {
            sendEvent.completed({
              allowEvent: false,
              errorMessage: "Envio cancelado. Escolha uma opção no Vshare."
            });
          }
        }
      );
    }
  );
}

// ── Recebe a decisão tomada no diálogo ───────────────────────
function onDialogMessage(arg) {
  let msg;
  try {
    msg = JSON.parse(arg.message);
  } catch {
    msg = { action: "send" };
  }

  if (dialog) { dialog.close(); dialog = null; }

  switch (msg.action) {
    case "archive_and_send":
      // O arquivamento já foi feito DENTRO do diálogo (que tem o token MSAL).
      // Aqui apenas liberamos o envio.
      if (msg.success) {
        sendEvent.completed({ allowEvent: true });
      } else {
        // Se o arquivamento falhou, bloqueia o envio e avisa
        sendEvent.completed({
          allowEvent: false,
          errorMessage: "Falha ao arquivar no SharePoint: " + (msg.error || "erro desconhecido")
        });
      }
      break;

    case "send_only":
      // Apenas envia, sem arquivar
      sendEvent.completed({ allowEvent: true });
      break;

    case "cancel":
    default:
      // Cancela o envio — o usuário volta a editar o email
      sendEvent.completed({
        allowEvent: false,
        errorMessage: "Envio cancelado pelo usuário."
      });
      break;
  }
}

// Expor globalmente para o Office encontrar o handler
if (typeof window !== "undefined") {
  window.validateBeforeSend = validateBeforeSend;
}

// Registrar a ação para o runtime de eventos (exigido pelo Office)
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("validateBeforeSend", validateBeforeSend);
}
