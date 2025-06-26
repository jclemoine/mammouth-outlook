Office.onReady(() => {
  document.getElementById("generateBtn").onclick = generateMail;
});

const API_KEY = "INSERE_TA_CLE_API_ICI"; // Remplace par ta clé Mammouth

async function generateMail() {
  try {
    const item = Office.context.mailbox.item;
    let text = "";

    // Essaye de récupérer le texte sélectionné, sinon tout le corps
    if (item.getSelectedDataAsync) {
      item.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
          text = result.value;
          callMammouthAPI(text);
        } else {
          // Si pas de sélection, récupère le corps complet
          item.body.getAsync("text", (bodyResult) => {
            if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
              text = bodyResult.value;
              callMammouthAPI(text);
            } else {
              showStatus("Erreur récupération du corps du mail.");
            }
          });
        }
      });
    } else {
      showStatus("Fonction getSelectedDataAsync non supportée.");
    }
  } catch (e) {
    showStatus("Erreur : " + e.message);
  }
}

async function callMammouthAPI(promptText) {
  showStatus("Appel à Mammouth en cours...");

  // Exemple d’appel API (à adapter selon ta doc Mammouth)
  try {
    const response = await fetch("https://api.mammouth.ai/generate", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${API_KEY}`,
      },
      body: JSON.stringify({
        prompt: `Tu es assistant dans un cabinet d'assurance. Génére un mail pro clair et courtois à partir du texte : ${promptText}`,
        max_tokens: 400,
      }),
    });

    if (!response.ok) {
      throw new Error("Erreur API : " + response.status);
    }

    const data = await response.json();
    if (data && data.generated_text) {
      insertTextInBody(data.generated_text);
      showStatus("Mail généré !");
    } else {
      showStatus("Réponse invalide de Mammouth.");
    }
  } catch (err) {
    showStatus("Erreur lors de l’appel API : " + err.message);
  }
}

function insertTextInBody(text) {
  Office.context.mailbox.item.body.setSelectedDataAsync(
    text,
    { coercionType: Office.CoercionType.Html },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        showStatus("Texte inséré dans le mail.");
      } else {
        showStatus("Erreur insertion texte.");
      }
    }
  );
}

function showStatus(msg) {
  document.getElementById("status").innerText = msg;
}
