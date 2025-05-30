const msalConfig = {
  auth: {
    clientId: "9cfa18cc-35a7-4b23-8616-e256aad79914",
    authority: "https://login.microsoftonline.com/62345b7a-94ed-4671-b8f2-624e28c8253a",
    redirectUri: window.location.origin + "/atestado.html",
  },
};

const loginScopes = ["User.Read"];
const graphScopes = ["Mail.Send", "User.Read"];

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginBtn = document.getElementById("btnLogin");
const form = document.getElementById("atestadoForm");
const funcionarioInput = document.getElementById("funcionario");
const dataAusenciaInput = document.getElementById("dataAusencia");
const statusDiv = document.getElementById("status");

loginBtn.addEventListener("click", async () => {
  try {
    const loginResponse = await msalInstance.loginPopup({ scopes: loginScopes });
    const account = loginResponse.account;
    loginBtn.style.display = "none";
    form.classList.remove("d-none");
    funcionarioInput.value = account.name || "";
  } catch (err) {
    alert("Erro no login: " + err.message);
  }
});

// Função para converter arquivo em base64
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]); // remove o prefixo data:...
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  statusDiv.style.color = "black";
  statusDiv.textContent = "Enviando atestado por email...";

  const funcionario = funcionarioInput.value;
  const dataAusencia = dataAusenciaInput.value;
  const fileInput = document.getElementById("fileAtestado");
  if (!fileInput.files.length) {
    alert("Por favor, escolha um arquivo.");
    return;
  }
  const file = fileInput.files[0];

  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({ scopes: graphScopes })
      .catch(() => msalInstance.acquireTokenPopup({ scopes: graphScopes }));
    const accessToken = tokenResponse.accessToken;

    const base64File = await fileToBase64(file);

    // Montar o corpo do email
    const email = {
      message: {
        subject: `Atestado de ${funcionario} - ${new Date().toLocaleDateString()}`,
        body: {
          contentType: "Text",
          content: `Funcionário: ${funcionario}\nData da Ausência: ${dataAusencia}`,
        },
        toRecipients: [
          {
            emailAddress: {
              address: "informatica@gsilva.com.br",
            },
          },
        ],
        attachments: [
          {
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: file.name,
            contentType: file.type,
            contentBytes: base64File,
          },
        ],
      },
      saveToSentItems: "true",
    };

    // Enviar email
    const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(email),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error?.message || "Erro ao enviar email");
    }

    statusDiv.style.color = "green";
    statusDiv.textContent = "Atestado enviado por email com sucesso!";
    form.reset();
    loginBtn.style.display = "block";
    form.classList.add("d-none");
  } catch (err) {
    statusDiv.style.color = "red";
    statusDiv.textContent = "Erro: " + err.message;
  }
});
