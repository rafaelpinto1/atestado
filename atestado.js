const msalConfig = {
  auth: {
    clientId: "9cfa18cc-35a7-4b23-8616-e256aad79914",
    authority: "https://login.microsoftonline.com/62345b7a-94ed-4671-b8f2-624e28c8253a",
    redirectUri: window.location.origin + "/atestado.html",
  },
};

const loginScopes = ["User.Read"];
const graphScopes = ["Mail.Send", "User.Read"]; // Precisa do Mail.Send para enviar email

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

// Função para converter arquivo para Base64
function toBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      // Remove o prefixo "data:<mime>;base64," para deixar só o conteúdo base64
      const base64 = reader.result.split(',')[1];
      resolve(base64);
    };
    reader.onerror = error => reject(error);
  });
}

async function sendEmailWithAttachment(accessToken, file, funcionario, dataAusencia) {
  const base64File = await toBase64(file);

  const email = {
    message: {
      subject: `Atestado Médico - ${funcionario}`,
      body: {
        contentType: "Text",
        content: `Segue em anexo o atestado médico referente à ausência em ${dataAusencia}.`,
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
          contentBytes: base64File,
        },
      ],
    },
  };

  const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(email),
  });

  if (!response.ok) {
    const err = await response.json();
    throw new Error("Erro ao enviar email: " + (err.error?.message || response.statusText));
  }

  return true;
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();

  statusDiv.style.color = "black";
  statusDiv.textContent = "Enviando atestado por e-mail...";

  const funcionario = funcionarioInput.value;
  const dataAusencia = dataAusenciaInput.value;
  const fileInput = document.getElementById("fileAtestado");
  if (!fileInput.files.length) {
    alert("Por favor, escolha um arquivo.");
    return;
  }
  const file = fileInput.files[0];

  try {
    const tokenResponse = await msalInstance
      .acquireTokenSilent({ scopes: graphScopes })
      .catch(() => msalInstance.acquireTokenPopup({ scopes: graphScopes }));
    const accessToken = tokenResponse.accessToken;

    await sendEmailWithAttachment(accessToken, file, funcionario, dataAusencia);

    statusDiv.style.color = "green";
    statusDiv.textContent = "Atestado enviado por e-mail com sucesso!";
    form.reset();
    loginBtn.style.display = "block";
    form.classList.add("d-none");
  } catch (err) {
    statusDiv.style.color = "red";
    statusDiv.textContent = "Erro: " + err.message;
  }
});
