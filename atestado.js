const msalConfig = {
  auth: {
    clientId: "9cfa18cc-35a7-4b23-8616-e256aad79914",
    authority: "https://login.microsoftonline.com/62345b7a-94ed-4671-b8f2-624e28c8253a",
    redirectUri: window.location.origin + "/atestado.html",
  },
};

const loginScopes = ["User.Read"];
const graphScopes = ["Sites.ReadWrite.All"];

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

async function getSiteId(accessToken) {
  const response = await fetch(
    "https://graph.microsoft.com/v1.0/sites/gsilvainfo.sharepoint.com:/sites/Adm",
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!response.ok) throw new Error("Erro ao buscar site: " + response.statusText);
  const data = await response.json();
  return data.id;
}

async function getListId(siteId, accessToken) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!response.ok) throw new Error("Erro ao buscar listas: " + response.statusText);
  const data = await response.json();
  const list = data.value.find((l) => l.name === "Atestados");
  if (!list) throw new Error("Lista 'Atestados' não encontrada.");
  return list.id;
}

async function uploadFileToSharePoint(siteId, listId, accessToken, file, funcionario, dataAusencia) {
  // Cria o item na lista com os campos
  const itemFields = {
    Title: `Atestado de ${funcionario} - ${new Date().toLocaleDateString()}`,
    Funcionario: funcionario,
    DataAusencia: dataAusencia,
  };

  const createItemResponse = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ fields: itemFields }),
    }
  );

  if (!createItemResponse.ok) {
    const err = await createItemResponse.json();
    throw new Error("Erro ao criar item: " + (err.error?.message || createItemResponse.statusText));
  }

  const createdItem = await createItemResponse.json();
  const itemId = createdItem.id;

  // Pega o driveId do site para upload do arquivo
  const driveResponse = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );
  if (!driveResponse.ok) throw new Error("Erro ao buscar drive: " + driveResponse.statusText);
  const driveData = await driveResponse.json();
  const driveId = driveData.id;

  // Upload do arquivo na raiz da biblioteca Documents
  const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${file.name}:/content`;

  const uploadResponse = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": file.type,
    },
    body: file,
  });

  if (!uploadResponse.ok) {
    const err = await uploadResponse.json();
    throw new Error("Erro ao enviar arquivo: " + (err.error?.message || uploadResponse.statusText));
  }

  const uploadedFile = await uploadResponse.json();

  // Atualiza o item da lista com o link do arquivo no campo Atestado
  const updateResponse = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
    {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ Atestado: uploadedFile.webUrl }),
    }
  );

  if (!updateResponse.ok) {
    const err = await updateResponse.json();
    throw new Error("Erro ao atualizar item: " + (err.error?.message || updateResponse.statusText));
  }

  return true;
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();

  statusDiv.style.color = "black";
  statusDiv.textContent = "Enviando atestado...";

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

    const siteId = await getSiteId(accessToken);
    const listId = await getListId(siteId, accessToken);

    await uploadFileToSharePoint(siteId, listId, accessToken, file, funcionario, dataAusencia);

    statusDiv.style.color = "green";
    statusDiv.textContent = "Atestado enviado com sucesso!";
    form.reset();
    loginBtn.style.display = "block";
    form.classList.add("d-none");
  } catch (err) {
    statusDiv.style.color = "red";
    statusDiv.textContent = "Erro: " + err.message;
  }
});
