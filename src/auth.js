import { PublicClientApplication } from "@azure/msal-browser";

const CLIENT_ID = "5afc0f8c-2edf-4b08-87ff-02afd798c442";
const TENANT_ID = "cefdbb51-3f9b-42f6-9035-321cad67be3b";
const FILE_NAME = "Controle de estruturas - DFG.xlsx";

export const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
};

export const msalInstance = new PublicClientApplication(msalConfig);

export const loginRequest = {
  scopes: ["Files.Read", "User.Read"],
};

export async function getAccessToken() {
  await msalInstance.initialize();
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) throw new Error("Não autenticado");
  const response = await msalInstance.acquireTokenSilent({
    ...loginRequest,
    account: accounts[0],
  });
  return response.accessToken;
}

export async function fetchPlanilha() {
  const token = await getAccessToken();

  // Busca o arquivo pelo nome usando search
  const searchUrl = `https://graph.microsoft.com/v1.0/me/drive/search(q='${encodeURIComponent(FILE_NAME)}')`;
  const searchRes = await fetch(searchUrl, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const searchData = await searchRes.json();
  console.log("Resultado busca:", JSON.stringify(searchData?.value?.map(f => ({ name: f.name, id: f.id, path: f.parentReference?.path }))));

  const file = searchData?.value?.find(f => f.name === FILE_NAME);
  if (!file) throw new Error(`Arquivo "${FILE_NAME}" não encontrado no OneDrive`);

  console.log("Arquivo encontrado, id:", file.id);
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${file.id}/content`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Erro ao baixar arquivo: ${res.status} - ${errText}`);
  }
  return await res.arrayBuffer();
}
