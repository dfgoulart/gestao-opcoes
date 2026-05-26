import { PublicClientApplication } from "@azure/msal-browser";

const CLIENT_ID = "5afc0f8c-2edf-4b08-87ff-02afd798c442";
const TENANT_ID = "cefdbb51-3f9b-42f6-9035-321cad67be3b";
const ONEDRIVE_PATH = "Negócios/Opções/Carteira Daniel/Controle de estruturas - DFG.xlsx";

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

export async function fetchPlanilha() {
  const token = await getAccessToken();
  const encodedPath = ONEDRIVE_PATH.split("/").map(s => encodeURIComponent(s)).join("/");
  
  // Tenta primeiro o drive pessoal, depois o drive corporativo
  const urls = [
    `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/content`,
    `https://graph.microsoft.com/v1.0/me/drives`,
  ];

  // Busca a lista de drives disponíveis
  const drivesRes = await fetch(urls[1], {
    headers: { Authorization: `Bearer ${token}` },
  });
  const drivesData = await drivesRes.json();
  
  // Tenta cada drive até encontrar o arquivo
  const drives = [{ id: "me" }, ...(drivesData.value || [])];
  for (const drive of drives) {
    const url = drive.id === "me"
      ? `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/content`
      : `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${encodedPath}:/content`;
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.ok) return await res.arrayBuffer();
  }
  throw new Error("Arquivo não encontrado em nenhum drive disponível");
}
