import { PublicClientApplication } from "@azure/msal-browser";

const CLIENT_ID = "5afc0f8c-2edf-4b08-87ff-02afd798c442";
const TENANT_ID = "cefdbb51-3f9b-42f6-9035-321cad67be3b";
const DRIVE_ID = "b!KYjiZUP18kuYhFviktQw722rXRO99dFAhHinzGeP281GkMu_gx-LQ6hCCacE2WZ8";
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
  const encodedPath = ONEDRIVE_PATH.split("/").map(s => encodeURIComponent(s)).join("/");
  const url = `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:/${encodedPath}:/content`;
  console.log("Tentando URL:", url);
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  console.log("Status:", res.status);
  if (!res.ok) {
    const errText = await res.text();
    console.log("Erro:", errText);
    throw new Error(`Erro
