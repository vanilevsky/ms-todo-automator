import { OAuth } from "@raycast/api";
import fetch from "node-fetch";

// Create an OAuth client ID via https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
// There is a tutorial here https://docs.microsoft.com/en-us/azure/active-directory/develop/tutorial-v2-javascript-auth-code
// const clientId = "f02a9afa-795c-428d-92f6-b5b4a426ced4"; // todo-automator-native
const clientId = "383a3ddf-293a-4d99-8b8d-8acb6b004ed5"; // todo-automator
// const clientSecret = "IBd8Q~k~68x1xWRtnABB1s~pcbQkJEhChaAiMaaz";

const client = new OAuth.PKCEClient({
  redirectMethod: OAuth.RedirectMethod.Web,
  providerName: "Microsoft",
  providerIcon: "microsoft-logo.png",
  providerId: "microsoft",
  description: "Connect your Microsoft account",
});

// Authorization

export async function authorize(): Promise<void> {
  const tokenSet = await client.getTokens();
  if (tokenSet?.accessToken) {
    if (tokenSet.refreshToken && tokenSet.isExpired()) {
      await client.setTokens(await refreshTokens(tokenSet.refreshToken));
    }
    return;
  }

  const authRequest = await client.authorizationRequest({
    endpoint: "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize",
    clientId: clientId,
    scope:
      "User.Read,Tasks.Read,Tasks.Read.Shared,Tasks.ReadWrite,Tasks.ReadWrite.Shared,offline_access",
    
  });
  const { authorizationCode } = await client.authorize(authRequest);
  await client.setTokens(await fetchTokens(authRequest, authorizationCode));
}

async function fetchTokens(authRequest: OAuth.AuthorizationRequest, authCode: string): Promise<OAuth.TokenResponse> {
  const params = new URLSearchParams();
  params.append("client_id", clientId);
  // params.append("client_secret", clientSecret);
  params.append("code", authCode);
  params.append("code_verifier", authRequest.codeVerifier);
  params.append("grant_type", "authorization_code");
  params.append("redirect_uri", authRequest.redirectURI);

  const response = await fetch("https://login.microsoftonline.com/consumers/oauth2/v2.0/token", {
    method: "POST",
    body: params,
    headers: { "Origin": "*" }
  });
  
  if (!response.ok) {
    console.error("fetch tokens error:", await response.text());
    throw new Error(response.statusText);
  }
  return (await response.json()) as OAuth.TokenResponse;
}

async function refreshTokens(refreshToken: string): Promise<OAuth.TokenResponse> {
  const params = new URLSearchParams();
  params.append("client_id", clientId);
  // params.append("client_secret", clientSecret);
  params.append("refresh_token", refreshToken);
  params.append("grant_type", "refresh_token");

  const response = await fetch("https://login.microsoftonline.com/consumers/oauth2/v2.0/token", {
    method: "POST",
    body: params,
    headers: { "Origin": "*" }
  });
  if (!response.ok) {
    console.error("refresh tokens error:", await response.text());
    throw new Error(response.statusText);
  }
  const tokenResponse = (await response.json()) as OAuth.TokenResponse;
  tokenResponse.refresh_token = tokenResponse.refresh_token ?? refreshToken;
  return tokenResponse;
}

// API

export async function fetchItems(): Promise<{ id: string; title: string }[]> {

  const params = new URLSearchParams();
  // params.append("q", "trashed = false");
  // params.append("fields", "files(id, name, mimeType, iconLink, modifiedTime, webViewLink, webContentLink, size)");
  // params.append("orderBy", "recency desc");
  // params.append("pageSize", "100");

  const response = await fetch("https://graph.microsoft.com/v1.0/me/planner/tasks?" + params.toString(), {
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${(await client.getTokens())?.accessToken}`,
    },
  });

  console.log("fetch items response:", response.body);

  if (!response.ok) {
    console.error("fetch items error:", await response.text());
    throw new Error(response.statusText);
  }

  console.log(await response.json())

  const json = (await response.json()) as { files: { id: string; name: string }[] };

  return json.files.map((item) => ({ id: item.id, title: item.name }));

}
