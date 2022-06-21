import { getPreferenceValues, OAuth } from "@raycast/api";
import fetch from "node-fetch";
import { CreateTaskForm, msApiBaseUrl, TaskListItem } from "../const";
import { URLSearchParams } from "url";
import { assign } from "lodash";

const preferences = getPreferenceValues();

// Create an OAuth client ID via https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
// There is a tutorial here https://docs.microsoft.com/en-us/azure/active-directory/develop/tutorial-v2-javascript-auth-code
const clientId = preferences.clientId;
const clientSecret = preferences.clientSecret;

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
        try {
          await client.setTokens(await refreshTokens(tokenSet.refreshToken));
          return;
        } catch (error) {
          console.log("Refresh token failed: ", error);
        }
      }
      return;
  }

  const authRequest = await client.authorizationRequest({
    clientId: clientId,
    endpoint: "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize",
    scope: "User.Read,Tasks.Read,Tasks.Read.Shared,Tasks.ReadWrite,Tasks.ReadWrite.Shared,offline_access"
  });
  const { authorizationCode } = await client.authorize(authRequest);
  await client.setTokens(await fetchTokens(authRequest, authorizationCode));
}

async function fetchTokens(authRequest: OAuth.AuthorizationRequest, authCode: string): Promise<OAuth.TokenResponse> {
  const params = new URLSearchParams();
  params.append("client_id", clientId);
  params.append("client_secret", clientSecret);
  params.append("code", authCode);
  params.append("code_verifier", authRequest.codeVerifier);
  params.append("grant_type", "authorization_code");
  params.append("redirect_uri", authRequest.redirectURI);

  const response = await fetch("https://login.microsoftonline.com/consumers/oauth2/v2.0/token", {
    method: "POST",
    body: params,
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
  params.append("client_secret", clientSecret);
  params.append("refresh_token", refreshToken);
  params.append("grant_type", "refresh_token");

  const response = await fetch("https://login.microsoftonline.com/consumers/oauth2/v2.0/token", {
    method: "POST",
    body: params,
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

export async function fetchLists(): Promise<TaskListItem[]> {

  const response = await fetch(`${msApiBaseUrl}/me/todo/lists`, {
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${(await client.getTokens())?.accessToken}`
    }
  });

  if (!response.ok) {
    console.error("fetch items error:", await response.text());
    throw new Error(response.statusText);
  }

  const json = (await response.json()) as { value: TaskListItem[] };

  return json.value.map((item) => ({
    id: item.id,
    displayName: item.displayName,
    wellknownListName: item.wellknownListName,
    isOwner: item.isOwner,
    isShared: item.isShared,
  }));
}

export async function createTask(values: CreateTaskForm): Promise<void> {

  const todoTaskListId = values.listId;

  if (!values.title) {
    throw new Error("Title is required");
  }

  if (!todoTaskListId) {
    throw new Error("No list selected");
  }

  let taskBody = {
    title: values.title,
  }

  if (values.body !== '') {
    Object.assign(taskBody, {
      body: {
        content: values.body,
        contentType: "html",
      },
    });
  }

  if (values.dueDateTime !== null) {
    Object.assign(taskBody, {
      dueDateTime: {
        dateTime: values.dueDateTime,
        timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
      }
    });
  }

  if (values.reminderDateTime !== null) {
    taskBody = assign(taskBody, {
      reminderDateTime: {
        dateTime: values.reminderDateTime,
        timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
      }
    });
  }

  const response = await fetch(`${msApiBaseUrl}/me/todo/lists/${todoTaskListId}/tasks`, {
    method: 'POST',
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${(await client.getTokens())?.accessToken}`,
    },
    body: JSON.stringify(taskBody)
  });

  if (!response.ok) {
    console.error("fetch items error:", await response.text());
    throw new Error(response.statusText);
  }

  console.log("Task created");
  console.log("values:", values);
  console.log("taskBody:", taskBody);
}