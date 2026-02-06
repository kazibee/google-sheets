import { google } from 'googleapis';

/**
 * Creates an authenticated OAuth2 client using credentials from process.env.
 * The sandbox worker injects CLIENT_ID, CLIENT_SECRET, and REFRESH_TOKEN
 * into process.env before loading the tool module.
 */
export function createAuthClient() {
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const refreshToken = process.env.REFRESH_TOKEN;

  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error(
      'Missing required credentials. Set CLIENT_ID, CLIENT_SECRET, and REFRESH_TOKEN via workerbee tool env.',
    );
  }

  const auth = new google.auth.OAuth2(clientId, clientSecret);
  auth.setCredentials({ refresh_token: refreshToken });

  return auth;
}
