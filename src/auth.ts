import { google } from 'googleapis';

export interface Env {
  CLIENT_ID: string;
  CLIENT_SECRET: string;
  REFRESH_TOKEN: string;
}

export function createAuthClient(env: Env) {
  const auth = new google.auth.OAuth2(env.CLIENT_ID, env.CLIENT_SECRET);
  auth.setCredentials({ refresh_token: env.REFRESH_TOKEN });
  return auth;
}
