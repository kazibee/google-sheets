import { auth as googleAuth } from '@googleapis/sheets';

/** Scoped secrets provided by the tool runtime for Google OAuth2 authentication. */
export interface Env {
  /** Google OAuth2 client ID. */
  CLIENT_ID: string;
  /** Google OAuth2 client secret. */
  CLIENT_SECRET: string;
  /** Long-lived refresh token obtained during `kazibee google-sheets login`. */
  REFRESH_TOKEN: string;
}

/** Build an OAuth2Client from the scoped secrets provided by the runtime. */
export function createAuthClient(env: Env) {
  const oauth2 = new googleAuth.OAuth2(env.CLIENT_ID, env.CLIENT_SECRET);
  oauth2.setCredentials({ refresh_token: env.REFRESH_TOKEN });
  return oauth2;
}
