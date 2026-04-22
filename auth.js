import fs from "fs";
import path from "path";
import os from "os";
import { createCipheriv, createDecipheriv, randomBytes, scryptSync } from "crypto";

// Replace with your own Azure app registration client ID.
// See README.md for instructions on creating one.
const DEFAULT_CLIENT_ID = process.env.TODO_MCP_CLIENT_ID?.trim() || "YOUR_CLIENT_ID_HERE";

// "common" allows any Entra ID tenant + personal Microsoft accounts.
// Override per-account with the tenant_id argument on authenticate_account,
// or globally with the TODO_MCP_TENANT_ID env var.
const DEFAULT_TENANT_ID = "common";

const DEFAULT_SCOPE = "Tasks.Read Tasks.ReadWrite Group.Read.All User.ReadBasic.All offline_access";
const CONFIG_PATH = process.env.TODO_MCP_CONFIG_PATH?.trim() || path.join(os.homedir(), ".todo-mcp-config.json");

const tokenCache = new Map();

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function getDefaultAccountName() {
  return process.env.TODO_MCP_DEFAULT_ACCOUNT?.trim() || "default";
}

function getDefaultAuthConfig() {
  return {
    client_id: process.env.TODO_MCP_CLIENT_ID?.trim() || DEFAULT_CLIENT_ID,
    tenant_id: process.env.TODO_MCP_TENANT_ID?.trim() || DEFAULT_TENANT_ID,
    scope: process.env.TODO_MCP_SCOPE?.trim() || DEFAULT_SCOPE,
  };
}

function getRequiredMasterKey() {
  const masterKey = process.env.TODO_MCP_MASTER_KEY;
  if (!masterKey || !masterKey.trim()) {
    throw new Error(
      "TODO_MCP_MASTER_KEY is required to encrypt refresh tokens. Configure it in `.vscode/mcp.json` or your environment."
    );
  }
  return masterKey;
}

function getAuthErrorMessage(resp) {
  return resp.error_description || resp.error || "Unknown authentication error";
}

async function postForm(url, body) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 15000);
  try {
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(body).toString(),
      signal: controller.signal,
    });
    return await resp.json();
  } finally {
    clearTimeout(timer);
  }
}

function normalizeConfig(rawConfig) {
  const raw = isRecord(rawConfig) ? rawConfig : {};

  return {
    version: 2,
    default_account: typeof raw.default_account === "string" && raw.default_account.trim()
      ? raw.default_account.trim()
      : getDefaultAccountName(),
    accounts: isRecord(raw.accounts) ? raw.accounts : {},
    legacy_refresh_token: typeof raw.refresh_token === "string" && raw.refresh_token ? raw.refresh_token : null,
  };
}

function loadConfig() {
  let rawText;

  try {
    rawText = fs.readFileSync(CONFIG_PATH, "utf8");
  } catch (error) {
    if (error?.code === "ENOENT") return normalizeConfig({});
    throw new Error(`Failed to read auth config at ${CONFIG_PATH}: ${error.message}`);
  }

  try {
    return normalizeConfig(JSON.parse(rawText));
  } catch (error) {
    throw new Error(`Failed to parse auth config at ${CONFIG_PATH}: ${error.message}`);
  }
}

function saveConfig(config) {
  const payload = {
    version: 2,
    default_account: config.default_account,
    accounts: config.accounts,
  };

  fs.mkdirSync(path.dirname(CONFIG_PATH), { recursive: true });
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(payload, null, 2), { encoding: "utf8", mode: 0o600 });
}

function resolveAccountName(requestedAccount, config) {
  if (typeof requestedAccount === "string" && requestedAccount.trim()) return requestedAccount.trim();
  if (typeof config?.default_account === "string" && config.default_account.trim()) return config.default_account.trim();
  return getDefaultAccountName();
}

function getStoredAccount(config, accountName) {
  const account = config.accounts[accountName];
  if (account === undefined) return null;
  if (!isRecord(account)) throw new Error(`Stored account "${accountName}" is invalid.`);
  return account;
}

function buildAuthConfig(overrides = {}, existingAccount = null) {
  const defaults = getDefaultAuthConfig();

  return {
    client_id: typeof overrides.client_id === "string" && overrides.client_id.trim()
      ? overrides.client_id.trim()
      : (existingAccount?.client_id || defaults.client_id),
    tenant_id: typeof overrides.tenant_id === "string" && overrides.tenant_id.trim()
      ? overrides.tenant_id.trim()
      : (existingAccount?.tenant_id || defaults.tenant_id),
    scope: typeof overrides.scope === "string" && overrides.scope.trim()
      ? overrides.scope.trim()
      : (existingAccount?.scope || defaults.scope),
  };
}

function encryptRefreshToken(refreshToken) {
  const masterKey = getRequiredMasterKey();
  const salt = randomBytes(16);
  const iv = randomBytes(12);
  const key = scryptSync(masterKey, salt, 32);
  const cipher = createCipheriv("aes-256-gcm", key, iv);
  const ciphertext = Buffer.concat([cipher.update(refreshToken, "utf8"), cipher.final()]);

  return {
    salt: salt.toString("base64"),
    iv: iv.toString("base64"),
    tag: cipher.getAuthTag().toString("base64"),
    ciphertext: ciphertext.toString("base64"),
  };
}

function decryptRefreshToken(accountName, encryptedToken) {
  if (!isRecord(encryptedToken)) {
    throw new Error(`Stored refresh token for account "${accountName}" is invalid.`);
  }

  const { salt, iv, tag, ciphertext } = encryptedToken;
  if (![salt, iv, tag, ciphertext].every((part) => typeof part === "string" && part)) {
    throw new Error(`Stored refresh token for account "${accountName}" is incomplete.`);
  }

  try {
    const masterKey = getRequiredMasterKey();
    const key = scryptSync(masterKey, Buffer.from(salt, "base64"), 32);
    const decipher = createDecipheriv("aes-256-gcm", key, Buffer.from(iv, "base64"));
    decipher.setAuthTag(Buffer.from(tag, "base64"));

    return Buffer.concat([
      decipher.update(Buffer.from(ciphertext, "base64")),
      decipher.final(),
    ]).toString("utf8");
  } catch {
    throw new Error(`Unable to decrypt refresh token for account "${accountName}". Check TODO_MCP_MASTER_KEY.`);
  }
}

function getRefreshTokenOrThrow(resp, accountName) {
  if (typeof resp.refresh_token !== "string" || !resp.refresh_token) {
    throw new Error(
      `Authentication succeeded for account "${accountName}", but no refresh token was returned. Ensure offline_access is included in the scope.`
    );
  }

  return resp.refresh_token;
}

function setCachedToken(accountName, accessToken, expiresIn) {
  tokenCache.set(accountName, {
    access_token: accessToken,
    expires_at: Date.now() + (expiresIn || 3600) * 1000,
  });
}

function getCachedToken(accountName) {
  const cached = tokenCache.get(accountName);
  if (!cached) return null;
  if (Date.now() >= cached.expires_at - 120000) {
    tokenCache.delete(accountName);
    return null;
  }
  return cached.access_token;
}

function upsertAccount(config, accountName, authConfig, refreshToken) {
  const now = new Date().toISOString();
  const existing = getStoredAccount(config, accountName) || {};

  config.accounts[accountName] = {
    ...existing,
    client_id: authConfig.client_id,
    tenant_id: authConfig.tenant_id,
    scope: authConfig.scope,
    refresh_token: encryptRefreshToken(refreshToken),
    created_at: typeof existing.created_at === "string" ? existing.created_at : now,
    updated_at: now,
    last_used_at: now,
  };

  if (!config.default_account) config.default_account = accountName;
  return config;
}

function migrateLegacyRefreshToken(config, accountName) {
  if (!config.legacy_refresh_token || getStoredAccount(config, accountName)) return config;

  process.stderr.write(`Migrating legacy refresh token into encrypted storage for account "${accountName}".\n`);
  const authConfig = buildAuthConfig();
  const nextConfig = upsertAccount(config, accountName, authConfig, config.legacy_refresh_token);
  saveConfig(nextConfig);
  return normalizeConfig(nextConfig);
}

async function refreshAccessToken(refreshToken, authConfig) {
  const resp = await postForm(
    `https://login.microsoftonline.com/${authConfig.tenant_id}/oauth2/v2.0/token`,
    {
      grant_type: "refresh_token",
      client_id: authConfig.client_id,
      refresh_token: refreshToken,
      scope: authConfig.scope,
    }
  );

  if (resp.error) throw new Error(getAuthErrorMessage(resp));
  return resp;
}

function toAccountSummary(accountName, account, defaultAccount) {
  return {
    account: accountName,
    is_default: defaultAccount === accountName,
    client_id: account.client_id || null,
    tenant_id: account.tenant_id || null,
    scope: account.scope || null,
    created_at: account.created_at || null,
    updated_at: account.updated_at || null,
    last_used_at: account.last_used_at || null,
  };
}

export function listStoredAccounts() {
  const config = loadConfig();

  return {
    default_account: config.default_account,
    accounts: Object.entries(config.accounts)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([accountName, account]) => {
        if (!isRecord(account)) throw new Error(`Stored account "${accountName}" is invalid.`);
        return toAccountSummary(accountName, account, config.default_account);
      }),
  };
}

// Phase 1: Request a device code and return it immediately (non-blocking).
// Returns { accountName, authConfig, dcResp, set_default, hadAccounts }
export async function startDeviceCode({ account, client_id, tenant_id, scope, set_default = false }) {
  getRequiredMasterKey();

  const config = loadConfig();
  const accountName = resolveAccountName(account, config);
  const existingAccount = getStoredAccount(config, accountName);
  const hadAccounts = Object.keys(config.accounts).length > 0;
  const authConfig = buildAuthConfig({ client_id, tenant_id, scope }, existingAccount);

  const dcResp = await postForm(
    `https://login.microsoftonline.com/${authConfig.tenant_id}/oauth2/v2.0/devicecode`,
    { client_id: authConfig.client_id, scope: authConfig.scope }
  );

  if (dcResp.error) throw new Error(getAuthErrorMessage(dcResp));

  return { accountName, authConfig, dcResp, set_default, hadAccounts };
}

// Phase 2: Poll for token completion. Saves to disk + in-memory cache when done.
// Run this in a fire-and-forget background Promise.
export async function pollDeviceCode({ accountName, authConfig, dcResp, set_default, hadAccounts }) {
  let interval = (dcResp.interval || 5) * 1000;
  const deadline = Date.now() + (dcResp.expires_in || 900) * 1000;

  while (Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, interval));

    const resp = await postForm(
      `https://login.microsoftonline.com/${authConfig.tenant_id}/oauth2/v2.0/token`,
      {
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: authConfig.client_id,
        device_code: dcResp.device_code,
      }
    );

    if (resp.access_token) {
      const refreshToken = getRefreshTokenOrThrow(resp, accountName);
      setCachedToken(accountName, resp.access_token, resp.expires_in);

      const config = loadConfig();
      const nextConfig = upsertAccount(config, accountName, authConfig, refreshToken);
      if (set_default || !hadAccounts) nextConfig.default_account = accountName;
      saveConfig(nextConfig);

      process.stderr.write(`Authentication complete for account "${accountName}".\n`);
      return toAccountSummary(accountName, nextConfig.accounts[accountName], nextConfig.default_account);
    }

    if (resp.error === "authorization_pending") continue;
    if (resp.error === "slow_down") { interval += 5000; continue; }
    if (resp.error) throw new Error(`Authentication failed: ${getAuthErrorMessage(resp)}`);
  }

  throw new Error("Device code flow timed out");
}

export function setDefaultStoredAccount(account) {
  const config = loadConfig();
  const accountName = resolveAccountName(account, config);
  const existingAccount = getStoredAccount(config, accountName);

  if (!existingAccount) {
    throw new Error(`Account "${accountName}" is not configured.`);
  }

  config.default_account = accountName;
  saveConfig(config);

  return toAccountSummary(accountName, existingAccount, config.default_account);
}

export function deleteStoredAccount(account) {
  const config = loadConfig();
  const accountName = resolveAccountName(account, config);

  if (!getStoredAccount(config, accountName)) {
    throw new Error(`Account "${accountName}" is not configured.`);
  }

  delete config.accounts[accountName];
  tokenCache.delete(accountName);

  if (config.default_account === accountName) {
    const remainingAccounts = Object.keys(config.accounts).sort((left, right) => left.localeCompare(right));
    config.default_account = remainingAccounts[0] || getDefaultAccountName();
  }

  saveConfig(config);

  return {
    deleted_account: accountName,
    default_account: config.default_account,
    remaining_accounts: Object.keys(config.accounts).sort((left, right) => left.localeCompare(right)),
  };
}

export async function getAccessToken(account) {
  getRequiredMasterKey();

  let config = loadConfig();
  const accountName = resolveAccountName(account, config);

  const cachedToken = getCachedToken(accountName);
  if (cachedToken) return cachedToken;

  config = migrateLegacyRefreshToken(config, accountName);

  const storedAccount = getStoredAccount(config, accountName);
  const authConfig = buildAuthConfig({}, storedAccount);

  if (storedAccount?.refresh_token) {
    const refreshToken = decryptRefreshToken(accountName, storedAccount.refresh_token);

    try {
      const resp = await refreshAccessToken(refreshToken, authConfig);
      const nextRefreshToken = typeof resp.refresh_token === "string" && resp.refresh_token
        ? resp.refresh_token
        : refreshToken;

      setCachedToken(accountName, resp.access_token, resp.expires_in);
      const nextConfig = upsertAccount(config, accountName, authConfig, nextRefreshToken);
      saveConfig(nextConfig);
      return resp.access_token;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      process.stderr.write(`Token refresh failed for account "${accountName}": ${message}\n`);
    }
  }

  throw new Error(`Account "${accountName}" is not authenticated. Use authenticate_account to log in.`);
}
