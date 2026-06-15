import { requestUrl } from "obsidian";
import type { TokenState } from "./types";

const SCOPE = "offline_access Files.ReadWrite.AppFolder";

interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message: string;
}

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
}

interface OAuthError {
  error: string;
  error_description?: string;
}

export class MicrosoftAuth {
  constructor(
    private readonly clientId: string,
    private readonly tenant: string,
    private token: TokenState | null,
    private readonly onToken: (token: TokenState) => Promise<void>
  ) {}

  get signedIn(): boolean {
    return this.token !== null;
  }

  async beginDeviceCode(): Promise<DeviceCodeResponse> {
    const response = await requestUrl({
      url: `${this.baseUrl}/devicecode`,
      method: "POST",
      contentType: "application/x-www-form-urlencoded",
      body: form({ client_id: this.clientId, scope: SCOPE }),
      throw: false
    });
    if (response.status !== 200) throw new Error(readOAuthError(response.json));
    return readDeviceCode(response.json);
  }

  async finishDeviceCode(code: DeviceCodeResponse): Promise<void> {
    const deadline = Date.now() + code.expires_in * 1000;
    let interval = Math.max(code.interval, 5) * 1000;

    while (Date.now() < deadline) {
      await sleep(interval);
      const response = await requestUrl({
        url: `${this.baseUrl}/token`,
        method: "POST",
        contentType: "application/x-www-form-urlencoded",
        body: form({
          grant_type: "urn:ietf:params:oauth:grant-type:device_code",
          client_id: this.clientId,
          device_code: code.device_code
        }),
        throw: false
      });
      if (response.status === 200) {
        await this.acceptToken(readToken(response.json));
        return;
      }
      const error = readOAuthResponse(response.json);
      if (error.error === "authorization_pending") continue;
      if (error.error === "slow_down") {
        interval += 5000;
        continue;
      }
      throw new Error(readOAuthError(error));
    }
    throw new Error("登录代码已过期，请重试");
  }

  async accessToken(): Promise<string> {
    if (!this.token) throw new Error("尚未登录 Microsoft 账户");
    if (this.token.expiresAt > Date.now() + 60_000) return this.token.accessToken;

    const response = await requestUrl({
      url: `${this.baseUrl}/token`,
      method: "POST",
      contentType: "application/x-www-form-urlencoded",
      body: form({
        client_id: this.clientId,
        grant_type: "refresh_token",
        refresh_token: this.token.refreshToken,
        scope: SCOPE
      }),
      throw: false
    });
    if (response.status !== 200) throw new Error(readOAuthError(response.json));
    await this.acceptToken(readToken(response.json));
    if (!this.token) throw new Error("Microsoft 登录状态丢失");
    return this.token.accessToken;
  }

  private get baseUrl(): string {
    return `https://login.microsoftonline.com/${encodeURIComponent(this.tenant || "common")}/oauth2/v2.0`;
  }

  private async acceptToken(response: TokenResponse): Promise<void> {
    const refreshToken = response.refresh_token ?? this.token?.refreshToken;
    if (!refreshToken) throw new Error("Microsoft 未返回刷新令牌");
    this.token = {
      accessToken: response.access_token,
      refreshToken,
      expiresAt: Date.now() + response.expires_in * 1000
    };
    await this.onToken(this.token);
  }
}

function form(values: Record<string, string>): string {
  return Object.entries(values)
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
    .join("&");
}

function readOAuthError(value: unknown): string {
  const error = readOAuthResponse(value);
  return error?.error_description ?? error?.error ?? "Microsoft 登录请求失败";
}

function readDeviceCode(value: unknown): DeviceCodeResponse {
  if (!isRecord(value)) throw new Error("Microsoft 登录响应格式无效");
  const deviceCode = value.device_code;
  const userCode = value.user_code;
  const verificationUri = value.verification_uri;
  const expiresIn = value.expires_in;
  const interval = value.interval;
  const message = value.message;
  if (
    typeof deviceCode !== "string" ||
    typeof userCode !== "string" ||
    typeof verificationUri !== "string" ||
    typeof expiresIn !== "number" ||
    typeof interval !== "number" ||
    typeof message !== "string"
  ) {
    throw new Error("Microsoft 登录响应缺少必要字段");
  }
  return {
    device_code: deviceCode,
    user_code: userCode,
    verification_uri: verificationUri,
    expires_in: expiresIn,
    interval,
    message
  };
}

function readToken(value: unknown): TokenResponse {
  if (!isRecord(value) || typeof value.access_token !== "string" || typeof value.expires_in !== "number") {
    throw new Error("Microsoft 令牌响应格式无效");
  }
  return {
    access_token: value.access_token,
    refresh_token: typeof value.refresh_token === "string" ? value.refresh_token : undefined,
    expires_in: value.expires_in
  };
}

function readOAuthResponse(value: unknown): OAuthError {
  if (!isRecord(value) || typeof value.error !== "string") return { error: "Microsoft 登录请求失败" };
  return {
    error: value.error,
    error_description: typeof value.error_description === "string" ? value.error_description : undefined
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}
