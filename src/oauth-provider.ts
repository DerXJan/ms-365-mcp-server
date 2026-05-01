import { Response } from 'express';
import { ProxyOAuthServerProvider } from '@modelcontextprotocol/sdk/server/auth/providers/proxyProvider.js';
import type { OAuthRegisteredClientsStore } from '@modelcontextprotocol/sdk/server/auth/clients.js';
import type { OAuthClientInformationFull, OAuthTokens } from '@modelcontextprotocol/sdk/shared/auth.js';
import type { AuthInfo } from '@modelcontextprotocol/sdk/server/auth/types.js';
import type { AuthorizationParams } from '@modelcontextprotocol/sdk/server/auth/provider.js';
import logger from './logger.js';
import AuthManager from './auth.js';
import type { AppSecrets } from './secrets.js';
import { getCloudEndpoints } from './cloud-config.js';
import crypto from 'node:crypto';

const _clientRegistry = new Map<string, OAuthClientInformationFull>();

export class MicrosoftOAuthProvider extends ProxyOAuthServerProvider {
  private authManager: AuthManager;
  private _azureClientId: string;
  private _azureClientSecret: string | undefined;

  constructor(authManager: AuthManager, secrets: AppSecrets) {
    const tenantId = secrets.tenantId || 'common';
    const clientId = secrets.clientId;
    const cloudEndpoints = getCloudEndpoints(secrets.cloudType);

    super({
      endpoints: {
        authorizationUrl: `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/authorize`,
        tokenUrl: `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/token`,
        revocationUrl: `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/logout`,
      },
      verifyAccessToken: async (token: string): Promise<AuthInfo> => {
        try {
          const response = await fetch(`${cloudEndpoints.graphApi}/v1.0/me`, {
            headers: { Authorization: `Bearer ${token}` },
          });

          if (response.ok) {
            const userData = await response.json();
            logger.info(`OAuth token verified for user: ${userData.userPrincipalName}`);
            await authManager.setOAuthToken(token);
            return { token, clientId, scopes: [] };
          } else {
            throw new Error(`Token verification failed: ${response.status}`);
          }
        } catch (error) {
          logger.error(`OAuth token verification error: ${error}`);
          throw error;
        }
      },
      getClient: async (client_id: string) => {
        return _clientRegistry.get(client_id) ?? { client_id, redirect_uris: [] };
      },
    });

    this.authManager = authManager;
    this._azureClientId = clientId;
    this._azureClientSecret = secrets.clientSecret;
  }

  override get clientsStore(): OAuthRegisteredClientsStore {
    const base = super.clientsStore;
    return {
      ...base,
      registerClient: async (
        clientMetadata: Omit<OAuthClientInformationFull, 'client_id' | 'client_id_issued_at'>
      ): Promise<OAuthClientInformationFull> => {
        const client: OAuthClientInformationFull = {
          ...clientMetadata,
          client_id: crypto.randomUUID(),
          client_id_issued_at: Math.floor(Date.now() / 1000),
        };
        _clientRegistry.set(client.client_id, client);
        logger.info(
          `OAuth client registered: ${client.client_id} redirect_uris=${JSON.stringify(clientMetadata.redirect_uris)}`
        );
        return client;
      },
    };
  }

  override async authorize(
    _client: OAuthClientInformationFull,
    params: AuthorizationParams,
    res: Response
  ): Promise<void> {
    const targetUrl = new URL(this._endpoints.authorizationUrl);
    const searchParams = new URLSearchParams({
      client_id: this._azureClientId,
      response_type: 'code',
      redirect_uri: params.redirectUri,
      code_challenge: params.codeChallenge,
      code_challenge_method: 'S256',
    });
    if (params.state) searchParams.set('state', params.state);
    if (params.scopes?.length) searchParams.set('scope', params.scopes.join(' '));
    if (params.resource) searchParams.set('resource', params.resource.href);
    targetUrl.search = searchParams.toString();
    res.redirect(targetUrl.toString());
  }

  override async exchangeAuthorizationCode(
    _client: OAuthClientInformationFull,
    authorizationCode: string,
    codeVerifier?: string,
    redirectUri?: string,
    resource?: URL
  ): Promise<OAuthTokens> {
    const params = new URLSearchParams({
      grant_type: 'authorization_code',
      client_id: this._azureClientId,
      code: authorizationCode,
    });
    if (this._azureClientSecret) {
      params.append('client_secret', this._azureClientSecret);
    }
    if (codeVerifier) {
      params.append('code_verifier', codeVerifier);
    }
    if (redirectUri) {
      params.append('redirect_uri', redirectUri);
    }
    if (resource) {
      params.append('resource', resource.href);
    }
    const response = await fetch(this._endpoints.tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });
    if (!response.ok) {
      const error = await response.text();
      logger.error(`Token exchange failed: ${error}`);
      throw new Error(`Token exchange failed: ${response.status} ${error}`);
    }
    return response.json();
  }

  override async exchangeRefreshToken(
    _client: OAuthClientInformationFull,
    refreshToken: string,
    scopes?: string[],
    resource?: URL
  ): Promise<OAuthTokens> {
    const params = new URLSearchParams({
      grant_type: 'refresh_token',
      client_id: this._azureClientId,
      refresh_token: refreshToken,
    });
    if (this._azureClientSecret) {
      params.set('client_secret', this._azureClientSecret);
    }
    if (scopes?.length) {
      params.set('scope', scopes.join(' '));
    }
    if (resource) {
      params.set('resource', resource.href);
    }
    const response = await fetch(this._endpoints.tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });
    if (!response.ok) {
      const error = await response.text();
      logger.error(`Token refresh failed: ${error}`);
      throw new Error(`Token refresh failed: ${response.status} ${error}`);
    }
    return response.json();
  }
}
