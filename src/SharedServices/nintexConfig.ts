export interface INintexConfig {
  apiBaseUrl: string;
  tokenEndpoint: string;
  clientId?: string;
  clientSecret?: string;
  tokenSourceUrl?: string;
  tokenListName?: string;
  tokenItemTitle?: string;
  directAccessToken?: string;
}

export class NintexConfig {
  private static readonly DEFAULT_CONFIG: INintexConfig = {
    apiBaseUrl: 'https://au.nintex.io',
    tokenEndpoint: 'https://au.nintex.io/authentication/v1/token',
    clientId: '',
    clientSecret: '',
    tokenSourceUrl: '',
    tokenListName: 'Nintex_AccessToken',
    tokenItemTitle: 'AccessToken',
    directAccessToken: ''
  };

  private readonly config: INintexConfig;

  constructor(customConfig?: Partial<INintexConfig>) {
    this.config = {
      ...NintexConfig.DEFAULT_CONFIG,
      ...(customConfig || {})
    };
  }

  public getApiBaseUrl(): string {
    return this.config.apiBaseUrl;
  }

  public getTokenEndpoint(): string {
    return this.config.tokenEndpoint;
  }

  public getTokenSourceUrl(): string {
    return this.config.tokenSourceUrl || '';
  }

  public getClientId(): string {
    return this.config.clientId || '';
  }

  public getClientSecret(): string {
    return this.config.clientSecret || '';
  }

  public getTokenListName(): string {
    return this.config.tokenListName || '';
  }

  public getTokenItemTitle(): string {
    return this.config.tokenItemTitle || 'AccessToken';
  }

  public getDirectAccessToken(): string {
    return this.config.directAccessToken || '';
  }

  public getConfig(): INintexConfig {
    return { ...this.config };
  }
}
