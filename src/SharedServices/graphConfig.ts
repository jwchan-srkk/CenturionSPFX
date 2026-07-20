/**
 * Configuration for Microsoft Graph API access
 * Manages token source location and provides configuration hierarchy:
 * 1. Default configuration (hardcoded)
 * 2. Webpart property overrides (if provided)
 */

export interface IGraphConfig {
  tokenSourceUrl: string;
  tokenListName: string;
}

export class GraphConfig {
  /**
   * Default configuration for token source
   * Can be overridden at webpart level via constructor
   */
  private static readonly DEFAULT_CONFIG: IGraphConfig = {
    tokenSourceUrl: 'https://cl.sharepoint.com/sites/App_CapitaLandWorkflowSystemUAT',
    tokenListName: 'MSGraph_AccessToken'
  };

  private config: IGraphConfig;

  /**
   * Creates a new GraphConfig instance
   * @param customConfig Optional custom configuration to override defaults
   */
  constructor(customConfig?: Partial<IGraphConfig>) {
    // Merge default config with custom config (custom takes precedence)
    this.config = {
      ...GraphConfig.DEFAULT_CONFIG,
      ...(customConfig || {})
    };

    // Filter out undefined/null values from custom config
    if (customConfig?.tokenSourceUrl === undefined || customConfig?.tokenSourceUrl === null || customConfig?.tokenSourceUrl === '') {
      this.config.tokenSourceUrl = GraphConfig.DEFAULT_CONFIG.tokenSourceUrl;
    }
    if (customConfig?.tokenListName === undefined || customConfig?.tokenListName === null || customConfig?.tokenListName === '') {
      this.config.tokenListName = GraphConfig.DEFAULT_CONFIG.tokenListName;
    }
  }

  /**
   * Gets the SharePoint site URL where the access token is stored
   * @returns Full URL to the token source site
   */
  public getTokenSourceUrl(): string {
    return this.config.tokenSourceUrl;
  }

  /**
   * Gets the SharePoint list name containing the access token
   * @returns List name (e.g., 'MSGraph_AccessToken')
   */
  public getTokenListName(): string {
    return this.config.tokenListName;
  }

  /**
   * Gets the complete configuration object
   * @returns Complete configuration
   */
  public getConfig(): IGraphConfig {
    return { ...this.config };
  }
}
