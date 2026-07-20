/**
 * Microsoft Graph API Service for SharePoint data access
 * Replaces SharePointRestService with Graph API calls
 * Uses token from SharePoint list instead of MSAL authentication
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointRestService } from './sharePointRestService';
import { GraphConfig } from './graphConfig';
import { GraphFieldMapper } from './graphFieldMapper';

/**
 * Cache item structure for localStorage
 */
interface CacheItem {
  data: any[];
  timestamp: number;
  expiryTime: number;
}

/**
 * Site ID cache entry
 */
interface SiteIdCacheEntry {
  siteId: string;
  timestamp: number;
}

/**
 * List ID cache entry
 */
interface ListIdCacheEntry {
  listId: string;
  timestamp: number;
}

export class MicrosoftGraphService {
  private context: WebPartContext;
  private sharePointService: SharePointRestService; // For token retrieval ONLY
  private config: GraphConfig;

  // ID Caches (in-memory, 24-hour expiry)
  private siteIdCache: Map<string, SiteIdCacheEntry> = new Map();
  private listIdCache: Map<string, ListIdCacheEntry> = new Map();

  // Cache expiry times
  private readonly SITE_ID_CACHE_HOURS = 24;
  private readonly LIST_ID_CACHE_HOURS = 24;

  constructor(context: WebPartContext, config?: GraphConfig) {
    this.context = context;
    this.sharePointService = new SharePointRestService(context);
    this.config = config || new GraphConfig();
  }

  // ==================== TOKEN MANAGEMENT ====================

  /**
   * Retrieves Microsoft Graph access token from SharePoint list
   * NO CACHING - Always fetches fresh token
   * @returns Access token string
   */
  private async getAccessToken(): Promise<string> {
    try {
      const tokenSourceUrl = this.config.getTokenSourceUrl();
      const tokenListName = this.config.getTokenListName();

      // Use SharePointRestService to fetch token from list
      const items = await this.sharePointService.getListItems(
        tokenSourceUrl,
        tokenListName,
        'Title,AccessToken',
        "Title eq 'AccessToken'"
      );

      if (!items || items.length === 0) {
        throw new Error('Microsoft Graph access token not found in list. Please ensure the MSGraph_AccessToken list contains an item with Title="AccessToken".');
      }

      const token = items[0].AccessToken;

      if (!token || token.trim() === '') {
        throw new Error('Microsoft Graph access token is empty. Please contact IT administrator to update the token.');
      }

      return token;

    } catch (error) {
      console.error('❌ Failed to retrieve Graph access token:', error);
      throw new Error(`Failed to retrieve Microsoft Graph access token: ${error.message}. Please contact IT administrator.`);
    }
  }

  // ==================== ID RESOLUTION WITH CACHING ====================

  /**
   * Resolves SharePoint site URL to Graph API site ID
   * Caches result for 24 hours
   * @param siteUrl Full SharePoint site URL
   * @returns Site ID for use in Graph API calls
   */
  private async getSiteId(siteUrl: string): Promise<string> {
    // Check cache first
    const cached = this.siteIdCache.get(siteUrl);
    if (cached && this.isCacheValid(cached.timestamp, this.SITE_ID_CACHE_HOURS)) {
      return cached.siteId;
    }

    try {
      // Parse site URL
      const url = new URL(siteUrl);
      const hostname = url.hostname;
      const sitePath = url.pathname;

      // Get token
      const token = await this.getAccessToken();

      // Call Graph API to resolve site ID
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`,
        {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/json'
          }
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API error ${response.status}: ${errorText}`);
      }

      const data = await response.json();
      const siteId = data.id.split(",")[1];

      // Cache for 24 hours
      this.siteIdCache.set(siteUrl, {
        siteId: data.id.split(",")[1],
        timestamp: Date.now()
      });

      return data.id.split(",")[1];

    } catch (error) {
      console.error(`❌ Failed to resolve Site ID for ${siteUrl}:`, error);
      throw new Error(`Failed to resolve Site ID for ${siteUrl}: ${error.message}`);
    }
  }

  /**
   * Resolves SharePoint list name to Graph API list ID
   * Caches result for 24 hours
   * @param siteUrl SharePoint site URL
   * @param listName SharePoint list display name
   * @returns List ID for use in Graph API calls
   */
  private async getListId(siteUrl: string, listName: string): Promise<string> {
    const cacheKey = `${siteUrl}|${listName}`;

    // Check cache first
    const cached = this.listIdCache.get(cacheKey);
    if (cached && this.isCacheValid(cached.timestamp, this.LIST_ID_CACHE_HOURS)) {
      return cached.listId;
    }


    try {
      // Get site ID first
      const siteId = await this.getSiteId(siteUrl);
      const token = await this.getAccessToken();

      // Call Graph API to get list by display name
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$filter=displayName eq '${encodeURIComponent(listName)}'`,
        {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/json'
          }
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API error ${response.status}: ${errorText}`);
      }

      let data = await response.json();

      // Fall back to internal (URL) name if display name lookup returns nothing
      if (!data.value || data.value.length === 0) {
        const fallbackResponse = await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$filter=name eq '${encodeURIComponent(listName)}'`,
          {
            headers: {
              'Authorization': `Bearer ${token}`,
              'Accept': 'application/json'
            }
          }
        );
        if (fallbackResponse.ok) {
          data = await fallbackResponse.json();
        }
      }

      if (!data.value || data.value.length === 0) {
        throw new Error(`List '${listName}' not found on site ${siteUrl}`);
      }

      const listId = data.value[0].id;

      // Cache for 24 hours
      this.listIdCache.set(cacheKey, {
        listId: listId,
        timestamp: Date.now()
      });

      return listId;

    } catch (error) {
      console.error(`❌ Failed to resolve List ID for ${listName}:`, error);
      throw new Error(`Failed to resolve List ID for '${listName}': ${error.message}`);
    }
  }

  /**
   * Check if cache entry is still valid
   */
  private isCacheValid(timestamp: number, expiryHours: number): boolean {
    const expiryMs = expiryHours * 60 * 60 * 1000;
    return (Date.now() - timestamp) < expiryMs;
  }

  // ==================== CORE CRUD OPERATIONS ====================

  /**
   * Get list items with optional filtering, sorting, and pagination
   * @param siteUrl SharePoint site URL
   * @param listName List display name
   * @param selectFields Comma-separated field names to select
   * @param filter OData filter expression
   * @param orderBy OData orderby expression
   * @param top Maximum number of items to return
   * @returns Array of list items
   */
  public async getListItems(
    siteUrl: string,
    listName: string,
    selectFields?: string,
    filter?: string,
    orderBy?: string,
    top?: number
  ): Promise<any[]> {
    try {
      const siteId = await this.getSiteId(siteUrl);
      const listId = await this.getListId(siteUrl, listName);
      const token = await this.getAccessToken();

      // Build Graph API URL
      const url = GraphFieldMapper.buildGraphItemsUrl(
        siteId,
        listId,
        selectFields,
        filter,
        orderBy,
        top || 5000
      );


      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json'
        }
      });

      if (response.status === 401) {
        throw new Error('Access token has expired. Please contact IT administrator to refresh the token.');
      }

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API error ${response.status}: ${errorText}`);
      }

      const data = await response.json();

      // Transform items from Graph format to SharePoint REST format
      const transformedItems = data.value.map((item: any) =>
        GraphFieldMapper.transformItemFromGraph(item)
      );


      // Client-side sorting to ensure correct order (if orderBy was specified)
      // This ensures data is properly sorted even if Graph API $orderby doesn't work reliably
      if (orderBy && transformedItems.length > 0) {
        const sortedItems = this._applySorting(transformedItems, orderBy);
        return sortedItems;
      }

      return transformedItems;

    } catch (error) {
      console.error(`❌ Error fetching items from ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Get ALL list items with automatic pagination handling
   * @param siteUrl SharePoint site URL
   * @param listName List display name
   * @param selectFields Comma-separated field names
   * @param filter OData filter expression
   * @param orderBy OData orderby expression
   * @returns Array of all list items (handles pagination)
   */
  public async getAllListItems(
    siteUrl: string,
    listName: string,
    selectFields?: string,
    filter?: string,
    orderBy?: string
  ): Promise<any[]> {
    try {
      const siteId = await this.getSiteId(siteUrl);
      const listId = await this.getListId(siteUrl, listName);
      const token = await this.getAccessToken();

      let allItems: any[] = [];
      let nextLink: string | null = GraphFieldMapper.buildGraphItemsUrl(
        siteId,
        listId,
        selectFields,
        filter,
        orderBy,
        5000 // Max per page
      );


      while (nextLink) {
        const response: Response = await fetch(nextLink, {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/json',
            'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly'
          }
        });

        if (response.status === 401) {
          throw new Error('Access token has expired. Please contact IT administrator to refresh the token.');
        }

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Graph API error ${response.status}: ${errorText}`);
        }

        const data: any = await response.json();

        // Transform and add items
        const transformedItems = data.value.map((item: any) =>
          GraphFieldMapper.transformItemFromGraph(item)
        );
        allItems = allItems.concat(transformedItems);

        // Get next page URL
        nextLink = data['@odata.nextLink'] || null;
      }


      // Client-side sorting to ensure correct order (if orderBy was specified)
      // This ensures data is properly sorted even if Graph API $orderby doesn't work reliably
      if (orderBy && allItems.length > 0) {
        const sortedItems = this._applySorting(allItems, orderBy);
        return sortedItems;
      }

      return allItems;

    } catch (error) {
      console.error(`❌ Error fetching all items from ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Get single list item by ID
   * @param siteUrl SharePoint site URL
   * @param listName List display name
   * @param itemId Item ID
   * @param selectFields Optional fields to select
   * @returns Single list item
   */
  public async getListItemById(
    siteUrl: string,
    listName: string,
    itemId: number,
    selectFields?: string
  ): Promise<any> {
    try {
      const siteId = await this.getSiteId(siteUrl);
      const listId = await this.getListId(siteUrl, listName);
      const token = await this.getAccessToken();

      let url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`;

      const graphSelect = GraphFieldMapper.transformSelectForGraph(selectFields);
      if (graphSelect) {
        url += `&$select=${encodeURIComponent(graphSelect)}`;
      }

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json'
        }
      });

      if (response.status === 401) {
        throw new Error('Access token has expired. Please contact IT administrator to refresh the token.');
      }

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API error ${response.status}: ${errorText}`);
      }

      const data = await response.json();

      return GraphFieldMapper.transformItemFromGraph(data);

    } catch (error) {
      console.error(`❌ Error fetching item ${itemId} from ${listName}:`, error);
      throw error;
    }
  }

  // ==================== FIELD DISCOVERY ====================

  /**
   * Get list columns (field schema)
   * @param siteUrl SharePoint site URL
   * @param listName List display name
   * @returns Array of field definitions
   */
  public async getListColumns(siteUrl: string, listName: string): Promise<any[]> {
    try {
      const siteId = await this.getSiteId(siteUrl);
      const listId = await this.getListId(siteUrl, listName);
      const token = await this.getAccessToken();

      const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns`;

      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json'
        }
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API error ${response.status}: ${errorText}`);
      }

      const data = await response.json();
      const allColumns = data.value;

      // Filter out system and complex fields (match SharePointRestService logic)
      const filteredColumns = allColumns.filter((column: any) => {
        // Exclude hidden or readonly
        if (column.hidden || column.readOnly) return false;

        // Exclude system fields
        const systemFields = [
          'Attachments', 'GUID', 'ContentTypeId', 'Edit', 'ComplianceAssetId',
          'ItemChildCount', 'FolderChildCount', '_UIVersionString', 'AppAuthor',
          'AppEditor', 'SelectTitle', 'SelectFilename', 'Edit', 'DocIcon'
        ];
        if (systemFields.includes(column.name)) return false;

        // Exclude fields starting with underscore
        if (column.name.startsWith('_')) return false;

        // Exclude complex field types
        if (column.lookup || column.personOrGroup || column.taxonomy) {
          return false;
        }

        return true;
      });

      // Transform to match SharePoint REST format
      const transformedColumns = filteredColumns.map((col: any) => ({
        InternalName: col.name,
        Title: col.displayName || col.name,
        TypeAsString: this.getSharePointFieldType(col),
        Hidden: col.hidden || false,
        ReadOnlyField: col.readOnly || false
      }));

      return transformedColumns;

    } catch (error) {
      console.error(`❌ Error fetching columns from ${listName}:`, error);
      throw error;
    }
  }

  /**
   * Map Graph column type to SharePoint field type string
   */
  private getSharePointFieldType(column: any): string {
    if (column.text) return 'Text';
    if (column.number) return 'Number';
    if (column.boolean) return 'Boolean';
    if (column.dateTime) return 'DateTime';
    if (column.choice) return 'Choice';
    if (column.lookup) return 'Lookup';
    if (column.personOrGroup) return 'User';
    if (column.currency) return 'Currency';
    if (column.hyperlinkOrPicture) return 'URL';
    return 'Text';
  }

  /**
   * Get list items with automatic field discovery
   * Combines getListColumns() + getAllListItems()
   * @param dataSiteUrl SharePoint site URL
   * @param listName List display name
   * @param filter OData filter
   * @param orderBy OData orderby
   * @param top Max items
   * @returns Object with items, fields, and selectFields string
   */
  public async getCrossSiteListItemsWithAllFields(
    dataSiteUrl: string,
    listName: string,
    filter?: string,
    orderBy?: string,
    top?: number
  ): Promise<{ items: any[], fields: any[], selectFields: string }> {
    try {
      // Get field schema
      const fields = await this.getListColumns(dataSiteUrl, listName);

      // Build select fields from discovered fields
      const essentialFields = ['Id', 'Title', 'Created', 'Modified', 'ContentType'];
      const discoveredFields = fields.map(f => f.InternalName);

      // Combine and deduplicate
      const allFields = [...new Set([...essentialFields, ...discoveredFields])];
      const selectFields = allFields.join(',');

      // Get items
      const items = top
        ? await this.getListItems(dataSiteUrl, listName, '', filter, orderBy, top)
        : await this.getAllListItems(dataSiteUrl, listName, '', filter, orderBy);

      return {
        items,
        fields,
        selectFields
      };

    } catch (error) {
      console.error(`❌ Error in getCrossSiteListItemsWithAllFields:`, error);
      throw error;
    }
  }

  /**
   * Auto-detect common filter fields using pattern matching
   * @param fields Array of field definitions
   * @returns Object with detected field names
   */
  public getCommonFilterFields(fields: any[]): {
    categoryField?: string,
    subcategoryField?: string,
    statusField?: string,
    dateField?: string
  } {
    const result: any = {};

    const patterns = {
      category: ['category', 'cat', 'type', 'classification'],
      subcategory: ['subcategory', 'subcat', 'subtype', 'detail'],
      status: ['status', 'state', 'condition', 'stage'],
      date: ['created', 'submitted', 'date', 'time']
    };

    fields.forEach(field => {
      const name = (field.InternalName || field.name || '').toLowerCase();
      const title = (field.Title || field.displayName || '').toLowerCase();

      for (const [key, patternList] of Object.entries(patterns)) {
        if (!result[`${key}Field`]) {
          if (patternList.some((p: string) => name.includes(p) || title.includes(p))) {
            result[`${key}Field`] = field.InternalName || field.name;
          }
        }
      }
    });

    return result;
  }

  // ==================== CACHING METHODS ====================

  /**
   * Get list data with localStorage caching
   * @param siteUrl SharePoint site URL
   * @param listName List name
   * @param cacheKey Cache key for localStorage
   * @param selectFields Fields to select
   * @param filter OData filter
   * @param orderBy OData orderby
   * @param expiryHours Cache expiry in hours (default 24)
   * @returns Cached or fresh list items
   */
  public async getCachedListData(
    siteUrl: string,
    listName: string,
    cacheKey: string,
    selectFields?: string,
    filter?: string,
    orderBy?: string,
    expiryHours: number = 24
  ): Promise<any[]> {
    try {
      // Check cache
      const cached = this.getCacheItem(cacheKey);

      if (cached && !this.isCacheExpired(cached)) {
        return cached.data;
      }

      // Fetch fresh data
      const data = await this.getAllListItems(siteUrl, listName, selectFields, filter, orderBy);

      // Cache the data
      this.setCacheItem(cacheKey, data, expiryHours);

      return data;

    } catch (error) {
      console.error(`❌ Error in getCachedListData:`, error);

      // Try to return expired cache as fallback
      const cached = this.getCacheItem(cacheKey);
      if (cached) {
        console.warn(`⚠️ Returning expired cache for ${cacheKey} due to error`);
        return cached.data;
      }

      throw error;
    }
  }

  /**
   * Get workflow details (TD_iRequest_Workflow_Details)
   * NO CACHING - Always fetches fresh
   */
  public async getTDWorkflowDetails(siteUrl: string): Promise<any[]> {
    return this.getAllListItems(
      siteUrl,
      'TD_iRequest_Workflow_Details',
      '',
      '',
      'fields/Created desc'
    );
  }

  /**
   * Get TD_iStaff_Workflow_Details list items
   * NO CACHING - Always fetches fresh
   */
  public async getTDiStaffWorkflowDetails(siteUrl: string): Promise<any[]> {
    return this.getAllListItems(
      siteUrl,
      'TD_iStaff_Workflow_Details',
      '',
      '',
      'fields/Created desc'
    );
  }

  /**
   * Get combined workflow details from both TD_iRequest and TD_iStaff lists
   * NO CACHING - Always fetches fresh from both lists
   */
  public async getCombinedWorkflowDetails(siteUrl: string): Promise<any[]> {

    try {
      // Fetch both lists in parallel
      const [iRequestData, iStaffData] = await Promise.all([
        this.getTDWorkflowDetails(siteUrl),
        this.getTDiStaffWorkflowDetails(siteUrl)
      ]);

      // Combine both arrays
      const combinedData = [...iRequestData, ...iStaffData];

      return combinedData;
    } catch (error) {
      console.error('❌ Error fetching combined workflow details:', error);
      return [];
    }
  }

  /**
   * Get iRequest forms (CD_iRequest_Form)
   * 5-minute cache
   */
  public async getCDiRequestForms(siteUrl: string): Promise<any[]> {
    return this.getCachedListData(
      siteUrl,
      'CD_iRequest_Form',
      'listCache_CD_iRequest_Form',
      '',
      '',
      'Created desc',
      5 / 60 // 5 minutes in hours
    );
  }

  /**
   * Get iStaff forms (CD_iStaff_Form)
   * 5-minute cache
   */
  public async getCDiStaffForms(siteUrl: string): Promise<any[]> {
    return this.getCachedListData(
      siteUrl,
      'CD_iStaff_Form',
      'listCache_CD_iStaff_Form',
      '',
      '',
      'Created desc',
      5 / 60 // 5 minutes in hours
    );
  }

  /**
   * Find workflow details by iRequest number
   */
  public findWorkflowDetailsByRequestNo(workflowData: any[], iRequestNo: string): any | null {
    if (!workflowData || !iRequestNo) return null;

   if (iRequestNo.startsWith('S')){
    return workflowData.find(item =>
      item.iStaff_x0020_No === iRequestNo &&
      item.Category.startsWith('iStaff')
    ) || null;
   }else{
    return workflowData.find(item =>
      item.iRequest_x0020_No === iRequestNo ||
      item.iRequestNo === iRequestNo ||
      item['iRequest No'] === iRequestNo
    ) || null;
   }
  }

  /**
   * Get requestor and status by iRequest number
   */
  public async getRequestorAndStatus(
    siteUrl: string,
    iRequestNo: string
  ): Promise<{ onBehalfOf: string, status: string } | null> {
    try {
      const workflowData = await this.getTDWorkflowDetails(siteUrl);
      const workflowItem = this.findWorkflowDetailsByRequestNo(workflowData, iRequestNo);

      if (!workflowItem) return null;

      return {
        onBehalfOf: workflowItem.On_x0020_Behalf_x0020_of_x0020_R || workflowItem.OnBehalfOfR || '',
        status: workflowItem.Status || ''
      };

    } catch (error) {
      console.error('❌ Error getting requestor and status:', error);
      return null;
    }
  }

  // ==================== CACHE MANAGEMENT ====================

  private getCacheItem(cacheKey: string): CacheItem | null {
    try {
      const cached = localStorage.getItem(cacheKey);
      if (!cached) return null;

      return JSON.parse(cached) as CacheItem;
    } catch (error) {
      console.error(`❌ Error reading cache ${cacheKey}:`, error);
      return null;
    }
  }

  private setCacheItem(cacheKey: string, data: any[], expiryHours: number): void {
    try {
      const cacheItem: CacheItem = {
        data,
        timestamp: Date.now(),
        expiryTime: Date.now() + (expiryHours * 60 * 60 * 1000)
      };

      localStorage.setItem(cacheKey, JSON.stringify(cacheItem));

    } catch (error) {
      console.error(`❌ Error setting cache ${cacheKey}:`, error);

      // Try to clear expired caches if storage is full
      this.clearExpiredCaches();
    }
  }

  private isCacheExpired(cacheItem: CacheItem): boolean {
    return Date.now() > cacheItem.expiryTime;
  }

  /**
   * Clear all expired caches from localStorage
   */
  public clearExpiredCaches(): void {
    try {
      const keysToRemove: string[] = [];

      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith('listCache_')) {
          const cached = this.getCacheItem(key);
          if (cached && this.isCacheExpired(cached)) {
            keysToRemove.push(key);
          }
        }
      }

      keysToRemove.forEach(key => {
        localStorage.removeItem(key);
      });

    } catch (error) {
      console.error('❌ Error clearing expired caches:', error);
    }
  }

  /**
   * Clear specific cache by key
   */
  public clearCache(cacheKey: string): void {
    try {
      localStorage.removeItem(cacheKey);
    } catch (error) {
      console.error(`❌ Error clearing cache ${cacheKey}:`, error);
    }
  }

  /**
   * Get cache information for diagnostics
   */
  public getCacheInfo(cacheKey: string): any {
    const cached = this.getCacheItem(cacheKey);

    if (!cached) {
      return {
        exists: false,
        itemCount: 0,
        timestamp: null,
        expiryTime: null,
        isExpired: null,
        sizeKB: 0
      };
    }

    return {
      exists: true,
      itemCount: cached.data.length,
      timestamp: new Date(cached.timestamp),
      expiryTime: new Date(cached.expiryTime),
      isExpired: this.isCacheExpired(cached),
      sizeKB: Math.round(JSON.stringify(cached).length / 1024)
    };
  }

  /**
   * Apply client-side sorting to items array
   * @param items Array of items to sort
   * @param orderBy Sort expression (e.g., "Id desc" or "Created asc")
   * @returns Sorted array
   */
  private _applySorting(items: any[], orderBy: string): any[] {
    if (!orderBy || items.length === 0) {
      return items;
    }

    // Parse orderBy string (e.g., "Id desc" or "Created asc")
    const parts = orderBy.trim().split(/\s+/);
    const fieldName = parts[0];
    const direction = (parts[1] || 'asc').toLowerCase();

    // Create a copy and sort
    return items.slice().sort((a: any, b: any) => {
      const valueA = a[fieldName];
      const valueB = b[fieldName];

      // Handle null/undefined
      if (valueA === null || valueA === undefined) return 1;
      if (valueB === null || valueB === undefined) return -1;

      // Numeric comparison (for Id fields)
      if (typeof valueA === 'number' && typeof valueB === 'number') {
        return direction === 'desc' ? valueB - valueA : valueA - valueB;
      }

      // Try to parse as numbers (for string representations of numbers like Id)
      const numA = parseInt(valueA, 10);
      const numB = parseInt(valueB, 10);
      if (!isNaN(numA) && !isNaN(numB)) {
        return direction === 'desc' ? numB - numA : numA - numB;
      }

      // String comparison
      const strA = String(valueA).toLowerCase();
      const strB = String(valueB).toLowerCase();

      if (direction === 'desc') {
        return strB.localeCompare(strA);
      } else {
        return strA.localeCompare(strB);
      }
    });
  }
}
