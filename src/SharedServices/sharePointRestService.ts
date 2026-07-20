import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface CacheItem {
  data: any[];
  timestamp: number;
  expiryTime: number;
}

export interface ISharePointRestService {
  getListItems(siteUrl: string, listName: string, selectFields?: string, filter?: string, orderBy?: string, top?: number): Promise<any[]>;
  getListItemById(siteUrl: string, listName: string, itemId: number, selectFields?: string): Promise<any>;
  createListItem(siteUrl: string, listName: string, itemData: any): Promise<any>;
  updateListItem(siteUrl: string, listName: string, itemId: number, itemData: any): Promise<any>;
  deleteListItem(siteUrl: string, listName: string, itemId: number): Promise<void>;
}

export class SharePointRestService implements ISharePointRestService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async getListItems(
    siteUrl: string, 
    listName: string, 
    selectFields?: string, 
    filter?: string, 
    orderBy?: string, 
    top?: number
  ): Promise<any[]> {
    try {
      // Build the REST API URL
      let restUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
      
      // Build query parameters
      const queryParams: string[] = [];
      
      if (selectFields) {
        //queryParams.push(`$select=${selectFields}`);
      }
      
      if (filter) {
        queryParams.push(`$filter=${encodeURIComponent(filter)}`);
      }
      
      if (orderBy) {
        queryParams.push(`$orderby=${orderBy}`);
      }
      
      if (top) {
        queryParams.push(`$top=${top}`);
      }
      
      if (queryParams.length > 0) {
        restUrl += `?${queryParams.join('&')}`;
      }

      // Make the REST API call
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        restUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status} - ${response.statusText}`);
      }

      const jsonResponse = await response.json();
      return jsonResponse.value || [];
      
    } catch (error) {
      console.error('Error fetching list items:', error);
      throw error;
    }
  }

  public async getListItemById(
    siteUrl: string, 
    listName: string, 
    itemId: number, 
    selectFields?: string
  ): Promise<any> {
    try {
      let restUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`;
      
      if (selectFields) {
        restUrl += `?$select=${selectFields}`;
      }

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        restUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status} - ${response.statusText}`);
      }

      return await response.json();
      
    } catch (error) {
      console.error('Error fetching list item:', error);
      throw error;
    }
  }

  public async createListItem(siteUrl: string, listName: string, itemData: any): Promise<any> {
    try {
      const restUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        restUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
          },
          body: JSON.stringify(itemData)
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status} - ${response.statusText}`);
      }

      return await response.json();
      
    } catch (error) {
      console.error('Error creating list item:', error);
      throw error;
    }
  }

  public async updateListItem(
    siteUrl: string, 
    listName: string, 
    itemId: number, 
    itemData: any
  ): Promise<any> {
    try {
      const restUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        restUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify(itemData)
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status} - ${response.statusText}`);
      }

      return await response.json();
      
    } catch (error) {
      console.error('Error updating list item:', error);
      throw error;
    }
  }

  public async deleteListItem(siteUrl: string, listName: string, itemId: number): Promise<void> {
    try {
      const restUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        restUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status} - ${response.statusText}`);
      }
      
    } catch (error) {
      console.error('Error deleting list item:', error);
      throw error;
    }
  }

  // Helper method to get all items with pagination support
  public async getAllListItems(
    siteUrl: string, 
    listName: string, 
    selectFields?: string, 
    filter?: string, 
    orderBy?: string
  ): Promise<any[]> {
    let allItems: any[] = [];
    let skipToken = '';
    const batchSize = 5000;
    let hasMore = true;

    while (hasMore) {
      try {
        // Build the REST API URL with skiptoken for pagination
        let restUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
        
        // Build query parameters
        const queryParams: string[] = [];
        
        if (selectFields) {
          queryParams.push(`$select=${selectFields}`);
        }
        
        if (filter) {
          queryParams.push(`$filter=${encodeURIComponent(filter)}`);
        }
        
        if (orderBy) {
          queryParams.push(`$orderby=${orderBy}`);
        }
        
        queryParams.push(`$top=${batchSize}`);
        
        if (skipToken) {
          queryParams.push(`$skiptoken=${encodeURIComponent(skipToken)}`);
        }
        
        if (queryParams.length > 0) {
          restUrl += `?${queryParams.join('&')}`;
        }

        // Make the REST API call
        const response: SPHttpClientResponse = await this.context.spHttpClient.get(
          restUrl,
          SPHttpClient.configurations.v1
        );

        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status} - ${response.statusText}`);
        }

        const jsonResponse = await response.json();
        const items = jsonResponse.value || [];
        
        if (items.length === 0) {
          hasMore = false;
        } else {
          allItems = allItems.concat(items);
          
          // Check if there's a next link for pagination
          if (jsonResponse['odata.nextLink']) {
            // Extract skiptoken from the nextLink
            skipToken = this.extractSkipToken(jsonResponse['odata.nextLink']);
            console.log(`📄 Loaded ${items.length} items, continuing with skiptoken...`);
          } else {
            hasMore = false;
            console.log(`📄 Loaded final batch of ${items.length} items`);
          }
        }
      } catch (error) {
        console.error('Error in pagination:', error);
        hasMore = false;
      }
    }

    console.log(`✅ Total items loaded: ${allItems.length}`);
    return allItems;
  }

  // Enhanced method to get list schema/fields with better filtering
  public async getListFields(siteUrl: string, listName: string): Promise<any[]> {
    try {
      const restUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        restUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status} - ${response.statusText}`);
      }

      const jsonResponse = await response.json();
      
      // Filter out system fields and complex field types
      const filteredFields = (jsonResponse.value || []).filter((field: any) => {
        const excludedFields = [
          'Attachments', 'GUID', 'WorkflowVersion', 'WorkflowInstanceID', 
          'ParentVersionString', 'ParentLeafName', 'DocIcon', 'ServerUrl',
          'EncodedAbsUrl', 'BaseName', 'MetaInfo', 'TimeLastModified', 
          'TimeCreated', 'CheckoutUser', 'CheckedOutUserId', 'IsCheckedoutToLocal',
          'ContentTypeId', 'InstanceID', 'Order', 'ScopeId', 'HTML_x0020_File_x0020_Type',
          'Edit', 'owshiddenversion', 'WorkflowAssociation', 'NoExecute',
          'ContentVersion', 'ItemChildCount', 'FolderChildCount', 'AppAuthor',
          'AppEditor', 'SelectTitle', 'SelectFilename', 'Edit', 'Type', 'UrlNoMenu',
          'ServerRedirectedEmbedUri', 'ServerRedirectedEmbedUrl', 'FileSizeDisplay',
          'StreamHash', 'CheckoutUser', 'ComplianceAssetId'
        ];

        const excludedTypes = [
          12, // Lookup (complex)
          20, // User (complex) 
          7,  // GridChoice (complex)
          15, // File (system)
          16, // Computed (system)
          24, // AllDayEvent (system)
          25, // Recurrence (system)
          26, // Attachments (system)
          27, // User Multi (complex)
          28, // Lookup Multi (complex)
          29, // Moderation Status (system)
          30  // Cross Project Link (system)
        ];

        return !field.ReadOnlyField && 
               !field.Hidden && 
               !excludedFields.includes(field.InternalName) &&
               !excludedTypes.includes(field.FieldTypeKind) &&
               !field.InternalName.startsWith('_') && // System fields often start with _
               !field.InternalName.includes('x0020') && // Encoded space often indicates system field
               field.InternalName !== 'ContentType'; // ContentType is special
      });

      return filteredFields;
      
    } catch (error) {
      console.error('Error fetching list fields:', error);
      throw error;
    }
  }

  // Helper method for cross-site collection calls with intelligent field discovery
  public async getCrossSiteListItemsWithAllFields(
    dataSiteUrl: string,
    listName: string,
    filter?: string,
    orderBy?: string,
    top?: number
  ): Promise<{ items: any[], fields: any[], selectFields: string }> {
    try {
      // First get the filtered list fields
      const fields = await this.getListFields(dataSiteUrl, listName);
      
      // Create select string from available fields
      const selectableFields = fields.map(field => field.InternalName);
      
      // Always include essential fields that might be useful
      const essentialFields = ['Id', 'Title', 'Created', 'Modified', 'Author/Title', 'Editor/Title'];
      
      // Combine and remove duplicates using native ES6 Set
      const allSelectFields = [...new Set([...essentialFields, ...selectableFields])];
      
      const selectString = allSelectFields.join(',');

      console.log(`Loading list '${listName}' with ${fields.length} fields:`, 
        fields.map(f => `${f.InternalName} (${f.Title}) - ${f.TypeAsString}`));

      // Get the list items with dynamic fields
      const items = await this.getListItems(dataSiteUrl, listName, selectString, filter, orderBy, top);
      
      return {
        items,
        fields,
        selectFields: selectString
      };
      
    } catch (error) {
      console.error('Error fetching list items with fields:', error);
      throw error;
    }
  }

  // Helper method to auto-detect common field patterns for filtering
  public getCommonFilterFields(fields: any[]): { categoryField?: string, subcategoryField?: string, statusField?: string, dateField?: string } {
    const result: any = {};
    
    // Common patterns for category fields
    const categoryPatterns = ['category', 'cat', 'type', 'classification'];
    const subcategoryPatterns = ['subcategory', 'subcat', 'subtype', 'detail'];
    const statusPatterns = ['status', 'state', 'condition', 'stage'];
    const datePatterns = ['created', 'submitted', 'date', 'time'];

    fields.forEach(field => {
      const fieldName = field.InternalName.toLowerCase();
      const fieldTitle = (field.Title || '').toLowerCase();
      
      // Check for category field
      if (!result.categoryField && categoryPatterns.some(pattern => 
        fieldName.includes(pattern) || fieldTitle.includes(pattern))) {
        result.categoryField = field.InternalName;
      }
      
      // Check for subcategory field
      if (!result.subcategoryField && subcategoryPatterns.some(pattern => 
        fieldName.includes(pattern) || fieldTitle.includes(pattern))) {
        result.subcategoryField = field.InternalName;
      }
      
      // Check for status field
      if (!result.statusField && statusPatterns.some(pattern => 
        fieldName.includes(pattern) || fieldTitle.includes(pattern))) {
        result.statusField = field.InternalName;
      }
      
      // Check for date field (prefer Created if exists)
      if (!result.dateField && datePatterns.some(pattern => 
        fieldName.includes(pattern) || fieldTitle.includes(pattern))) {
        result.dateField = field.InternalName;
      }
    });

    return result;
  }

  // ========== CACHING FUNCTIONALITY ==========

  /**
   * Get cached list data or fetch fresh data if cache is expired/missing
   * @param siteUrl - SharePoint site URL
   * @param listName - Name of the SharePoint list
   * @param cacheKey - Unique key for caching
   * @param selectFields - Fields to select
   * @param filter - OData filter
   * @param orderBy - Order by clause
   * @param expiryHours - Hours until cache expires (default: 24)
   * @returns Promise<any[]> - Array of list items
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
    // Check if data exists in localStorage
    const cachedItem = this.getCacheItem(cacheKey);
    
    if (cachedItem && !this.isCacheExpired(cachedItem) && !(cachedItem.data.length == 0)) {
      console.log(`📦 Using cached data for ${listName} (${cachedItem.data.length} items)`);
      return cachedItem.data;
    }
    
    console.log(`🔄 Fetching fresh data for ${listName} (cache ${cachedItem ? 'expired' : 'not found'})`);
    
    try {
      // Fetch all items from SharePoint (handles pagination automatically)
      const freshData = await this.getAllListItems(siteUrl, listName, selectFields, filter, orderBy);
      
      // Store in cache
      this.setCacheItem(cacheKey, freshData, expiryHours);
      
      console.log(`✅ Cached ${freshData.length} items for ${listName}`);
      return freshData;
      
    } catch (error) {
      console.error(`❌ Error fetching data for ${listName}:`, error);
      
      // Return cached data if available, even if expired, as fallback
      if (cachedItem) {
        console.log(`⚠️ Returning expired cache as fallback for ${listName}`);
        return cachedItem.data;
      }
      
      throw error;
    }
  }


  /**
   * Get TD_iRequest_Workflow_Details data without caching
   * @param siteUrl - SharePoint site URL
   * @returns Promise<any[]> - All workflow details
   */
  public async getTDWorkflowDetails(siteUrl: string): Promise<any[]> {
    const selectFields = '*';
    
    try {
      // Fetch fresh data directly without caching
      const freshData = await this.getAllListItems(
        siteUrl,
        'TD_iRequest_Workflow_Details',
        selectFields,
        undefined, // no filter - get all items
        'Created desc' // order by created date
      );
      
      console.log(`✅ Fetched ${freshData.length} items from TD_iRequest_Workflow_Details (no cache)`);
      return freshData;
      
    } catch (error) {
      console.error('❌ Error fetching TD_iRequest_Workflow_Details data:', error);
      throw error;
    }
  }

  /**
   * Get cached CD_iRequest_Forms data
   * @param siteUrl - SharePoint site URL
   * @returns Promise<any[]> - All iRequest forms
   */
  public async getCDiRequestForms(siteUrl: string): Promise<any[]> {
    const cacheKey = 'listCache_CD_iRequest_Form';
    const selectFields = '*'; // Get all fields for forms
    
    return await this.getCachedListData(
      siteUrl,
      'CD_iRequest_Form',
      cacheKey,
      selectFields,
      undefined, // no filter - get all items
      'Created desc', // order by created date
      5/60
    );
  }

  /**
   * Find workflow details by iRequest number
   * @param workflowData - The cached workflow data
   * @param iRequestNo - The iRequest number to search for
   * @returns any | null - The matching workflow item or null
   */
  public findWorkflowDetailsByRequestNo(workflowData: any[], iRequestNo: string): any | null {
    if (!workflowData || !iRequestNo) return null;
    
    const found = workflowData.find((item: any) => 
      item.iRequest_x0020_No && item.iRequest_x0020_No.toString() === iRequestNo.toString()
    );
    
    return found || null;
  }

  /**
   * Get On Behalf of Requestor and Status for a specific iRequest number
   * @param siteUrl - SharePoint site URL  
   * @param iRequestNo - The iRequest number to lookup
   * @returns Promise<{onBehalfOf: string, status: string} | null>
   */
  public async getRequestorAndStatus(siteUrl: string, iRequestNo: string): Promise<{onBehalfOf: string, status: string} | null> {
    try {
      const workflowData = await this.getTDWorkflowDetails(siteUrl);
      const workflowItem = this.findWorkflowDetailsByRequestNo(workflowData, iRequestNo);
      
      if (!workflowItem) {
        return null;
      }
      
      return {
        onBehalfOf: workflowItem.On_x0020_Behalf_x0020_of_x0020_R || '',
        status: workflowItem.Status || ''
      };
    } catch (error) {
      console.error(`Error getting requestor and status for ${iRequestNo}:`, error);
      return null;
    }
  }

  // ========== CACHE HELPER METHODS ==========

  /**
   * Extract skip token from OData next URL
   */
  private extractSkipToken(nextUrl: string): string {
    const match = nextUrl.match(/\$skiptoken=([^&]+)/);
    return match ? decodeURIComponent(match[1]) : '';
  }

  /**
   * Get cached item from localStorage
   */
  private getCacheItem(cacheKey: string): CacheItem | null {
    try {
      const cached = localStorage.getItem(cacheKey);
      if (!cached) return null;
      
      return JSON.parse(cached) as CacheItem;
    } catch (error) {
      console.error('Error reading from cache:', error);
      return null;
    }
  }

  /**
   * Store data in localStorage with expiry
   */
  private setCacheItem(cacheKey: string, data: any[], expiryHours: number): void {
    try {
      const now = Date.now();
      const cacheItem: CacheItem = {
        data: data,
        timestamp: now,
        expiryTime: now + (expiryHours * 60 * 60 * 1000)
      };
      
      localStorage.setItem(cacheKey, JSON.stringify(cacheItem));
    } catch (error) {
      console.error('Error storing to cache:', error);
      // If localStorage is full, try to clear old caches
      this.clearExpiredCaches();
    }
  }

  /**
   * Check if cache item is expired
   */
  private isCacheExpired(cacheItem: CacheItem): boolean {
    return Date.now() > cacheItem.expiryTime;
  }

  /**
   * Clear expired caches to free up localStorage space
   */
  public clearExpiredCaches(): void {
    try {
      const keysToRemove: string[] = [];
      
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (!key || !key.startsWith('listCache_')) continue;
        
        const item = this.getCacheItem(key);
        if (item && this.isCacheExpired(item)) {
          keysToRemove.push(key);
        }
      }
      
      keysToRemove.forEach(key => {
        localStorage.removeItem(key);
        console.log(`🗑️ Removed expired cache: ${key}`);
      });
      
    } catch (error) {
      console.error('Error clearing expired caches:', error);
    }
  }

  /**
   * Manually clear specific cache
   */
  public clearCache(cacheKey: string): void {
    try {
      localStorage.removeItem(cacheKey);
      console.log(`🗑️ Manually cleared cache: ${cacheKey}`);
    } catch (error) {
      console.error('Error clearing cache:', error);
    }
  }

  /**
   * Get cache statistics for debugging
   */
  public getCacheInfo(cacheKey: string): any {
    const cacheItem = this.getCacheItem(cacheKey);
    if (!cacheItem) {
      return { exists: false };
    }
    
    return {
      exists: true,
      itemCount: cacheItem.data.length,
      timestamp: new Date(cacheItem.timestamp),
      expiryTime: new Date(cacheItem.expiryTime),
      isExpired: this.isCacheExpired(cacheItem),
      sizeKB: Math.round(JSON.stringify(cacheItem).length / 1024)
    };
  }
}
