/**
 * Utility class for transforming data between Microsoft Graph API and SharePoint REST API formats
 * Handles field name transformations, data flattening, and format normalization
 */

export class GraphFieldMapper {
  /**
   * Known field name mappings between SharePoint internal names and Graph API names
   * Add mappings here as they are discovered during testing
   */
  private static readonly FIELD_MAPPINGS = new Map<string, string>([
    // Common SharePoint field encodings
    ['Submitter_x0020_Email', 'SubmitterEmail'],
    ['On_x0020_Behalf_x0020_of_x0020_R', 'OnBehalfOfR'],
    ['iRequest_x0020_No', 'iRequestNo'],
    ['iRequest_x0020_Request_x0020_ID', 'iRequestRequestID'],
    ['Submitted_x0020_On', 'SubmittedOn'],
    ['SBU_x0020_Code', 'SBUCode'],
    ['SBU_x0020_Name', 'SBUName'],
    // Add more as discovered
  ]);

  /**
   * Transform a Graph API item response to SharePoint REST format
   * @param graphItem Graph API item with fields object
   * @returns Flattened object matching SharePoint REST format
   */
  public static transformItemFromGraph(graphItem: any): any {
    if (!graphItem) return graphItem;

    const transformed: any = {
      Id: graphItem.id || graphItem.Id
    };

    // Flatten the fields object
    if (graphItem.fields) {
      for (const [fieldName, fieldValue] of Object.entries(graphItem.fields)) {
        // Handle lookup fields (convert {lookupId, lookupValue} to separate fields)
        if (fieldValue && typeof fieldValue === 'object' && 'lookupValue' in fieldValue) {
          const lookupData = this.flattenLookupField(fieldName, fieldValue);
          Object.assign(transformed, lookupData);
        }
        // Handle person/group fields (extract display name)
        else if (fieldValue && typeof fieldValue === 'object' && ('displayName' in fieldValue || 'email' in fieldValue)) {
          transformed[fieldName] = this.flattenPersonField(fieldValue);
        }
        // Handle multi-lookup fields (array of lookup objects)
        else if (Array.isArray(fieldValue) && fieldValue.length > 0 && 'lookupValue' in fieldValue[0]) {
          transformed[fieldName] = fieldValue.map(lv => lv.lookupValue).join('; ');
        }
        // Handle date fields (normalize format)
        else if (typeof fieldValue === 'string' && this.isISODate(fieldValue)) {
          transformed[fieldName] = this.normalizeDateField(fieldValue);
        }
        // Handle all other fields
        else {
          transformed[fieldName] = fieldValue;
        }
      }
    }

    // Preserve other top-level properties (webUrl, contentType, etc.)
    for (const [key, value] of Object.entries(graphItem)) {
      if (key !== 'fields' && key !== 'id' && !transformed[key]) {
        transformed[key] = value;
      }
    }

    return transformed;
  }

  /**
   * Flatten a lookup field from Graph format to SharePoint REST format
   * Graph: { lookupId: 5, lookupValue: "IT" }
   * SharePoint: { CategoryId: 5, Category: "IT" }
   */
  private static flattenLookupField(fieldName: string, lookupField: any): any {
    const result: any = {};

    if (lookupField.lookupId !== undefined) {
      result[`${fieldName}Id`] = lookupField.lookupId;
    }

    if (lookupField.lookupValue !== undefined) {
      result[fieldName] = lookupField.lookupValue;
    }

    return result;
  }

  /**
   * Flatten a person/group field to display name
   * Graph: { displayName: "John Doe", email: "john@example.com", id: "..." }
   * SharePoint: "John Doe"
   */
  private static flattenPersonField(personField: any): string {
    if (!personField) return '';
    if (typeof personField === 'string') return personField;

    return personField.displayName || personField.email || personField.userPrincipalName || personField.title || '';
  }

  /**
   * Normalize date field format (remove milliseconds if present)
   * Graph: 2024-01-15T10:30:00.000Z
   * SharePoint: 2024-01-15T10:30:00Z
   */
  private static normalizeDateField(dateValue: string): string {
    if (!dateValue) return dateValue;

    // Remove milliseconds for consistency
    return dateValue.replace(/\.000Z$/, 'Z');
  }

  /**
   * Check if a string is an ISO date format
   */
  private static isISODate(value: string): boolean {
    if (!value || typeof value !== 'string') return false;

    const isoDateRegex = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{3})?Z?$/;
    return isoDateRegex.test(value);
  }

  /**
   * Transform SharePoint REST filter expression to Graph API format
   * SharePoint: Status eq 'Active'
   * Graph: fields/Status eq 'Active'
   */
  public static transformFilterForGraph(spFilter: string | undefined): string | undefined {
    if (!spFilter) return undefined;

    // Add "fields/" prefix to field names in filter expression
    // This regex finds field names before operators (eq, ne, gt, lt, ge, le, startswith, etc.)
    let graphFilter = spFilter;

    // Handle common OData filter patterns
    const operators = ['eq', 'ne', 'gt', 'lt', 'ge', 'le', 'and', 'or', 'not'];
    const functions = ['startswith', 'endswith', 'substringof', 'contains'];

    // Add fields/ prefix to field names (avoid prefixing already prefixed fields)
    graphFilter = graphFilter.replace(/\b([A-Za-z_][A-Za-z0-9_]*)\s+(eq|ne|gt|lt|ge|le)\s+/g, (match, fieldName, operator) => {
      if (fieldName.startsWith('fields/')) return match;
      return `fields/${fieldName} ${operator} `;
    });

    // Handle function calls like startswith(Category, 'IT')
    functions.forEach(func => {
      const funcRegex = new RegExp(`${func}\\(([A-Za-z_][A-Za-z0-9_]*)\\s*,`, 'g');
      graphFilter = graphFilter.replace(funcRegex, (match, fieldName) => {
        if (fieldName.startsWith('fields/')) return match;
        return `${func}(fields/${fieldName},`;
      });
    });

    return graphFilter;
  }

  /**
   * Transform SharePoint REST select fields to Graph API format
   * SharePoint: Title,Category,Status
   * Graph: id,fields/Title,fields/Category,fields/Status
   */
  public static transformSelectForGraph(spSelect: string | undefined): string | undefined {
    if (!spSelect) return undefined;

    const fields = spSelect.split(',').map(f => f.trim());
    const graphFields = fields.map(field => {
      // Preserve id field
      if (field.toLowerCase() === 'id') return 'id';

      // Add fields/ prefix
      if (!field.startsWith('fields/')) {
        return `fields/${field}`;
      }

      return field;
    });

    // Always include id
    if (!graphFields.includes('id')) {
      graphFields.unshift('id');
    }

    return graphFields.join(',');
  }

  /**
   * Transform SharePoint REST orderBy to Graph API format
   * SharePoint: Created desc
   * Graph: fields/Created desc
   */
  public static transformOrderByForGraph(spOrderBy: string | undefined): string | undefined {
    if (!spOrderBy) return undefined;

    // Split by comma for multiple sort fields
    const orderBys = spOrderBy.split(',').map(ob => ob.trim());

    const graphOrderBys = orderBys.map(orderBy => {
      const parts = orderBy.split(/\s+/);
      const fieldName = parts[0];
      const direction = parts[1] || ''; // asc/desc

      if (fieldName.toLowerCase() === 'id') {
        return orderBy; // Keep id as is
      }

      if (!fieldName.startsWith('fields/')) {
        return `fields/${fieldName}${direction ? ' ' + direction : ''}`;
      }

      return orderBy;
    });

    return graphOrderBys.join(',');
  }

  /**
   * Build complete Graph API URL for list items
   */
  public static buildGraphItemsUrl(
    siteId: string,
    listId: string,
    selectFields?: string,
    filter?: string,
    orderBy?: string,
    top?: number
  ): string {
    let url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields`;

    // Add select
    const graphSelect = this.transformSelectForGraph(selectFields);
    if (graphSelect) {
      url += `&$select=${encodeURIComponent(graphSelect)}`;
    }

    // Add filter
    const graphFilter = this.transformFilterForGraph(filter);
    if (graphFilter) {
      url += `&$filter=${encodeURIComponent(graphFilter)}`;
    }

    // Add orderby
    const graphOrderBy = this.transformOrderByForGraph(orderBy);
    if (graphOrderBy) {
      url += `&$orderby=${encodeURIComponent(graphOrderBy)}`;
    }

    // Add top
    if (top) {
      url += `&$top=${top}`;
    }

    return url;
  }
}
