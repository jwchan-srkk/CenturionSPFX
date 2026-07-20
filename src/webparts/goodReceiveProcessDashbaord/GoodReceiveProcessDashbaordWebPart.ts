import { escape as escapeHtml } from '@microsoft/sp-lodash-subset';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { Version } from '@microsoft/sp-core-library';
import * as $ from 'jquery';
import 'datatables.net';

import styles from './GoodReceiveProcessDashbaordWebPart.module.scss';
import * as strings from 'GoodReceiveProcessDashbaordWebPartStrings';
import { SharePointRestService } from '../../SharedServices/sharePointRestService';

const TABLE_ID = 'goodReceiveProcessDashboardTable';

export interface IGoodReceiveProcessDashbaordWebPartProps {
  description: string;
  sharePointSiteUrl: string;
  listName: string;
  actionLink: string;
}

interface IGoodReceiveRow {
  id: number;
  poNumber: string;
  doNumber: string;
  invoiceDate: string;
  invoiceDateValue: number;
  dateReceived: string;
  dateReceivedValue: number;
  vendorName: string;
  emailSubject: string;
  status: string;
  rawItem: Record<string, unknown>;
}

export default class GoodReceiveProcessDashbaordWebPart extends BaseClientSideWebPart<IGoodReceiveProcessDashbaordWebPartProps> {
  private _sharePointRestService!: SharePointRestService;
  private _rows: IGoodReceiveRow[] = [];
  private _dataTable: any = null;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sharePointRestService = new SharePointRestService(this.context);
  }

  private _ensureFullWidthLayout(): void {
    const styleId = 'srkk-dashboard-canvas-override';
    if (document.getElementById(styleId)) {
      this._tagFullWidthContainers();
      return;
    }

    const style = document.createElement('style');
    style.id = styleId;
    style.innerHTML = `
      @media screen and (min-width: 1024px) {
        .srkk-full-width-layout,
        .srkk-full-width-section,
        .srkk-full-width-zone,
        .srkk-full-width-control {
          max-width: 100% !important;
          width: 100% !important;
          flex-basis: 100% !important;
        }
      }
    `;
    document.head.appendChild(style);
    this._tagFullWidthContainers();
  }

  private _tagFullWidthContainers(): void {
    const sectionContainer = this.domElement.closest('[data-automation-id="CanvasZone-SectionContainer"], .CanvasZone-SectionContainer');
    const section = this.domElement.closest('[data-automation-id="CanvasSection"], .CanvasSection');
    const control = this.domElement.closest('[data-automation-id="CanvasControl"], .ControlZone');
    const zone = this.domElement.closest('[data-automation-id="CanvasZone"], .CanvasZone');
    const layout = this.domElement.closest('[data-automation-id="CanvasLayout"], .CanvasLayout');

    if (sectionContainer instanceof HTMLElement) {
      sectionContainer.classList.add('srkk-full-width-layout');
      sectionContainer.style.width = '100%';
      sectionContainer.style.maxWidth = '100%';
    }

    if (layout instanceof HTMLElement) {
      layout.classList.add('srkk-full-width-layout');
      layout.style.width = '100%';
      layout.style.maxWidth = '100%';
    }

    if (section instanceof HTMLElement) {
      section.classList.add('srkk-full-width-section');
      section.style.width = '100%';
      section.style.maxWidth = '100%';
      section.style.flexBasis = '100%';
    }

    if (control instanceof HTMLElement) {
      control.classList.add('srkk-full-width-control');
      control.style.width = '100%';
      control.style.maxWidth = '100%';
      control.style.flexBasis = '100%';
    }

    if (zone instanceof HTMLElement) {
      zone.classList.add('srkk-full-width-zone');
      zone.style.width = '100%';
      zone.style.maxWidth = '100%';
      zone.style.flexBasis = '100%';
    }

    this.domElement.style.width = '100%';
    this.domElement.style.maxWidth = '100%';
    this.domElement.style.display = 'block';
    this.domElement.style.flexBasis = '100%';
  }

  public render(): void {
    this._ensureFullWidthLayout();
    this._renderShell('Loading dashboard records...');
    void this._loadDashboardData();
  }

  protected onDispose(): void {
    this._destroyDataTable();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (
      oldValue !== newValue &&
      (propertyPath === 'sharePointSiteUrl' ||
        propertyPath === 'listName' ||
        propertyPath === 'actionLink')
    ) {
      this.render();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: 'Dashboard Settings',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('sharePointSiteUrl', {
                  label: 'SharePoint Site URL'
                }),
                PropertyPaneTextField('listName', {
                  label: 'List Name'
                }),
                PropertyPaneTextField('actionLink', {
                  label: 'Action Link (use {ColumnInternalName} for dynamic values, e.g. https://example.com?id={Id})'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // ─── Data loading ────────────────────────────────────────────────────────────

  private async _loadDashboardData(): Promise<void> {
    try {
      const selectFields = [
        'Id',
        'PONumber',
        'DoNumber',
        'InvoiceDate',
        'DateReceived',
        'VendorName',
        'EmailSubject',
        'Status'
      ].join(',');

      const items: any[] = await this._sharePointRestService.getListItems(
        this._getConfiguredSiteUrl(),
        this._getConfiguredListName(),
        selectFields,
        '',
        'Id desc',
        5000
      );

      this._rows = items.map((item: any) => this._mapItemToRow(item));
      this._renderDashboard();
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error occurred while loading the dashboard.';
      this._renderShell(`Unable to load the dashboard. ${escapeHtml(message)}`, true);
    }
  }

  private _mapItemToRow(item: any): IGoodReceiveRow {
    const invoiceDate = item.InvoiceDate ? new Date(item.InvoiceDate) : undefined;
    const dateReceived = item.DateReceived ? new Date(item.DateReceived) : undefined;

    return {
      id: Number(item.Id) || 0,
      poNumber: this._getTextValue(item.PONumber),
      doNumber: this._getTextValue(item.DoNumber),
      invoiceDate: invoiceDate ? this._formatDate(invoiceDate) : '-',
      invoiceDateValue: invoiceDate ? invoiceDate.getTime() : 0,
      dateReceived: dateReceived ? this._formatDateTime(dateReceived) : '-',
      dateReceivedValue: dateReceived ? dateReceived.getTime() : 0,
      vendorName: this._getTextValue(item.VendorName),
      emailSubject: this._getTextValue(item.EmailSubject),
      status: this._getTextValue(item.Status),
      rawItem: item as Record<string, unknown>
    };
  }

  // ─── Rendering ───────────────────────────────────────────────────────────────

  private _renderDashboard(): void {
    const bodyMarkup = this._rows.length > 0
      ? `
        <div class="${styles.tableContainer}">
          <table id="${TABLE_ID}" class="${styles.dashboardTable}">
            <thead>
              <tr>
                <th>PO Number</th>
                <th>DO Number</th>
                <th>Date Received</th>
                <th>Vendor Name</th>
                <th>Email Subject</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>`
      : `<div class="${styles.emptyState}">No records were found in the configured list.</div>`;

    this.domElement.innerHTML = `
      <section class="${styles.goodReceiveProcessDashbaord}">
        <div class="${styles.dashboardCard}">
          <div class="${styles.headerRow}">
            <div>
              <h2 class="${styles.title}">Good Receive Process Dashboard</h2>
            </div>
            <div class="${styles.summaryBlock}">
              <span class="${styles.summaryLabel}">Total records</span>
              <span class="${styles.summaryValue}">${this._rows.length}</span>
            </div>
          </div>
          ${bodyMarkup}
        </div>
      </section>`;

    if (this._rows.length > 0) {
      this._initializeDataTable();
    } else {
      this._destroyDataTable();
    }
  }

  private _renderShell(message: string, isError: boolean = false): void {
    this._destroyDataTable();

    this.domElement.innerHTML = `
      <section class="${styles.goodReceiveProcessDashbaord}">
        <div class="${styles.dashboardCard}">
          <div class="${styles.headerRow}">
            <div>
              <h2 class="${styles.title}">Good Receive Process Dashboard</h2>
            </div>
            <div class="${styles.summaryBlock}">
              <span class="${styles.summaryLabel}">Total records</span>
              <span class="${styles.summaryValue}">0</span>
            </div>
          </div>
          <div class="${isError ? styles.errorState : styles.loadingState}">${message}</div>
        </div>
      </section>`;
  }

  // ─── DataTable ───────────────────────────────────────────────────────────────

  private _initializeDataTable(): void {
    this._destroyDataTable();

    const tableSelector = `#${TABLE_ID}`;
    const table = $(this.domElement).find(tableSelector);

    this._dataTable = (table as any).DataTable({
      autoWidth: false,
      data: this._rows,
      deferRender: true,
      dom: '<"top"lf>rt<"bottom"ip>',
      lengthMenu: [[10, 25, 50, 100], [10, 25, 50, 100]],
      order: [[2, 'desc']],
      pageLength: 10,
      paging: true,
      searching: true,
      info: true,
      ordering: true,
      columns: [
        {
          data: 'poNumber',
          width: '140px',
          render: (_data: string, type: string, row: IGoodReceiveRow): string => {
            if (type !== 'display') {
              return row.poNumber || '';
            }

            if (row.poNumber === '-' || !row.poNumber) {
              return '-';
            }

            const href = this._buildActionLink(row);
            if (!href) {
              return escapeHtml(row.poNumber);
            }

            return `<a class="${styles.referenceLink}" href="${escapeHtml(href)}" target="_blank" rel="noreferrer">${escapeHtml(row.poNumber)}</a>`;
          }
        },
        { data: 'doNumber', width: '130px' },
        {
          data: 'dateReceived',
          width: '160px',
          render: (_data: string, type: string, row: IGoodReceiveRow): string | number => {
            if (type === 'sort' || type === 'type') {
              return row.dateReceivedValue;
            }

            return row.dateReceived || '-';
          }
        },
        { data: 'vendorName', width: '200px' },
        { data: 'emailSubject', width: '260px' },
        {
          data: 'status',
          width: '140px',
          render: (data: string, type: string): string => {
            if (type !== 'display') {
              return data || '';
            }

            const safeStatus = data || 'N/A';
            return `<span class="${styles.statusBadge} ${this._getStatusBadgeClassName(safeStatus)}">${escapeHtml(safeStatus)}</span>`;
          }
        }
      ],
      language: {
        info: 'Showing _START_ to _END_ of _TOTAL_ entries',
        lengthMenu: 'Show _MENU_ entries',
        search: 'Search:',
        paginate: {
          previous: 'Previous',
          next: 'Next'
        }
      }
    });
  }

  private _destroyDataTable(): void {
    if (this._dataTable) {
      this._dataTable.destroy(true);
      this._dataTable = null;
    }
  }

  // ─── Helpers ─────────────────────────────────────────────────────────────────

  private _getConfiguredSiteUrl(): string {
    return (this.properties.sharePointSiteUrl || '').trim() || this.context.pageContext.web.absoluteUrl;
  }

  private _getConfiguredListName(): string {
    return (this.properties.listName || '').trim();
  }

  /**
   * Resolves the action link template, substituting {ColumnInternalName} tokens
   * with the actual values from the SharePoint item. E.g. if the template is
   * "https://example.com?id={Id}&po={PONumber}" it produces the correct URL.
   */
  private _buildActionLink(row: IGoodReceiveRow): string {
    const template = (this.properties.actionLink || '').trim();
    if (!template) {
      return '';
    }

    return template.replace(/\{([^}]+)\}/g, (_match: string, token: string): string => {
      const value = (row.rawItem as Record<string, unknown>)[token];
      if (value === null || value === undefined) {
        return '';
      }

      return encodeURIComponent(String(value));
    });
  }

  private _getTextValue(value: unknown): string {
    if (value === null || value === undefined) {
      return '-';
    }

    const text = String(value).trim();
    return text ? text : '-';
  }

  /** Formats a Date as "dd MMM yyyy" e.g. 07 Jan 2025 */
  private _formatDate(date: Date): string {
    const formatter = new Intl.DateTimeFormat('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    });

    return formatter.format(date);
  }

  /** Formats a Date as "dd MMM yyyy hh:mm AM/PM" e.g. 07 Jan 2025 02:30 PM */
  private _formatDateTime(date: Date): string {
    const day = this._pad(date.getDate());
    const month = date.toLocaleString('en-GB', { month: 'short' });
    const year = date.getFullYear();

    let hours = date.getHours();
    const minutes = this._pad(date.getMinutes());
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12 || 12;
    const hoursStr = this._pad(hours);

    return `${day} ${month} ${year} ${hoursStr}:${minutes} ${ampm}`;
  }

  private _pad(value: number): string {
    return value < 10 ? `0${value}` : `${value}`;
  }

  private _getStatusBadgeClassName(status: string): string {
    const s = status.toLowerCase().trim();

    if (s === 'completed') {
      return styles.statusSuccess;
    }

    if (s === 'require human action') {
      return styles.statusDanger;
    }

    return styles.statusWarning;
  }
}
