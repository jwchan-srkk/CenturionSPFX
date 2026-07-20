import { escape } from '@microsoft/sp-lodash-subset';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { Version } from '@microsoft/sp-core-library';
import * as $ from 'jquery';
import 'datatables.net';

import styles from './VendorInformationFormRequestorDashboardWebPart.module.scss';
import * as strings from 'VendorInformationFormRequestorDashboardWebPartStrings';
import { SharePointRestService } from '../../SharedServices/sharePointRestService';

const DEFAULT_SITE_URL = 'https://centurioncorporation.sharepoint.com/sites/FinanceRPAProcess-UAT';
const DEFAULT_LIST_NAME = 'Vendor Information';
const FORM_URL = 'https://centurion.workflowcloud.com/forms/f53621e8-33a7-4a24-8f25-57dbbc3f11e7';
const TABLE_ID = 'vendorInformationRequestorDashboardTable';

export interface IVendorInformationFormRequestorDashboardWebPartProps {
  description: string;
  sharePointSiteUrl: string;
  listName: string;
}

interface IVendorDashboardRow {
  id: number;
  requestorName: string;
  formNo: string;
  createdDate: string;
  createdSortValue: number;
  company: string;
  department: string;
  companyRefNo: string;
  vendorEmail: string;
  status: string;
}

export default class VendorInformationFormRequestorDashboardWebPart extends BaseClientSideWebPart<IVendorInformationFormRequestorDashboardWebPartProps> {
  private _sharePointRestService!: SharePointRestService;
  private _dataTable: unknown;

  public async onInit(): Promise<void> {
    await super.onInit();
    this._sharePointRestService = new SharePointRestService(this.context);
  }

  public render(): void {
    this._ensureFullWidthLayout();
    this._renderShell();
    this._bindUiEvents();
    this._loadDashboardData().catch((error: Error) => {
      this._renderError(error.message || 'An unexpected error occurred while loading the dashboard.');
    });
  }

  protected onDispose(): void {
    this._destroyDataTable();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (oldValue !== newValue && (propertyPath === 'sharePointSiteUrl' || propertyPath === 'listName')) {
      this.render();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('sharePointSiteUrl', {
                  label: 'SharePoint Site URL'
                }),
                PropertyPaneTextField('listName', {
                  label: 'List Name'
                })
              ]
            }
          ]
        }
      ]
    };
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

  private _renderShell(): void {
    const css: Record<string, string> = styles as unknown as Record<string, string>;

    this.domElement.innerHTML = `
      <section class="${css.vendorInformationFormRequestorDashboard}">
        <div class="${css.dashboardShell}">
          <div class="${css.dashboardHeader}">
            <div>
              <h2 class="${css.dashboardTitle}">Vendor Information Form Requestor Dashboard</h2>
            </div>
            <div class="${css.dashboardHeaderActions}">
              <button type="button" class="${css.refreshButton}" data-action="refresh">Refresh</button>
              <div class="${css.dashboardMeta}">
                <span class="${css.totalLabel}">TOTAL RECORDS</span>
                <span class="${css.totalValue}" data-dashboard-total>0</span>
              </div>
            </div>
          </div>
          <div class="${css.tableSection}">
            <div class="${css.infoMessage}">Loading records...</div>
          </div>
        </div>
      </section>
    `;
  }

  private _bindUiEvents(): void {
    const refreshButton = this.domElement.querySelector('[data-action="refresh"]') as HTMLButtonElement | null;
    if (refreshButton) {
      refreshButton.onclick = () => {
        this._setLoadingState();
        this._loadDashboardData().catch((error: Error) => {
          this._renderError(error.message || 'An unexpected error occurred while loading the dashboard.');
        });
      };
    }
  }

  private async _loadDashboardData(): Promise<void> {
    const siteUrl: string = this._getConfiguredSiteUrl();
    const listName: string = this._getConfiguredListName();
    const selectFields: string[] = [
      'Id',
      'Requestor_x0020_Name',
      'Reference_x0020_Number',
      'Created',
      'Company_x0020_Name',
      'Departments_x002f__x0020_Entitie',
      'Company_x0020_Ref_x0020__x0020_N',
      'Vendor_x0020_Email',
      'Status'
    ];

    const items = await this._sharePointRestService.getAllListItems(
      siteUrl,
      listName,
      selectFields.join(','),
      '',
      'Created desc'
    ) as Array<Record<string, unknown>>;

    const rows: IVendorDashboardRow[] = items.map((item: Record<string, unknown>) => this._mapItemToRow(item));
    this._renderDashboard(rows);
  }

  private _renderDashboard(rows: IVendorDashboardRow[]): void {
    const css: Record<string, string> = styles as unknown as Record<string, string>;
    const totalElement: HTMLElement | null = this.domElement.querySelector('[data-dashboard-total]');

    if (totalElement) {
      totalElement.textContent = rows.length.toString();
    }

    if (!rows.length) {
      const section: Element | null = this.domElement.querySelector(`.${css.tableSection}`);
      if (section) {
        section.innerHTML = `<div class="${css.emptyMessage}">No records were found in the configured list.</div>`;
      }
      return;
    }

    const tableRows: string = rows.map((row: IVendorDashboardRow) => `
      <tr>
        <td>${row.id}</td>
        <td>${escape(row.requestorName)}</td>
        <td>
          <a class="${css.formLink}" href="${escape(this._buildFormUrl(row.id))}" target="_blank" rel="noreferrer">
            ${escape(row.formNo)}
          </a>
        </td>
        <td data-order="${row.createdSortValue}">${escape(row.createdDate)}</td>
        <td>${escape(row.company)}</td>
        <td>${escape(row.department)}</td>
        <td>${escape(row.companyRefNo)}</td>
        <td>${escape(row.vendorEmail)}</td>
        <td>
          <span class="${css.statusBadge} ${this._getStatusBadgeClassName(row.status)}">${escape(row.status)}</span>
        </td>
      </tr>
    `).join('');

    const section: Element | null = this.domElement.querySelector(`.${css.tableSection}`);
    if (!section) {
      return;
    }

    section.innerHTML = `
      <div class="${css.tableWrapper}">
        <table id="${TABLE_ID}" class="${css.dashboardTable}">
          <thead>
            <tr>
              <th>ID</th>
              <th>Requestor Name</th>
              <th>Form No.</th>
              <th>Created Date</th>
              <th>Company</th>
              <th>Department</th>
              <th>Company Ref No.</th>
              <th>Vendor Email</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${tableRows}
          </tbody>
        </table>
      </div>
    `;

    this._initializeDataTable();
  }

  private _initializeDataTable(): void {
    this._destroyDataTable();

    const tableSelector: string = `#${TABLE_ID}`;
    if (!this.domElement.querySelector(tableSelector)) {
      return;
    }

    this._dataTable = ($(tableSelector) as unknown as { DataTable: (options: unknown) => unknown }).DataTable({
      order: [[3, 'desc']],
      paging: true,
      searching: true,
      ordering: true,
      info: true,
      lengthChange: true,
      pageLength: 10,
      autoWidth: false,
      destroy: true,
      columnDefs: [
        { orderable: true, targets: '_all' }
      ]
    });
  }

  private _destroyDataTable(): void {
    if (this._dataTable) {
      ((this._dataTable as { destroy: () => void })).destroy();
      this._dataTable = undefined;
    }
  }

  private _setLoadingState(): void {
    this._destroyDataTable();

    const css: Record<string, string> = styles as unknown as Record<string, string>;
    const totalElement: HTMLElement | null = this.domElement.querySelector('[data-dashboard-total]');
    const section: Element | null = this.domElement.querySelector(`.${css.tableSection}`);

    if (totalElement) {
      totalElement.textContent = '0';
    }

    if (section) {
      section.innerHTML = `<div class="${css.infoMessage}">Loading records...</div>`;
    }
  }

  private _renderError(message: string): void {
    const css: Record<string, string> = styles as unknown as Record<string, string>;
    const totalElement: HTMLElement | null = this.domElement.querySelector('[data-dashboard-total]');
    const section: Element | null = this.domElement.querySelector(`.${css.tableSection}`);

    if (totalElement) {
      totalElement.textContent = '0';
    }

    if (section) {
      section.innerHTML = `<div class="${css.errorMessage}">${escape(message)}</div>`;
    }
  }

  private _mapItemToRow(item: Record<string, unknown>): IVendorDashboardRow {
    const createdDateValue: string = this._getTextValue(item.Created);
    const createdDate: Date = createdDateValue ? new Date(createdDateValue) : new Date(0);
    const fallbackFormNo: string = `VIF-${this._getTextValue(item.Id)}`;

    return {
      id: Number(item.Id || 0),
      requestorName: this._getTextValue(item.Requestor_x0020_Name, '-'),
      formNo: this._getTextValue(item.Reference_x0020_Number, fallbackFormNo),
      createdDate: this._formatDate(createdDateValue),
      createdSortValue: createdDate.getTime(),
      company: this._getTextValue(item.Company_x0020_Name, '-'),
      department: this._getTextValue(item.Departments_x002f__x0020_Entitie, '-'),
      companyRefNo: this._getTextValue(item.Company_x0020_Ref_x0020__x0020_N, '-'),
      vendorEmail: this._getTextValue(item.Vendor_x0020_Email, '-'),
      status: this._getStatusValue(item)
    };
  }

  private _getConfiguredSiteUrl(): string {
    const configuredValue: string = (this.properties.sharePointSiteUrl || '').trim();
    return configuredValue || DEFAULT_SITE_URL;
  }

  private _getConfiguredListName(): string {
    const configuredValue: string = (this.properties.listName || '').trim();
    return configuredValue || DEFAULT_LIST_NAME;
  }

  private _buildFormUrl(itemId: number): string {
    return `${FORM_URL}?id=${encodeURIComponent(itemId.toString())}`;
  }

  private _getStatusValue(item: Record<string, unknown>): string {
    const candidates: unknown[] = [
      item.Status,
      item.OData__Status,
      item.Task_x0020_Outcome
    ];

    for (const candidate of candidates) {
      const value = this._getChoiceTextValue(candidate);
      if (value) {
        return value;
      }
    }

    return 'Pending';
  }

  private _getChoiceTextValue(value: unknown): string {
    const directValue = this._getTextValue(value);
    if (directValue) {
      return directValue;
    }

    if (Array.isArray(value)) {
      const joinedValues = value
        .map((entry: unknown) => this._getTextValue(entry))
        .filter((entry: string) => !!entry)
        .join(', ');

      if (joinedValues) {
        return joinedValues;
      }
    }

    if (value && typeof value === 'object') {
      const choiceRecord = value as Record<string, unknown>;

      const objectCandidates: unknown[] = [
        choiceRecord.Value,
        choiceRecord.Label,
        choiceRecord.value,
        choiceRecord.label,
        choiceRecord.Title,
        choiceRecord.results
      ];

      for (const candidate of objectCandidates) {
        const resolvedValue = this._getChoiceTextValue(candidate);
        if (resolvedValue) {
          return resolvedValue;
        }
      }
    }

    return '';
  }

  private _getTextValue(value: unknown, fallback: string = ''): string {
    if (value === null || value === undefined) {
      return fallback;
    }

    const stringValue: string = String(value).trim();
    return stringValue || fallback;
  }

  private _formatDate(value: string): string {
    if (!value) {
      return '-';
    }

    const dateValue: Date = new Date(value);
    if (isNaN(dateValue.getTime())) {
      return '-';
    }

    return dateValue.toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    });
  }

  private _getStatusBadgeClassName(status: string): string {
    const css: Record<string, string> = styles as unknown as Record<string, string>;
    const normalizedStatus: string = status.toLowerCase();

    if (normalizedStatus.indexOf('completed') >= 0 || normalizedStatus.indexOf('approve') >= 0 || normalizedStatus.indexOf('success') >= 0) {
      return css.statusSuccess;
    }

    if (normalizedStatus.indexOf('pending rpa') >= 0 || normalizedStatus.indexOf('reject') >= 0 || normalizedStatus.indexOf('cancel') >= 0 || normalizedStatus.indexOf('fail') >= 0) {
      return css.statusDanger;
    }

    if (normalizedStatus.indexOf('draft') >= 0) {
      return css.statusDraft;
    }

    if (normalizedStatus.indexOf('pending vendor submission') >= 0 || normalizedStatus.indexOf('pending initiator action') >= 0) {
      return css.statusWarning;
    }

    if (
      normalizedStatus.indexOf('pending dormitory manager approval') >= 0 ||
      normalizedStatus.indexOf('pending finance approval') >= 0 ||
      normalizedStatus.indexOf('pending finance ap action') >= 0
    ) {
      return css.statusInfo;
    }

    if (normalizedStatus.indexOf('pending') >= 0 || normalizedStatus.indexOf('review') >= 0 || normalizedStatus.indexOf('progress') >= 0 || normalizedStatus.indexOf('submit') >= 0) {
      return css.statusWarning;
    }

    return css.statusNeutral;
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
}
