import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './TaskDashboardWebPart.module.scss';
import * as strings from 'TaskDashboardWebPartStrings';
import { NintexService, INintexTask } from '../../SharedServices/nintexService';

interface ITaskRow {
  task: string;
  workflow: string;
  status: string;
  assignedWhen: string;
  completedWhen: string;
  taskUrl?: string;
  activityDate?: Date;
}

interface IUserIdentity {
  email: string;
  loginName: string;
  displayName: string;
  normalizedValues: string[];
}

export interface ITaskDashboardWebPartProps {
  description: string;
  dataSiteUrl: string;
  nintexClientId: string;
  nintexClientSecret: string;
}

export default class TaskDashboardWebPart extends BaseClientSideWebPart<ITaskDashboardWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _nintexService!: NintexService;
  private _allRows: ITaskRow[] = [];
  private _currentPage: number = 1;
  private readonly _pageSize: number = 10;
  private _currentUserIdentity!: IUserIdentity;

  public async render(): Promise<void> {
    this._ensureFullWidthLayout();
    this._renderShell();
    await this._loadDashboard();
  }

  protected async onInit(): Promise<void> {
    this._nintexService = this._createNintexService();
    this._currentUserIdentity = this._getCurrentUserIdentity();
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

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--bodyBackground', semanticColors.bodyBackground || '#ffffff');
      this.domElement.style.setProperty('--disabledBodyText', semanticColors.disabledBodyText || '#605e5c');
      this.domElement.style.setProperty('--inputBorder', semanticColors.inputBorder || '#c8c6c4');
      this.domElement.style.setProperty('--inputBackground', semanticColors.inputBackground || '#ffffff');
      this.domElement.style.setProperty('--buttonBackground', semanticColors.buttonBackground || '#f3f2f1');
      this.domElement.style.setProperty('--buttonBackgroundHovered', semanticColors.buttonBackgroundHovered || '#edebe9');
      this.domElement.style.setProperty('--cardBackground', semanticColors.bodyBackground || '#ffffff');
      this.domElement.style.setProperty('--cardBorder', semanticColors.variantBorder || '#edebe9');
      this.domElement.style.setProperty('--headerBackground', semanticColors.bodyStandoutBackground || '#f8f7f6');
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('dataSiteUrl', {
                  label: strings.DataSiteUrlFieldLabel,
                  description: strings.DataSiteUrlFieldDescription
                }),
                PropertyPaneTextField('nintexClientId', {
                  label: strings.NintexClientIdFieldLabel,
                  description: strings.NintexClientIdFieldDescription
                }),
                PropertyPaneTextField('nintexClientSecret', {
                  label: strings.NintexClientSecretFieldLabel,
                  description: strings.NintexClientSecretFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected async onAfterPropertyPaneChangesApplied(): Promise<void> {
    this._nintexService = this._createNintexService();
    await this.render();
  }

  private _renderShell(): void {
    this.domElement.innerHTML = `
      <section class="${styles.taskDashboard} ${this._isDarkTheme ? styles.dark : ''}">
        <div class="${styles.header}">
          <div>
            <div class="${styles.eyebrow}">Tasks</div>
            <h2 class="${styles.title}">Task Dashboard</h2>
          </div>
          <button type="button" class="${styles.refreshButton}" data-action="refresh">Refresh</button>
        </div>

        <div class="${styles.filterBar}">
          <label class="${styles.filterGroup}">
            <span>Status</span>
            <select class="${styles.filterControl}" data-filter="status">
              <option value="active" selected>Active</option>
              <option value="in progress">In Progress</option>
            </select>
          </label>

          <label class="${styles.filterGroup}">
            <span>Filter Date</span>
            <select class="${styles.filterControl}" data-filter="date">
              <option value="30" selected>Past 30 days</option>
              <option value="7">Past 7 days</option>
              <option value="1">Today</option>
              <option value="all">All dates</option>
            </select>
          </label>

          <label class="${styles.filterGroup} ${styles.searchGroup}">
            <span>Search task name</span>
            <input type="search" class="${styles.filterControl}" data-filter="search" placeholder="Search task name" />
          </label>
        </div>

        <div class="${styles.metaRow}">
          <span class="${styles.resultInfo}" data-role="result-count">Loading tasks...</span>
        </div>

        <div class="${styles.tableWrapper}">
          <table class="${styles.taskTable}">
            <thead>
              <tr>
                <th>Task</th>
                <th>Workflow</th>
                <th>Status</th>
                <th>Assigned (when)</th>
                <th>Completed (when)</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody data-role="table-body">
              <tr>
                <td colspan="6" class="${styles.emptyState}">Loading tasks...</td>
              </tr>
            </tbody>
          </table>
        </div>

        <div class="${styles.paginationBar}" data-role="pagination"></div>
      </section>
    `;

    this._bindUiEvents();
  }

  private _bindUiEvents(): void {
    const refreshButton = this.domElement.querySelector('[data-action="refresh"]') as HTMLButtonElement | null;
    const statusFilter = this.domElement.querySelector('[data-filter="status"]') as HTMLSelectElement | null;
    const dateFilter = this.domElement.querySelector('[data-filter="date"]') as HTMLSelectElement | null;
    const searchFilter = this.domElement.querySelector('[data-filter="search"]') as HTMLInputElement | null;

    if (refreshButton) {
      refreshButton.onclick = async () => {
        await this._loadDashboard();
      };
    }

    if (statusFilter) {
      statusFilter.onchange = () => this._applyFilters();
    }

    if (dateFilter) {
      dateFilter.onchange = () => this._applyFilters();
    }

    if (searchFilter) {
      searchFilter.oninput = () => this._applyFilters();
    }
  }

  private async _loadDashboard(): Promise<void> {
    const tableBody = this.domElement.querySelector('[data-role="table-body"]') as HTMLTableSectionElement | null;
    const resultCount = this.domElement.querySelector('[data-role="result-count"]') as HTMLSpanElement | null;

    if (tableBody) {
      tableBody.innerHTML = `<tr><td colspan="6" class="${styles.emptyState}">Loading tasks...</td></tr>`;
    }

    if (resultCount) {
      resultCount.textContent = 'Loading tasks...';
    }

    try {
      const tasks = await this._nintexService.getTasksAcrossStatuses(['all'], 30);
      this._allRows = tasks
        .filter((task: INintexTask) => this._isTaskAssignedToCurrentUser(task))
        .map((task: INintexTask) => this._mapTaskRow(task))
        .filter((row: ITaskRow) => !!row.task || !!row.workflow || !!row.status)
        .sort((left: ITaskRow, right: ITaskRow) => {
          const leftValue = left.activityDate ? left.activityDate.getTime() : 0;
          const rightValue = right.activityDate ? right.activityDate.getTime() : 0;
          return rightValue - leftValue;
        });

      this._currentPage = 1;
      this._populateStatusOptions();
      this._applyFilters();
    } catch (error) {
      console.error('Failed to load task dashboard:', error);

      if (tableBody) {
        const message = error instanceof Error ? error.message : 'Unable to load Nintex task data.';
        tableBody.innerHTML = `<tr><td colspan="6" class="${styles.emptyState}">${this._escapeHtml(message)}</td></tr>`;
      }

      if (resultCount) {
        resultCount.textContent = '0 tasks';
      }
    }
  }

  private _populateStatusOptions(): void {
    const statusFilter = this.domElement.querySelector('[data-filter="status"]') as HTMLSelectElement | null;
    if (!statusFilter) {
      return;
    }

    statusFilter.innerHTML = `
      <option value="active" selected>Active</option>
      <option value="in progress">In Progress</option>
    `;
    statusFilter.value = 'active';
  }

  private _applyFilters(): void {
    const statusFilter = this.domElement.querySelector('[data-filter="status"]') as HTMLSelectElement | null;
    const dateFilter = this.domElement.querySelector('[data-filter="date"]') as HTMLSelectElement | null;
    const searchFilter = this.domElement.querySelector('[data-filter="search"]') as HTMLInputElement | null;
    const tableBody = this.domElement.querySelector('[data-role="table-body"]') as HTMLTableSectionElement | null;
    const resultCount = this.domElement.querySelector('[data-role="result-count"]') as HTMLSpanElement | null;
    const pagination = this.domElement.querySelector('[data-role="pagination"]') as HTMLDivElement | null;

    if (!tableBody) {
      return;
    }

    const selectedStatus = statusFilter ? statusFilter.value.trim().toLowerCase() : 'active';
    const selectedDateRange = dateFilter ? dateFilter.value : '30';
    const searchTerm = searchFilter ? searchFilter.value.trim().toLowerCase() : '';
    const cutoffDate = this._getCutoffDate(selectedDateRange);

    const filteredRows = this._allRows.filter((row: ITaskRow) => {
      const rowStatus = row.status.toLowerCase();
      const matchesStatus = rowStatus === selectedStatus;
      const matchesSearch = !searchTerm || row.task.toLowerCase().indexOf(searchTerm) >= 0;
      const matchesDate = !cutoffDate || (!!row.activityDate && row.activityDate >= cutoffDate);

      return matchesStatus && matchesSearch && matchesDate;
    });

    const totalPages = Math.max(1, Math.ceil(filteredRows.length / this._pageSize));
    if (this._currentPage > totalPages) {
      this._currentPage = totalPages;
    }
    if (this._currentPage < 1) {
      this._currentPage = 1;
    }

    const startIndex = (this._currentPage - 1) * this._pageSize;
    const pagedRows = filteredRows.slice(startIndex, startIndex + this._pageSize);

    if (!filteredRows.length) {
      tableBody.innerHTML = `<tr><td colspan="6" class="${styles.emptyState}">No tasks found for the current filters.</td></tr>`;
    } else {
      tableBody.innerHTML = pagedRows.map((row: ITaskRow) => `
        <tr>
          <td>${this._escapeHtml(row.task)}</td>
          <td>${this._escapeHtml(row.workflow)}</td>
          <td><span class="${styles.statusPill}">${this._escapeHtml(row.status)}</span></td>
          <td>${this._escapeHtml(row.assignedWhen)}</td>
          <td>${this._escapeHtml(row.completedWhen)}</td>
          <td>${this._renderActionCell(row)}</td>
        </tr>
      `).join('');
    }

    if (resultCount) {
      resultCount.textContent = `${filteredRows.length} task${filteredRows.length === 1 ? '' : 's'}`;
    }

    if (pagination) {
      pagination.innerHTML = this._renderPagination(filteredRows.length, totalPages);
      this._bindPaginationEvents(totalPages);
    }
  }

  private _mapTaskRow(task: INintexTask): ITaskRow {
    const assignedDate = this._parseDate(task.assignedWhen);
    const completedDate = this._parseDate(task.completedWhen);

    return {
      task: task.name || 'Untitled Task',
      workflow: task.workflowName || 'Workflow not specified',
      status: task.status || 'Pending',
      assignedWhen: this._formatDateTime(assignedDate),
      completedWhen: this._formatDateTime(completedDate),
      taskUrl: task.taskUrl,
      activityDate: completedDate || assignedDate
    };
  }

  private _renderActionCell(row: ITaskRow): string {
    if (!row.taskUrl) {
      return '<span class="' + styles.noLink + '">Unavailable</span>';
    }

    const safeUrl = this._escapeHtml(row.taskUrl);
    return `<a class="${styles.taskLink}" href="${safeUrl}" target="_blank" rel="noreferrer">Open Task</a>`;
  }

  private _formatDateTime(value?: Date): string {
    if (!value) {
      return '-';
    }

    const day = this._padNumber(value.getDate());
    const month = this._padNumber(value.getMonth() + 1);
    const year = value.getFullYear();
    const hours = this._padNumber(value.getHours());
    const minutes = this._padNumber(value.getMinutes());

    return `${day}/${month}/${year} ${hours}:${minutes}`;
  }

  private _getCurrentUserIdentity(): IUserIdentity {
    const email = (this.context.pageContext.user.email || '').trim();
    const loginName = (this.context.pageContext.user.loginName || '').trim();
    const displayName = (this.context.pageContext.user.displayName || '').trim();

    return {
      email,
      loginName,
      displayName,
      normalizedValues: this._buildIdentityVariants([email, loginName, displayName])
    };
  }

  private _isTaskAssignedToCurrentUser(task: INintexTask): boolean {
    const currentUser = this._currentUserIdentity;
    if (!currentUser || !currentUser.normalizedValues.length) {
      return true;
    }

    const assignmentCandidates = this._getTaskAssignmentCandidates(task);
    if (!assignmentCandidates.length) {
      return false;
    }

    return assignmentCandidates.some((candidate: string) => currentUser.normalizedValues.indexOf(candidate) >= 0);
  }

  private _getTaskAssignmentCandidates(task: INintexTask): string[] {
    const rawCandidates: string[] = [];
    rawCandidates.push(task.assignedTo || '');

    const stack: unknown[] = [task.raw];
    while (stack.length > 0) {
      const current = stack.pop();
      if (!current || typeof current !== 'object') {
        continue;
      }

      const record = current as Record<string, unknown>;
      Object.keys(record).forEach((key: string) => {
        const value = record[key];
        const normalizedKey = key.toLowerCase();

        if (typeof value === 'string') {
          if (this._isAssignmentKey(normalizedKey)) {
            rawCandidates.push(value);
          }
          return;
        }

        if (Array.isArray(value)) {
          value.forEach((item: unknown) => {
            if (typeof item === 'string' && this._isAssignmentKey(normalizedKey)) {
              rawCandidates.push(item);
            } else if (item && typeof item === 'object') {
              stack.push(item);
            }
          });
          return;
        }

        if (value && typeof value === 'object') {
          stack.push(value);
        }
      });
    }

    return this._buildIdentityVariants(rawCandidates);
  }

  private _isAssignmentKey(key: string): boolean {
    return key.indexOf('assign') >= 0 ||
      key.indexOf('assignee') >= 0 ||
      key.indexOf('owner') >= 0 ||
      key.indexOf('participant') >= 0 ||
      key.indexOf('user') >= 0 ||
      key.indexOf('email') >= 0 ||
      key.indexOf('principal') >= 0;
  }

  private _buildIdentityVariants(values: string[]): string[] {
    const uniqueValues = new Set<string>();

    values.forEach((value: string) => {
      const normalized = this._normalizeIdentityValue(value);
      if (!normalized) {
        return;
      }

      uniqueValues.add(normalized);

      const emailMatch = normalized.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
      if (emailMatch) {
        uniqueValues.add(emailMatch[0].toLowerCase());
      }

      const claimsMatch = normalized.match(/i:0#\.f\|membership\|(.+)/i);
      if (claimsMatch && claimsMatch[1]) {
        uniqueValues.add(claimsMatch[1].toLowerCase());
      }

      if (normalized.indexOf('\\') >= 0) {
        uniqueValues.add(normalized.split('\\').pop() || normalized);
      }
    });

    return Array.from(uniqueValues.values());
  }

  private _normalizeIdentityValue(value: string): string {
    return value.replace(/\s+/g, ' ').trim().toLowerCase();
  }

  private _getCutoffDate(range: string): Date | null {
    if (!range || range === 'all') {
      return null;
    }

    const days = parseInt(range, 10);
    if (isNaN(days)) {
      return null;
    }

    const cutoff = new Date();
    cutoff.setHours(0, 0, 0, 0);
    cutoff.setDate(cutoff.getDate() - (days - 1));
    return cutoff;
  }

  private _getDataSiteUrl(): string {
    return (this.properties.dataSiteUrl || this.context.pageContext.web.absoluteUrl || '').trim();
  }

  private _createNintexService(): NintexService {
    return new NintexService(this.context, {
      apiBaseUrl: 'https://au.nintex.io',
      clientId: (this.properties.nintexClientId || '').trim(),
      clientSecret: (this.properties.nintexClientSecret || '').trim(),
      tokenSourceUrl: this._getDataSiteUrl(),
      tokenListName: 'Nintex_AccessToken',
      tokenItemTitle: 'AccessToken'
    });
  }

  private _parseDate(value?: string): Date | undefined {
    if (!value) {
      return undefined;
    }

    const date = new Date(value);
    return isNaN(date.getTime()) ? undefined : date;
  }

  private _padNumber(value: number): string {
    return value < 10 ? `0${value}` : value.toString();
  }

  private _renderPagination(totalItems: number, totalPages: number): string {
    if (totalItems <= this._pageSize) {
      return '';
    }

    const maxVisiblePages = 5;
    const windowStart = Math.floor((this._currentPage - 1) / maxVisiblePages) * maxVisiblePages + 1;
    const windowEnd = Math.min(windowStart + maxVisiblePages - 1, totalPages);
    const pages: string[] = [];
    for (let page = windowStart; page <= windowEnd; page++) {
      pages.push(`
        <button
          type="button"
          class="${styles.pageButton} ${page === this._currentPage ? styles.pageButtonActive : ''}"
          data-page="${page}">
          ${page}
        </button>
      `);
    }

    return `
      <div class="${styles.paginationSummary}">
        Page ${this._currentPage} of ${totalPages}
      </div>
      <div class="${styles.paginationButtons}">
        <button
          type="button"
          class="${styles.pageButton}"
          data-page="${this._currentPage - 1}"
          ${this._currentPage === 1 ? 'disabled' : ''}>
          Previous
        </button>
        ${pages.join('')}
        <button
          type="button"
          class="${styles.pageButton}"
          data-page="${this._currentPage + 1}"
          ${this._currentPage === totalPages ? 'disabled' : ''}>
          Next
        </button>
      </div>
    `;
  }

  private _bindPaginationEvents(totalPages: number): void {
    const buttons = this.domElement.querySelectorAll('[data-role="pagination"] [data-page]');
    buttons.forEach((button: Element) => {
      (button as HTMLButtonElement).onclick = () => {
        const pageValue = parseInt((button as HTMLButtonElement).getAttribute('data-page') || '1', 10);
        if (isNaN(pageValue)) {
          return;
        }

        this._currentPage = Math.min(Math.max(pageValue, 1), totalPages);
        this._applyFilters();
      };
    });
  }

  private _escapeHtml(value: string): string {
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  
}
