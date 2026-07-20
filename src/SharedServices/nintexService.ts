import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointRestService } from './sharePointRestService';
import { INintexConfig, NintexConfig } from './nintexConfig';

export interface INintexTask {
  id: string;
  name: string;
  workflowName: string;
  status: string;
  assignedTo: string;
  assignedWhen?: string;
  completedWhen?: string;
  taskUrl?: string;
  raw: Record<string, unknown>;
}

type INintexTaskResponse = {
  tasks?: Record<string, unknown>[];
};

export class NintexService {
  private readonly sharePointService: SharePointRestService;
  private readonly config: NintexConfig;

  constructor(context: WebPartContext, config?: Partial<INintexConfig>) {
    this.sharePointService = new SharePointRestService(context);
    this.config = new NintexConfig(config);
  }

  public async getTasks(
    status: string = 'active',
    daysBack: number = 30
  ): Promise<INintexTask[]> {
    const token = await this.getAccessToken();
    const responses = await this.getTasksByDateRange(token, status, daysBack);

    return responses.map((task: Record<string, unknown>) => this.mapTask(task));
  }

  public async getTasksAcrossStatuses(
    statuses: string[] = ['all'],
    daysBack: number = 30
  ): Promise<INintexTask[]> {
    const normalizedStatuses = statuses
      .map((status: string) => status.trim().toLowerCase())
      .filter((status: string, index: number, source: string[]) => !!status && source.indexOf(status) === index);

    const allTasks = await Promise.all(
      normalizedStatuses.map((status: string) => this.getTasks(status, daysBack))
    );

    const deduped = new Map<string, INintexTask>();
    const mergedTasks = allTasks.reduce((acc: INintexTask[], current: INintexTask[]) => acc.concat(current), []);

    mergedTasks.forEach((task: INintexTask) => {
      const key = task.id || `${task.name}|${task.workflowName}|${task.assignedWhen || ''}|${task.status}`;
      if (!deduped.has(key)) {
        deduped.set(key, task);
      }
    });

    return Array.from(deduped.values());
  }

  private async getAccessToken(): Promise<string> {
    const directToken = this.config.getDirectAccessToken().trim();
    if (directToken) {
      return directToken;
    }

    const clientId = this.config.getClientId().trim();
    const clientSecret = this.config.getClientSecret().trim();
    if (clientId && clientSecret) {
      return this.requestAccessToken(clientId, clientSecret);
    }

    const tokenSourceUrl = this.config.getTokenSourceUrl().trim();
    const tokenListName = this.config.getTokenListName().trim();
    const tokenItemTitle = this.config.getTokenItemTitle().trim();

    if (!tokenSourceUrl || !tokenListName) {
      throw new Error('Nintex credentials are not configured. Provide client credentials for DEV testing or configure a token source site/list.');
    }

    const items = await this.sharePointService.getListItems(
      tokenSourceUrl,
      tokenListName,
      'Title,AccessToken',
      `Title eq '${tokenItemTitle}'`
    );

    if (!items || !items.length || !items[0].AccessToken) {
      throw new Error(`Nintex access token was not found in ${tokenListName}.`);
    }

    return items[0].AccessToken;
  }

  private async requestAccessToken(clientId: string, clientSecret: string): Promise<string> {
    const requestHeaders = new Headers();
    requestHeaders.append('Content-Type', 'application/x-www-form-urlencoded');

    const requestBody = new URLSearchParams();
    requestBody.append('client_id', clientId);
    requestBody.append('client_secret', clientSecret);
    requestBody.append('grant_type', 'client_credentials');

    const response = await fetch(this.config.getTokenEndpoint(), {
      method: 'POST',
      headers: requestHeaders,
      body: requestBody
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to retrieve Nintex access token: HTTP ${response.status} ${errorText}`);
    }

    const tokenData = await response.json() as { access_token?: string };
    if (!tokenData.access_token) {
      throw new Error('Nintex authentication response did not include an access_token.');
    }

    return tokenData.access_token;
  }

  private async getTasksByDateRange(
    token: string,
    status: string,
    daysBack: number
  ): Promise<Record<string, unknown>[]> {
    const requestHeaders = new Headers();
    requestHeaders.append('Content-Type', 'application/x-www-form-urlencoded');
    requestHeaders.append('Authorization', `Bearer ${token}`);

    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(endDate.getDate() - Math.max(daysBack, 1));

    let cursorDate = new Date(startDate);
    const chunkDays = 7;
    const requestPromises: Array<Promise<Record<string, unknown>[]>> = [];

    while (cursorDate < endDate) {
      const fromDate = new Date(cursorDate);
      const toDate = this.addDays(fromDate, chunkDays);
      if (toDate > endDate) {
        toDate.setTime(endDate.getTime());
      }

      const url = `${this.config.getApiBaseUrl()}/workflows/v2/tasks?status=${encodeURIComponent(status)}&from=${encodeURIComponent(fromDate.toISOString())}&to=${encodeURIComponent(toDate.toISOString())}`;

      requestPromises.push(
        fetch(url, {
          method: 'GET',
          headers: requestHeaders
        })
          .then(async (response: Response) => {
            if (!response.ok) {
              throw new Error(`HTTP ${response.status} while loading Nintex tasks.`);
            }

            const data = await response.json() as INintexTaskResponse;
            return Array.isArray(data.tasks) ? data.tasks : [];
          })
          .catch((error: Error) => {
            console.error('Nintex task request failed:', error);
            return [];
          })
      );

      cursorDate = new Date(toDate);
      cursorDate.setSeconds(cursorDate.getSeconds() + 1);
    }

    const allResults = await Promise.all(requestPromises);
    return allResults.reduce((acc: Record<string, unknown>[], current: Record<string, unknown>[]) => acc.concat(current), []);
  }

  private mapTask(task: Record<string, unknown>): INintexTask {
    const linkCandidates = this.pickNestedString(task, [
      ['taskUrl'],
      ['url'],
      ['link'],
      ['taskLink'],
      ['response'],
      ['responseUrl'],
      ['completeUrl'],
      ['formUrl'],
      ['htmlUrl'],
      ['links', 'task'],
      ['_links', 'self', 'href'],
      ['_links', 'html', 'href'],
      ['_links', 'web', 'href'],
      ['_links', 'response', 'href']
    ]) || this.findFirstUrl(task);

    const assignedWhen = this.pickString(task, [
      'createdAt',
      'assignedAt',
      'dateCreated',
      'startDate',
      'dateInitiated',
      'initiatedAt',
      'taskDate'
    ]) || this.findFirstMatchingString(task, [
      'initiated',
      'assigned',
      'created',
      'start'
    ]);

    const completedWhen = this.pickString(task, [
      'completedAt',
      'dateCompleted',
      'lastModifiedAt',
      'endDate',
      'completedDate'
    ]) || this.findFirstMatchingString(task, [
      'completed',
      'finished',
      'closed'
    ]);

    return {
      id: this.pickString(task, ['id', 'taskId']) || this.pickString(task, ['instanceId']) || '',
      name: this.pickString(task, ['name', 'taskName', 'displayName', 'title']) || 'Untitled Task',
      workflowName: this.pickString(task, ['workflowName', 'processName', 'formName', 'workflowDisplayName']) || 'Workflow not specified',
      status: this.pickString(task, ['status', 'taskStatus']) || 'Unknown',
      assignedTo: this.pickString(task, ['assigneeDisplayName', 'assignee', 'assignedTo']) || '',
      assignedWhen,
      completedWhen,
      taskUrl: linkCandidates || undefined,
      raw: task
    };
  }

  private pickString(task: Record<string, unknown>, keys: string[]): string {
    for (const key of keys) {
      const value = task[key];
      if (typeof value === 'string' && value.trim()) {
        return value.trim();
      }
    }

    return '';
  }

  private pickNestedString(task: Record<string, unknown>, keyPaths: string[][]): string {
    for (const keyPath of keyPaths) {
      let current: unknown = task;

      for (const key of keyPath) {
        if (!current || typeof current !== 'object' || !(key in current)) {
          current = undefined;
          break;
        }

        current = (current as Record<string, unknown>)[key];
      }

      if (typeof current === 'string' && current.trim()) {
        return current.trim();
      }
    }

    return '';
  }

  private findFirstMatchingString(task: Record<string, unknown>, keyFragments: string[]): string {
    const stack: unknown[] = [task];
    let found = '';

    while (stack.length > 0) {
      const current = stack.pop();
      if (!current || typeof current !== 'object') {
        continue;
      }

      Object.keys(current as Record<string, unknown>).forEach((key: string) => {
        if (found) {
          return;
        }

        const value = (current as Record<string, unknown>)[key];

        if (typeof value === 'string') {
          const lowerKey = key.toLowerCase();
          const matches = keyFragments.some((fragment: string) => lowerKey.indexOf(fragment) >= 0);
          if (matches && value.trim()) {
            found = value.trim();
            return;
          }
        }

        if (value && typeof value === 'object') {
          stack.push(value);
        }
      });
      if (found) {
        return found;
      }
    }

    return '';
  }

  private findFirstUrl(task: Record<string, unknown>): string {
    const stack: unknown[] = [task];

    while (stack.length > 0) {
      const current = stack.pop();
      if (!current || typeof current !== 'object') {
        continue;
      }

      for (const key of Object.keys(current as Record<string, unknown>)) {
        const value = (current as Record<string, unknown>)[key];
        if (typeof value === 'string' && /^https?:\/\//i.test(value.trim())) {
          return value.trim();
        }

        if (value && typeof value === 'object') {
          stack.push(value);
        }
      }
    }

    return '';
  }

  private addDays(date: Date, days: number): Date {
    const result = new Date(date);
    result.setUTCDate(result.getUTCDate() + days);
    return result;
  }
}
