declare interface ITaskDashboardWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DataSiteUrlFieldLabel: string;
  DataSiteUrlFieldDescription: string;
  NintexClientIdFieldLabel: string;
  NintexClientIdFieldDescription: string;
  NintexClientSecretFieldLabel: string;
  NintexClientSecretFieldDescription: string;
}

declare module 'TaskDashboardWebPartStrings' {
  const strings: ITaskDashboardWebPartStrings;
  export = strings;
}
