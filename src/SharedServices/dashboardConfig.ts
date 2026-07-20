// Legacy configuration file - now replaced by dynamic field discovery
// This file is kept for reference only and is no longer used

// The webpart now uses:
// 1. Dynamic site URL configuration through webpart properties  
// 2. Automatic field discovery from SharePoint lists
// 3. Intelligent field mapping based on common naming patterns

// No more static field mappings needed! 🎉

export interface ILegacyDashboardConfig {
  dataSiteUrl: string;
  lists: {
    CD_iRequest_Form: string;
    pendingRequests: string;
    completedRequests: string;
    userSubmissions: string;
    workflowTasks: string;
  };
}

// This configuration is now replaced by webpart properties
export const legacyDashboardConfig: ILegacyDashboardConfig = {
  dataSiteUrl: "https://srkksingapore.sharepoint.com/sites/CapitaLandDemo",
  lists: {
    CD_iRequest_Form: "CD_iRequest_Form",
    pendingRequests: "Pending Requests", 
    completedRequests: "Completed Requests",
    userSubmissions: "User Submissions",
    workflowTasks: "Workflow Tasks",
  }
};