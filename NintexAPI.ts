export async function getNintexToken(): Promise<string | undefined>{
    const requestHeaders = new Headers();
    requestHeaders.append("Content-Type", "application/x-www-form-urlencoded");

    const urlencoded = new URLSearchParams();
    urlencoded.append("client_id", "b80c527d-5fb9-485e-872d-4937f4a3bd0a");
    urlencoded.append("client_secret", "sNtSsLPtSsKOtSsNtWsQtTVSWVsI2HsPtWsJRK2RsILKPIKtPsRQKtWsNtVsQtRsFtSsPQOFMtV2p2K2vtSsFOLtWsROOxmD2XtW");
    urlencoded.append("grant_type", "client_credentials");

    const requestOptions = {
      method: "POST",
      headers: requestHeaders,
      body: urlencoded,
      redirect: "follow" as RequestRedirect
    };

    try{
      const response = await fetch("https://au.nintex.io/authentication/v1/token", requestOptions)
      const tokenData = await response.json();
      return tokenData.access_token;
    } catch(error) {
      console.error('Error fetching token:', error);
    }
}

export async function getNintexTask(token: string | undefined, status: string = 'active'): Promise<any>{
    const requestHeaders = new Headers();
    requestHeaders.append("Content-Type", "application/x-www-form-urlencoded");
    requestHeaders.append("Authorization", "Bearer " + token);

    var startDate = new Date("2025-06-01T00:00:00.000Z");
    var endDate = new Date("2025-06-15T00:00:00.000Z");
    const range = 3; // Days
    const today = new Date();
    const requestPromises = [];
    const requestOptions = {
      method: "GET",
      headers: requestHeaders,
      async: false
    };

    while (startDate < today)
    {
      const fromParam = startDate.toISOString();
      const toParam = endDate.toISOString();

      const url = `https://au.nintex.io/workflows/v2/tasks?status=${status}&from=${fromParam}&to=${toParam}`

      try{
        const response = fetch(url, requestOptions)
                          .then(async (response) => {
                            if( !response.ok){
                              throw new Error(`HTTP Error ${response.status} at ${url}`);
                            }

                            return (await response.json())
                          })
                          .catch((err: Error): [] => {
                              console.error(`Request failed: ${err.message}`);
                              // Return empty array on failure so Promise.all succeeds
                              return []; 
                          });;
        
        requestPromises.push(response);

        startDate = new Date(endDate);
        endDate = addDays(startDate, range);

      } catch(error) {
        console.error('Error fetching token:', error);
        return [];
      }
    }

    const allResults: any[] = await Promise.all(requestPromises);
    const allTasks: any[] = allResults.reduce((acc, result) => {
        // Check if tasks exists and is an array to avoid errors
        if (result && Array.isArray(result.tasks)) {
            return acc.concat(result.tasks);
        }
        return acc;
    }, []);

    return allTasks;

    
    /*
    try{
      const response = await fetch(`https://au.nintex.io/workflows/v2/tasks?status=${status}`, requestOptions)
      const data = await response.json();
      return data.tasks;
    } catch(error) {
      console.error('Error fetching token:', error);
    }
      */
}

// Helper to add days to a Date object without mutating the original
function addDays(date: any, days: any) {
    const result = new Date(date);
    result.setUTCDate(result.getUTCDate() + days); // Use UTC methods
    return result;
}

