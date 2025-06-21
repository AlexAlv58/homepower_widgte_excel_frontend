export interface SearchResult {
  found: boolean;
  data?: any;
  error?: string;
}

export const searchRecord = async (
  entity: string,
  searchType: string,
  query: string
): Promise<SearchResult> => {
  //try {
    if (!query) {
      return { found: false, error: "Search query is required" };
    }

    const response = await ZOHO.CRM.API.searchRecord({
      Entity: entity,
      Type: searchType,
      Query: query,
    });

    if (response) {
        if (response.data && response.data.length > 0) {
            return { found: true, data: response.data };
        } else {
            return { found: false, data: [] };
        }
    } else {
      const errorMessage = response;
      console.error("Zoho API error:", errorMessage);
      return { found: false, error: errorMessage };
    }
  /* } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
    console.error(`Error searching for ${entity} by ${searchType}:`, errorMessage);
    return { found: false, error: errorMessage };
  } */
};

export const searchContactByEmail = async (email: string): Promise<SearchResult> => {
  return searchRecord("Contacts", "email", email);
};

export const searchAccountByName = async (account_name: string): Promise<SearchResult> => {
    const escapedAccountName = encodeURIComponent(account_name);
    return searchRecord("Accounts", "criteria", `((Account_Name:equals:${escapedAccountName}))`);
};


export interface RecordInsertResult {
  success: boolean;
  data?: any;
  error?: string;
}

export const insertRecord = async (entity: string, recordData: any): Promise<RecordInsertResult> => {
  //try {
    if (!recordData) {
      return { success: false, error: "Record data is required" };
    }

    const response = await ZOHO.CRM.API.insertRecord({
      Entity: entity,
      APIData: [recordData]
    })

    console.log("insertRecord", {
        Entity: entity,
        APIData: [recordData]
    }, response);

    if (response) {
      if (response.data && response.data.length > 0) {
        return { success: true, data: response.data };
      } else {
        return { success: false, error: "No data returned after insert" };
      }
    } else {
      const errorMessage = response?.message || "Unknown error occurred";
      console.error("Zoho API error:", errorMessage);
      return { success: false, error: errorMessage };
    }
  /* } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
    console.error(`Error inserting ${entity} record:`, errorMessage);
    return { success: false, error: errorMessage };
  } */
};

export const insertAccountRecord = async (accountData: any): Promise<RecordInsertResult> => {
  return insertRecord("Accounts", accountData);
};

export const insertSubmoduleRecord = async (subModuleData: any): Promise<RecordInsertResult> => {
  return insertRecord("Submodule_Commercial", subModuleData);
};


export const insertContactRecord = async (contactData: any): Promise<RecordInsertResult> => {
  return insertRecord("Contacts", contactData);
};

export const insertDealRecord = async (dealData: any): Promise<RecordInsertResult> => {
  return insertRecord("Deals", dealData);
};


export const updateRecord = async (entity: string, recordData: any): Promise<RecordInsertResult> => {
    //try {
      if (!recordData) {
        return { success: false, error: "Record data is required" };
      }

      const response = await ZOHO.CRM.API.updateRecord({
        Entity: entity,
        APIData: [recordData]
      })

      if (response) {
        if (response.data && response.data.length > 0) {
          return { success: true, data: response.data };
        } else {
          return { success: false, error: "No data returned after insert" };
        }
      } else {
        const errorMessage = response?.message || "Unknown error occurred";
        console.error("Zoho API error:", errorMessage);
        return { success: false, error: errorMessage };
      }
    /* } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
      console.error(`Error inserting ${entity} record:`, errorMessage);
      return { success: false, error: errorMessage };
    } */
};

export const updateContactRecord = async (contactData: any): Promise<RecordInsertResult> => {
    return updateRecord("Contacts", contactData);
};

