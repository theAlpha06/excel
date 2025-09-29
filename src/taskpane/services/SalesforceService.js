class SalesforceService {
  static async getAccessToken() {
    await Office.context.document.settings.refreshAsync();
    const accessToken = Office.context.document.settings.get("salesforce_access_token");

    if (!accessToken) {
      throw new Error("Not authenticated with Salesforce");
    }

    return accessToken;
  }

  static async getInstanceUrl() {
    await Office.context.document.settings.refreshAsync();
    const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

    if (!instanceUrl) {
      throw new Error("Salesforce instance URL not found");
    }

    return instanceUrl;
  }

  static async getObjectsFromSalesforce(endpoint, method = "GET", data = null) {
    try {
      const accessToken = await this.getAccessToken();
      const instanceUrl = await this.getInstanceUrl();

      const url = `http://localhost:5000/salesforce${endpoint}`;

      const headers = {
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
        "Content-Type": "application/json",
      };

      let options = {
        method,
        headers,
      };

      if (data && (method === "POST" || method === "PATCH" || method === "PUT")) {
        options.body = JSON.stringify(data);
      }
      const response = await fetch(url, options);

      if (!response.ok) {
        if (response.status === 401) {
          throw new Error("Authentication failed. Please reconnect to Salesforce.");
        }

        const errorData = await response.json();
        throw new Error(errorData.message || `API error: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error("Salesforce API call failed:", error);
      throw error;
    }
  }

  static async callSalesforceApi(endpoint, method = "GET", data = null) {
    try {
      const accessToken = await this.getAccessToken();
      const instanceUrl = await this.getInstanceUrl();

      const url = `http://localhost:5000/salesforce${endpoint}`;

      const headers = {
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
        "Content-Type": "application/json",
      };

      let options = {
        method,
        headers,
      };

      if (data && (method === "POST" || method === "PATCH" || method === "PUT")) {
        options.body = JSON.stringify(data);
      }

      const response = await fetch(url, options);

      if (!response.ok) {
        if (response.status === 401) {
          throw new Error("Authentication failed. Please reconnect to Salesforce.");
        }

        const errorData = await response.json();
        throw new Error(errorData.message || `API error: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error("Salesforce API call failed:", error);
      throw error;
    }
  }

  static async getObjects() {
    return this.getObjectsFromSalesforce("/services/data/v58.0/sobjects");
  }

  static async getObjectMetadata(objectName) {
    return this.callSalesforceApi(`/services/data/v58.0/sobjects/${objectName}/describe`);
  }

  static async query(soqlQuery) {
    const accessToken = await this.getAccessToken();
    const instanceUrl = await this.getInstanceUrl();
    const response = await fetch("http://localhost:5000/salesforce/services/data/v58.0/query", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
      body: JSON.stringify({ q: soqlQuery }),
    });
    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.message || `SOQL query failed: ${response.status}`);
    }
    return response.json();
  }

  static async createRecords(objectName, records) {
    const accessToken = await this.getAccessToken();
    const instanceUrl = await this.getInstanceUrl();

    const response = await fetch("http://localhost:5000/salesforce/create", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
      body: JSON.stringify({ objectName, records }),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.message || `Failed to create records: ${response.status}`);
    }

    return response.json();
  }

  static async updateRecords(objectName, records) {
    try {
      if (!Array.isArray(records)) {
        if (!records.Id) {
          throw new Error("Id field is required for update operations");
        }

        const id = records.Id;
        const { Id, ...recordWithoutId } = records;

        await this.callSalesforceApi(
          `/services/data/v58.0/sobjects/${objectName}/${id}`,
          "PATCH",
          recordWithoutId
        );

        return [
          {
            id: id,
            success: true,
            errors: [],
          },
        ];
      }

      const batchSize = 200;
      const results = [];

      for (let i = 0; i < records.length; i += batchSize) {
        const batch = records.slice(i, i + batchSize);

        const missingIds = batch.filter((record) => !record.Id);
        if (missingIds.length > 0) {
          throw new Error(`${missingIds.length} records are missing Id field required for update`);
        }

        const compositeRequest = {
          allOrNone: false,
          records: batch.map((record) => {
            const { Id, ...recordWithoutId } = record;
            return {
              attributes: { type: objectName },
              id: Id,
              ...recordWithoutId,
            };
          }),
        };

        const response = await this.callSalesforceApi(
          "/services/data/v58.0/composite/sobjects",
          "PATCH",
          compositeRequest
        );

        results.push(...response);
      }

      return results;
    } catch (error) {
      console.error("Error updating records:", error);
      throw error;
    }
  }

  static async upsertRecords(objectName, externalIdField, records) {
    try {
      const accessToken = await this.getAccessToken();
      const instanceUrl = await this.getInstanceUrl();

      const response = await fetch("http://localhost:5000/salesforce/upsert", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "sf-access-token": accessToken,
          "sf-instance-url": instanceUrl,
        },
        body: JSON.stringify({
          objectName,
          externalIdField,
          records,
        }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.message || `Failed to upsert records: ${response.status}`);
      }

      return response.json();
    } catch (error) {
      console.error("Error upserting records:", error);
      throw error;
    }
  }

  static async deleteRecords(objectName, recordIds) {
    try {
      if (!Array.isArray(recordIds)) {
        await this.callSalesforceApi(
          `/services/data/v58.0/sobjects/${objectName}/${recordIds}`,
          "DELETE"
        );

        return [
          {
            id: recordIds,
            success: true,
            errors: [],
          },
        ];
      }

      const batchSize = 200;
      const results = [];

      for (let i = 0; i < recordIds.length; i += batchSize) {
        const batch = recordIds.slice(i, i + batchSize);

        const response = await this.callSalesforceApi(
          `/services/data/v58.0/composite/sobjects?ids=${batch.join(",")}&allOrNone=false`,
          "DELETE"
        );

        results.push(...response);
      }

      return results;
    } catch (error) {
      console.error("Error deleting records:", error);
      throw error;
    }
  }

  static async createCompositeTreeWithLines(data) {
    try {
      const accessToken = await this.getAccessToken();
      const instanceUrl = await this.getInstanceUrl();
  
      const response = await fetch(
        `http://localhost:5000/salesforce/composite-tree`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "sf-access-token": accessToken,
            "sf-instance-url": instanceUrl,
          },
          body: JSON.stringify(data),
        }
      );
      if (!response.ok) {
        const errorData = await response.json();
        throw new Error((errorData.message));
      }
  
      return await response.json();
    } catch (error) {
      throw error;
    }
  }

  static async uploadToSalesforce (fileName, base64Data, parentId) {
    const accessToken = await this.getAccessToken();
    const instanceUrl = await this.getInstanceUrl();

    const response = await fetch("http://localhost:5000/salesforce/uploadWorkbook", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "sf-access-token": accessToken,
        "sf-instance-url": instanceUrl,
      },
      body: JSON.stringify({
        fileName: fileName,
        base64Data,
        parentId: parentId,
      }),
    });
    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.message || `Upload failed: ${response.status}`);
    }
    return response.json();
  }
}

export default SalesforceService;