import React, { useState, useEffect } from 'react';
import { 
  Stack, 
  Dropdown, 
  PrimaryButton, 
  Spinner, 
  SpinnerSize, 
  MessageBar, 
  MessageBarType,
  Text,
  TextField,
  ChoiceGroup,
  DetailsList,
  SelectionMode,
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton
} from '@fluentui/react';
import SalesforceService from '../services/SalesforceService';

const ExcelToSalesforce = () => {
  const [isConnected, setIsConnected] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [successMessage, setSuccessMessage] = useState(null);
  const [objects, setObjects] = useState([]);
  const [selectedObject, setSelectedObject] = useState(null);
  const [fields, setFields] = useState([]);
  const [excelData, setExcelData] = useState([]);
  const [excelHeaders, setExcelHeaders] = useState([]);
  const [fieldMappings, setFieldMappings] = useState({});
  const [operationType, setOperationType] = useState('insert');
  const [previewRecords, setPreviewRecords] = useState([]);
  const [showPreview, setShowPreview] = useState(false);
  const [confirmUpload, setConfirmUpload] = useState(false);
  const [uploadResults, setUploadResults] = useState([]);
  const [showResults, setShowResults] = useState(false);

  useEffect(() => {
    checkConnection();
  }, []);

  const checkConnection = async () => {
    try {
      await Office.context.document.settings.refreshAsync();
      const accessToken = Office.context.document.settings.get("salesforce_access_token");
      
      if (accessToken) {
        setIsConnected(true);
        loadObjects();
      } else {
        setIsConnected(false);
      }
    } catch (error) {
      console.error("Error checking connection:", error);
      setIsConnected(false);
    }
  };

  const loadObjects = async () => {
    try {
      setIsLoading(true);
      setError(null);
      
      const result = await SalesforceService.getObjects();
      
      if (result && result.sobjects) {
        const writableObjects = result.sobjects
          .filter(obj => 
            obj.createable && 
            obj.updateable && 
            !obj.deprecatedAndHidden
          )
          .map(obj => ({
            key: obj.name,
            text: obj.label
          }));
          
        setObjects(writableObjects);
      }
    } catch (error) {
      console.error("Error loading objects:", error);
      setError("Failed to load Salesforce objects. Please try reconnecting.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleObjectChange = async (_, option) => {
    setSelectedObject(option.key);
    setFieldMappings({});
    await loadFields(option.key);
    readExcelData();
  };

  const loadFields = async (objectName) => {
    try {
      setIsLoading(true);
      setError(null);
      
      const metadata = await SalesforceService.getObjectMetadata(objectName);
      
      if (metadata && metadata.fields) {
        const writableFields = metadata.fields
          .filter(field => 
            field.createable && 
            field.updateable && 
            !field.deprecatedAndHidden &&
            field.name !== 'Id'
          )
          .map(field => ({
            key: field.name,
            text: `${field.label} (${field.type})`,
            required: field.nillable === false,
            type: field.type
          }));
          
        setFields(writableFields);
      }
    } catch (error) {
      console.error("Error loading fields:", error);
      setError("Failed to load Salesforce fields. Please try again.");
    } finally {
      setIsLoading(false);
    }
  };

  const readExcelData = async () => {
    try {
      setIsLoading(true);
      setError(null);
      
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load("values");
        
        await context.sync();
        
        if (usedRange.values.length > 0) {
          const headers = usedRange.values[0];
          setExcelHeaders(headers);
          
          const data = usedRange.values.slice(1);
          setExcelData(data);
        } else {
          setExcelHeaders([]);
          setExcelData([]);
          setError("No data found in the active worksheet.");
        }
      });
    } catch (error) {
      console.error("Error reading Excel data:", error);
      setError("Failed to read data from Excel. Please make sure you have data in the active worksheet.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleFieldMappingChange = (excelHeader, salesforceField) => {
    setFieldMappings(prev => ({
      ...prev,
      [excelHeader]: salesforceField
    }));
  };

  const handleOperationTypeChange = (_, option) => {
    setOperationType(option.key);
  };

  const generatePreview = () => {
    try {
      const preview = [];
      
      for (let i = 0; i < Math.min(5, excelData.length); i++) {
        const row = excelData[i];
        const recordPreview = {};
        
        Object.entries(fieldMappings).forEach(([excelHeader, sfField]) => {
          const headerIndex = excelHeaders.indexOf(excelHeader);
          if (headerIndex >= 0 && sfField) {
            recordPreview[sfField] = row[headerIndex];
          }
        });
        
        preview.push(recordPreview);
      }
      
      setPreviewRecords(preview);
      setShowPreview(true);
    } catch (error) {
      console.error("Error generating preview:", error);
      setError("Failed to generate preview. Please check your field mappings.");
    }
  };

  const prepareRecords = () => {
    const records = [];
    
    for (let i = 0; i < excelData.length; i++) {
      const row = excelData[i];
      const record = {};
      
      Object.entries(fieldMappings).forEach(([excelHeader, sfField]) => {
        const headerIndex = excelHeaders.indexOf(excelHeader);
        if (headerIndex >= 0 && sfField) {
          record[sfField] = row[headerIndex];
        }
      });
      
      if (Object.keys(record).length > 0) {
        records.push(record);
      }
    }
    
    return records;
  };

  const uploadToSalesforce = async () => {
    try {
      setIsLoading(true);
      setError(null);
      setSuccessMessage(null);
      
      const records = prepareRecords();
      
      if (records.length === 0) {
        setError("No valid records to upload. Please check your field mappings.");
        return;
      }
      
      if (operationType === 'update' && !fieldMappings.Id) {
        setError("Update operation requires an Id field mapping.");
        return;
      }
      
      const batchSize = 200;
      const results = [];
      
      for (let i = 0; i < records.length; i += batchSize) {
        const batch = records.slice(i, i + batchSize);
        
        let batchResults;
        if (operationType === 'insert') {
          batchResults = await SalesforceService.createRecords(selectedObject, batch);
        } else if (operationType === 'update') {
          batchResults = await SalesforceService.updateRecords(selectedObject, batch);
        } else if (operationType === 'upsert') {
          const externalIdField = 'Id';
          batchResults = await SalesforceService.upsertRecords(selectedObject, externalIdField, batch);
        }
        
        results.push(...batchResults);
      }
      
      setUploadResults(results);
      setShowResults(true);
      
      const successCount = results.filter(r => r.success).length;
      const failureCount = results.filter(r => !r.success).length;
      
      setSuccessMessage(`Upload completed: ${successCount} records successful, ${failureCount} records failed.`);
    } catch (error) {
      console.error("Error uploading to Salesforce:", error);
      setError(`Failed to upload data to Salesforce: ${error.message}`);
    } finally {
      setIsLoading(false);
      setConfirmUpload(false);
    }
  };

  const previewColumns = [
    { key: 'field', name: 'Salesforce Field', fieldName: 'field', minWidth: 100, maxWidth: 200 },
    { key: 'value', name: 'Excel Value', fieldName: 'value', minWidth: 100, maxWidth: 200 }
  ];

  const resultsColumns = [
    { key: 'status', name: 'Status', fieldName: 'success', minWidth: 100, maxWidth: 100, 
      onRender: (item) => item.success ? 'Success' : 'Error' },
    { key: 'message', name: 'Message', fieldName: 'message', minWidth: 200, maxWidth: 300 },
    { key: 'id', name: 'Record ID', fieldName: 'id', minWidth: 200, maxWidth: 300 }
  ];

  if (!isConnected) {
    return (
      <div className="ms-welcome__data">
        <MessageBar messageBarType={MessageBarType.info}>
          Please connect to Salesforce first in the Connect tab.
        </MessageBar>
      </div>
    );
  }

  return (
    <div className="ms-welcome__data">
      <Stack tokens={{ childrenGap: 15 }}>
        {error && (
          <MessageBar 
            messageBarType={MessageBarType.error} 
            isMultiline={false} 
            onDismiss={() => setError(null)}
          >
            {error}
          </MessageBar>
        )}
        
        {successMessage && (
          <MessageBar 
            messageBarType={MessageBarType.success} 
            isMultiline={false} 
            onDismiss={() => setSuccessMessage(null)}
          >
            {successMessage}
          </MessageBar>
        )}
        
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <Dropdown
            placeholder="Select a Salesforce object"
            label="Salesforce Object"
            options={objects}
            onChange={handleObjectChange}
            disabled={isLoading}
            style={{ width: 200 }}
          />
          
          <PrimaryButton 
            text="Refresh" 
            onClick={readExcelData} 
            disabled={!selectedObject || isLoading}
            style={{ marginTop: 29 }}
          />
        </Stack>
        
        <ChoiceGroup
          label="Operation Type"
          options={[
            { key: 'insert', text: 'Insert (Create new records)' },
            // { key: 'update', text: 'Update (Update existing records - requires Id)' },
            { key: 'upsert', text: 'Upsert (Insert or Update based on Id)' }
          ]}
          selectedKey={operationType}
          onChange={handleOperationTypeChange}
          disabled={isLoading}
        />
        
        {selectedObject && excelHeaders.length > 0 && (
          <Stack tokens={{ childrenGap: 10 }}>
            <Text variant="large">Field Mappings</Text>
            <Text>Map your Excel columns to Salesforce fields:</Text>
            
            {excelHeaders.map((header, index) => (
              <Stack horizontal tokens={{ childrenGap: 10 }} key={index}>
                <TextField 
                  label="Excel Column" 
                  value={header} 
                  disabled 
                  style={{ width: 200 }}
                />
                <Dropdown
                  placeholder="Select Salesforce Field"
                  options={[
                    { key: '', text: '-- None --' },
                    ...(operationType !== 'insert' ? [{ key: 'Id', text: 'Id (Record ID)' }] : []),
                    ...fields
                  ]}
                  selectedKey={fieldMappings[header] || ''}
                  onChange={(_, option) => handleFieldMappingChange(header, option.key)}
                  style={{ width: 300 }}
                />
              </Stack>
            ))}
            
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton
                text="Preview Data"
                onClick={generatePreview}
                disabled={isLoading || Object.keys(fieldMappings).length === 0}
              />
              
              <PrimaryButton
                text="Upload to Salesforce"
                onClick={() => setConfirmUpload(true)}
                disabled={isLoading || Object.keys(fieldMappings).length === 0}
              />
            </Stack>
          </Stack>
        )}
        
        {isLoading && (
          <Spinner size={SpinnerSize.large} label="Processing..." />
        )}
        
        <Dialog
          hidden={!showPreview}
          onDismiss={() => setShowPreview(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Data Preview',
            subText: 'First 5 records that will be uploaded to Salesforce'
          }}
          modalProps={{
            isBlocking: false,
            styles: { main: { maxWidth: 600 } }
          }}
        >
          {previewRecords.map((record, recordIndex) => (
            <div key={recordIndex} style={{ marginBottom: 20 }}>
              <Text variant="mediumPlus">Record {recordIndex + 1}</Text>
              <DetailsList
                items={Object.entries(record).map(([field, value]) => ({ field, value: String(value) }))}
                columns={previewColumns}
                selectionMode={SelectionMode.none}
              />
            </div>
          ))}
          <DialogFooter>
            <DefaultButton onClick={() => setShowPreview(false)} text="Close" />
          </DialogFooter>
        </Dialog>
        
        <Dialog
          hidden={!confirmUpload}
          onDismiss={() => setConfirmUpload(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Confirm Upload',
            subText: `You are about to ${operationType} ${excelData.length} records to Salesforce. Do you want to continue?`
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={uploadToSalesforce} text="Upload" />
            <DefaultButton onClick={() => setConfirmUpload(false)} text="Cancel" />
          </DialogFooter>
        </Dialog>
        
        <Dialog
          hidden={!showResults}
          onDismiss={() => setShowResults(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Upload Results',
            subText: 'Results of the upload operation'
          }}
          modalProps={{
            isBlocking: false,
            styles: { main: { maxWidth: 700 } }
          }}
        >
          <DetailsList
            items={uploadResults.map(result => ({
              success: result.success,
              message: result.success ? 'Success' : result.errors.map(e => e.message).join(', '),
              id: result.id || 'N/A'
            }))}
            columns={resultsColumns}
            selectionMode={SelectionMode.none}
          />
          <DialogFooter>
            <DefaultButton onClick={() => setShowResults(false)} text="Close" />
          </DialogFooter>
        </Dialog>
      </Stack>
    </div>
  );
};

export default ExcelToSalesforce;