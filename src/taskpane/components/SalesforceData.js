import React, { useState, useEffect } from 'react';
import { 
  DetailsList, 
  SelectionMode, 
  Dropdown, 
  PrimaryButton, 
  Spinner, 
  SpinnerSize, 
  MessageBar,
  MessageBarType,
  Stack,
  Text,
  StackItem,
  useTheme,
  Checkbox
} from '@fluentui/react';
import SalesforceService from '../services/SalesforceService';

const SalesforceData = ({ isAuthenticated }) => {
  const [isConnected, setIsConnected] = useState(isAuthenticated || false);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [objects, setObjects] = useState([]);
  const [selectedObject, setSelectedObject] = useState(null);
  const [data, setData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [availableFields, setAvailableFields] = useState([]);
  const [selectedFields, setSelectedFields] = useState([]);
  const [showFieldSelection, setShowFieldSelection] = useState(false);
  const theme = useTheme();

  const [isSmallScreen, setIsSmallScreen] = useState(window.innerWidth <= 600);

  useEffect(() => {
    const handleResize = () => {
      setIsSmallScreen(window.innerWidth <= 600);
    };

    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  useEffect(() => {
    if (isAuthenticated) {
      setIsConnected(true);
      loadObjects();
    } else {
      checkConnection();
    }
  }, [isAuthenticated]);

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
        const commonObjects = result.sobjects
          .filter(obj => 
            ['Account', 'Contact', 'Opportunity', 'Lead', 'Case', 'Campaign', 'Product2', 'User', 'Item', 'Sale_Order__c', 'Sale_Order_Line__c'].includes(obj.name)
          )
          .map(obj => ({
            key: obj.name,
            text: obj.label
          }));
        setObjects(commonObjects);
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
    setData([]);
    setColumns([]);
    setSelectedFields([]);
    await loadObjectFields(option.key);
  };

  const loadObjectFields = async (objectName) => {
    try {
      setIsLoading(true);
      setError(null);
      
      const metadata = await SalesforceService.getObjectMetadata(objectName);
      
      const fields = metadata.fields
        .filter(field => 
          !field.deprecatedAndHidden && 
          field.filterable && 
          ['string', 'boolean', 'double', 'int', 'date', 'datetime', 'picklist', 'reference'].includes(field.type)
        )
        .map(field => ({
          key: field.name,
          text: field.label || field.name,
          fieldName: field.name,
          type: field.type
        }));
      
      setAvailableFields(fields);
      setShowFieldSelection(true);
    } catch (error) {
      console.error("Error loading object fields:", error);
      setError("Failed to load object fields. Please try again.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleFieldSelectionChange = (fieldKey, isChecked) => {
    if (isChecked) {
      setSelectedFields(prev => [...prev, fieldKey]);
    } else {
      setSelectedFields(prev => prev.filter(f => f !== fieldKey));
    }
  };

  const selectAllFields = () => {
    setSelectedFields(availableFields.map(f => f.key));
  };

  const clearAllFields = () => {
    setSelectedFields([]);
  };

  const loadRecords = async () => {
    if (!selectedObject) return;
    
    try {
      setIsLoading(true);
      setError(null);
      setData([]);
      setColumns([]);
      
      const fieldsToQuery = selectedFields.length > 0 
        ? selectedFields 
        : availableFields.map(f => f.key);
      
      const soqlQuery = `SELECT ${fieldsToQuery.join(',')} FROM ${selectedObject}`;
      
      const result = await SalesforceService.query(soqlQuery);
      
      if (result && result.records && result.records.length > 0) {
        const detailsColumns = fieldsToQuery
          .map(field => {
            const fieldInfo = availableFields.find(f => f.key === field);
            return {
              key: field,
              name: fieldInfo?.text || field,
              fieldName: field,
              minWidth: 100,
              maxWidth: 200,
              isResizable: true
            };
          });
        
        const formattedData = result.records.map(record => {
          const item = {};
          fieldsToQuery.forEach(field => {
            if (field !== 'attributes') {
              item[field] = record[field];
            }
          });
          return item;
        });
        
        setColumns(detailsColumns);
        setData(formattedData);
      } else {
        setData([]);
        setError("No records found for this object.");
      }
    } catch (error) {
      console.error("Error loading records:", error);
      setError("Failed to load Salesforce records. Please try again.");
    } finally {
      setIsLoading(false);
    }
  };

  const refreshData = () => {
    loadRecords();
  };

  const exportToExcel = async () => {
    try {
      setIsLoading(true);
      
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        sheet.getRange().clear();
        
        const headerRange = sheet.getRange("A1").getResizedRange(0, columns.length - 1);
        headerRange.values = [columns.map(col => col.name)];
        headerRange.format.font.bold = true;
        
        if (data.length > 0) {
          const dataRange = sheet.getRange("A2").getResizedRange(data.length - 1, columns.length - 1);
          const dataValues = data.map(item => 
            columns.map(col => item[col.fieldName] || '')
          );
          dataRange.values = dataValues;
        }
        
        sheet.getUsedRange().format.autofitColumns();
        
        await context.sync();
      });
      
      setIsLoading(false);
    } catch (error) {
      console.error("Error exporting to Excel:", error);
      setError("Failed to export data to Excel. Please try again.");
      setIsLoading(false);
    }
  };
  
  const containerStyles = {
    root: {
      width: '100%',
      padding: '10px'
    }
  };

  const stackTokens = { 
    childrenGap: 15 
  };

  const dropdownStyles = {
    root: {
      width: '100%',
      minWidth: 150,
      maxWidth: 300
    }
  };

  const buttonStyles = {
    root: {
      minWidth: 80
    }
  };

  const fieldSelectionStyles = {
    root: {
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: '4px',
      padding: '15px',
      maxHeight: '300px',
      overflowY: 'auto',
      backgroundColor: theme.palette.neutralLighterAlt
    }
  };

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
      <Stack tokens={stackTokens} styles={containerStyles}>
        {error && (
          <MessageBar 
            messageBarType={MessageBarType.error} 
            isMultiline={false} 
            onDismiss={() => setError(null)}
            styles={{ root: { width: '100%' } }}
          >
            {error}
          </MessageBar>
        )}
        
        <Stack 
          horizontal={!isSmallScreen} 
          wrap={true} 
          tokens={{ childrenGap: isSmallScreen ? 10 : 15 }}
        >
          <StackItem grow={1} styles={{ root: { minWidth: 200, maxWidth: 300 } }}>
            <Dropdown
              placeholder="Select a Salesforce object"
              label="Salesforce Object"
              options={objects}
              onChange={handleObjectChange}
              disabled={isLoading}
              styles={dropdownStyles}
            />
          </StackItem>
        </Stack>

        {showFieldSelection && availableFields.length > 0 && (
          <Stack tokens={{ childrenGap: 10 }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: '600' } }}>
              Select Fields to Retrieve (leave empty to fetch all fields):
            </Text>
            
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton 
                text="Select All" 
                onClick={selectAllFields} 
                disabled={isLoading}
                styles={{ root: { minWidth: 100 } }}
              />
              <PrimaryButton 
                text="Clear All" 
                onClick={clearAllFields} 
                disabled={isLoading}
                styles={{ root: { minWidth: 100 } }}
              />
            </Stack>
            
            <div style={fieldSelectionStyles.root}>
              <Stack tokens={{ childrenGap: 8 }}>
                {availableFields.map(field => (
                  <Checkbox
                    key={field.key}
                    label={`${field.text} (${field.key})`}
                    checked={selectedFields.includes(field.key)}
                    onChange={(_, isChecked) => handleFieldSelectionChange(field.key, isChecked)}
                    disabled={isLoading}
                  />
                ))}
              </Stack>
            </div>
            
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton 
                text="Load Records" 
                onClick={loadRecords} 
                disabled={isLoading}
                styles={buttonStyles}
              />
              
              <Text styles={{ root: { alignSelf: 'center', fontSize: '12px', color: theme.palette.neutralSecondary } }}>
                {selectedFields.length > 0 
                  ? `${selectedFields.length} fields selected` 
                  : 'All fields will be loaded'
                }
              </Text>
            </Stack>
          </Stack>
        )}
        
        {data.length > 0 && (
          <Stack 
            horizontal={!isSmallScreen} 
            wrap={true} 
            tokens={{ childrenGap: isSmallScreen ? 10 : 15 }}
          >
            <StackItem align={isSmallScreen ? "start" : "end"}>
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <PrimaryButton 
                  text="Refresh" 
                  onClick={refreshData} 
                  disabled={isLoading}
                  styles={buttonStyles}
                />
                
                <PrimaryButton 
                  text="Export" 
                  onClick={exportToExcel} 
                  disabled={!data.length || isLoading}
                  styles={buttonStyles}
                />
              </Stack>
            </StackItem>
          </Stack>
        )}
        
        {isLoading && (
          <Spinner 
            size={SpinnerSize.large} 
            label="Loading data..." 
            styles={{ root: { margin: '20px auto' } }}
          />
        )}
        
        {!isLoading && data.length > 0 && (
          <div style={{ overflowX: 'auto', width: '100%' }}>
            <DetailsList
              items={data}
              columns={columns}
              selectionMode={SelectionMode.none}
              layoutMode={1}
              isHeaderVisible={true}
              styles={{
                root: {
                  width: '100%',
                  overflowX: 'auto'
                }
              }}
            />
          </div>
        )}
        
        {!isLoading && !data.length && selectedObject && showFieldSelection && (
          <Text styles={{ root: { padding: '10px 0' } }}>
            Select fields and click "Load Records" to view data.
          </Text>
        )}
      </Stack>
    </div>
  );
};

export default SalesforceData;