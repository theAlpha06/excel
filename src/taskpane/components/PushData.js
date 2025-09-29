import React, { useEffect, useState } from "react";
import {
  PrimaryButton,
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton,
  DetailsList,
  SelectionMode,
  Text,
} from "@fluentui/react";
import { refSheet, fieldHeaderMapping, fieldLineMapping } from "../../utils/data";
import SalesforceService from "../services/SalesforceService";
import WorkBookAPI from "./WorkBookAPI";

const PushData = () => {
  const [sheetName, setSheetName] = useState("");
  const [excelData, setExcelData] = useState([]);
  const [loadingMsg, setLoadingMsg] = useState("Pushing data to Salesforce...");
  const [error, setError] = useState(null);
  const [showPreview, setShowPreview] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [pushResults, setPushResults] = useState([]);
  const [showSuccessMsg, setShowSuccessMsg] = useState(null);
  const partProductMap = new Map();
  const productCode = new Set();
  let rangeString = "";

  const fetchWorksheets = async () => {
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load("name");
        await context.sync();
        setSheetName(activeSheet.name);
      });
    } catch (error) {
      console.error("Error fetching worksheets:", error);
    }
  };

  useEffect(() => {
    fetchWorksheets();
  }, []);

  useEffect(() => {
    if (sheetName) {
      readExcelData();
    }
  }, [sheetName]);

  const readExcelData = async () => {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const usedRange = sheet.getUsedRange();
        usedRange.load("formulas");
        await context.sync();
        if (usedRange.formulas.length > 0) {
          const data = usedRange.formulas;
          setExcelData(data);
          setError(null);
        } else {
          setExcelData([]);
          setError("No data found in the active worksheet.");
        }
      });
    } catch (error) {
      console.error("Error reading Excel data:", error);
    }
  };

  const rowHeader = (val, rowIndices) => {
    let start = 0,
      end = rowIndices.length - 1;
    if (val < rowIndices[0]) return 0;
    let index;
    while (start <= end) {
      let mid = (start + end) >> 1;
      if (rowIndices[mid] == val) return mid;
      else if (rowIndices[mid] < val) {
        index = mid;
        start = mid + 1;
      } else if (rowIndices[mid] > val) {
        end = mid - 1;
      }
    }
    return index;
  };

  function expandRangeString(rangeStr, startCol = "A", endCol = "K") {
    return rangeStr.split(",").map((item) => {
      const [sheet, rows] = item.split("!");
      const [rowStart, rowEnd = rowStart] = rows.split(":");
      return { sheet, range: `${startCol}${rowStart - 1}:${endCol}${rowEnd}` };
    });
  }

  async function fetchRangeChunk(rangeStr) {
    const expanded = expandRangeString(rangeStr);
    return Excel.run(async (context) => {
      const ranges = expanded.map(({ sheet, range }) => {
        const ws = context.workbook.worksheets.getItem(sheet);
        const rng = ws.getRange(range);
        rng.load("values");
        return rng;
      });
      await context.sync();
      return ranges.flatMap((r) => r.values);
    });
  }

  async function fetchAllRangeChunks(chunks) {
    const promises = chunks.map((chunk) => fetchRangeChunk(chunk));
    return Promise.all(promises);
  }

  const rowIndex = (workSheetName, num) => {
    const rowIndices = refSheet[workSheetName];
    const index = rowHeader(num, rowIndices);
    return rowIndices[index];
  };

  const excelDateToJSDate = (serial) => {
    if (typeof serial !== "number") return serial;
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    return dateInfo.toISOString().split("T")[0];
  };

  const prepareRecords = async () => {
    const headers = {};
    const lines = [];
    let currentSection = null;
    let currentQuoteLineDescription = "";
    let currentQuoteLine = null;
    let linesHeaders = [];
    for (let i = 0; i < excelData.length; i++) {
      const row = excelData[i];
      const firstCell = (row[0] || "").toString().toLowerCase().trim();

      if (firstCell.startsWith("1.")) {
        currentSection = "header";
        continue;
      } else if (firstCell.startsWith("2.")) {
        currentSection = "schedule";
        continue;
      } else if (firstCell.startsWith("3.")) {
        currentSection = "lines";
        continue;
      }

      if (currentSection === "header") {
        let j = 0;
        while (j < row.length) {
          const cell = row[j];
          if (cell && cell.toString().trim() !== "") {
            const key = cell.toString().trim();
            let value = null;
            let foundValueIndex = -1;

            for (let k = j + 1; k < row.length; k++) {
              if (row[k] != undefined && row[k].toString().trim() !== "") {
                value = row[k];
                foundValueIndex = k;
                break;
              }
            }

            if (value !== null) {
              const isDateField = key.toLowerCase().includes("date");
              if (fieldHeaderMapping.has(key)) {
                headers[fieldHeaderMapping.get(key)] = isDateField
                  ? excelDateToJSDate(value)
                  : value;
              }
              j = foundValueIndex + 1;
            } else {
              j++;
            }
          } else {
            j++;
          }
        }
      }

      const isSNoRow = row.some(
        (cell) => typeof cell === "string" && cell.trim().toLowerCase() === "s no."
      );

      if (isSNoRow) {
        const quoteLineRow = excelData[i - 1];
        currentQuoteLine = {
          attributes: { type: "QuoteLineItem", referenceId: `refQL${lines.length}` },
          ManualSalesPrice__c: 1,
          Description: quoteLineRow[2] || "",
          Quantity: 2,
          Product2Id: "01t2s0000012QjJAAU",
          Quote_Cost_Items__r: { records: [] },
        };
        lines.push(currentQuoteLine);
        linesHeaders = row;
        continue;
      }
      if (currentQuoteLine && currentQuoteLine.Description !== currentQuoteLineDescription) {
        currentQuoteLineDescription = currentQuoteLine.Description;
        rangeString += '"';
      }
      if (currentQuoteLine) {
        const nonEmpty = row.filter((cell) => cell !== null && cell !== "").length;
        if (nonEmpty >= 5) {
          for (let j = 0; j < row.length && j < linesHeaders.length; j++) {
            const headerLabel = linesHeaders[j];
            const cellValue = row[j];
            if (headerLabel == "Rate" && cellValue[0] == "=") {
              const value = cellValue.slice(1).replaceAll("'", "").split("+");
              let str = "";
              for (const val of value) {
                const refSheet = val.split("!")[0];
                setLoadingMsg(`Fetching data from ${refSheet}...`);
                const vale = val.split("!")[1].slice(1);
                str += `${refSheet}!${rowIndex(refSheet, vale) + 1}:${vale - 1},`;
              }
              rangeString += str;
            }
          }
        }
      }
    }

    return {
      records: [
        {
          attributes: { type: "Quote", referenceId: "refQuote1" },
          ...headers,
          OpportunityId: "0067100000GL5I1AAL",
          QuoteLineItems: {
            records: lines,
          },
        },
      ],
    };
  };

  const getProductIds = async (productCode) => {
    const queryStr = `SELECT Id, ProductCode FROM Product2 WHERE ProductCode IN ('${[...productCode].join("','")}')`;
    const result = await SalesforceService.query(queryStr);
    if (result && result.records) {
      result.records.forEach((record) => {
        partProductMap.set(record.ProductCode, record.Id);
      });
    }
  };

  const pushDataToSalesforce = async () => {
    try {
      let startTime, dataEndTime, endTime;
      startTime = performance.now();
      setIsLoading(true);
      setError(null);
      setPushResults([]);
      const data = await prepareRecords();
      if (!data) {
        setError("No valid records to upload. Please check your field mappings.");
        setIsLoading(false);
        return;
      }
      setLoadingMsg("Creating records in Salesforce...");
      const chunks = rangeString.replaceAll(',"', '","').slice(1).slice(0, -1).split('","');
      let worksheetData = [];
      await fetchAllRangeChunks(chunks).then((sheetData) => {
        for (const data of sheetData) {
          let chunkRecords = [];
          let currentHeaders = [];

          for (const row of data) {
            const isHeaderRow = row.some(
              (cell) => cell && fieldLineMapping.has(cell.toString().trim())
            );
            if (isHeaderRow) {
              currentHeaders = row.map((cell) => (cell ? cell.toString().trim() : ""));
              continue;
            }
            if (currentHeaders.length === 0) continue;
            const nonEmptyCount = row.filter(
              (cell) => cell !== null && cell !== undefined && cell !== ""
            ).length;

            if (nonEmptyCount < 3) continue;
            const firstCell = row[0] ? row[0].toString().trim() : "";
            if (firstCell === "Sl No" || firstCell === "Part No.") continue;
            const mappedRecord = {
              attributes: {
                type: "QuoteCostItem__c",
                referenceId: `refQCI${worksheetData.length}_${chunkRecords.length}`,
              },
            };

            let hasValidData = false;

            for (let i = 0; i < currentHeaders.length && i < row.length; i++) {
              const header = currentHeaders[i];
              const cellValue = row[i];
              if (
                fieldLineMapping.has(header) &&
                cellValue !== null &&
                cellValue !== undefined &&
                cellValue !== ""
              ) {
                const salesforceField = fieldLineMapping.get(header);
                if (salesforceField === "LocalPartNo__c") {
                  productCode.add(cellValue);
                }
                mappedRecord[salesforceField] = cellValue;
                hasValidData = true;
              }
            }
            if (hasValidData) {
              chunkRecords.push(mappedRecord);
            }
          }
          worksheetData.push(chunkRecords);
        }
      });
      let worksheetIndex = 0;
      for (const record of data.records) {
        for (const line of record.QuoteLineItems.records) {
          if (worksheetIndex < worksheetData.length) {
            const currentWorksheetChunk = worksheetData[worksheetIndex];
            line.Quote_Cost_Items__r.records = [];
            for (const worksheetRecord of currentWorksheetChunk) {
              line.Quote_Cost_Items__r.records.push(worksheetRecord);
            }
            worksheetIndex++; 
          }
        }
      }
      const productIds = await getProductIds(productCode);
      for (const record of data.records) {
        for (const line of record.QuoteLineItems.records) {
          for (const costItem of line.Quote_Cost_Items__r.records) {
            if (costItem.LocalPartNo__c && partProductMap.has(`${costItem.LocalPartNo__c}`)) {
              costItem.Product2Id__c = partProductMap.get(`${costItem.LocalPartNo__c}`);
            }
          }
        }
      }
      dataEndTime = performance.now();
      const result = await SalesforceService.createCompositeTreeWithLines(data);
      endTime = performance.now();
      console.log(`Data Preparation Time: ${dataEndTime - startTime} ms`);
      console.log(`Salesforce Operation Time: ${endTime - dataEndTime} ms`);
      console.log(`Total Time: ${endTime - startTime} ms`);
      setShowSuccessMsg(`Quote(${result.results[0].id}) and related records created successfully!`);
      setTimeout(() => {
        setShowSuccessMsg(null);
      }, 10000);
      // setPushResults(result.results || []);
      // setShowPreview(true);
    } catch (error) {
      setError(error.message || "Something went wrong while uploading data.");
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: "20px" }}>
      <PrimaryButton
        text={isLoading ? `${loadingMsg}` : "Push Records"}
        onClick={pushDataToSalesforce}
        styles={{ root: { minWidth: 100 } }}
        disabled={isLoading}
      />
      {error && <div style={{ color: "red", marginTop: 10 }}>{error}</div>}
      {showSuccessMsg && <div style={{ color: "green", marginTop: 10 }}>{showSuccessMsg}</div>}
      {/* <Dialog
        hidden={!showPreview}
        onDismiss={() => setShowPreview(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Push Results Preview",
          subText: "All records of the push results to Salesforce",
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 600 } },
        }}
      >
        {pushResults.map((result, idx) => (
          <div key={idx} style={{ marginBottom: 20 }}>
            <Text variant="mediumPlus">Result {idx + 1}</Text>
            <DetailsList
              items={Object.entries(result).map(([field, value]) => ({
                field,
                value: typeof value === "object" ? JSON.stringify(value) : String(value),
              }))}
              columns={[
                { key: "field", name: "Field", fieldName: "field", minWidth: 100, maxWidth: 200 },
                { key: "value", name: "Value", fieldName: "value", minWidth: 100, maxWidth: 300 },
              ]}
              selectionMode={SelectionMode.none}
            />
          </div>
        ))}
        <DialogFooter>
          <DefaultButton onClick={() => setShowPreview(false)} text="Close" />
        </DialogFooter>
      </Dialog> */}
      <WorkBookAPI />
    </div>
  );
};

export default PushData;