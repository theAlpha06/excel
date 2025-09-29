import React from "react";
import SalesforceService from "../services/SalesforceService";
import { DefaultButton, Dialog, DialogFooter, PrimaryButton } from "@fluentui/react";

const WorkBookAPI = () => {
  const [results, setResults] = React.useState(null);
  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState(null);
  const pushToSalesforce = async () => {
    setIsLoading(true);
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4194304 },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const file = result.value;
          let sliceIndex = 0;
          let slices = [];

          const getNextSlice = () => {
            file.getSliceAsync(sliceIndex, (sliceResult) => {
              if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                slices.push(new Uint8Array(sliceResult.value.data));
                sliceIndex++;
                if (sliceIndex < file.sliceCount) {
                  getNextSlice();
                } else {
                  file.closeAsync();
                  const blob = new Blob(slices, {
                    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  });
                  const reader = new FileReader();
                  reader.onload = async function () {
                    const base64Data = reader.result.split(",")[1];
                    const result = await SalesforceService.uploadToSalesforce(
                      "Workbook.xlsx",
                      base64Data,
                      "0Q071000000sVObCAM"
                    );
                    if(!result.ok){
                      setError(result[0].message);
                      setTimeout(() => {
                        setError(null);
                      }, 3000);
                    } else {
                      setResults(result);
                    }
                    setIsLoading(false);
                  };
                  reader.readAsDataURL(blob);
                }
              } else {
                console.error("getSliceAsync failed:", sliceResult.error.message);
              }
            });
          };

          getNextSlice();
        } else {
          console.error("getFileAsync failed:", result.error.message);
        }
      }
    );
  };

  return (
    <div style={{ width: "100%" }}>
      {results && (
        <Dialog
          hidden={!results}
          onDismiss={() => setResults(null)}
          dialogContentProps={{
            title: "Push Workbook Results",
            subText: results.message,
          }}
        >
          {results && (
            <div>
              <h4>Results:</h4>
              <pre>{results?.id }</pre>
            </div>
          )}
          <DialogFooter>
            <DefaultButton onClick={() => setResults(null)} text="Close" />
          </DialogFooter>
        </Dialog>
      )}
      <PrimaryButton style={{ width: "100%" }} text="Push Workbook" onClick={pushToSalesforce} disabled={isLoading} />
      {error && <div style={{ color: "red", marginTop: 10 }}>{error}</div>}
    </div>
  );
};

export default WorkBookAPI;
