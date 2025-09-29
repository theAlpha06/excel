export const refSheet = {
  "VTA1710L": [3, 10, 24],
  "NTA855R": [3, 10, 23],
  "ANX-G": [3],
  "ANX-H": [3],
  "Anx-I": [3, 34],
  "ANX-J": [3],
  "ANX-K": [3],
  "ANX-L": [3],
  "ANX-M": [3],
};

export const fieldHeaderMapping = new Map();
fieldHeaderMapping.set("Name of Work", "Name");
fieldHeaderMapping.set("Are Joint Venture (JV) firms allowed to bid", "Clone__c");
fieldHeaderMapping.set("Bidding Start Date", "ContractStartDate__c");

export const fieldLineMapping = new Map();
fieldLineMapping.set("Qty", "ManualQuantity__c");
fieldLineMapping.set("Qty / Engine", "ManualQuantity__c");
fieldLineMapping.set("2nd Year Rate Including GST", "ManualCostPrice__c");
fieldLineMapping.set("Part No.", "LocalPartNo__c");
fieldLineMapping.set("Part Description", "ProductDescription__c");