/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "block";
      document.getElementById("run").onclick = getExtractedData;
  }
});

function log(message) {
  console.log(message);
  document.getElementById("debugLog").innerText += `${message}\n`;
}

function getExtractedData() {
  const mailbox = Office.context.mailbox;
  const itemId = mailbox.item.itemId;

  log(`Starting EWS request for item ID: ${itemId}`);

  const ewsRequest =
      `<?xml version="1.0" encoding="utf-8"?> 
      <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
                  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"> 
      <soap:Header>
          <t:RequestServerVersion Version="Exchange2013" />
      </soap:Header>
      <soap:Body> 
          <m:GetItem> 
          <m:ItemShape> 
              <t:BaseShape>IdOnly</t:BaseShape> 
              <t:AdditionalProperties>
              <t:ExtendedFieldURI PropertySetId="00062008-0000-0000-C000-000000000046" PropertyName="EntityExtraction/HasTask" PropertyType="Boolean" />
              </t:AdditionalProperties>
          </m:ItemShape> 
          <m:ItemIds> 
              <t:ItemId Id="${itemId}" /> 
          </m:ItemIds> 
          </m:GetItem> 
      </soap:Body> 
      </soap:Envelope>`;

  log(`EWS Request: ${ewsRequest}`);

  mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          log("EWS request succeeded.");
          const response = asyncResult.value;
          log(`EWS Response: ${response}`);

          // Directly extract the value between <t:Value> tags
          const valueMatch = response.match(/<t:Value>(.*?)<\/t:Value>/);
          const hasTaskValue = valueMatch ? valueMatch[1] : "Not found";

          document.getElementById("extractedData").innerHTML = `<strong>EntityExtraction/HasTask:</strong> ${hasTaskValue}`;

          log(`Extracted data: EntityExtraction/HasTask: ${hasTaskValue}`);
      } else {
          log(`EWS request failed. Error: ${asyncResult.error.message}`);
          document.getElementById("extractedData").textContent = "Error retrieving the extracted data.";
      }
  });
}
