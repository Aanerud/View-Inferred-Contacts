/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = getEmailHeader;
  }
});

function getEmailHeader() {
  const mailbox = Office.context.mailbox;
  const itemId = mailbox.item.itemId;

  const ewsRequest =
    `<?xml version="1.0" encoding="utf-8"?> 
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                   xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
                   xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" 
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> 
        <soap:Body> 
            <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"> 
                <ItemShape> 
                    <t:BaseShape>IdOnly</t:BaseShape> 
                    <t:IncludeMimeContent>true</t:IncludeMimeContent> 
                    <t:AdditionalProperties>
                        <t:ExtendedFieldURI PropertyTag="0x007D" PropertyType="Binary"/>
                    </t:AdditionalProperties>
                </ItemShape> 
                <ItemIds> 
                    <t:ItemId Id="${itemId}"/> 
                </ItemIds> 
            </GetItem> 
        </soap:Body> 
    </soap:Envelope>`;

  mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const response = $.parseXML(asyncResult.value);
      const extendedProps = response.getElementsByTagName("t:ExtendedProperty");
      let headerContent = "";

      for (let i = 0; i < extendedProps.length; i++) {
        const extendedProp = extendedProps[i];
        const propertyTag = extendedProp.getElementsByTagName("t:ExtendedFieldURI")[0].getAttribute("PropertyTag");
        if (propertyTag === "0x007D") { // Example for PR_TRANSPORT_MESSAGE_HEADERS property
          headerContent = extendedProp.getElementsByTagName("t:Value")[0].textContent;
          break;
        }
      }

      document.getElementById("emailHeader").textContent = headerContent || "No email header found.";
    } else {
      console.error("Failed to fetch the email header.", asyncResult.error);
      document.getElementById("emailHeader").textContent = "Error retrieving the email header.";
    }
  });
}
