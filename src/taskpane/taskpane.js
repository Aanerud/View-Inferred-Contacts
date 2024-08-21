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
    getHasTaskProperty();
    getContactInformationProperty();
}

function getHasTaskProperty() {
    const mailbox = Office.context.mailbox;
    const itemId = mailbox.item.itemId;

    log(`Starting EWS request for HasTask property with item ID: ${itemId}`);

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

    log(`EWS Request for HasTask: ${ewsRequest}`);

    mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            log("EWS request for HasTask succeeded.");
            const response = asyncResult.value;
            log(`EWS Response for HasTask: ${response}`);

            const hasTaskValue = extractValue(response, 'EntityExtraction/HasTask');
            document.getElementById("extractedData").innerHTML = `<strong>EntityExtraction/HasTask:</strong> ${hasTaskValue}`;

            log(`Extracted HasTask: ${hasTaskValue}`);
        } else {
            log(`EWS request for HasTask failed. Error: ${asyncResult.error.message}`);
            document.getElementById("extractedData").textContent = "Error retrieving the HasTask property.";
        }
    });
}

function getContactInformationProperty() {
    const mailbox = Office.context.mailbox;
    const itemId = mailbox.item.itemId;

    log(`Starting EWS request for ContactInformation property with item ID: ${itemId}`);

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
                    <t:ExtendedFieldURI PropertySetId="00062008-0000-0000-C000-000000000046" PropertyName="EntityExtraction/ContactInformation" PropertyType="String" />
                </t:AdditionalProperties>
            </m:ItemShape> 
            <m:ItemIds> 
                <t:ItemId Id="${itemId}" /> 
            </m:ItemIds> 
            </m:GetItem> 
        </soap:Body> 
        </soap:Envelope>`;

    log(`EWS Request for ContactInformation: ${ewsRequest}`);

    mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            log("EWS request for ContactInformation succeeded.");
            const response = asyncResult.value;
            log(`EWS Response for ContactInformation: ${response}`);

            const contactInfoRaw = extractValue(response, 'EntityExtraction/ContactInformation');
            const vCard = createVCard(contactInfoRaw);
            document.getElementById("vCardOutput").innerHTML = `<strong>vCard:</strong><br/><pre>${vCard}</pre>`;

            log(`Extracted ContactInformation: ${contactInfoRaw}`);
        } else {
            log(`EWS request for ContactInformation failed. Error: ${asyncResult.error.message}`);
            document.getElementById("vCardOutput").textContent = "Error retrieving the ContactInformation property.";
        }
    });
}

function extractValue(response, propertyName) {
    const regex = new RegExp(`<t:ExtendedFieldURI[^>]+PropertyName="${propertyName}"[^>]*><t:Value>(.*?)</t:Value>`, 'i');
    const match = response.match(regex);
    return match ? match[1] : "Not found";
}

function createVCard(contactInfoRaw) {
    try {
        if (contactInfoRaw === "Not found") {
            return "No contact information found.";
        }

        let contactInfoData = JSON.parse(contactInfoRaw);

        // Check if the value itself is a JSON string and needs further parsing
        if (typeof contactInfoData[0].entities[0].value === "string") {
            contactInfoData[0].entities[0].value = JSON.parse(contactInfoData[0].entities[0].value);
        }

        let vCard = "";

        // Iterate over the entities to extract the ContactInformation
        contactInfoData[0].entities.forEach(entity => {
            if (entity.value && entity.value.ContactInformation && entity.value.ContactInformation.length > 0) {
                entity.value.ContactInformation.forEach(contact => {
                    const name = contact.Name && contact.Name.Value ? contact.Name.Value : "Unknown";
                    const firstName = contact.Name && contact.Name.FirstName ? contact.Name.FirstName : "";
                    const lastName = contact.Name && contact.Name.LastName ? contact.Name.LastName : "";
                    const role = contact.Role || "No role specified";
                    const email = contact.EmailAddress && contact.EmailAddress.length > 0 ? contact.EmailAddress.join(', ') : "No email";
                    const phone = contact.PhoneNumber && contact.PhoneNumber.length > 0 ? contact.PhoneNumber.map(phone => phone.Value).join(', ') : "No phone number";

                    vCard += `
BEGIN:VCARD
VERSION:3.0
FN:${name}
N:${lastName};${firstName};;;
TITLE:${role}
EMAIL:${email}
TEL:${phone}
END:VCARD
`;
                });
            }
        });

        return vCard || "No valid contact information found.";
    } catch (error) {
        log(`Error parsing contact information: ${error.message}`);
        return "Error parsing contact information.";
    }
}

