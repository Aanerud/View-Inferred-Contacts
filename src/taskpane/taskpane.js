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
    const debugLogElement = document.getElementById("debugLog");
    if (debugLogElement) {
        debugLogElement.innerText += `${message}\n`;
    }
}

function getExtractedData() {
    getIsM2HProperty(); // Call IsM2H first to determine if other properties should be shown
    getHasTaskProperty();
    getContactInformationProperty();
    getScalableExtractionClassifierProperty();
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
            displayHasTaskAsTable(hasTaskValue);

            log(`Extracted HasTask: ${hasTaskValue}`);
        } else {
            log(`EWS request for HasTask failed. Error: ${asyncResult.error.message}`);
            document.getElementById("extractedData").textContent = "Error retrieving the HasTask property.";
        }
    });
}

function displayHasTaskAsTable(hasTaskValue) {
    const hasTaskHtml = `
        <table border="0">
            <tbody>
                <tr><td>HasTask</td><td>${hasTaskValue}</td></tr>
            </tbody>
        </table>
        <p><strong>EntityExtraction/HasTask:</strong> Indicates if the email contains a task. Useful for tracking tasks across other applications.</p>`;
    document.getElementById("extractedData").innerHTML = hasTaskHtml;
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

        // Check if the ContactInfo array exists
        if (!contactInfoData.ContactInfo || !Array.isArray(contactInfoData.ContactInfo)) {
            throw new Error("Unexpected structure in ContactInformation data.");
        }

        let vCard = "";

        // Iterate over the ContactInfo array to extract and create vCards
        contactInfoData.ContactInfo.forEach(contact => {
            const name = contact.name && contact.name.displayName ? contact.name.displayName : "Unknown";
            const firstName = contact.name && contact.name.additionalInfo && contact.name.additionalInfo.first ? contact.name.additionalInfo.first : "";
            const lastName = contact.name && contact.name.additionalInfo && contact.name.additionalInfo.last ? contact.name.additionalInfo.last : "";
            const jobTitle = contact.position && contact.position.jobTitle ? contact.position.jobTitle : "No job title";
            const email = contact.emails && contact.emails.length > 0 ? contact.emails.map(e => e.address).join(', ') : "No email";
            const phone = contact.phones && contact.phones.length > 0 ? contact.phones.map(p => p.number).join(', ') : "No phone number";
            const address = contact.addresses && contact.addresses.length > 0 ? contact.addresses.map(a => a.address).join(', ') : "No address";

            vCard += `
BEGIN:VCARD
VERSION:3.0
FN:${name}
N:${lastName};${firstName};;;
TITLE:${jobTitle}
EMAIL:${email}
TEL:${phone}
ADR:${address}
END:VCARD
`;
        });

        return vCard || "No valid contact information found.";
    } catch (error) {
        log(`Error parsing contact information: ${error.message}`);
        return "Error parsing contact information.";
    }
}

function getIsM2HProperty() {
    const mailbox = Office.context.mailbox;
    const itemId = mailbox.item.itemId;

    log(`Starting EWS request for IsM2H property with item ID: ${itemId}`);

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
                    <t:ExtendedFieldURI PropertySetId="00062008-0000-0000-C000-000000000046" PropertyName="EntityExtraction/IsM2H" PropertyType="Boolean" />
                </t:AdditionalProperties>
            </m:ItemShape> 
            <m:ItemIds> 
                <t:ItemId Id="${itemId}" /> 
            </m:ItemIds> 
            </m:GetItem> 
        </soap:Body> 
        </soap:Envelope>`;

    log(`EWS Request for IsM2H: ${ewsRequest}`);

    mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            log("EWS request for IsM2H succeeded.");
            const response = asyncResult.value;
            log(`EWS Response for IsM2H: ${response}`);

            const isM2HValue = extractValue(response, 'EntityExtraction/IsM2H');
            displayIsM2HAsTable(isM2HValue || "no data");

            log(`Extracted IsM2H: ${isM2HValue}`);
        } else {
            log(`EWS request for IsM2H failed. Error: ${asyncResult.error.message}`);
            document.getElementById("extractedDataOutput").textContent += "Error retrieving the IsM2H property.";
        }
    });
}

function displayIsM2HAsTable(isM2HValue) {
    const isM2HHtml = `
        <table border="0">
            <tbody>
                <tr><td>IsM2H</td><td>${isM2HValue}</td></tr>
            </tbody>
        </table>
        <p><strong>EntityExtraction/IsM2H:</strong> Identifies if the email is machine-generated, helping to decide if extraction should proceed.</p>`;
    document.getElementById("extractedDataOutput").innerHTML = isM2HHtml;

    // Hide all other properties if IsM2H is "true"
    if (isM2HValue.trim().toLowerCase() === "true") {
        document.getElementById("extractedData").style.display = "none";
        document.getElementById("scalableExtractionClassifierOutput").style.display = "none";
        document.getElementById("vCardOutput").style.display = "none";
    } else {
        document.getElementById("extractedData").style.display = "block";
        document.getElementById("scalableExtractionClassifierOutput").style.display = "block";
        document.getElementById("vCardOutput").style.display = "block";
    }
}

function displayScalableExtractionClassifierData(scalableExtractionData) {
    try {
        // Check if data is available
        if (scalableExtractionData === "Not found" || !scalableExtractionData) {
            return "No data available.";
        }

        // Parse the JSON data
        const parsedData = JSON.parse(scalableExtractionData);

        // Ensure the structure matches what we expect
        if (!Array.isArray(parsedData) || !parsedData[0] || !parsedData[0].entities) {
            throw new Error("Unexpected structure in ScalableExtractionClassifier data.");
        }

        // Extract the value field containing the JSON-like string
        const valueData = JSON.parse(parsedData[0].entities[0].value);

        // Start building the HTML table
        let tableHtml = "<table border='0'><thead><tr><th>Type</th><th>Score</th><th>Threshold</th></tr></thead><tbody>";

        // Iterate over the valueData to populate the table rows
        valueData.forEach(item => {
            tableHtml += `<tr><td>${item.classification_entity_type}</td><td>${item.score}</td><td>${item.threshold}</td></tr>`;
        });

        tableHtml += "</tbody></table>";
        tableHtml += `<p><strong>EntityExtraction/ScalableExtractionClassifier:</strong> Pre-filter score for detecting contact information. If above 0.5, the contact extraction process runs.</p>`;

        return tableHtml;
    } catch (error) {
        log(`Error parsing ScalableExtractionClassifier data: ${error.message}`);
        return "Error parsing ScalableExtractionClassifier data.";
    }
}

function getScalableExtractionClassifierProperty() {
    const mailbox = Office.context.mailbox;
    const itemId = mailbox.item.itemId;

    log(`Starting EWS request for ScalableExtractionClassifier property with item ID: ${itemId}`);

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
                    <t:ExtendedFieldURI PropertySetId="00062008-0000-0000-C000-000000000046" PropertyName="EntityExtraction/ScalableExtractionClassifier" PropertyType="String" />
                </t:AdditionalProperties>
            </m:ItemShape> 
            <m:ItemIds> 
                <t:ItemId Id="${itemId}" /> 
            </m:ItemIds> 
            </m:GetItem> 
        </soap:Body> 
        </soap:Envelope>`;

    log(`EWS Request for ScalableExtractionClassifier: ${ewsRequest}`);

    mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            log("EWS request for ScalableExtractionClassifier succeeded.");
            const response = asyncResult.value;
            log(`EWS Response for ScalableExtractionClassifier: ${response}`);

            const scalableExtractionData = extractValue(response, 'EntityExtraction/ScalableExtractionClassifier');
            const tableHtml = displayScalableExtractionClassifierData(scalableExtractionData);
            document.getElementById("scalableExtractionClassifierOutput").innerHTML = tableHtml;

            log(`Extracted ScalableExtractionClassifier data: ${scalableExtractionData}`);
        } else {
            log(`EWS request for ScalableExtractionClassifier failed. Error: ${asyncResult.error.message}`);
            document.getElementById("scalableExtractionClassifierOutput").textContent = "Error retrieving the ScalableExtractionClassifier property.";
        }
    });
}

// Function to handle the data after it's been extracted
function handleScalableExtractionClassifier(response) {
    const scalableExtractionData = extractValue(response, 'EntityExtraction/ScalableExtractionClassifier');
    const tableHtml = displayScalableExtractionClassifierData(scalableExtractionData);
    document.getElementById("scalableExtractionClassifierOutput").innerHTML = tableHtml;
}
