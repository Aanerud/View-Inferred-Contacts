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

function toggleDebugLog() {
    const debugLogElement = document.getElementById("debugLog");
    const toggleButton = document.getElementById("toggleDebugBtn");

    if (debugLogElement.style.display === "none") {
        debugLogElement.style.display = "block";
        toggleButton.textContent = "Hide EWS debug output";
    } else {
        debugLogElement.style.display = "none";
        toggleButton.textContent = "Show EWS debug output";
    }
}

function getExtractedData() {
    // Hide properties text when data is being loaded
    document.getElementById("propertiesInfo").style.display = "none";
    
    getIsM2HProperty(); // Check IsM2H first
    getHasTaskProperty();
    getContactInformationProperty();
    getContactInformationFTSLMProperty(); // New FTSLM extraction
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

function getContactInformationFTSLMProperty() {
    const mailbox = Office.context.mailbox;
    const itemId = mailbox.item.itemId;

    log(`Starting EWS request for ContactInformationFTSLM property with item ID: ${itemId}`);

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
                    <t:ExtendedFieldURI PropertySetId="00062008-0000-0000-C000-000000000046" PropertyName="EntityExtraction/ContactInformationFTSLM" PropertyType="String" />
                </t:AdditionalProperties>
            </m:ItemShape> 
            <m:ItemIds> 
                <t:ItemId Id="${itemId}" /> 
            </m:ItemIds> 
            </m:GetItem> 
        </soap:Body> 
        </soap:Envelope>`;

    log(`EWS Request for ContactInformationFTSLM: ${ewsRequest}`);

    mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            log("EWS request for ContactInformationFTSLM succeeded.");
            const response = asyncResult.value;
            log(`EWS Response for ContactInformationFTSLM: ${response}`);

            const contactInformationFTSLMValue = extractValue(response, 'EntityExtraction/ContactInformationFTSLM');
            displayContactInformationFTSLMAsTable(contactInformationFTSLMValue);

            log(`Extracted ContactInformationFTSLM: ${contactInformationFTSLMValue}`);
        } else {
            log(`EWS request for ContactInformationFTSLM failed. Error: ${asyncResult.error.message}`);
            document.getElementById("extractedData").textContent = "Error retrieving the ContactInformationFTSLM property.";
        }
    });
}

function displayHasTaskAsTable(hasTaskValue) {
    const hasTaskHtml = `
        <table>
            <tbody>
                <tr><td>HasTask</td><td>${hasTaskValue}</td></tr>
            </tbody>
        </table>
        <p><strong>EntityExtraction/HasTask:</strong> Indicates if the email contains a task. Useful for tracking tasks across other applications.</p>
    `;
    document.getElementById("extractedData").innerHTML += hasTaskHtml;
}

function displayContactInformationFTSLMAsTable(contactInformationFTSLMValue) {
    const contactInformationFTSLMHtml = `
        <table>
            <tbody>
                <tr><td>ContactInformationFTSLM</td><td>${contactInformationFTSLMValue}</td></tr>
            </tbody>
        </table>
        <p><strong>EntityExtraction/ContactInformationFTSLM:</strong> This contains fine-tuned contact information data for extraction testing purposes.</p>
    `;
    document.getElementById("extractedData").innerHTML += contactInformationFTSLMHtml;
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

        if (typeof contactInfoData === 'object' && contactInfoData.ContactInfo) {
            let vCard = "";

            contactInfoData.ContactInfo.forEach(contact => {
                const name = contact.name && contact.name.displayName ? contact.name.displayName : "Unknown";
                const firstName = contact.name && contact.name.additionalInfo.first ? contact.name.additionalInfo.first : "";
                const lastName = contact.name && contact.name.additionalInfo.last ? contact.name.additionalInfo.last : "";
                const role = contact.position && contact.position.jobTitle ? contact.position.jobTitle : "No role specified";
                const email = contact.emails && contact.emails.length > 0 ? contact.emails.map(email => email.address).join(', ') : "No email";
                const phone = contact.phones && contact.phones.length > 0 ? contact.phones.map(phone => phone.number).join(', ') : "No phone number";
                const address = contact.addresses && contact.addresses.length > 0 ? contact.addresses.map(addr => addr.address).join(', ') : "No address";

                vCard += `
BEGIN:VCARD
VERSION:3.0
FN:${name}
N:${lastName};${firstName};;;
TITLE:${role}
EMAIL:${email}
TEL:${phone}
ADR:${address}
END:VCARD
`;
            });

            return vCard || "No valid contact information found.";
        } else {
            throw new Error("Unexpected structure in ContactInformation data.");
        }
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
        <table>
            <tbody>
                <tr><td>IsM2H</td><td>${isM2HValue}</td></tr>
            </tbody>
        </table>
        <p><strong>EntityExtraction/IsM2H:</strong> Identifies if the email is machine-generated, helping to decide if extraction should proceed.</p>
    `;
    document.getElementById("extractedDataOutput").innerHTML = isM2HHtml;
}

function displayScalableExtractionClassifierData(scalableExtractionData) {
    try {
        if (scalableExtractionData === "Not found" || !scalableExtractionData) {
            return "No data available.";
        }

        const parsedData = JSON.parse(scalableExtractionData);

        if (!Array.isArray(parsedData) || !parsedData[0] || !parsedData[0].entities) {
            throw new Error("Unexpected structure in ScalableExtractionClassifier data.");
        }

        const valueData = JSON.parse(parsedData[0].entities[0].value);

        let tableHtml = "<table><thead><tr><th>Type</th><th>Score</th><th>Threshold</th></tr></thead><tbody>";

        valueData.forEach(item => {
            tableHtml += `<tr><td>${item.classification_entity_type}</td><td>${item.score}</td><td>${item.threshold}</td></tr>`;
        });

        tableHtml += "</tbody></table><p><strong>EntityExtraction/ScalableExtractionClassifier:</strong> Pre-filter score for detecting contact information. If above 0.5, the contact extraction process runs.</p>";

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
/* open bug dialog */

function openReportBugDialog() {
    document.getElementById("bugReportDialog").style.display = "flex";
}

function closeReportBugDialog() {
    document.getElementById("bugReportDialog").style.display = "none";
}

// Function to extract the value of a property from the debug log

function extractPropertyValueFromLog(debugLog, propertyName) {
    const regex = new RegExp(`<t:ExtendedFieldURI[^>]+PropertyName="${propertyName}"[^>]*><t:Value>(.*?)</t:Value>`, 'i');
    const match = debugLog.match(regex);
    return match ? match[1] : "Not found";
}

// Function to send the bug report
function sendBugReport() {
    const description = document.getElementById("bugDescription").value;
    const includeOriginalEmail = document.getElementById("includeOriginalEmail").checked;
    const debugLog = document.getElementById("debugLog").value;

    const isM2HValue = extractPropertyValueFromLog(debugLog, 'EntityExtraction/IsM2H');
    const scalableExtractionValue = extractPropertyValueFromLog(debugLog, 'EntityExtraction/ScalableExtractionClassifier');
    const contactInformationValue = extractPropertyValueFromLog(debugLog, 'EntityExtraction/ContactInformation');
    const contactInformationFTSLMValue = extractPropertyValueFromLog(debugLog, 'EntityExtraction/ContactInformationFTSLM'); // New property

    let originalEmailContent = "";
    const mailbox = Office.context.mailbox;

    if (includeOriginalEmail) {
        mailbox.item.body.getAsync("html", { asyncContext: { description, debugLog } }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                originalEmailContent = result.value;
                createDraftEmail(description, isM2HValue, scalableExtractionValue, contactInformationValue, contactInformationFTSLMValue, originalEmailContent);
            } else {
                console.error("Failed to retrieve email body:", result.error.message);
                createDraftEmail(description, isM2HValue, scalableExtractionValue, contactInformationValue, contactInformationFTSLMValue, "Could not retrieve original email content.");
            }
        });
    } else {
        createDraftEmail(description, isM2HValue, scalableExtractionValue, contactInformationValue, contactInformationFTSLMValue, "");
    }

    closeReportBugDialog();
}

// Function to create a draft email with the bug report information
function createDraftEmail(description, isM2HValue, scalableExtractionValue, contactInformationValue, contactInformationFTSLMValue, originalEmailContent) {
    const mailbox = Office.context.mailbox;
    const subject = "Bug from plugin";
    const toRecipients = ["sedsats@microsoft.com"];

    let message = `
        <p><strong>Bug Report</strong></p>
        <br>
        <p><strong>Description:</strong> ${description}</p>
        <br>
        <p><strong>IsM2H:</strong> ${isM2HValue}</p>
        <br>
        <p><strong>ScalableExtractionClassifier:</strong> ${scalableExtractionValue}</p>
        <br>
        <p><strong>ContactInformation:</strong> ${contactInformationValue}</p>
        <br>
        <p><strong>ContactInformationFTSLM:</strong> ${contactInformationFTSLMValue}</p>
        <br><br>
        ${originalEmailContent ? `<p><strong>Original Email Content:</strong></p><pre>=== ORIGINAL EMAIL BEGINS ===
${originalEmailContent}
=== ORIGINAL EMAIL ENDS ===</pre><br><br><br>` : ""}
    `;

    mailbox.displayNewMessageForm({
        toRecipients,
        subject,
        htmlBody: message
    });
}

