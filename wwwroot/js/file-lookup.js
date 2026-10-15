function processExcel(DocCategory) {
    const form = document.getElementById("excelUploadForm");
    if (!form) {
        console.error("excelUploadForm was not found.");
        showUploadError("The Excel upload form was not found.");
        return;
    }
    let fileInput;
    let docCategoryInput;
    if (DocCategory === 1) {    
        fileInput = document.getElementById("FULoanDoc");
        docCategoryInput = document.getElementById("DocCategory1");
    }
    else if (DocCategory === 2) {         
        fileInput = document.getElementById("FUTitleDoc");
        docCategoryInput = document.getElementById("DocCategory2");
    }
    else {
        console.error("Invalid DocCategory.");
        showUploadError("The document category is invalid.");
        return;
    }
    if (!fileInput || !docCategoryInput || fileInput.files.length === 0) {
        showUploadError("Please select an Excel file.");
        return;
    }
    const selectedDocCategory = Number(docCategoryInput.value);
    if (!Number.isInteger(selectedDocCategory) || selectedDocCategory <= 0) {
        showUploadError("Please enter a valid Doc Category.");
        return;
    }   
    const clientIdInput = document.getElementById("ClientId");
    const collRequestIdInput = document.getElementById("CollRequestId");
    const userIdInput = document.getElementById("UserID");
    if (!clientIdInput || !collRequestIdInput || !userIdInput) {
        console.error("One or more required form fields were not found.");
        showUploadError("One or more required form fields were not found.");
        return;
    }
    const formData = new FormData(form);    
    formData.delete("ExcelFile");
    formData.delete("DocCategory");
    formData.set("ClientId", clientIdInput.value);
    formData.set("CollRequestId", collRequestIdInput.value);
    formData.set("UserID", userIdInput.value);    
    formData.append("ExcelFile", fileInput.files[0]);
    formData.append("DocCategory", selectedDocCategory);    
    fetch("/Home/UploadAndProcessExcelFile", {
        method: "POST", body: formData
    })
        .then(async response => {
            const responseText = await response.text();
            const responseDocument = new DOMParser().parseFromString(responseText, "text/html");
            const responseMessage = responseDocument.querySelector(".upload-message");
            if (!response.ok) {
                throw new Error(responseMessage?.textContent.trim() || "The Excel upload failed.");
            }
            return responseText;
        })
        .then(result => {
            const resultDocument = new DOMParser().parseFromString(result, "text/html");           
            const resultMessage = resultDocument.querySelector(".upload-message");
            if (resultMessage) {
                const category = DocCategory === 1 ? "loan" : "title";
                const messageType = resultMessage.classList.contains("success") ? "success" : "error";
                if (messageType === "error") {
                    showUploadError(resultMessage.textContent);
                } else {
                    showUploadNotification(category, resultMessage.textContent, messageType);
                }
            }
        })
        .catch(error => {
            console.error("Excel processing error:", error);
            showUploadError(error instanceof Error ? error.message : String(error));
        });
}
const uploadNotificationTimers = {
    loan: { fade: null, hide: null },
    title: { fade: null, hide: null }
};
function showUploadNotification(category, message, messageType) {
    const notification = document.getElementById(`${category}UploadNotification`);
    const timers = uploadNotificationTimers[category];
    if (!notification || !timers) {
        return;
    }
    clearTimeout(timers.fade);
    clearTimeout(timers.hide);
    notification.className = `upload-message upload-notification ${messageType}`;
    notification.textContent = message;
    notification.style.display = "block";
    timers.fade = setTimeout(() => notification.classList.add("fade-out"), 4000);
    timers.hide = setTimeout(() => {
        notification.style.display = "none";
        notification.textContent = "";
        notification.classList.remove("fade-out");
    }, 4500);
}
function showUploadError(message) {
    const popup = document.getElementById("uploadErrorPopup");
    const messageContainer = document.getElementById("uploadErrorMessage");
    const fallbackMessage = "The Excel upload failed. Please try again.";
    if (!popup || !messageContainer) {
        console.error("Upload error popup elements are missing.");
        return;
    }
    messageContainer.textContent = typeof message === "string" && message.trim()? message.trim(): fallbackMessage;
    console.log("Showing upload error popup");
    popup.style.display = "flex";
    popup.setAttribute("aria-hidden", "false");
}
function hideUploadError() {
    const popup = document.getElementById("uploadErrorPopup");
    if (!popup) {
        return;
    }
    popup.style.display = "none";
    popup.setAttribute("aria-hidden", "true");
}
document.querySelectorAll(".static-upload-notification").forEach(notification => {
    notification.style.display = "block";
    setTimeout(() => notification.classList.add("fade-out"), 4000);
    setTimeout(() => {
        notification.style.display = "none";
        notification.textContent = "";
        notification.classList.remove("fade-out");
    }, 4500);
});
const serverUploadError = document.getElementById("serverUploadError");
if (serverUploadError?.textContent.trim()) {
    showUploadError(serverUploadError.textContent);
}
async function uploadLoanDocument(fileInput) {
    const status = document.getElementById("loanUploadNotification");
    const uploadButton = document.getElementById("uploadLoanDocumentButton");
    const form = document.getElementById("excelUploadForm");
    const clientIdInput = document.getElementById("ClientId");
    const collRequestIdInput = document.getElementById("CollRequestId");
    const userIdInput = document.getElementById("UserID");
    if (!status || !uploadButton || !form || !clientIdInput || !collRequestIdInput || !userIdInput) {
        console.error("The Loan upload controls were not found.");
        showUploadError("The Loan upload controls were not found.");
        return;
    }
    const file = fileInput?.files?.[0];
    const clientId = Number(clientIdInput.value);
    const collRequestId = Number(collRequestIdInput.value);
    const userId = Number(userIdInput.value);
    if (!file) {
        showUploadError("Please select a Loan Excel file.");
        return;
    }
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
        showUploadError("Only Excel files (.xlsx) are allowed.");
        fileInput.value = "";
        return;
    }
    if (!Number.isInteger(clientId) || clientId <= 0 ||
        !Number.isInteger(collRequestId) || collRequestId <= 0 ||
        !Number.isInteger(userId) || userId < 0) {
        showUploadError("Please enter valid Client ID, Coll Request ID, and User ID values.");
        fileInput.value = "";
        return;
    }
    const formData = new FormData();
    formData.append("excelFile", file);
    formData.append("clientId", String(clientId));
    formData.append("collRequestId", String(collRequestId));
    formData.append("userId", String(userId));
    uploadButton.disabled = true;
    try {
        console.log("Upload request started");
        const response = await fetch("/Documents/UploadLoanDocument", {
            method: "POST", body: formData
        });
        console.log("Upload HTTP status:", response.status);
        const responseText = await response.text();
        let result = null;
        try {
            result = responseText ? JSON.parse(responseText) : null;
        } catch {
            result = null;
        }
        if (!response.ok || !result || !result.success) {
            throw new Error(result?.message || responseText.trim() || "The loan document could not be processed.");
        }
        showUploadNotification("loan", result.message, "success");
        await loadLoanDocuments(clientId, collRequestId);
    } catch (error) {
        console.error("Loan document upload error:", error);
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.log("Upload error message:", errorMessage);
        showUploadError(errorMessage || "The loan document could not be processed.");
    } finally {
        fileInput.value = "";
        uploadButton.disabled = false;
    }
}
async function uploadTitleDocument(fileInput) {
    const status = document.getElementById("titleUploadNotification");
    const uploadButton = document.getElementById("uploadTitleDocumentButton");
    const form = document.getElementById("excelUploadForm");
    const clientIdInput = document.getElementById("ClientId");
    const collRequestIdInput = document.getElementById("CollRequestId");
    const userIdInput = document.getElementById("UserID");
    if (!status || !uploadButton || !form || !clientIdInput || !collRequestIdInput || !userIdInput) {
        console.error("The Title upload controls were not found.");
        showUploadError("The Title upload controls were not found.");
        return;
    }
    const file = fileInput?.files?.[0];
    const clientId = Number(clientIdInput.value);
    const collRequestId = Number(collRequestIdInput.value);
    const userId = Number(userIdInput.value);
    if (!file) {
        showUploadError("Please select a Title Excel file.");
        return;
    }
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
        showUploadError("Only Excel files (.xlsx) are allowed.");
        fileInput.value = "";
        return;
    }
    if (!Number.isInteger(clientId) || clientId <= 0 ||
        !Number.isInteger(collRequestId) || collRequestId <= 0 ||
        !Number.isInteger(userId) || userId < 0) {
        showUploadError("Please enter valid Client ID, Coll Request ID, and User ID values.");
        fileInput.value = "";
        return;
    }
    const formData = new FormData();
    formData.append("excelFile", file);
    formData.append("clientId", String(clientId));
    formData.append("collRequestId", String(collRequestId));
    formData.append("userId", String(userId));
    uploadButton.disabled = true;
    try {
        console.log("Upload request started");
        const response = await fetch("/Documents/UploadTitleDocument", {
            method: "POST", body: formData
        });
        console.log("Upload HTTP status:", response.status);
        const responseText = await response.text();
        let result = null;
        try {
            result = responseText ? JSON.parse(responseText) : null;
        } catch {
            result = null;
        }
        if (!response.ok || !result || !result.success) {
            throw new Error(result?.message || responseText.trim() || "The title document could not be processed.");
        }
        showUploadNotification("title", result.message, "success");
        await loadTitleDocuments(clientId, collRequestId);
    } catch (error) {
        console.error("Title document upload error:", error);
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.log("Upload error message:", errorMessage);
        showUploadError(errorMessage || "The title document could not be processed.");
    } finally {
        fileInput.value = "";
        uploadButton.disabled = false;
    }
}
async function loadLoanDocuments(clientId, collRequestId) {    
    const form = document.getElementById("excelUploadForm");  
    const loanTableContainer = document.getElementById("loanDocumentsTableContainer");    
    if (!form || !loanTableContainer) {
        throw new Error("The document loading controls were not found.");
    }    
    const result = await fetchDocumentsByCategory(collRequestId, clientId, 1);      
    if (!Array.isArray(result)) {
        throw new Error("The Loan documents response did not contain an array of rows.");
    }
    const rows = result;
    const loanRows = rows.filter(row => Number(row.DocCategory ?? row.docCategory) === 1);    
    loanTableContainer.replaceChildren();    
    appendDocumentTable(loanTableContainer, loanRows, "loan-documents-table", "No loan documents found.");
}
async function loadTitleDocuments(clientId, collRequestId) {
    const form = document.getElementById("excelUploadForm");
    const titleTableContainer = document.getElementById("titleDocumentsTableContainer");
    if (!form || !titleTableContainer) {
        throw new Error("The document loading controls were not found.");
    }
    const result = await fetchDocumentsByCategory(collRequestId, clientId, 2);
    if (!Array.isArray(result)) {
        throw new Error("The Title documents response did not contain an array of rows.");
    }
    const rows = result;
    const titleRows = rows.filter(row => Number(row.DocCategory ?? row.docCategory) === 2);
    titleTableContainer.replaceChildren();
    appendDocumentTable(titleTableContainer, titleRows, "title-documents-table", "No title documents found.");
}
async function fetchDocumentsByCategory(collRequestId, clientId, docCategory) {   
    const endpoint = docCategory === 1 ? "/Documents/LoadLoanDocuments" : "/Documents/LoadTitleDocuments";
    const query = new URLSearchParams({
        clientId: String(clientId),
        collRequestId: String(collRequestId)
    });
    const response = await fetch(`${endpoint}?${query.toString()}`);
    const responseText = await response.text();
    let result;
    try {
        result = responseText ? JSON.parse(responseText) : null;
    } catch {
        throw new Error("The document response was not valid JSON.");
    }
    if (!response.ok) {
        throw new Error(result?.message || "Unable to load documents.");
    }
    if (!result || result.success !== true) {
        throw new Error(result?.message || "Unable to load documents.");
    }

    if (!Array.isArray(result.data)) {
        throw new Error("The document response did not contain an array of rows.");
    }
    return result.data;
}
function appendDocumentTable(container, rows, tableClass, emptyMessage) {
    if (!Array.isArray(rows)) {
        throw new TypeError("Document table rows must be an array.");
    }
    if (rows.length === 0) {
        appendMessage(container, emptyMessage);
        return;
    }
    const tableWrapper = document.createElement("div");
    tableWrapper.className = "document-table-wrapper";    
    const table = document.createElement("table");
    table.className = `document-table ${tableClass}`;
    const headers = [
        ["Client ID", "ClientId"],
        ["Coll Request ID", "CollRequestId"],
        ["Doc ID", "DocId"],
        ["Doc Category", "DocCategory"],
        ["Doc Type", "DocType"],
        ["Doc Name", "DocName"],
        ["Doc Date", "DocDate"],
        ["Doc Mode", "DocMode"],
        ["No Of Pages", "NoOfPages"],
        ["Remark", "Remark"],
        ["PNumbering", "PNumbering"]
    ]; 
    const tableHead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    headers.forEach(([header]) => {
        const cell = document.createElement("th");
        cell.scope = "col";
        cell.textContent = header;
        headerRow.appendChild(cell);
    });
    tableHead.appendChild(headerRow);    
    table.appendChild(tableHead);  
    const tableBody = document.createElement("tbody"); 
    rows.forEach(row => {       
        const tableRow = document.createElement("tr");         
        headers.forEach(([, property]) => {
            const propertyName = property.charAt(0).toLowerCase() + property.slice(1);
            const rawValue = row[property] ?? row[propertyName];
            const value = property === "DocDate" ? formatDate(rawValue) : rawValue;            
            const cell = document.createElement("td");           
            cell.textContent = value === null || value === undefined? "": value; tableRow.appendChild(cell);
        });
        tableBody.appendChild(tableRow);
    });
    table.appendChild(tableBody);
    tableWrapper.appendChild(table);
    container.appendChild(tableWrapper);
}
function appendMessage(container, messageText) {
    const message = document.createElement("p");
    message.textContent = messageText;
    container.appendChild(message);
}
function formatDate(value) {
    if (!value) {
        return "";
    }
    const date = new Date(value);
    if (isNaN(date.getTime())) {
        return value;
    }
    return date.toISOString().split("T")[0];
}
async function loadInitialDocuments()
{
    const clientIdInput = document.getElementById("ClientId");
    const collRequestIdInput = document.getElementById("CollRequestId");
    const loanTableContainer = document.getElementById("loanDocumentsTableContainer");
    const titleTableContainer = document.getElementById("titleDocumentsTableContainer");

    if (!clientIdInput || !collRequestIdInput || !loanTableContainer || !titleTableContainer) {
        console.error("The initial document loading controls were not found.");
        return;
    }
    const clientId = Number(clientIdInput.value);
    const collRequestId = Number(collRequestIdInput.value);

    if (!Number.isInteger(clientId) || clientId <= 0 ||
        !Number.isInteger(collRequestId) || collRequestId <= 0) {
        loanTableContainer.replaceChildren();
        titleTableContainer.replaceChildren();
        appendMessage(loanTableContainer, "Enter valid Client ID and Coll Request ID values.");
        appendMessage(titleTableContainer, "Enter valid Client ID and Coll Request ID values.");
        return;
    }
    loanTableContainer.replaceChildren();
    titleTableContainer.replaceChildren();
    const documentLoads = [
        loadLoanDocuments(clientId, collRequestId),
        loadTitleDocuments(clientId, collRequestId)
    ];
    const results = await Promise.allSettled(documentLoads);
    results.forEach((result, index) => {
        if (result.status !== "rejected") {
            return;
        }
        const container = index === 0 ? loanTableContainer : titleTableContainer;
        container.replaceChildren();
        appendMessage(container, result.reason?.message || "Unable to load documents.");
    });
}
loadInitialDocuments();

