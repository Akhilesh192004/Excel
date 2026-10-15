//ShowExcelUploadPage
//UploadAndProcessExcelFile
//GetUploadedFileMetadata
//GetUploadedFileContents
//LoadDocuments

using Excel.Services;
using Excel.ViewModels;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System;

namespace Excel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ExcelUploadService _uploadService;

        public HomeController(ExcelUploadService uploadService)
        {
            _uploadService = uploadService;
        }

        [HttpGet]
        [ActionName("Index")]
        public IActionResult ShowExcelUploadPage()
        {
            return View("Index", new ExcelUploadViewModel());
        }

        [HttpPost]
        public async Task<IActionResult> UploadAndProcessExcelFile(ExcelUploadViewModel model, int clientId, int collRequestId, int docCategory, int userId)
        {
            if (model.ExcelFile == null || model.ExcelFile.Length == 0)
            {
                model.ErrorMessage = "Please select an Excel file.";
                return View("Index", model);
            }
            string extension = Path.GetExtension(model.ExcelFile.FileName).ToLowerInvariant();
            if (extension != ".xlsx")
            {
                model.ErrorMessage = "Only Excel files (.xlsx) are allowed.";
                return View("Index", model);
            }
            try
            {
                model.UploadId = await _uploadService.UploadExcelFileAndImportDetailsAsync(model.ExcelFile, clientId, collRequestId, docCategory, userId);
                long uploadId = model.UploadId.Value;
                DataTable dataTable = await _uploadService.GetUploadedExcelDetailRowsAsync(uploadId);
                if (dataTable.Rows.Count == 0)
                {
                    model.ErrorMessage = $"Upload completed, but no Excel detail rows were found. Upload ID: {model.UploadId}";
                    return View("Index", model);
                }
                dataTable = await _uploadService.AddItemIdToDataTableAsync(dataTable);
                DataTable uploadTable = await _uploadService.GetUploadMetadataAsync(uploadId);
                if (uploadTable.Rows.Count == 0)
                {
                    throw new Exception($"Upload record was not found for UploadId: {model.UploadId}.");
                }
                DataRow uploadRow = uploadTable.Rows[0];
                int storedClientId = Convert.ToInt32(uploadRow["ClientId"]);
                int storedCollRequestId = Convert.ToInt32(uploadRow["CollRequestId"]);
                int storedDocCategory = Convert.ToInt32(uploadRow["DocCategory"]);
                long maxDocId = await _uploadService.GetMaxDocIdAsync(storedClientId, storedCollRequestId);
                int insertedRows = await _uploadService.InsertDataTableToCollRequestDocsAsync(dataTable, storedClientId, storedCollRequestId, storedDocCategory, maxDocId);
                model.SuccessMessage = $"Excel processing worked correctly. Upload ID: {model.UploadId}. " + $"Rows inserted into: {insertedRows}";
            }
            catch (Exception ex)
            {
                model.ErrorMessage = ex.Message;
            }
            return View("Index", model);
        }

        [HttpPost]
        public async Task<IActionResult> GetUploadedFileMetadata([FromBody] long uploadId)
        {
            if (uploadId <= 0)
            {
                return BadRequest(new { success = false, message = "Enter a valid upload ID.", rowCount = 0 });
            }
            try
            {
                DataTable dataTable = await _uploadService.GetUploadMetadataAsync(uploadId);
                if (dataTable.Rows.Count == 0)
                {
                    return Json(new { success = false, message = "No matching uploaded file was found.", rowCount = 0 });
                }
                int rowCount = await _uploadService.ValidateUploadMetadataAsync(dataTable);
                return Json(new { success = true, message = "Record found successfully.", rowCount = rowCount });
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { success = false, message = ex.Message, rowCount = 0 });
            }
        }

        [HttpPost]
        public async Task<IActionResult> GetUploadedFileContents([FromBody] long uploadId)
        {
            if (uploadId <= 0)
            {
                return BadRequest(new { success = false, message = "Enter a valid upload ID.", rowCount = 0 });
            }
            try
            {
                DataTable dataTable = await _uploadService.GetUploadedExcelDetailRowsAsync(uploadId);
                if (dataTable.Rows.Count == 0)
                {
                    return Json(new { success = false, message = "No contents were found for this upload ID.", rowCount = 0, data = new List<object>() });
                }
                dataTable = await _uploadService.AddItemIdToDataTableAsync(dataTable);
                DataTable uploadTable = await _uploadService.GetUploadMetadataAsync(uploadId);
                if (uploadTable.Rows.Count == 0)
                {
                    throw new Exception($"Upload record was not found for UploadId: {uploadId}.");
                }
                DataRow uploadRow = uploadTable.Rows[0];
                int clientId = Convert.ToInt32(uploadRow["ClientId"]);
                int collRequestId = Convert.ToInt32(uploadRow["CollRequestId"]);
                int docCategory = Convert.ToInt32(uploadRow["DocCategory"]);
                long maxDocId = await _uploadService.GetMaxDocIdAsync(clientId, collRequestId);
                int insertedRows = await _uploadService.InsertDataTableToCollRequestDocsAsync(dataTable, clientId, collRequestId, docCategory, maxDocId);
                var data = dataTable.AsEnumerable().Select(row => new { Id = row["Id"], SrNo = row["SrNo"], DocName = row["DocName"], DocDate = row["DocDate"], DocModeName = row["DocModeName"], DocModeId = row["DocModeId"], NoOfPages = row["NoOfPages"], Remark = row["Remark"] }).ToList();
                return Json(new { success = true, message = "File contents retrieved successfully.", rowCount = dataTable.Rows.Count, insertedRows = insertedRows, data = data });
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { success = false, message = ex.Message, rowCount = 0 });
            }
        }
        [HttpPost]
        public async Task<IActionResult> LoadDocuments(int collRequestId, int clientId, int docCategory)
        {
            if (collRequestId <= 0 || clientId <= 0 || docCategory <= 0)
            {
                return BadRequest(new { success = false, message = "Client ID, Coll Request ID, and Doc Category must be valid.", rowCount = 0, data = new List<object>() });
            }
            try
            {
                DataTable dataTable = await _uploadService.GetDocumentsByCategoryAsync(collRequestId, clientId, docCategory);
                if (dataTable.Rows.Count == 0)
                {
                    return Json(new { success = false, message = "No documents were found.", rowCount = 0, data = new List<object>() });
                }
                string[] columns = { "ClientId", "CollRequestId", "DocId", "DocCategory", "DocType", "DocName", "DocDate", "DocMode", "NoOfPages", "Remark", "PNumbering" };
                var data = dataTable.AsEnumerable().Select(row => columns.ToDictionary(column => column, column => row.IsNull(column) ? null : row[column])).ToList();
                return Json(new { success = true, message = "Documents loaded successfully.", rowCount = data.Count, data });
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { success = false, message = ex.Message, rowCount = 0, data = new List<object>() });
            }
        }

        [HttpGet("/Documents/LoadLoanDocuments")]
        public Task<IActionResult> LoadLoanDocuments(int clientId, int collRequestId)
        {
            return LoadDocumentsByCategoryAsync(clientId, collRequestId, 1, "Loan");
        }

        [HttpGet("/Documents/LoadTitleDocuments")]
        public Task<IActionResult> LoadTitleDocuments(int clientId, int collRequestId)
        {
            return LoadDocumentsByCategoryAsync(clientId, collRequestId, 2, "Title");
        }
        private async Task<IActionResult> LoadDocumentsByCategoryAsync(int clientId, int collRequestId, int docCategory, string documentType)
        {
            if (clientId <= 0 || collRequestId <= 0)
            {
                return BadRequest(new { success = false, message = "Client Id and Coll request Id must be valid.", rowCount = 0, data = Array.Empty<object>() });
            }
            try
            {
                DataTable dataTable = await _uploadService.GetDocumentsByCategoryAsync(collRequestId, clientId, docCategory);
                string[] columns = { "ClientId", "CollRequestId", "DocId", "DocCategory", "DocType", "DocName", "DocDate", "DocMode", "NoOfPages", "Remark", "PNumbering" };
                var data = dataTable.AsEnumerable().Select(row => columns.ToDictionary(column => column, column => row.IsNull(column) ? null : row[column])).ToList();
                return Json(new { success = true, message = data.Count == 0 ? $"No {documentType} documents were found." : "Documents loaded successfully.", rowCount = data.Count, data });
            }
            catch
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { success = false, message = $"Unable to load {documentType} documents.", rowCount = 0, data = Array.Empty<object>() });
            }
        }
       
        [HttpPost("/Documents/UploadLoanDocument")]
        public async Task<IActionResult> UploadLoanDocument(IFormFile? excelFile, int clientId, int collRequestId, int userId)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { success = false, message = "Please select an Excel file.", uploadId = (long?)null, insertedRows = 0 });
            }
            if (!string.Equals(Path.GetExtension(excelFile.FileName), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new { success = false, message = "Only Excel files (.xlsx) are allowed.", uploadId = (long?)null, insertedRows = 0 });
            }
            if (clientId <= 0 || collRequestId <= 0 || userId < 0)
            {
                return BadRequest(new { success = false, message = "Client ID and Coll Request ID must be positive integers, and User ID must be a valid non-negative integer.", uploadId = (long?)null, insertedRows = 0 });
            }
            try
            {
                const int loanCategory = 1;
                long uploadId = await _uploadService.UploadExcelFileAndImportDetailsAsync(excelFile, clientId, collRequestId, loanCategory, userId);
                DataTable dataTable = await _uploadService.GetUploadedExcelDetailRowsAsync(uploadId);
                if (dataTable.Rows.Count == 0)
                {
                    return UnprocessableEntity(new { success = false, message = "Upload completed, but no Excel detail rows were found.", uploadId, insertedRows = 0 });
                }
                dataTable = await _uploadService.AddItemIdToDataTableAsync(dataTable);
                DataTable uploadMetadata = await _uploadService.GetUploadMetadataAsync(uploadId);
                if (uploadMetadata.Rows.Count == 0)
                {
                    return UnprocessableEntity(new { success = false, message = "The uploaded file metadata could not be found.", uploadId, insertedRows = 0 });
                }
                long maxDocId = await _uploadService.GetMaxDocIdAsync(clientId, collRequestId);
                int insertedRows = await _uploadService.InsertDataTableToCollRequestDocsAsync(dataTable, clientId, collRequestId, loanCategory, maxDocId);
                return Json(new { success = true, message = $"Loan document processed successfully. Rows inserted: {insertedRows}.", uploadId, insertedRows });
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { success = false, message = ex.InnerException?.Message ?? ex.Message, uploadId = (long?)null, insertedRows = 0 });
            }
        }

        [HttpPost("/Documents/UploadTitleDocument")]
        public async Task<IActionResult> UploadTitleDocument(IFormFile? excelFile, int clientId, int collRequestId, int userId)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { success = false, message = "Please select an Excel file.", uploadId = (long?)null, insertedRows = 0 });
            }
            if (!string.Equals(Path.GetExtension(excelFile.FileName), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new { success = false, message = "Only Excel files (.xlsx) are allowed.", uploadId = (long?)null, insertedRows = 0 });
            }
            if (clientId <= 0 || collRequestId <= 0 || userId < 0)
            {
                return BadRequest(new { success = false, message = "Client ID and Coll Request ID must be positive integers, and User ID must be a valid non-negative integer.", uploadId = (long?)null, insertedRows = 0 });
            }
            try
            {
                const int titleCategory = 2;
                long uploadId = await _uploadService.UploadExcelFileAndImportDetailsAsync(excelFile, clientId, collRequestId, titleCategory, userId);
                DataTable dataTable = await _uploadService.GetUploadedExcelDetailRowsAsync(uploadId);
                if (dataTable.Rows.Count == 0)
                {
                    return UnprocessableEntity(new { success = false, message = "Upload completed, but no Excel detail rows were found.", uploadId, insertedRows = 0 });
                }
                dataTable = await _uploadService.AddItemIdToDataTableAsync(dataTable);
                DataTable uploadMetadata = await _uploadService.GetUploadMetadataAsync(uploadId);
                if (uploadMetadata.Rows.Count == 0)
                {
                    return UnprocessableEntity(new { success = false, message = "The uploaded file metadata could not be found.", uploadId, insertedRows = 0 });
                }
                long maxDocId = await _uploadService.GetMaxDocIdAsync(clientId, collRequestId);
                int insertedRows = await _uploadService.InsertDataTableToCollRequestDocsAsync(dataTable, clientId, collRequestId, titleCategory, maxDocId);
                return Json(new { success = true, message = $"Title document processed successfully. Rows inserted: {insertedRows}.", uploadId, insertedRows });
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { success = false, message = ex.InnerException?.Message ?? ex.Message, uploadId = (long?)null, insertedRows = 0 });
            }
        }
    }
}
