//UploadExcelFileAndImportDetailsAsync
//SaveUploadedExcelFileAsync
//InsertExcelUploadRecordAsync
//ValidateUploadedExcelHeaders

using ExcelDataReader;
using Excel.ViewModels;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Text;

namespace Excel.Services
{
    public sealed class ExcelUploadService
    {
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _environment;
        public ExcelUploadService(IConfiguration configuration, IWebHostEnvironment environment)
        {
            _configuration = configuration;
            _environment = environment;
        }

        public async Task<long> UploadExcelFileAndImportDetailsAsync(IFormFile excelFile, int clientId, int collRequestId, int docCategory, int userId)
        {
            string savedFilePath = await SaveUploadedExcelFileAsync(excelFile);
            if (!File.Exists(savedFilePath))
            {
                throw new Exception($"FILE WAS NOT CREATED.\n\n" +
                                    $"Expected path:\n{savedFilePath}");
            }
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using FileStream stream = new FileStream(savedFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                });
                if (dataSet.Tables.Count == 0)
                {
                    throw new Exception("The Excel file does not contain any worksheet.");
                }
                DataTable table = dataSet.Tables[0];
                ValidateUploadedExcelHeaders(table, docCategory);
                string connectionString = _configuration.GetConnectionString("DefaultConnection") ?? throw new Exception("Database connection string was not found.");
                using SqlConnection connection = new SqlConnection(connectionString);
                await connection.OpenAsync();
                using SqlTransaction transaction = connection.BeginTransaction();
                bool transactionCommitted = false;
                try
                {
                    long uploadId = await InsertExcelUploadRecordAsync(connection, transaction, Path.GetFileName(savedFilePath), clientId, collRequestId, docCategory, userId);
                    await transaction.CommitAsync();
                    transactionCommitted = true;
                    await ImportExcelDetailRowsAsync(uploadId, Path.GetFileName(savedFilePath));
                    return uploadId;
                }
                catch
                {
                    if (!transactionCommitted)
                    {
                        await transaction.RollbackAsync();
                    }

                    throw;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"UPLOAD FAILED.\n\n" + $"Saved file path:\n" + $"{savedFilePath}\n\n" + $"File exists:\n" + $"{File.Exists(savedFilePath)}\n\n" + $"Error:\n" + $"{ex.Message}", ex);
            }
        }

        public async Task<string> SaveUploadedExcelFileAsync(IFormFile excelFile)
        {
            if (excelFile == null)
            {
                throw new Exception("Excel file was not provided.");
            }

            if (excelFile.Length == 0)
            {
                throw new Exception("The uploaded Excel file is empty.");
            }
            string uploadsFolder = Path.Combine(_environment.ContentRootPath, "wwwroot", "Files");
            Console.WriteLine($"Excel upload directory: {uploadsFolder}");
            Directory.CreateDirectory(uploadsFolder);
            string fileName = Path.GetFileName(excelFile.FileName);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new Exception("Invalid Excel file name.");
            }
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            string filePath = Path.Combine(uploadsFolder, fileName);
            int fileNumber = 1;
            while (File.Exists(filePath))
            {
                string newFileName = $"{fileNameWithoutExtension}_{fileNumber}{extension}";
                filePath = Path.Combine(uploadsFolder, newFileName);
                fileNumber++;
            }
            await using (FileStream fileStream = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                await excelFile.CopyToAsync(fileStream);
            }
            if (!File.Exists(filePath))
            {

                throw new Exception($"FILE WAS NOT CREATED.\n\n" +
                                     $"Content root:\n" +
                                     $"{_environment.ContentRootPath}\n\n" +
                                    $"Files directory:\n" +
                                    $"{uploadsFolder}\n\n" +
                                    $"Expected file:\n" +
                                    $"{filePath}");
            }
            return filePath;
        }

        private static async Task<long> InsertExcelUploadRecordAsync(SqlConnection connection, SqlTransaction transaction, string fileName, int clientId, int collRequestId, int docCategory, int userId)
        {
            const string query = @"
                INSERT INTO dbo.ExcelUploads
                (
                    FileName,
                    UserId,
                    UploadedAt,
                    ClientId,
                    CollRequestId,
                    DocCategory
                )
                OUTPUT INSERTED.Id
                VALUES
                (
                    @FileName,
                    @UserID,
                    GETDATE(),
                    @ClientId,
                    @CollRequestId,
                    @DocCategory
                );";

            using SqlCommand command = new SqlCommand(query, connection, transaction);
            command.Parameters.AddWithValue("@FileName", fileName);
            command.Parameters.Add("@UserID", SqlDbType.Int).Value = userId;
            command.Parameters.Add("@ClientId", SqlDbType.Int).Value = clientId;
            command.Parameters.Add("@CollRequestId", SqlDbType.Int).Value = collRequestId;
            command.Parameters.Add("@DocCategory", SqlDbType.Int).Value = docCategory;
            object? result = await command.ExecuteScalarAsync();
            if (result == null || result == DBNull.Value)
            {
                throw new Exception("Unable to create upload record.");
            }
            return Convert.ToInt64(result);
        }

        private static void ValidateUploadedExcelHeaders(DataTable table, int docCategory)
        {
            string firstColumn = docCategory switch
            {
                1 => "Name of the Loan Documents", 2 => "Name of the Title Documents", _ => throw new Exception($"Unsupported document category '{docCategory}'. Use 1 for Loan or 2 for Title.")
            };
            string[] requiredHeaders =
            {
                firstColumn,
                "Date of Document",
                "Document Mode",
                "No. of Pages",
                "Remark"
            };
            foreach (string header in requiredHeaders)
            {
                if (!table.Columns.Contains(header))
                {
                    throw new Exception($"Required Excel column '{header}' was not found.");
                }
            }
        }

        private static string ConvertToTrimmedString(object? value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }
            return value.ToString()?.Trim() ?? string.Empty;
        }

        private static DateTime? ConvertToNullableDate(object? value)
        {
            if (value == null || value == DBNull.Value)
            {
                return null;
            }
            if (value is DateTime date)
            {
                return date;
            }
            return DateTime.TryParse(value.ToString(), out DateTime parsedDate) ? parsedDate : null;
        }

        private static int? ConvertToNullableInteger(object? value)
        {
            if (value == null || value == DBNull.Value)
            {
                return null;
            }
            if (int.TryParse(value.ToString(), out int result))
            {
                return result;
            }
            return double.TryParse(value.ToString(), out double doubleResult) ? Convert.ToInt32(doubleResult) : null;
        }

        public async Task<int> ImportExcelDetailRowsAsync(long uploadId, string fileName)
        {
            if (uploadId <= 0)
            {
                throw new Exception("Invalid Upload ID.");
            }
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new Exception("Excel file name was not provided.");
            }
            string filesFolder = Path.Combine(_environment.ContentRootPath, "wwwroot", "Files");
            if (!Directory.Exists(filesFolder))
            {
                throw new Exception($"Files folder was not found:\n{filesFolder}");
            }
            string safeFileName = Path.GetFileName(fileName);
            string filePath = Path.Combine(filesFolder, safeFileName);
            if (!File.Exists(filePath))
            {
                throw new Exception($"Excel file was not found.\n\n" + $"Expected path:\n{filePath}");
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            await using FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
            DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });
            if (dataSet.Tables.Count == 0)
            {
                throw new Exception("The Excel file does not contain any worksheet.");
            }
            DataTable table = dataSet.Tables[0];
            DataTable uploadMetadata = await GetUploadMetadataAsync(uploadId);
            if (uploadMetadata == null || uploadMetadata.Rows.Count == 0)
            {
                throw new Exception("Upload metadata was not found.");
            }
            int docCategory = Convert.ToInt32(uploadMetadata.Rows[0]["DocCategory"]);
            ValidateExcelDetailHeaders(table, docCategory);
            string connectionString = _configuration.GetConnectionString("DefaultConnection") ?? throw new Exception("Database connection string was not found.");
            using SqlConnection connection = new SqlConnection(connectionString);
            await connection.OpenAsync();
            using SqlTransaction transaction = connection.BeginTransaction();
            try
            {
                const string query = @"
INSERT INTO dbo.ExcelUploadDets
(
    Id,
    SrNo,
    DocName,
    DocDate,
    DocModeName,
    NoOfPages,
    Remark
)
VALUES
(
    @Id,
    @SrNo,
    @DocName,
    @DocDate,
    @DocModeName,
    @NoOfPages,
    @Remark
);";
                int insertedRowCount = 0;
                long srNo = 1;
                string firstColumnName = docCategory switch
                {
                    1 => "Name of the Loan Documents",
                    2 => "Name of the Title Documents",
                    _ => throw new Exception($"Unsupported document category '{docCategory}'. Use 1 for Loan or 2 for Title.")
                };
                int excelRowNumber = 2;
                foreach (DataRow row in table.Rows)
                {
                    string docName = ConvertToTrimmedString(row[firstColumnName]);
                    DateTime? docDate = ConvertToNullableDate(row["Date of Document"]);
                    string docModeName = ConvertToTrimmedString(row["Document Mode"]);
                    int? noOfPages = ConvertToNullableInteger(row["No. of Pages"]);
                    string remark = ConvertToTrimmedString(row["Remark"]);
                    if (
                        string.IsNullOrWhiteSpace(docName) && docDate == null && string.IsNullOrWhiteSpace(docModeName) && noOfPages == null && string.IsNullOrWhiteSpace(remark))
                    {
                        excelRowNumber++;
                        continue;
                    }
                    string? normalizedDocModeName = NormalizeDocumentMode(docModeName);
                    if (normalizedDocModeName == null)
                    {
                        string suppliedValue = string.IsNullOrWhiteSpace(docModeName) ? "(empty)" : $"'{docModeName}'";
                        throw new Exception($"Invalid Document Mode {suppliedValue} at Excel row {excelRowNumber}. Valid values are Original, Certified, Photo, PhotoCopy, CertifiedCopy, True, and TrueCopy.");
                    }

                    // Create INSERT command
                    using SqlCommand command = new SqlCommand(query, connection, transaction);
                    command.Parameters.Add("@Id", SqlDbType.BigInt).Value = uploadId;
                    command.Parameters.Add("@SrNo", SqlDbType.Decimal).Value = srNo;
                    command.Parameters.Add("@DocName", SqlDbType.NVarChar, 200).Value = docName;
                    command.Parameters.Add("@DocDate", SqlDbType.Date).Value = docDate ?? (object)DBNull.Value;
                    command.Parameters.Add("@DocModeName", SqlDbType.NVarChar, 200).Value = normalizedDocModeName;
                    command.Parameters.Add("@NoOfPages", SqlDbType.Decimal).Value = noOfPages ?? 0;
                    command.Parameters.Add("@Remark", SqlDbType.VarChar, 200).Value = remark;
                    // Execute INSERT
                    await command.ExecuteNonQueryAsync();
                    insertedRowCount++;
                    srNo++;
                    excelRowNumber++;
                }
                await transaction.CommitAsync();
                return insertedRowCount;
            }
            catch
            {
                await transaction.RollbackAsync();
                throw;
            }
        }

        public async Task<DataTable> GetUploadMetadataAsync(long uploadId)
        {
            const string query = @"
        SELECT
            Id AS UploadId,
            FileName,
            ClientId,
            CollRequestId,
            DocCategory
        FROM dbo.ExcelUploads
        WHERE Id = @UploadId;";

            DataTable dataTable = new DataTable();

            string connectionString =
                _configuration.GetConnectionString("DefaultConnection")
                ?? throw new Exception("Database connection string was not found.");

            using SqlConnection connection = new SqlConnection(connectionString);
            using SqlCommand command = new SqlCommand(query, connection);

            command.Parameters.Add("@UploadId", SqlDbType.BigInt).Value = uploadId;

            await connection.OpenAsync();

            using SqlDataReader reader = await command.ExecuteReaderAsync();
            dataTable.Load(reader);

            return dataTable;
        }

        private static void ValidateExcelDetailHeaders(DataTable table, int docCategory)
        {
            string firstColumn = docCategory switch
            {
                1 => "Name of the Loan Documents",
                2 => "Name of the Title Documents",
                _ => throw new Exception(
                    $"Unsupported document category '{docCategory}'. Use 1 for Loan or 2 for Title.")
            };

            string[] requiredHeaders =
            {
        firstColumn,
        "Date of Document",
        "Document Mode",
        "No. of Pages",
        "Remark"
    };

            foreach (string header in requiredHeaders)
            {
                if (!table.Columns.Contains(header))
                {
                    throw new Exception(
                        $"Required Excel column '{header}' was not found.");
                }
            }
        }

        private static string? NormalizeDocumentMode(string? documentMode)
        {
            if (string.IsNullOrWhiteSpace(documentMode))
                return null;

            string mode = new string(
                documentMode.Where(character => !char.IsWhiteSpace(character)).ToArray()
            ).ToUpperInvariant();

            return mode switch
            {
                "ORIGINAL" => "Original",
                "PHOTO" or "PHOTOCOPY" => "PhotoCopy",
                "CERTIFIED" or "CERTIFIEDCOPY" => "CertifiedCopy",
                "TRUE" or "TRUECOPY" => "TrueCopy",
                _ => null
            };
        }

        public async Task<DataTable> GetDocumentsByCategoryAsync(int collRequestId, int clientId, int docCategory)
        {
            const string query = @"
SELECT
    ClientId,
    CollRequestId,
    DocId,
    DocCategory,
    DocType,
    DocName,
    DocDate,
    mode.ItemId AS DocMode,
    NoOfPages,
    Remark,
    PNumbering
FROM dbo.T_CollRequestDocs AS docs
LEFT JOIN dbo.M_Master AS mode
    ON mode.MasterId = 6
   AND mode.ItemId = docs.DocMode
WHERE docs.CollRequestId = @CollRequestId
  AND docs.ClientId = @ClientId
  AND docs.DocCategory = @DocCategory
ORDER BY docs.DocId;";

            DataTable dataTable = new DataTable();
            string connectionString = _configuration.GetConnectionString("DefaultConnection") ?? throw new Exception("Database connection string was not found.");
            using SqlConnection connection = new SqlConnection(connectionString);
            using SqlCommand command = new SqlCommand(query, connection);
            command.Parameters.Add("@CollRequestId", SqlDbType.Int).Value = collRequestId;
            command.Parameters.Add("@ClientId", SqlDbType.Int).Value = clientId;
            command.Parameters.Add("@DocCategory", SqlDbType.Int).Value = docCategory;
            await connection.OpenAsync();
            using SqlDataReader reader = await command.ExecuteReaderAsync();
            dataTable.Load(reader);
            return dataTable;
        }

        public async Task<DataTable> GetUploadedExcelDetailRowsAsync(long uploadId)
        {
            const string query = @"
SELECT Id, SrNo, DocName, DocDate, DocModeName, NoOfPages, Remark
FROM dbo.ExcelUploadDets
WHERE Id = @UploadId
ORDER BY SrNo;";

            DataTable dataTable = new DataTable();
            string connectionString = _configuration.GetConnectionString("DefaultConnection") ?? throw new Exception("Database connection string was not found.");
            using SqlConnection connection = new SqlConnection(connectionString);
            using SqlCommand command = new SqlCommand(query, connection);
            command.Parameters.Add("@UploadId", SqlDbType.BigInt).Value = uploadId;
            await connection.OpenAsync();
            using SqlDataReader reader = await command.ExecuteReaderAsync();
            dataTable.Load(reader);
            return dataTable;
        }

        public async Task<DataTable> GetDocumentModeMasterItemsAsync()
        {
            const string query = @"
                                    SELECT ItemId, Name
                                    FROM dbo.M_Master
                                    WHERE MasterId = 6;";
            DataTable dataTable = new DataTable();
            string connectionString = _configuration.GetConnectionString("DefaultConnection") ?? throw new Exception("Database connection string was not found.");
            using SqlConnection connection = new SqlConnection(connectionString);
            using SqlCommand command = new SqlCommand(query, connection);
            await connection.OpenAsync();
            using SqlDataReader reader = await command.ExecuteReaderAsync();
            dataTable.Load(reader);
            return dataTable;
        }

        public async Task<long> GetMaxDocIdAsync(int clientId, int collRequestId)
        {
            const string query = @"
        SELECT ISNULL(MAX(DocId), 0)
        FROM dbo.T_CollRequestDocs
        WHERE ClientId = @ClientId
          AND CollRequestId = @CollRequestId;";
            string connectionString = _configuration.GetConnectionString("DefaultConnection") ?? throw new Exception("Database connection string was not found.");
            using SqlConnection connection = new SqlConnection(connectionString);
            using SqlCommand command = new SqlCommand(query, connection);
            command.Parameters.Add("@ClientId", SqlDbType.Int).Value = clientId;
            command.Parameters.Add("@CollRequestId", SqlDbType.Int).Value = collRequestId;
            await connection.OpenAsync();
            object? result = await command.ExecuteScalarAsync();
            return Convert.ToInt64(result);
        }

        public async Task<int> InsertDataTableToCollRequestDocsAsync(DataTable dataTable, int clientId, int collRequestId, int docCategory, long maxDocId)
        {
            const string query = @"
INSERT INTO dbo.T_CollRequestDocs
(
    ClientId, CollRequestId, DocId, DocCategory, DocType,
    DocName, DocDate, DocMode, NoOfPages, Remark, PNumbering
)
VALUES
(
    @ClientId, @CollRequestId, @DocId, @DocCategory, @DocType,
    @DocName, @DocDate, @DocMode, @NoOfPages, @Remark, @PNumbering
);";

            string connectionString = _configuration.GetConnectionString("DefaultConnection") ?? throw new Exception("Database connection string was not found.");
            using SqlConnection connection = new SqlConnection(connectionString);
            await connection.OpenAsync();
            using SqlTransaction transaction = connection.BeginTransaction();
            try
            {
                int insertedRowCount = 0;
                foreach (DataRow row in dataTable.Rows)
                {
                    long srNo = Convert.ToInt64(row["SrNo"]);
                    long docId = maxDocId + srNo;
                    using SqlCommand command = new SqlCommand(query, connection, transaction);
                    command.Parameters.Add("@ClientId", SqlDbType.Int).Value = clientId;
                    command.Parameters.Add("@CollRequestId", SqlDbType.Int).Value = collRequestId;
                    command.Parameters.Add("@DocId", SqlDbType.BigInt).Value = docId;
                    command.Parameters.Add("@DocCategory", SqlDbType.Int).Value = docCategory;
                    command.Parameters.Add("@DocType", SqlDbType.Int).Value = 0;
                    command.Parameters.Add("@DocName", SqlDbType.NVarChar, 255).Value = row["DocName"] == DBNull.Value ? string.Empty : row["DocName"];
                    command.Parameters.Add("@DocDate", SqlDbType.Date).Value = row["DocDate"] == DBNull.Value ? DBNull.Value : row["DocDate"];
                    command.Parameters.Add("@DocMode", SqlDbType.BigInt).Value = row["DocModeId"] == DBNull.Value ? 0L : Convert.ToInt64(row["DocModeId"]);
                    command.Parameters.Add("@NoOfPages", SqlDbType.Int).Value = row["NoOfPages"] == DBNull.Value ? 0 : Convert.ToInt32(row["NoOfPages"]);
                    command.Parameters.Add("@Remark", SqlDbType.VarChar, 255).Value = row["Remark"] == DBNull.Value ? string.Empty : row["Remark"];
                    command.Parameters.Add("@PNumbering", SqlDbType.Int).Value = 1;
                    await command.ExecuteNonQueryAsync();
                    insertedRowCount++;
                }
                await transaction.CommitAsync();
                return insertedRowCount;
            }
            catch
            {
                await transaction.RollbackAsync();
                throw;
            }
        }

        public async Task<DataTable> AddItemIdToDataTableAsync(DataTable dataTable)
        {
            if (dataTable == null)
            {
                throw new ArgumentNullException(nameof(dataTable));
            }
            if (!dataTable.Columns.Contains("DocModeName"))
            {
                throw new Exception("The DataTable does not contain a DocModeName column.");
            }
            if (!dataTable.Columns.Contains("DocModeId"))
            {
                dataTable.Columns.Add("DocModeId", typeof(long));
            }
            DataTable masterTable = await GetDocumentModeMasterItemsAsync();
            foreach (DataRow row in dataTable.Rows)
            {
                string docModeName = row["DocModeName"]?.ToString() ?? string.Empty;
                string? normalizedDocModeName = NormalizeDocumentMode(docModeName);
                if (normalizedDocModeName == null)
                {
                    string suppliedValue = string.IsNullOrWhiteSpace(docModeName) ? "(empty)" : $"'{docModeName}'";
                    throw new Exception($"Invalid Document Mode {suppliedValue} at uploaded row {row["SrNo"]}. Valid values are Original, CertifiedCopy, PhotoCopy, and TrueCopy.");
                }

                long docModeId = 0;
                foreach (DataRow masterRow in masterTable.Rows)
                {
                    string masterName = masterRow["Name"]?.ToString() ?? string.Empty;
                    string? normalizedMasterName = NormalizeDocumentMode(masterName);
                    if (string.Equals(normalizedDocModeName, normalizedMasterName, StringComparison.Ordinal))
                    {
                        docModeId = Convert.ToInt64(masterRow["ItemId"]);

                        break;
                    }
                }
                if (docModeId == 0)
                {
                    throw new Exception($"Document Mode '{docModeName}' at uploaded row {row["SrNo"]} does not match a value in the document mode master table.");
                }

                row["DocModeId"] = docModeId;
            }
            return dataTable;
        }

        public async Task<int> ValidateUploadMetadataAsync(DataTable uploadMetadata)
        {
            if (uploadMetadata == null || uploadMetadata.Rows.Count == 0)
            {
                throw new Exception("No upload data was found.");
            }
            DataRow row = uploadMetadata.Rows[0];
            _ = Convert.ToInt64(row["UploadId"]);
            _ = row["FileName"]?.ToString() ?? string.Empty;
            _ = Convert.ToInt32(row["CollRequestId"]);
            _ = Convert.ToInt32(row["DocCategory"]);
            _ = Convert.ToInt32(row["ClientId"]);
            return 0;
        }
    }
}