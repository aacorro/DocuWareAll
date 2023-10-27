using DocuWare.Platform.ServerClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;


namespace Bel_DocuWare
{

    public class DocuWare
    {
        static string defaultDownloadLocation = @"C:\Users\Developer\Desktop\DownloadDefault";
        static ServiceConnection _SVCConnection = null;

        public ServiceConnection GetServiceConnection()
        {
            return _SVCConnection;
        }

        #region _SVCConnect
        public async Task<bool> ConnectAsync(Uri uri, string username, string password)
        {
            try
            {
                _SVCConnection = await ServiceConnection.CreateAsync(uri, username, password).ConfigureAwait(false);

                if (_SVCConnection != null)
                {
                    return true;
                }
                else
                {
                    Console.WriteLine("Failed to establish a connection to DocuWare.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                return false;
            }
        }

        #endregion


        #region _SVCDisconnect
        public bool Disconnect()
        {
            try
            {
                _SVCConnection.Disconnect();

                return true;
            }
            catch (Exception ex)
            {
                Console.Write($"An error occurred: {ex.Message}");

                return false;
            }

        }
        #endregion


        #region List Docs
        public static async Task<List<Document>> ListAllDocumentsAsync(ServiceConnection _SVCConnection, string fileCabinetId, int? count = 100)
        {
            try
            {
                Console.Write("Retrieving documents...");
                DocumentsQueryResult queryResult = await _SVCConnection.GetFromDocumentsForDocumentsQueryResultAsync(
                    fileCabinetId,
                    count: count)
                    .ConfigureAwait(false);

                List<Document> result = new List<Document>();
                await GetAllDocumentsAsync(queryResult, result);

                // Clear the "Retrieving documents..." informative message by overwriting it with an empty line
                Console.Write("\r" + new string(' ', Console.WindowWidth - 1) + "\r");

                Console.WriteLine("Documents:");
                foreach (Document document in result)
                {
                    Console.WriteLine($"Document with ID: {document.Id} created at {document.CreatedAt}");
                }
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while listing documents: " + ex.Message);
                return null;
            }
        }


        #endregion


        #region Get All
        public static async Task GetAllDocumentsAsync(DocumentsQueryResult queryResult, List<Document> documents)
        {
            try
            {
                documents.AddRange(queryResult.Items);

                if (queryResult.NextRelationLink != null)
                {
                    await GetAllDocumentsAsync(await queryResult.GetDocumentsQueryResultFromNextRelationAsync(), documents);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while getting documents: " + ex.Message);
            }
        }

        #endregion


        #region Download and Display
        public static async Task DownloadAndDisplayDocumentContentAsync(ServiceConnection _SVCConnection, string fileCabinetId, string documentId, string filePath)
        {
            try
            {
                var documents = await ListAllDocumentsAsync(_SVCConnection, fileCabinetId);

                if (documents.Count > 0)
                {
                    // Find the document by ID
                    var selectedDocument = documents.FirstOrDefault(doc => doc.Id == int.Parse(documentId));

                    if (selectedDocument != null)
                    {
                        Console.WriteLine("Downloading the selected document...");

                        // Download and save the document using the provided filePath
                        await DownloadAndSaveDocument(selectedDocument, filePath);

                        Console.WriteLine("Downloaded and saved the selected document.");
                    }
                    else
                    {
                        Console.WriteLine("Document with the specified ID not found.");
                    }
                }
                else
                {
                    Console.WriteLine("No documents found in the file cabinet.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while downloading and displaying document content: " + ex.Message);
            }
        }


        private static async Task DownloadAndSaveDocument(Document document, string downloadLocation)
        {
            var fileDownloadResult = await DownloadDocumentContentAsync(document);

            if (fileDownloadResult != null && fileDownloadResult.Stream != null)
            {
                // Decode the file name from URL encoding
                string decodedFileName = Uri.UnescapeDataString(fileDownloadResult.FileName);

                // Clean the file name to remove any problematic characters
                string cleanedFileName = Path.GetInvalidFileNameChars()
                    .Aggregate(decodedFileName, (current, c) => current.Replace(c.ToString(), string.Empty));

                // Construct the full file path with the cleaned file name
                string filePath = Path.Combine(downloadLocation, cleanedFileName);

                // Save the downloaded content to the specified location
                using (var fileStream = File.Create(filePath))
                {
                    await fileDownloadResult.Stream.CopyToAsync(fileStream);
                }

                Console.WriteLine($"Downloaded and saved: {cleanedFileName}");
            }
            else
            {
                Console.WriteLine($"Failed to download document content for ID: {document.Id}");
            }
        }


        #endregion


        #region Download Doc Content

        public class FileDownloadResult
        {
            public string ContentType { get; set; }
            public string FileName { get; set; }
            public long? ContentLength { get; set; }
            public System.IO.Stream Stream { get; set; }

            // Error properties
            public bool HasError { get; set; } // Indicates whether an error occurred
            public string ErrorMessage { get; set; } // Stores the error message

            public static FileDownloadResult FromError(string errorMessage)
            {
                return new FileDownloadResult
                {
                    HasError = true,
                    ErrorMessage = errorMessage
                };
            }
        }


        public static async Task<FileDownloadResult> DownloadDocumentContentAsync(Document document)
        {
            try
            {
                if (document.FileDownloadRelationLink == null)
                    document = await document.GetDocumentFromSelfRelationAsync().ConfigureAwait(false);

                var downloadResponse = await document.PostToFileDownloadRelationForStreamAsync(
                    new FileDownload()
                    {
                        TargetFileType = FileDownloadType.Auto
                    })
                    .ConfigureAwait(false);

                var contentHeaders = downloadResponse.ContentHeaders;
                return new FileDownloadResult
                {
                    Stream = downloadResponse.Content,
                    ContentLength = contentHeaders.ContentLength,
                    ContentType = contentHeaders.ContentType.MediaType,
                    FileName = contentHeaders.ContentDisposition.FileName
                };
            }
            catch (Exception ex)
            {
                return FileDownloadResult.FromError("An error occurred while downloading the document: " + ex.Message);
            }
        }


        #endregion


        #region Upload To File Cabinet
        public static async Task<Document> UploadSingleFileToFileCabinetAsync(ServiceConnection _SVCConnection, string fileCabinetId, string filePathUpload)
        {
            try
            {
                // Retrieve the FileCabinet using the provided fileCabinetId
                var org = _SVCConnection.Organizations[0]; 
                var fileCabinets = org.GetFileCabinetsFromFilecabinetsRelation().FileCabinet;

                // Find the FileCabinet with the matching ID
                FileCabinet selectedFileCabinet = fileCabinets.FirstOrDefault(fc => fc.Id == fileCabinetId);

                if (selectedFileCabinet == null)
                {
                    Console.WriteLine($"File Cabinet with ID {fileCabinetId} not found.");
                    return null;
                }

                if (File.Exists(filePathUpload))
                {
                    var fileInfo = new FileInfo(filePathUpload);

                    // Upload the document to the selected File Cabinet
                    Console.Write("Uploading document...");
                    var uploadedDocument = await selectedFileCabinet.UploadDocumentAsync(fileInfo).ConfigureAwait(false);
                    Console.WriteLine("\r" + new string(' ', Console.WindowWidth - 1) + "\rDocument uploaded!");
                    return uploadedDocument;
                }
                else
                {
                    Console.WriteLine($"File not found at the specified path: {filePathUpload}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while uploading the document: " + ex.Message);
                return null;
            }
        }


        #endregion


        #region Delete

        public static void DeleteDocumentById(ServiceConnection _SVCConnection, string fileCabinetId, string documentId)
        {
            try
            {

                // Retrieve the FileCabinet using the provided fileCabinetId
                var org = _SVCConnection.Organizations[0];
                var fileCabinets = org.GetFileCabinetsFromFilecabinetsRelation().FileCabinet;

                // Find the FileCabinet with the matching ID
                FileCabinet selectedFileCabinet = fileCabinets.FirstOrDefault(fc => fc.Id == fileCabinetId);

                if (selectedFileCabinet == null)
                {
                    Console.WriteLine($"File Cabinet with ID {fileCabinetId} not found.");
                    return;
                }

                Dialog search = selectedFileCabinet.GetDialogFromCustomSearchRelation();

                Console.Write("Deleting document...");
                Document documentToDelete = search.GetDialogFromSelfRelation().GetDocumentsResult(new DialogExpression()
                {
                    Start = 0,
                    Count = 1,
                    Condition = new List<DialogExpressionCondition>(new List<DialogExpressionCondition>
            {
                new DialogExpressionCondition
                {
                    Value = new List<string>() { documentId },
                    DBName = "DWDOCID"
                }
            }),
                    Operation = DialogExpressionOperation.And
                }).Items.FirstOrDefault();

                if (documentToDelete != null)
                {
                    documentToDelete.DeleteSelfRelation();
                    Console.Write("\r" + new string(' ', Console.WindowWidth - 1) + $"\rDocument with ID {documentId} deleted successfully.\n");
                }
                else
                {
                    Console.Write("\r" + new string(' ', Console.WindowWidth - 1) + "\rDocument with ID " + documentId + " not found or already deleted.\n");
                }
            }
            catch (Exception ex)
            {
                Console.Write("\r" + new string(' ', Console.WindowWidth - 1) + "\rAn error occurred while deleting the document: " + ex.Message + "\n");
            }
        }

        #endregion

    }
}