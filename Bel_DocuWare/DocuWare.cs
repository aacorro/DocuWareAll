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
        static ServiceConnection _SVCConnection = null;


        #region CONNECTION

        /// <summary>
        /// GetServiceConnection
        ///  - Provides access to the _SVCConnection field.
        /// </summary>
        /// <returns>
        /// An instance of the `ServiceConnection` class representing the connection to the DocuWare service.
        /// </returns>
        public ServiceConnection GetServiceConnection()
        {
            return _SVCConnection;
        }

        // CONNECT

        /// <summary>
        /// ConnectAsync
        ///  - Establish a connection to a DocuWare service using the provided credentials.
        /// </summary>
        /// <param name="uri"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns>
        /// A boolean value indicating whether the disconnection was successful(true) or not (false)</returns>
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


        //DISCONNECT

        /// <summary>
        /// Disconnect
        ///  - Disconnect from the DocuWare service.
        /// </summary>
        /// <returns>
        /// A boolean value indicating whether the disconnection was successful (true) or not (false)</returns>
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


        #region DOWNLOAD

        /// <summary>
        /// DownloadAndDisplayDocumentContentAsync
        ///  - Downnload and display document information
        /// </summary>
        /// <param name="_SVCConnection">Service connection to DocuWare.</param>
        /// <param name="fileCabinetId">ID of the file cabinet to search for documents.</param>
        /// <param name="documentId">ID of the document to download and display.</param>
        /// <param name="filePath">The directory path where the downloaded document will be saved.</param>
        /// <returns>
        /// No direct return value (void method).
        /// </returns>
        public static async Task<bool> DownloadAndDisplayDocumentContentAsync(ServiceConnection _SVCConnection, string fileCabinetId, string documentId, string filePath)
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
                        await DownloadAndSaveDocument(selectedDocument, filePath);
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// DownloadAndSaveDocument
        ///  - Download and save the content of a document to a specified file path.
        /// </summary>
        /// <param name="document">The Document to be downloaded.</param>
        /// <param name="downloadLocation">The directory path where the document will be saved.</param>
        /// <returns>
        /// An instance of the FileDownloadResult class.
        /// </returns>
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



        // Download File Result

        /// <summary>
        /// FileDownloadResult
        /// -A class to represent the result of a document download operation.
        /// </summary>
        /// <returns>
        /// An instance of the FileDownloadResult class.
        /// </returns>
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

        //Downnload Content

        /// <summary>
        /// DownloadDocumentContentAsync
        ///  - Download the content of a document as a stream.
        /// </summary>
        /// <param name="document">The Document to be downloaded.</param>
        /// <returns>
        /// A FileDownloadResult containing the downloaded content or an error message if the download fails.
        /// </returns>
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
            catch (Exception)
            {
                return null;
            }
        }


        #endregion


        #region UPLOAD

        /// <summary>
        /// UploadSingleFileToFileCabinetAsync
        /// - Upload a single file to a specified file cabinet in DocuWare.
        /// </summary>
        /// <param name="_SVCConnection">Service connection to DocuWare.</param>
        /// <param name="fileCabinetId">ID of the file cabinet to upload the document to.</param>
        /// <param name="filePathUpload">The file path of the document to be uploaded.</param>
        /// <returns>
        /// A Document object representing the uploaded document or null if the upload fails.
        /// </returns>

        public static async Task<bool> UploadSingleFileToFileCabinetAsync(ServiceConnection _SVCConnection, string fileCabinetId, string filePathUpload)
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
                    return false;
                }

                if (File.Exists(filePathUpload))
                {
                    var fileInfo = new FileInfo(filePathUpload);

                    // Upload the document to the selected File Cabinet
                    var uploadedDocument = await selectedFileCabinet.UploadDocumentAsync(fileInfo).ConfigureAwait(false);
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }


        #endregion


        #region DELETE

        /// <summary>
        /// DeleteDocumentById
        ///  - Delete a document from a specified file cabinet in DocuWare by its ID.
        /// </summary>
        /// <param name="_SVCConnection">Service connection to DocuWare.</param>
        /// <param name="fileCabinetId">ID of the file cabinet to delete the document from</param>
        /// <param name="documentId">ID of the document to be deleted.</param>
        /// <returns>
        /// No direct return value (void method).
        /// </returns>

        public static bool DeleteDocumentById(ServiceConnection _SVCConnection, string fileCabinetId, string documentId)
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
                    return false;
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
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion


        #region GET

        /// <summary>
        /// ListAllDocumentsAsync
        ///  - Retrieve a list of documents from a specified file cabinet in DocuWare.
        /// </summary>
        /// <param name="_SVCConnection">Service connection to DocuWare</param>
        /// <param name="fileCabinetId">ID of the file cabinet to list documents from</param>
        /// <param name="count">The maximum number of documents to retrieve (optional)</param>
        /// <returns>
        /// A list of Document objects representing the retrieved documents or null if an error occurs.
        /// </returns>
        public static async Task<List<Document>> ListAllDocumentsAsync(ServiceConnection _SVCConnection, string fileCabinetId, int? count = 100)
        {
            try
            {
                DocumentsQueryResult queryResult = await _SVCConnection.GetFromDocumentsForDocumentsQueryResultAsync(
                    fileCabinetId,
                    count: count)
                    .ConfigureAwait(false);

                List<Document> result = new List<Document>();
                await GetAllDocumentsAsync(queryResult, result);

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

        /// <summary>
        /// GetAllDocumentsAsync
        /// - Helper method for retrieving all documents recursively.
        /// </summary>
        /// <param name="queryResult">The DocumentsQueryResult containing document information.</param>
        /// <param name="documents">List to store the retrieved documents.</param>
        /// <returns>
        /// No direct return value (void method).
        /// </returns>
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

    }
}