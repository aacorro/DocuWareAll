//using DocuWare.Platform.ServerClient;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Threading.Tasks;

//public class DocuWareManager
//{
//    private ServiceConnection _svcConnection;

//    public DocuWareManager(ServiceConnection svcConnection)
//    {
//        _svcConnection = svcConnection;
//    }

//    public async Task DownloadAndDisplayDocumentContentAsync(string fileCabinetId, string documentId, string filePath)
//    {
//        await DocuWare.DownloadAndDisplayDocumentContentAsync(_svcConnection, fileCabinetId, documentId, filePath);
//    }

//    private async Task UploadDocumentToCabinet(string fileCabinetId, string sender, string fileName, string filePath, List<FileCabinet> fileCabinets)
//    {
//        try
//        {
//            FileCabinet selectedFileCabinet = fileCabinets.Find(cabinet => cabinet.Id == fileCabinetId);

//            if (selectedFileCabinet != null)
//            {
//                Console.WriteLine($"Uploading to File Cabinet: {selectedFileCabinet.Name}");

//                // Create the index data
//                var indexData = new Document()
//                {
//                    Fields = new List<DocumentIndexField>()
//                    {
//                        DocumentIndexField.Create("SENDER", sender),
//                        DocumentIndexField.Create("FILE_NAME", fileName),
//                    }
//                };

//                if (File.Exists(filePath))
//                {
//                    var fileInfo = new FileInfo(filePath);

//                    // Upload the document to the selected File Cabinet
//                    Console.Write("Uploading document...");
//                    var uploadedDocument = await selectedFileCabinet.UploadDocumentAsync(indexData, fileInfo).ConfigureAwait(false);
//                    Console.WriteLine("\r" + new string(' ', Console.WindowWidth - 1) + "\rDocument uploaded!");
//                }
//                else
//                {
//                    Console.WriteLine($"File not found at the specified path: {filePath}");
//                }
//            }
//            else
//            {
//                Console.WriteLine("File Cabinet not found.");
//            }
//        }
//        catch (Exception ex)
//        {
//            Console.WriteLine("An error occurred while uploading the document: " + ex.Message);
//        }
//    }

//    public async Task ListAllDocumentsAsync(string fileCabinetId)
//    {
//        try
//        {
//            var fileCabinet = _svcConnection.GetFileCabinet(fileCabinetId);
//            var documents = await GetAllDocumentsAsync(fileCabinet);

//            Console.WriteLine("Documents in File Cabinet:");
//            foreach (var document in documents)
//            {
//                Console.WriteLine($"Document ID: {document.Id}, Created At: {document.CreatedAt}");
//            }
//        }
//        catch (Exception ex)
//        {
//            Console.WriteLine("An error occurred while listing documents: " + ex.Message);
//        }
//    }

//    private async Task<List<Document>> GetAllDocumentsAsync(FileCabinet fileCabinet)
//    {
//        List<Document> allDocuments = new List<Document>();
//        int start = 0;
//        int count = 100;

//        while (true)
//        {
//            var queryResult = await fileCabinet.GetDocumentsAsync(start, count).ConfigureAwait(false);
//            var documents = queryResult.Items;

//            if (documents.Count == 0)
//            {
//                break;
//            }

//            allDocuments.AddRange(documents);
//            start += documents.Count;
//        }

//        return allDocuments;
//    }

//    public void DeleteDocumentById(string fileCabinetId, string documentId)
//    {
//        var fileCabinet = _svcConnection.GetFileCabinet(fileCabinetId);
//        DocuWare.DeleteDocumentById(fileCabinet, documentId);
//    }
//}
