using DocuWare.Platform.ServerClient;
using Bel_DocuWare;
using System;
using System.Diagnostics.Metrics;
using System.Threading.Tasks;
using Serilog;

class Program
{
    static ServiceConnection _SVCConnection = null;

    static async Task Main(string[] args)
    {
        Log.Logger = new LoggerConfiguration()
        .WriteTo.File("log.txt", rollingInterval: RollingInterval.Day)
        .CreateLogger();

        try
        {
            // Credentials and connection details
            string username = "Dennis.Kireyeu.admin";
            string password = "BelgioiosoCheese";
            Uri uri = new Uri("https://belgioioso-cheese.docuware.cloud/docuware/platform");

            // Instance of the DocuWare class
            Bel_DocuWare.DocuWare oBELDoc = new Bel_DocuWare.DocuWare();

            // Hard-coded values
            string fileCabinetId = "7a0017a5-524d-44eb-acb2-52f1f554e4cf";
            string documentId = "14";
            string filePathDowload = @"C:\Temp\Download";
            string filePathUpload = @"C:\Temp\Upload\Text.txt";

            // Connect to the DocuWare server
            bool isConnected = await oBELDoc.ConnectAsync(uri, username, password).ConfigureAwait(false);

            if (isConnected)
            {
               _SVCConnection = oBELDoc.GetServiceConnection();
                var org = _SVCConnection.Organizations[0];

                var fileCabinets = org.GetFileCabinetsFromFilecabinetsRelation().FileCabinet;

                foreach (var fc in fileCabinets)
                {
                    Console.WriteLine($"{fc.Name} (ID: {fc.Id})");
                }

                Console.WriteLine("Select an operation:");
                Console.WriteLine("1. Download and Display Document Content");
                Console.WriteLine("2. Upload a File to the File Cabinet");
                Console.WriteLine("3. Delete a Document by ID");
                Console.WriteLine("4. List Files");

                int choice;
                if (int.TryParse(Console.ReadLine(), out choice))
                {
                    switch (choice)
                    {
                        case 1:
                            await Bel_DocuWare.DocuWare.DownloadAndDisplayDocumentContentAsync(_SVCConnection, fileCabinetId, documentId, filePathDowload);
                            break;
                        case 2:
                            await Bel_DocuWare.DocuWare.UploadSingleFileToFileCabinetAsync(_SVCConnection, fileCabinetId, filePathUpload);
                            break;
                        case 3:
                            Bel_DocuWare.DocuWare.DeleteDocumentById(_SVCConnection, fileCabinetId, documentId);
                            break;
                        case 4:
                            await Bel_DocuWare.DocuWare.ListAllDocumentsAsync(_SVCConnection, fileCabinetId);
                            break;
                        default:
                            Console.WriteLine("Invalid choice. Please select a valid operation.");
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("Invalid input. Please enter a valid number.");
                }

                // Disconnect from the DocuWare server
                oBELDoc.Disconnect();

                Console.WriteLine("Operation completed. Press any key to exit.");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("Failed to connect to DocuWare.");
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "An error occurred: {ErrorMessage}", ex.Message);
            Console.WriteLine("An error occurred: " + ex.Message);
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }

}

