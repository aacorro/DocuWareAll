using DocuWare.Platform.ServerClient;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DocuWareExampleApp
{
    class Program
    {
        static ServiceConnection _SVCConnection = null;
        static Bel_DocuWare.DocuWare oBELDoc = new Bel_DocuWare.DocuWare();
        static async Task Main(string[] args)
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.File("app.log")
                .CreateLogger();

            try
            {
                Bel_DocuWare.DocuWare oBELDoc = new Bel_DocuWare.DocuWare();

                #region Credentials
                string username = "Dennis.Kireyeu.admin";
                string password = "BelgioiosoCheese";
                Uri uri = new Uri("https://belgioioso-cheese.docuware.cloud/docuware/platform");
                #endregion

                bool bConActive = await oBELDoc.ConnectAsync(uri, username, password).ConfigureAwait(false);

                if (bConActive)
                {
                    try
                    {
                        _SVCConnection = oBELDoc.GetServiceConnection();

                        if (bConActive)
                        {
                            var org = _SVCConnection.Organizations[0];
                            var fileCabinets = org.GetFileCabinetsFromFilecabinetsRelation().FileCabinet;

                            var defaultBasket = fileCabinets.FirstOrDefault(f => f.IsBasket && f.Default);

                            if (fileCabinets.Any())
                            {
                                Console.WriteLine("\u001b[32mFile cabinets available: \u001b[0m");
                                foreach (var fc in fileCabinets)
                                {
                                    Console.WriteLine($"{fc.Name} (ID: {fc.Id})");
                                }

                                Console.Write("Enter the ID of the file cabinet to perform operations: ");
                                string fileCabinetId = Console.ReadLine();

                                var selectedFileCabinet = fileCabinets.FirstOrDefault(fc => fc.Id == fileCabinetId);

                                if (selectedFileCabinet != null)
                                {
                                    Console.WriteLine($"Selected file cabinet: {selectedFileCabinet.Name}");

                                    do
                                    {
                                        Console.WriteLine("Choose an operation:");
                                        //Console.WriteLine("1. List Documents");
                                        Console.WriteLine("2. Upload a Single Document to File Cabinet");
                                        Console.WriteLine("3. Upload Multiple Documents to Default Basket");
                                        Console.WriteLine("4. Upload a Single Document to Basket");
                                        Console.WriteLine("5. Download a Document");
                                        Console.WriteLine("6. Delete a Document");
                                        Console.WriteLine("7. Exit the program");

                                        Console.Write("Enter the operation number: ");

                                        int operationChoice;

                                        if (int.TryParse(Console.ReadLine(), out operationChoice))
                                        {
                                            if (operationChoice == 7)
                                            {
                                                Console.WriteLine("Exiting the program...");
                                                return;
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("Invalid operation choice.");
                                        }


                                        switch (operationChoice)
                                        {
                                            //case 1:
                                            //    await Bel_DocuWare.DocuWare.ListAllDocumentsAsync(_SVCConnection, fileCabinetId);
                                            //    break;
                                            case 2:
                                                await Bel_DocuWare.DocuWare.UploadSingleFileToFileCabinetAsync(_SVCConnection, fileCabinets);
                                                break;

                                            case 5:
                                                await Bel_DocuWare.DocuWare.DownloadAndDisplayDocumentContentAsync(_SVCConnection, fileCabinetId);
                                                break;
                                            case 6:
                                                await Bel_DocuWare.DocuWare.ListAllDocumentsAsync(_SVCConnection, fileCabinetId);
                                                Console.Write("Enter the ID of the document to delete: ");
                                                string documentIdToDelete = Console.ReadLine();
                                                Bel_DocuWare.DocuWare.DeleteDocumentById(selectedFileCabinet, documentIdToDelete);
                                                break;
                                            default:
                                                Console.WriteLine("Invalid operation choice.");
                                                break;
                                        }

                                        Console.Write("Do you want to choose another operation? (yes/no): ");
                                        string continueChoice = Console.ReadLine();

                                        if (continueChoice.ToLower() == "no")
                                        {
                                            Console.WriteLine("Exiting the program...");
                                            return;
                                        }
                                    } while (true);
                                }
                                else
                                {
                                    Console.WriteLine("File cabinet not found with the provided ID.");
                                }
                            }
                            else
                            {
                                Console.WriteLine("No file cabinets available.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Connection to DocuWare is not established. Exiting...");
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "An error occurred: {ErrorMessage}");
                        oBELDoc.Disconnect();
                        Console.WriteLine("Disconnected from DocuWare.");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "An error occurred: {ErrorMessage}");
            }

            oBELDoc.Disconnect();
            Console.WriteLine("Disconnected from DocuWare.");

            Log.CloseAndFlush();
        }
    }
}
