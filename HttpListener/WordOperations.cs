using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Text;
using System.Threading.Tasks;
using HttpListenerService.Model;
using Microsoft.Extensions.Configuration;
using Microsoft.Office.Interop.Word;

namespace HttpListenerService
{
    public class WordOperations
    {
        public static string localFilePath;
        private readonly IConfiguration _config;
        private static Application wordApp;
        private static Document wordDocument;
        private static WordDocumentInputModel wordDocumentInputModel;
        private static TaskCompletionSource<string> _saveCompletionSource = new TaskCompletionSource<string>();

        public WordOperations(WordDocumentInputModel _wordDocumentInputModel)
        {
            wordDocumentInputModel= _wordDocumentInputModel;
            localFilePath= wordDocumentInputModel.MyDocumentsFolderPath;
            ProcessWordFile();
        }
        public void ProcessWordFile()
        {
            RootFolderCreation();
            if (!string.IsNullOrEmpty(wordDocumentInputModel.MyDocumentsFolderPath))
            {
                DownloadWordFileInLocal();
            }
        }
        public void RootFolderCreation()
        {
            try
            {
                if (!Directory.Exists(localFilePath))
                {
                    Directory.CreateDirectory(localFilePath);
                    Console.WriteLine($"Directory created at: {localFilePath}");
                }
                else
                {
                    Console.WriteLine("Directory already exists.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public void DownloadWordFileInLocal()
        {
            string wordDocLocalPath=Path.Combine(localFilePath, "Temp.Doc");
            if (File.Exists(wordDocLocalPath))
            {
                File.Delete(wordDocLocalPath);
            }
            using (WebClient client = new WebClient())
            {
                Console.WriteLine("Downloading file...");
                client.DownloadFile(wordDocumentInputModel.FileToDownload, wordDocLocalPath);
                Console.WriteLine($"File downloaded successfully at: {wordDocLocalPath}");
            }
            OpenWordAndTrackSave(wordDocLocalPath);
        }

        private void OpenWordAndTrackSave(string filePath)
        {
            wordApp = new Application();
            wordDocument = wordApp.Documents.Open(filePath);

            wordApp.DocumentBeforeSave += WordApp_DocumentBeforeSave;

            wordApp.Visible = true;
        }

        private void WordApp_DocumentBeforeSave(Document doc, ref bool saveAsUI, ref bool cancel)
        {
            Console.WriteLine("User pressed Ctrl+S - Document is being saved!");

            try
            {
                string savedFilePath = doc.FullName;  // Get the saved file path

                doc.Close(SaveChanges: true);  // ✅ Close the document properly
                wordApp.Quit();               // ✅ Quit Word application

                // Release COM objects properly
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);

                wordApp = null;
                wordDocument = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                System.Threading.Tasks.Task.Delay(1000).Wait(); 

                Console.WriteLine($"Word document saved and closed: {savedFilePath}");

                string base64Content = ConvertWordFileToBase64(savedFilePath);

                _saveCompletionSource.TrySetResult(base64Content);  // ✅ Correctly set result
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error closing Word: {ex.Message}");
                _saveCompletionSource.TrySetResult(null);  // Indicate failure
            }
        }
        public System.Threading.Tasks.Task WaitForSaveAsync()
        {
            return _saveCompletionSource.Task;
        }

        private string ConvertWordFileToBase64(string filePath)
        {
            try
            {
                byte[] fileBytes = File.ReadAllBytes(filePath);
                var test= Convert.ToBase64String(fileBytes);

                //Byte[] bytes = Convert.FromBase64String(b64Str);
                //File.WriteAllBytes("C:\\Users\\Admin\\Documents\\WordPlugin\\tets.doc", bytes);
                return test;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting file to Base64: {ex.Message}");
                return null;
            }
        }
        public async Task<string> ProcessWordFileAsync()
        {
            return await _saveCompletionSource.Task; 
        }
    }
}
