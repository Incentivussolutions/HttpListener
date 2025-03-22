using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using HttpListenerService.Model;
using Microsoft.Extensions.Configuration;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Task = System.Threading.Tasks.Task;

namespace HttpListenerService
{
    public class HttpListenerClass
    {
        private static HttpListener listener; 
        private static string urlPrefix;
        private static HttpListenerResponse response;
        private static Application wordApp;
        private static Document wordDocument;
        private static bool isSaved = false;
        private static IConfiguration _config;
        public HttpListenerClass(IConfiguration config)
        {
            urlPrefix = "http://localhost:4040/";  
            listener = new HttpListener();
            listener.Prefixes.Add(urlPrefix);
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        public async Task Start()
        {
            try
            {
                listener.Start();
                Console.WriteLine($"Listening for requests at {urlPrefix}");

                while (listener.IsListening)
                {
                    HttpListenerContext context = await listener.GetContextAsync(); 
                    _ = ProcessRequestAsync(context);  
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private async Task ProcessRequestAsync(HttpListenerContext context)
        {
            try
            {
                HttpListenerRequest request = context.Request;
                 response = context.Response;

                string requestBody = await new StreamReader(request.InputStream, request.ContentEncoding).ReadToEndAsync();
                Console.WriteLine($"Request body received: {requestBody}");

                var requestData = JsonConvert.DeserializeObject<InputParamters>(requestBody);

                var placeholders = requestData?.Placeholders;
                var filePath = requestData?.FilePath;

                if (placeholders != null && filePath != null)
                {
                    var input = GetWordFileDetails(filePath);
                    WordOperations wordOperations = new WordOperations(input);
                    string base64FileContent = await wordOperations.ProcessWordFileAsync();

                    response.ContentType = "application/json";
                    byte[] buffer = Encoding.UTF8.GetBytes(base64FileContent);

                    response.ContentLength64 = buffer.Length;
                    await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
                }
                else
                {
                    string errorResponse = "Invalid request. Expected placeholders and FilePath in the request body.";
                    byte[] errorBuffer = Encoding.UTF8.GetBytes(errorResponse);
                    response.ContentLength64 = errorBuffer.Length;
                    await response.OutputStream.WriteAsync(errorBuffer, 0, errorBuffer.Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing request: {ex.Message}");
                string errorResponse = "An error occurred while processing your request.";
                byte[] errorBuffer = Encoding.UTF8.GetBytes(errorResponse);
                response.ContentLength64 = errorBuffer.Length;
                await response.OutputStream.WriteAsync(errorBuffer, 0, errorBuffer.Length);
            }
            finally
            {
                response.OutputStream.Close(); 
            }
        }

        public WordDocumentInputModel GetWordFileDetails(string fileToDownload)
        {
            string myDoc = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            WordDocumentInputModel wordDocumentInputModel = new WordDocumentInputModel();
            wordDocumentInputModel.MyDocumentsFolderPath = Path.Combine(myDoc, _config["AppSettings:MyDocumentsFolderName"]);
            wordDocumentInputModel.AddInsConfigFileName = Path.Combine(wordDocumentInputModel.MyDocumentsFolderPath, _config["AppSettings:AddInsConfigFileName"]);
            wordDocumentInputModel.TempFileName = Path.Combine(wordDocumentInputModel.MyDocumentsFolderPath, _config["AppSettings:TempFileName"]);
            wordDocumentInputModel.FileToDownload = fileToDownload;
            return wordDocumentInputModel;
        }

        public void Stop()
        {
            try
            {
                listener.Stop();
                listener.Close();
                Console.WriteLine("Listener stopped.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error stopping listener: {ex.Message}");
            }
        }
    }
}
