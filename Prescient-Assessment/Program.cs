using HtmlAgilityPack;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Net.Http;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Prescient_Assessment
{
    internal class Program
    {
        static Task Main(string[] args)
        {
            string currentDirectory = Directory.GetCurrentDirectory();
            string relativePath = Path.Combine(currentDirectory, @"..\..\Downloads");
            string fullPath = Path.GetFullPath(relativePath);
            string url = "https://clientportal.jse.co.za/downloadable-files?RequestNode=/YieldX/Derivatives/Docs_DMTM";

            //SetDirectoryPermissions(fullPath);
            DownloadFilesAsync(fullPath, url).GetAwaiter().GetResult();
            ProcessFilesAsync(fullPath);
            return Task.CompletedTask;
        }

        static async Task DownloadFilesAsync(string fullPath, string url)
        {
            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(url);
                try
                {
                    // Get the HTML page content
                    HttpResponseMessage response = await client.GetAsync(url);
                    response.EnsureSuccessStatusCode();
                    string htmlContent = await response.Content.ReadAsStringAsync();

                    // Load the HTML document
                    var htmlDoc = new HtmlDocument();
                    htmlDoc.LoadHtml(htmlContent);

                    // Parse the document to find links to XLS and PDF files
                    var links = htmlDoc.DocumentNode.SelectNodes("//a[contains(@href, '2024') and contains(@href, '.xls')]");

                    if (links != null)
                    {
                        foreach (var link in links)
                        {
                            string fileUrl = link.GetAttributeValue("href", string.Empty);
                            string fileName = Path.GetFileName(fileUrl);

                            if (!string.IsNullOrEmpty(fileUrl) && !string.IsNullOrEmpty(fileName))
                            {
                                string filePath = Path.Combine(fullPath, fileName);

                                if (!File.Exists(filePath))
                                {
                                    await DownloadFile(client, fileUrl, filePath);
                                    Console.WriteLine($"Downloaded {fileName} to {filePath}");
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("No XLS or PDF for 2024 files found.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            }
        }

        static async Task DownloadFile(HttpClient client, string fileURL, string filePath)
        {
            HttpResponseMessage response = await client.GetAsync(fileURL);
            response.EnsureSuccessStatusCode();

            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                await response.Content.CopyToAsync(fs);
            }
        }

        static void ProcessFilesAsync(string fullPath)
        {
            string[] files = Directory.GetFiles(fullPath);

            foreach (string file in files)
            {
                Excel.Application app = new Excel.Application();
                Excel.Workbook workbook = app.Workbooks.Open(file);
                Excel.Worksheet sheet = workbook.Sheets[1];
                Excel.Range range = sheet.UsedRange;

                int numRows = range.Rows.Count;
                string connectionString = "";

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    //conn.Open();

                    for (int i = 6; i < numRows; i++) 
                    {
                        string contract = ((Excel.Range)sheet.Cells[i,1]).Value?.ToString();
                        string expiryDate = ((Excel.Range)sheet.Cells[i, 3]).Value?.ToString();
                        string classification = ((Excel.Range)sheet.Cells[i, 4]).Value?.ToString();
                        string strike = ((Excel.Range)sheet.Cells[i, 5]).Value?.ToString();
                        string callPut = ((Excel.Range)sheet.Cells[i, 6]).Value?.ToString();
                        string yield = ((Excel.Range)sheet.Cells[i, 7]).Value.ToString();
                        string price = ((Excel.Range)sheet.Cells[i, 8]).Value.ToString();
                        string spotRate = ((Excel.Range)sheet.Cells[i, 9]).Value.ToString();
                        string previousMTM = ((Excel.Range)sheet.Cells[i, 10]).Value.ToString();
                        string previousPrice = ((Excel.Range)sheet.Cells[i, 11]).Value.ToString();
                        string premiumOption = ((Excel.Range)sheet.Cells[i, 12]).Value.ToString();
                        string volatility = ((Excel.Range)sheet.Cells[i, 13]).Value.ToString();
                        string delta = ((Excel.Range)sheet.Cells[i, 14]).Value.ToString();
                        string deltaValue = ((Excel.Range)sheet.Cells[i, 15]).Value.ToString();
                        string contractsTraded = ((Excel.Range)sheet.Cells[i, 16]).Value.ToString();
                        string openInterest = ((Excel.Range)sheet.Cells[i, 17]).Value.ToString();

                        //using (SqlCommand cmd = new SqlCommand(, conn))
                        //{
                        //    cmd.Parameters.AddWithValue("", );
                        //}
                    }
                }
            }
        }
    }
}
