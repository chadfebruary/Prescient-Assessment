using HtmlAgilityPack;
using System;
using System.Data.SqlClient;
using System.Globalization;
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

                    // Load the document
                    var htmlDoc = new HtmlDocument();
                    htmlDoc.LoadHtml(htmlContent);

                    // Parse the document to find links to XLS files
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

                string fileName = Path.GetFileNameWithoutExtension(file);
                DateTime fileDate = DateTime.ParseExact(fileName.Substring(0,8), "yyyyMMdd", null);

                int numRows = range.Rows.Count;
                string connectionString = "Server=;Database=;User Id=;Password=;";

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    for (int i = 6; i < numRows; i++) 
                    {
                        string contract = ((Excel.Range)sheet.Cells[i,1]).Value?.ToString();
                        string expiryDate = ((Excel.Range)sheet.Cells[i, 3]).Value?.ToString();
                        string classification = ((Excel.Range)sheet.Cells[i, 4]).Value?.ToString();
                        decimal strike = decimal.Parse(((Excel.Range)sheet.Cells[i, 5]).Value?.ToString());
                        string callPut = ((Excel.Range)sheet.Cells[i, 6]).Value?.ToString();
                        decimal yield = decimal.Parse(((Excel.Range)sheet.Cells[i, 7]).Value.ToString());
                        decimal price = decimal.Parse(((Excel.Range)sheet.Cells[i, 8]).Value.ToString());
                        decimal spotRate = decimal.Parse(((Excel.Range)sheet.Cells[i, 9]).Value.ToString());
                        decimal previousMTM = decimal.Parse(((Excel.Range)sheet.Cells[i, 10]).Value.ToString());
                        decimal previousPrice = decimal.Parse(((Excel.Range)sheet.Cells[i, 11]).Value.ToString());
                        decimal premiumOption = decimal.Parse(((Excel.Range)sheet.Cells[i, 12]).Value.ToString());
                        decimal volatility = decimal.Parse(((Excel.Range)sheet.Cells[i, 13]).Value?.ToString().Trim(), CultureInfo.InvariantCulture);
                        decimal delta = decimal.Parse(((Excel.Range)sheet.Cells[i, 14]).Value.ToString());
                        decimal deltaValue = decimal.Parse(((Excel.Range)sheet.Cells[i, 15]).Value.ToString());
                        decimal contractsTraded = decimal.Parse(((Excel.Range)sheet.Cells[i, 16]).Value.ToString());
                        decimal openInterest = decimal.Parse(((Excel.Range)sheet.Cells[i, 17]).Value.ToString());

                        string query = "INSERT INTO DailyMTM (FileDate, Contract, ExpiryDate, Classification, Strike, CallPut, MTMYield, MarkPrice, SpotRate, PreviousMTM, PreviousPrice, PremiumOnOption, Volatility, Delta, DeltaValue, ContractsTraded, OpenInterest) " +
                                       "VALUES (@FileDate, @Contract, @ExpiryDate, @Classification, @Strike, @CallPut, @MTMYield, @MarkPrice, @SpotRate, @PreviousMTM, @PreviousPrice, @PremiumOnOption, @Volatility, @Delta, @DeltaValue, @ContractsTraded, @OpenInterest)";

                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@FileDate", fileDate);
                            cmd.Parameters.AddWithValue("@Contract", contract);
                            cmd.Parameters.AddWithValue("@ExpiryDate", expiryDate);
                            cmd.Parameters.AddWithValue("@Classification", classification);
                            cmd.Parameters.AddWithValue("@Strike", strike);
                            cmd.Parameters.AddWithValue("@CallPut", (object)callPut ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@MTMYield", yield);
                            cmd.Parameters.AddWithValue("@MarkPrice", price);
                            cmd.Parameters.AddWithValue("@SpotRate", spotRate);
                            cmd.Parameters.AddWithValue("@PreviousMTM", previousMTM);
                            cmd.Parameters.AddWithValue("@PreviousPrice", previousPrice);
                            cmd.Parameters.AddWithValue("@PremiumOnOption", premiumOption);
                            cmd.Parameters.AddWithValue("@Volatility", volatility);
                            cmd.Parameters.AddWithValue("@Delta", delta);
                            cmd.Parameters.AddWithValue("@DeltaValue", deltaValue);
                            cmd.Parameters.AddWithValue("@ContractsTraded", contractsTraded);
                            cmd.Parameters.AddWithValue("@OpenInterest", openInterest);

                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            }
        }
    }
}
