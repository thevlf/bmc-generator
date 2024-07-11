using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.SemanticKernel.Connectors.OpenAI;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Application = Microsoft.Office.Interop.Visio.Application;
using Page = Microsoft.Office.Interop.Visio.Page;
using Shape = Microsoft.Office.Interop.Visio.Shape;


namespace BMCGen
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {        
        // As Environment Variables
        static string? url = Environment.GetEnvironmentVariable("azureAIUrl");
        static string? key = Environment.GetEnvironmentVariable("azureAIKey");
        static string? aiModel = "gpt-4";

        // As AppSettings
        //static string? url = ConfigurationManager.AppSettings["azureAIUrl"];
        //static string? key = ConfigurationManager.AppSettings["azureAIKey"];
        //static string? aiModel = ConfigurationManager.AppSettings["azureAIModel"];

        static string visiosFolder = @"C:\Users\osmoral\Microsoft\Digital Sales Enterprise - Account Business Canvas";
        static string? customersFile;
        static string? metricsFile;
        static string? customerName;
        static double customerTPID;
        static string? nonEnglishName;
        static string? domain;
        static string rootSaveFolder = "On-Off Run";
        static string? pod;
        static BMC? bmCanvas;
        static string? azureRecommendations;
        static string? dynamicsRecommendations;
        static string? modernWorkRecommendations;
        static string? summary;
        static string? metrics;
        static DataSet? dsMetrics;
        static DataSet? dsAccounts;
        static int iSleep = 10;
        static bool incompleteBMC = false;
        static bool updateDataOnly = false;
        static bool createExportOnly = false;
        Microsoft.Office.Interop.Excel.Application? xlApp;

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Logging.LogCleanup();      
                
                Logging.WriteLog("Starting new run!");
                DateTime dtStart = DateTime.Now;

                await Semantisize();

                Logging.WriteLog("Run completed!");
                DateTime dtEnd = DateTime.Now;
                double dbTotalMin = dtEnd.Subtract(dtStart).TotalMinutes;
                Console.WriteLine("");
                Logging.WriteLog("Total Run Time: " + dbTotalMin.ToString("0.#") + " minutes");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Failure processing {customerName} ({customerTPID}): {ex.Message}");
                Logging.WriteLog($"Failure processing {customerName} ({customerTPID}): {ex.Message}");

                MessageBox.Show(ex.Message);
                Environment.Exit(-1);
            }
        }

        private async Task Semantisize()
        {
            btnProcess.IsEnabled = false;
            updateDataOnly = Convert.ToBoolean(ckUpdateOnly.IsChecked);
            createExportOnly = Convert.ToBoolean(ckGenerateFile.IsChecked);

            if (createExportOnly)            
            {
                GenerateExportFile();

                btnProcess.IsEnabled = true;
                return;
            }                                                  

            if (updateDataOnly)
            {
                if (string.IsNullOrEmpty(metricsFile)) { MessageBox.Show("Please specify metrics file!"); Environment.Exit(0); }
                dsMetrics = PopulateDataset($"{metricsFile}", out string error);
                ProcessMetrics();

                btnProcess.IsEnabled = true;
                return;
            }

            if (customersFile != null)
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook? xlWorkbook = xlApp.Workbooks.Open(customersFile);
                Worksheet? xlWorksheet = xlWorkbook.Worksheets[1];
                int rowCount = xlWorksheet.UsedRange.Rows.Count;

                Trace.WriteLine($"{rowCount} total rows in Spreadsheet");
                Logging.WriteLog($"{rowCount} total rows in Spreadsheet");

                for (int i = 2; i < rowCount + 1; i++)
                {
                    try
                    {
                        double tpid = (xlWorksheet.Cells[i, 1]).Value;
                        string accountName = (xlWorksheet.Cells[i, 2]).Value;
                        string salesUnit = (xlWorksheet.Cells[i, 3]).Value;
                        string country = (xlWorksheet.Cells[i, 4]).Value;
                        string website = (xlWorksheet.Cells[i, 5]).Value;
                        bool process = (xlWorksheet.Cells[i, 6]).Value;
                        bool completed = (xlWorksheet.Cells[i, 7]).Value;

                        if (process && !completed)
                        {
                            Trace.WriteLine($"Row {i} of {rowCount} being processed");
                            Logging.WriteLog($"Row {i} of {rowCount} being processed");

                            customerName = accountName;
                            customerTPID = tpid;
                            rootSaveFolder = country;
                            pod = salesUnit;

                            if (string.IsNullOrEmpty(website))
                            {
                                await GetDomain(customerName);
                                (xlWorksheet.Cells[i, 5]).Value = domain;
                            }
                            else { domain = website; }                            

                            Logging.WriteLog($"Processing {customerName} ({customerTPID})");
                            DateTime dtStart = DateTime.Now;

                            await GetCanvasAndRecommendations();
                            if (!incompleteBMC)
                            {
                                await GetCustomerSummary();
                                GetMetrics();
                                PopulateVisio();

                                //Set Completed to true
                                xlWorksheet.Cells[i, 7].Value = true;
                                xlWorkbook.Save();
                            }

                            DateTime dtEnd = DateTime.Now;
                            double dbTotalMin = dtEnd.Subtract(dtStart).TotalMinutes;
                            Logging.WriteLog($"Done processing {customerName} ({customerTPID}) in {dbTotalMin.ToString("0.#")} minutes");
                        }
                        else
                        {
                            Trace.WriteLine($"Row {i} not processed due to either value of Process column: {process} and/or Completed column: {completed}");
                            Logging.WriteLog($"Row {i} not processed due to either value of Process column: {process} and/or Completed column: {completed}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine($"Failure processing {customerName} ({customerTPID}): {ex.Message} continuing on to next customer...");
                        Logging.WriteLog($"Failure processing {customerName} ({customerTPID}): {ex.Message} continuing on to next customer...");
                    }
                }

                //Closing Excel
                xlWorksheet = null;
                xlWorkbook.Close();
                xlWorkbook = null;
                xlApp.Quit();
                xlApp = null;                
            }
            else
            {
                customerName = txtCustomer.Text.Trim();
                customerTPID = Convert.ToDouble(txtTPID.Text.Trim());
                domain = txtDomain.Text.Trim();

                if (string.IsNullOrEmpty(domain)) { await GetDomain(customerName); }

                Logging.WriteLog($"Processing {customerName}");

                await GetCanvasAndRecommendations();
                if (!incompleteBMC)
                {
                    await GetCustomerSummary();
                    GetMetrics();
                    PopulateVisio();
                }
            }

            btnProcess.IsEnabled = true;
        }

        private void GenerateExportFile()
        {
            visiosFolder = "C:\\Users\\osmoral\\Microsoft\\Digital Sales Enterprise - Account Business Canvas\\United States";

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlWorkbook = xlApp.Workbooks.Add();
            Worksheet xlWorksheet = xlWorkbook.Worksheets.Add();
            xlWorksheet.Name = "BMC Output";
            xlWorksheet.Cells[1, 1].Value = "Number";
            xlWorksheet.Cells[1, 2].Value = "Customer";
            xlWorksheet.Cells[1, 3].Value = "Category";
            xlWorksheet.Cells[1, 4].Value = "Text";
            xlWorksheet.Cells[1, 5].Value = "Region";

            xlWorkbook.SaveAs($"{Directory.GetCurrentDirectory()}\\BMC Output.xlsx");            

            string[] visioFiles = Directory.GetFiles(visiosFolder, $"*.vsdx", SearchOption.AllDirectories);
            Logging.WriteLog($"{visioFiles.Length} files to process...");

            int row = 2;
            int rowIndex = 1;

            for (int i = 0; i < visioFiles.Length; i++)
            {
                string visioFile = visioFiles[i];

                Logging.WriteLog($"Processing {i+1} of {visioFiles.Length}");
                Application vizApp = new Application();
                vizApp.Visible = false;
                vizApp.Documents.Open(visioFile);
                Page page = vizApp.ActivePage;

                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Customer Segments", "CustomerSegments"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Value Propositions", "ValuePropositions"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Channels", "Channels"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Customer Relationships", "CustomerRelationships"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Key Resources", "KeyResources"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Key Activities", "KeyActivities"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Key Partners", "KeyPartners"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Cost Structure", "CostStructure"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Revenue Streams", "RevenueStreams"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Summary", "Summary"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Metrics", "Metrics"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Azure", "Azure"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Dynamics", "Dynamics"); rowIndex++; row++;
                AddRowToWorkSheet(xlWorksheet, row, rowIndex, visioFile, page, "Modern Work", "ModernWork"); rowIndex++; row++;                

                vizApp.Quit();
                xlWorkbook.Save();
            }            

            xlWorkbook.Close();            
            xlApp.Quit();
        }

        private void AddRowToWorkSheet(Worksheet xlWorksheet, int row, int rowIndex, string visioFile, Page page, string categoryDisplay, string categoryName)
        {
            xlWorksheet.Cells[row, 1].Value = rowIndex;
            xlWorksheet.Cells[row, 2].Value = GetShapeValue(page, "Customer");
            xlWorksheet.Cells[row, 5].Value = visioFile.Split("\\")[6].Replace("USA - ", "");
            xlWorksheet.Cells[row, 3].Value = categoryDisplay;
            xlWorksheet.Cells[row, 4].Value = GetShapeValue(page, categoryName);
        }

        private void ProcessMetrics()
        {            
            customerName = txtCustomer.Text.Trim();
            double.TryParse(txtTPID.Text.Trim(), out customerTPID);

            if (customerTPID != 0)
            {
                PopulateMetrics();
                UpdateVisio();
            }
            else
            {
                Logging.WriteLog($"{dsMetrics.Tables[0].Rows.Count} rows to be evaluated");
                foreach (DataRow row in dsMetrics.Tables[0].Rows)
                {
                    customerTPID = Convert.ToDouble(row["TPID"]);
                    customerName = row["TopParent"].ToString();
                    PopulateMetrics();
                    UpdateVisio();
                }
            }
        }

        private void PopulateMetrics()
        {
            DataRow[] customerMetrics = dsMetrics.Tables[0].Select($"TPID = {customerTPID}");
            if (customerMetrics.Count() > 0)
            {
                string acr = customerMetrics[0]["ACR (LCM)"]?.ToString() ?? "not determined";
                string ads = customerMetrics[0]["ADS (LCM)"]?.ToString() ?? "not determined";

                decimal acrValue;
                if (acr != "not determined")
                {
                    decimal.TryParse(acr, out acrValue);
                    acr = acrValue.ToString("C", CultureInfo.GetCultureInfo("en-US"));
                }

                decimal adsValue;
                if (ads != "not determined")
                {
                    decimal.TryParse(ads, out adsValue);
                    ads = adsValue.ToString("C", CultureInfo.GetCultureInfo("en-US"));
                }

                metrics = $"ACR: {acr}\r\nADS: {ads}";
            }
            else
            { Logging.WriteLog($"TPID: {customerTPID} not found in metrics file for {customerName}"); }
        }

        private void UpdateVisio()
        {
            //1541329	BREMER FINANCIAL
            
            string? fileName = customerName;
            fileName = fileName.Replace("/", "_");
            fileName = fileName.Replace("\\", "_");

            string[] customerVisio = Directory.GetFiles(visiosFolder, $"*{fileName}*.vsdx", SearchOption.AllDirectories);

            if (customerVisio.Length > 0)
            {
                string filePath = customerVisio[0];
                Application vizApp = new Application();

                if (txtTPID.Text == "[TPID]") { vizApp.Visible = false; }
                
                vizApp.Documents.Open(filePath);

                Page page = vizApp.ActivePage;
                PopulateShape(page, "Metrics", metrics);

                vizApp.ActiveDocument.SaveAs(filePath);
                vizApp.Quit();

                Logging.WriteLog($"Exisiting Visio found and updated for {customerName} ({customerTPID})");
            }
            else
            {
                Logging.WriteLog($"Exisiting Visio not found for {customerName} ({customerTPID})");
            }
        }

        private async Task TranslateToEnglish(string accountName)
        {
            nonEnglishName = accountName;

            string translateQuery = $"Please translate the following to English: {accountName} and only return the translated text.";

            var builder = Kernel.CreateBuilder();
            builder.Services.AddAzureOpenAIChatCompletion(aiModel, url, key);

            builder.Plugins.AddFromType<DatePlugin>();
            builder.Plugins.AddFromType<BingPlugin>();

            var kernel = builder.Build();

            IChatCompletionService chat = kernel.GetRequiredService<IChatCompletionService>();

            ChatHistory history = new();
            history.AddUserMessage(translateQuery);

            OpenAIPromptExecutionSettings openAIPromptExecutionSettings = new()
            {
                ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions
            };

            Trace.WriteLine($"Querying {aiModel} to translate {accountName}");
            Logging.WriteLog($"Querying {aiModel} to translate {accountName}");

            // Get the response from the AI
            var result = await chat.GetChatMessageContentAsync(
                history,
                executionSettings: openAIPromptExecutionSettings,
                kernel: kernel);

            Trace.WriteLine(result);
            Logging.WriteLog($"Account info returned from query {result}");

            customerName = result.ToString();
        }

        private async Task GetDomain(string customer)
        {
            string domainQuery = $"What is the lastest most up to date web domain associated with {customer}.";

            var builder = Kernel.CreateBuilder();
            builder.Services.AddAzureOpenAIChatCompletion(aiModel, url, key);

            builder.Plugins.AddFromType<DatePlugin>();
            builder.Plugins.AddFromType<BingPlugin>();

            var kernel = builder.Build();

            IChatCompletionService chat = kernel.GetRequiredService<IChatCompletionService>();

            ChatHistory history = new();
            history.AddUserMessage(domainQuery);

            OpenAIPromptExecutionSettings openAIPromptExecutionSettings = new()
            {
                ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions
            };

            Trace.WriteLine($"Querying {aiModel} for Domain info");
            Logging.WriteLog($"Querying {aiModel} for Domain info");

            // Get the response from the AI
            var result = await chat.GetChatMessageContentAsync(
                history,
                executionSettings: openAIPromptExecutionSettings,
                kernel: kernel);

            Trace.WriteLine(result);
            Logging.WriteLog($"Domain info returned from query {result}");

            domain = CleantText(result.ToString(), true);

            Logging.WriteLog($"Cleaned domain info: {domain}");
        }

        private async Task GetCustomerSummary()
        {
            string summaryQuery = txtSummaryQuery.Text.Replace("[CUSTOMER]", customerName);

            var builder = Kernel.CreateBuilder();
            builder.Services.AddAzureOpenAIChatCompletion(aiModel, url, key);

            builder.Plugins.AddFromType<DatePlugin>();
            builder.Plugins.AddFromType<BingPlugin>();

            var kernel = builder.Build();

            IChatCompletionService chat = kernel.GetRequiredService<IChatCompletionService>();

            ChatHistory history = new();
            history.AddUserMessage(summaryQuery);

            OpenAIPromptExecutionSettings openAIPromptExecutionSettings = new()
            {
                ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions
            };

            Trace.WriteLine($"Querying {aiModel} for Summary info");
            Logging.WriteLog($"Querying {aiModel} for Summary info");

            // Get the response from the AI
            var result = await chat.GetChatMessageContentAsync(
                history,
                executionSettings: openAIPromptExecutionSettings,
                kernel: kernel);

            //Trace.WriteLine(result);

            summary = CleantText(result.ToString(), false);
        }

        private void GetMetrics()
        {
            metrics = "No metrics data available";

            if (dsMetrics == null && !string.IsNullOrEmpty(metricsFile))
            {
                dsMetrics = PopulateDataset($"{metricsFile}", out string error);
                DataRow[] customerMetrics = dsMetrics.Tables[0].Select($"TPID = {customerTPID}");
                if (customerMetrics.Count() > 0)
                {
                    string acr = customerMetrics[0]["ACR (LCM)"]?.ToString() ?? "not determined";
                    string ads = customerMetrics[0]["ADS (LCM)"]?.ToString() ?? "not determined";

                    decimal acrValue;
                    if (acr != "not determined")
                    {
                        decimal.TryParse(acr, out acrValue);
                        acr = acrValue.ToString("C", CultureInfo.GetCultureInfo("en-US"));
                    }

                    decimal adsValue;
                    if (ads != "not determined")
                    {
                        decimal.TryParse(ads, out adsValue);
                        ads = adsValue.ToString("C", CultureInfo.GetCultureInfo("en-US"));
                    }

                    metrics = $"ACR: {acr}\r\nADS: {ads}";
                }
            }
        }

        private string? CleantText(string text, bool isDomain)
        {
            text = text.Replace("### ", "");
            text = text.Replace("**", "");

            if (isDomain)
            {
                Regex regex = new Regex("[a-zA-Z0-9]{2,}(\\.[a-zA-Z0-9]{2,})(\\.[a-zA-Z0-9]{2,})?");
                Match match = regex.Match(text);
                if (match.Success)
                {
                    text = match.Value;
                    text = text.Replace("https://www.", "");
                    text = text.Replace("http://www.", "");
                    text = text.Replace("www.", "");
                    text = text.Replace("/", "");
                }
                else
                { text = "bing.com"; }

            }

            return text;
        }

        private async Task GetCanvasAndRecommendations()
        {
            //Trace.WriteLine($"Sleeping {iSleep} seconds before querying for BMC Canvas details");
            //Logging.WriteLog($"Sleeping {iSleep} seconds before querying for BMC Canvas details");
            //Thread.Sleep(iSleep * 1000);

            string bmcPrompt = txtBMCQuery.Text.Replace("[CUSTOMER]", customerName) + "Format response as JSON object.";
            string azurePrompt = txtAzureQuery.Text.Replace("[CUSTOMER]", customerName) + "Format response as text.";
            string dynamicsPrompt = txtDynamicsQuery.Text.Replace("[CUSTOMER]", customerName) + "Format response as text.";
            string modernWorkPrompt = txtModernWorkQuery.Text.Replace("[CUSTOMER]", customerName) + "Format response as text.";

            var builder = Kernel.CreateBuilder();
            builder.Services.AddAzureOpenAIChatCompletion(aiModel, url, key);
            builder.Plugins.AddFromType<DatePlugin>();
            builder.Plugins.AddFromType<BingPlugin>();

            var kernel = builder.Build();

            IChatCompletionService chat = kernel.GetRequiredService<IChatCompletionService>();

            ChatHistory history = new();
            history.AddSystemMessage(txtSystem.Text);
            history.AddUserMessage(bmcPrompt);

            // Enable auto function calling
#pragma warning disable SKEXP0010 // Type is for evaluation purposes only and is subject to change or removal in future updates. Suppress this diagnostic to proceed.
            OpenAIPromptExecutionSettings openAIPromptExecutionSettings = new()
            {
                ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions,
                ResponseFormat = "json_object"
            };
#pragma warning restore SKEXP0010 // Type is for evaluation purposes only and is subject to change or removal in future updates. Suppress this diagnostic to proceed.            

            Trace.WriteLine($"Querying {aiModel} for Business Model Canvas details");
            Logging.WriteLog($"Querying {aiModel} for Business Model Canvas details");

            // Get the response from the AI
            var result = await chat.GetChatMessageContentAsync(
                history,
                executionSettings: openAIPromptExecutionSettings,
                kernel: kernel);

            //Add the message from the agent to the chat history
            history.AddMessage(result.Role, result.Content ?? string.Empty);

            //Trace.WriteLine(result);

            string? responseString = result.ToString();
            ConvertResponseToObject(responseString);

            if (incompleteBMC)
            {
                Trace.WriteLine($"Business Model incomplete - this typically means that the LLM did not return properly formatted json, this will have to be re-run");
                Logging.WriteLog($"Business Model incomplete - this typically means that the LLM did not return properly formatted json, this will have to be re-run");
                Logging.WriteLog($"Response returned: {responseString}");
                return;
            }

            // Azure Recommendations //
            history.AddUserMessage(azurePrompt);

            Trace.WriteLine($"Querying {aiModel} for Azure Recommentations");
            Logging.WriteLog($"Querying {aiModel} for Azure Recommendations");
            //Thread.Sleep(iSleep * 1000);

            openAIPromptExecutionSettings = new()
            {
                ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions
            };

            result = await chat.GetChatMessageContentAsync(
                history,
                executionSettings: openAIPromptExecutionSettings,
                kernel: kernel);

            azureRecommendations = CleantText(result.ToString(), false);

            // Dynamics Recommendations // 
            history.AddUserMessage(dynamicsPrompt);

            Trace.WriteLine($"Querying {aiModel} for Dynamics Recommentations");
            Logging.WriteLog($"Querying {aiModel} for Dynamics Recommendations");
            //Thread.Sleep(iSleep * 1000);

            openAIPromptExecutionSettings = new()
            {
                ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions
            };

            result = await chat.GetChatMessageContentAsync(
                history,
                executionSettings: openAIPromptExecutionSettings,
                kernel: kernel);

            dynamicsRecommendations = CleantText(result.ToString(), false);

            // Modern Work Recommendations //
            history.AddUserMessage(modernWorkPrompt);

            Trace.WriteLine($"Querying {aiModel} for Modern Work Recommentations");
            Logging.WriteLog($"Querying {aiModel} for Modern Work Recommendations");
            //Thread.Sleep(iSleep * 1000);

            openAIPromptExecutionSettings = new()
            {
                ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions
            };

            result = await chat.GetChatMessageContentAsync(
                history,
                executionSettings: openAIPromptExecutionSettings,
                kernel: kernel);

            modernWorkRecommendations = CleantText(result.ToString(), false);
        }

        private void ConvertResponseToObject(string responseString)
        {
            string newJson = string.Empty;

            try
            {
                if (responseString.Contains("\": \""))
                {
                    newJson = responseString.Replace("\": \"", ": ");
                    newJson = newJson.Replace("{", "");
                    newJson = newJson.Replace("}", "");
                    newJson = "{" + newJson + "}";
                }

                if (string.IsNullOrEmpty(newJson))
                {
                    Regex regex = new Regex("\\\"(.*?)\\\":");
                    Match match = regex.Match(responseString);
                    if (match.Success)
                    {
                        if (match.Value.ToLower().Contains("canvas"))
                        {
                            newJson = responseString.Replace(match.Value, "");
                            newJson = newJson.Replace("{", "");
                            newJson = newJson.Replace("}", "");
                            newJson = "{" + newJson + "}";
                        }
                    }
                }
                else
                {
                    Regex regex = new Regex("\\\"(.*?)\\\":");
                    Match match = regex.Match(newJson);
                    if (match.Success)
                    {
                        if (match.Value.ToLower().Contains("canvas"))
                        {
                            newJson = newJson.Replace(match.Value, "");
                        }
                    }
                }

                if (string.IsNullOrEmpty(newJson))
                {
                    bmCanvas = JsonConvert.DeserializeObject<BMC>(responseString) ?? throw new InvalidOperationException();
                }
                else
                {
                    bmCanvas = JsonConvert.DeserializeObject<BMC>(newJson) ?? throw new InvalidOperationException();
                }

                if (CanvasIncomplete())
                {
                    Trace.WriteLine($"Reformatting attempt failed");
                    Logging.WriteLog($"Reformatting attempt failed");
                    Logging.WriteLog($"Original format: {responseString}");
                    Logging.WriteLog($"Attempted reformat: {newJson}");
                }
            }
            catch (Exception ex)
            {
                incompleteBMC = true;
                Logging.WriteLog($"{ex.Message}");
                Logging.WriteLog($"{responseString}");
                Logging.WriteLog($"{newJson}");
            }
        }

        private void PopulateVisio()
        {
            Application vizApp = new Application();
            if (!string.IsNullOrEmpty(customersFile)) { vizApp.Visible = false; }

            string docPath = $"{Directory.GetCurrentDirectory()}\\Template.vsdx";
            vizApp.Documents.Open(docPath);

            Page page = vizApp.ActivePage;
            page.Shapes["Customer"].Text = customerName;
            if (!string.IsNullOrEmpty(nonEnglishName)) { page.Shapes["Customer"].Text = $"{nonEnglishName} ({customerName})"; }

            PopulateShape(page, "Summary", summary);
            PopulateShape(page, "Metrics", metrics);
            PopulateShape(page, "Azure", azureRecommendations);
            PopulateShape(page, "Dynamics", dynamicsRecommendations);
            PopulateShape(page, "ModernWork", modernWorkRecommendations);

            if (bmCanvas != null)
            {
                foreach (PropertyInfo prop in bmCanvas.GetType().GetProperties())
                {
                    string propName = prop.Name;
                    var entries = prop.GetValue(bmCanvas) as List<string>;
                    if (entries != null && entries.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (var entry in entries)
                        {
                            sb.AppendLine(entry);
                        }

                        foreach (Shape shape in page.Shapes)
                        {
                            if (shape.Shapes.Count > 0)
                            {
                                foreach (Shape subShape in shape.Shapes)
                                {
                                    if (subShape.Name == propName)
                                    {
                                        subShape.Text = sb.ToString();
                                    }

                                }
                            }
                        }
                    }
                }
            }
            page.ResizeToFitContents();

            string folder = $"C:\\LOCAL\\BMC Output\\{rootSaveFolder}\\{pod}";
            if (!Directory.Exists(folder)) { Directory.CreateDirectory(folder); }

            string fileName = customerName;
            if (!string.IsNullOrEmpty(nonEnglishName)) { fileName = $"{nonEnglishName}({customerName}).vsdx"; }

            fileName = fileName.Replace("/", "_");
            fileName = fileName.Replace("\\", "_");

            vizApp.ActiveDocument.SaveAs($"{folder}\\{fileName}.vsdx");

            if (!string.IsNullOrEmpty(customersFile)) { vizApp.Quit(); }
        }

        private bool CanvasIncomplete()
        {
            bool emptyProperties = false;

            if (bmCanvas != null)
            {
                foreach (PropertyInfo prop in bmCanvas.GetType().GetProperties())
                {
                    string propName = prop.Name;
                    var entries = prop.GetValue(bmCanvas) as List<string>;
                    if (entries == null || entries.Count == 0)
                    {
                        Trace.WriteLine($"{propName} is empty, Visio will not be generated");
                        Logging.WriteLog($"{propName} is empty, Visio will not be generated");
                        emptyProperties = true;
                    }
                }
            }
            incompleteBMC = emptyProperties;
            return emptyProperties;
        }

        private void PopulateShape(Page page, string shapeName, string shapeValue)
        {
            foreach (Shape shape in page.Shapes)
            {
                if (shape.Shapes.Count > 0)
                {
                    foreach (Shape subShape in shape.Shapes)
                    {
                        if (subShape.Name == shapeName)
                        {
                            subShape.Text = shapeValue;
                        }
                    }
                }
            }
        }

        private string GetShapeValue(Page page, string shapeName)
        {
            string value = string.Empty;

            foreach (Shape shape in page.Shapes)
            {
                if (shape.Name == shapeName) { return shape.Text; }

                if (shape.Shapes.Count > 0)
                {
                    foreach (Shape subShape in shape.Shapes)
                    {
                        if (subShape.Name == shapeName)
                        {
                            value = subShape.Text;
                        }
                    }
                }
            }

            return value;
        }

        private void Browse_Customers(object sender, RoutedEventArgs e)
        {
            OpenFileDialog oFile = new OpenFileDialog();
            oFile.DefaultExt = ".xlsx";
            oFile.Filter = "Customers (*.xls)|*.xlsx";

            Nullable<bool> result = oFile.ShowDialog();

            if (result == true)
            {
                txtCustomersFile.Text = oFile.FileName;
                customersFile = oFile.FileName;
                txtCustomer.IsEnabled = false;
                txtDomain.IsEnabled = false;
            }

        }

        public static DataSet PopulateDataset(string filePath, out string error)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                error = string.Empty;
                DataSet? dataset;

                if (filePath.ToLower().Contains("csv"))
                {
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                        {
                            dataset = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                UseColumnDataType = true,
                                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                                {
                                    EmptyColumnNamePrefix = "Column",
                                    UseHeaderRow = true,
                                    FilterRow = (rowReader) =>
                                    {
                                        return true;
                                    },
                                    FilterColumn = (rowReader, columnIndex) =>
                                    {
                                        return true;
                                    }
                                }
                            });
                        }
                    }
                }
                else
                {
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            dataset = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                UseColumnDataType = true,
                                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                                {
                                    EmptyColumnNamePrefix = "Column",
                                    UseHeaderRow = true,
                                    FilterRow = (rowReader) =>
                                    {
                                        return true;
                                    },
                                    FilterColumn = (rowReader, columnIndex) =>
                                    {
                                        return true;
                                    }
                                }
                            });
                        }
                    }
                }
                return dataset;
            }
            catch (Exception ex)
            {
                error = "Error in PopulateDataset():\r\n" + ex.Message;
                return null;
            }
        }

        public static bool ContainsChinese(string input)
        {
            // This pattern covers the Basic Chinese range and can be extended to include other ranges
            string pattern = @"[\u4E00-\u9FFF\u3400-\u4DBF\u20000-\u2A6DF\u2A700-\u2B73F\u2B740-\u2B81F\u2B820-\u2CEAF\u2CEB0-\u2EBEF\u30000-\u3134F]+";
            return Regex.IsMatch(input, pattern);
        }

        public static bool ContainsJapanese(string input)
        {
            // This pattern covers Hiragana, Katakana, and common Kanji.
            // You might need to adjust it based on your specific needs.
            string pattern = @"[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]+";
            return Regex.IsMatch(input, pattern);
        }

        private void Browse_Metrics(object sender, RoutedEventArgs e)
        {
            OpenFileDialog oFile = new OpenFileDialog();
            oFile.DefaultExt = ".xlsx";
            oFile.Filter = "Metrics (*.xls)|*.xlsx";

            Nullable<bool> result = oFile.ShowDialog();

            if (result == true)
            {
                txtMetricsFile.Text = oFile.FileName;
                metricsFile = oFile.FileName;
            }
        }
    }
}