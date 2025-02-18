#region Using
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.SemanticKernel.Connectors.OpenAI;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Application = Microsoft.Office.Interop.Visio.Application;
using Page = Microsoft.Office.Interop.Visio.Page;
using Shape = Microsoft.Office.Interop.Visio.Shape;
#endregion

namespace BMCGen
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        #region Variables
        // Environment Variables
        static string? url = Environment.GetEnvironmentVariable("azureAIUrl");
        static string? key = Environment.GetEnvironmentVariable("azureAIKey");
        static string? aiModel = "gpt-4";

        // AppSettings
        static string? outputFolderPath = ConfigurationManager.AppSettings["outputFolderPath"];
        static string? rootSaveFolderName = ConfigurationManager.AppSettings["rootSaveFolderName"];
        static string? excelExportFileName = ConfigurationManager.AppSettings["excelExportFileName"];

        // Variables
        static string? customersFile;        
        static string? customerName;
        static double customerTPID;
        static string? nonEnglishName;
        static string? domain;        
        static string? pod;
        static BMC? bmCanvas;
        static string? azureRecommendations;
        static string? dynamicsRecommendations;
        static string? modernWorkRecommendations;
        static string? challengesRecommendations;
        static string? summary;        
        static DataSet? dsAccounts;
        static int iSleep = 10;
        static bool incompleteBMC = false;
        static bool generateExcel;
        static bool generateVisio;
        Microsoft.Office.Interop.Excel.Application? xlCustomerApp;
        Workbook? xlCustomerWorkbook;
        Worksheet? xlCustomerWorksheet;
        Microsoft.Office.Interop.Excel.Application? xlExportApp;
        Workbook? xlExportWorkbook;
        Worksheet? xlExportWorksheet;
        static int rowID;
        #endregion

        #region Main Functions
        public MainWindow()
        {
            InitializeComponent();
        }        

        private async Task Semantisize()
        {
            DateTime dtStart = DateTime.Now;

            btnProcess.IsEnabled = false;            

            if (customersFile != null)
            {
                await ProcessCustomers();                
            }
            else
            {
                customerName = txtCustomer.Text.Trim();
                Double.TryParse(txtTPID.Text.Trim(), out customerTPID);                
                domain = txtDomain.Text.Trim();                                

                Logging.WriteLog($"Processing {customerName}");

                await GetCanvasAndRecommendations();
                if (!incompleteBMC)
                {
                    await GetCustomerSummary();
                    if (generateVisio) { PopulateVisio(); }
                    if (generateExcel) { PopulateExcel(customerTPID, customerName, "Sales-Unit", "Country", 0); }
                }
            }

            if (generateExcel)
            { CloseExcel(xlExportWorksheet, xlExportWorkbook, xlExportApp); }
            CloseExcel(xlCustomerWorksheet, xlCustomerWorkbook, xlCustomerApp);

            btnProcess.IsEnabled = true;

            Logging.WriteLog("Run completed!");
            DateTime dtEnd = DateTime.Now;
            double dbTotalMin = dtEnd.Subtract(dtStart).TotalMinutes;
            Console.WriteLine("");
            Logging.WriteLog("Total Run Time: " + dbTotalMin.ToString("0.#") + " minutes");
        }

        private async Task ProcessCustomers()
        {
            xlCustomerApp = new Microsoft.Office.Interop.Excel.Application();
            xlCustomerWorkbook = xlCustomerApp.Workbooks.Open(customersFile);
            xlCustomerWorksheet = xlCustomerWorkbook.Worksheets[1];
            int rowCount = xlCustomerWorksheet.UsedRange.Rows.Count;

            Trace.WriteLine($"{rowCount} total rows in Spreadsheet");
            Logging.WriteLog($"{rowCount} total rows in Spreadsheet");

            for (int i = 2; i < rowCount + 1; i++)
            {
                try
                {
                    double tpid = (xlCustomerWorksheet.Cells[i, 1]).Value;
                    string accountName = (xlCustomerWorksheet.Cells[i, 2]).Value.Trim();
                    string country = (xlCustomerWorksheet.Cells[i, 3]).Value.Trim();
                    string salesUnit = (xlCustomerWorksheet.Cells[i, 4]).Value.Trim();
                    string website = (xlCustomerWorksheet.Cells[i, 5]).Value.Trim();
                    bool.TryParse(xlCustomerWorksheet.Cells[i, 6].Value?.ToString(), out bool process);
                    bool.TryParse(xlCustomerWorksheet.Cells[i, 7].Value?.ToString(), out bool completed);

                    if (process && !completed)
                    {
                        Trace.WriteLine($"Row {i} of {rowCount} being processed");
                        Logging.WriteLog($"Row {i} of {rowCount} being processed");

                        customerName = accountName;
                        customerTPID = tpid;
                        rootSaveFolderName = country;
                        pod = salesUnit;

                        if (string.IsNullOrEmpty(website))
                        {
                            await GetDomain(customerName, country);
                            (xlCustomerWorksheet.Cells[i, 5]).Value = domain;
                        }
                        else { domain = website; }

                        Logging.WriteLog($"Processing {customerName} ({customerTPID})");
                        DateTime dtStart = DateTime.Now;

                        await GetCanvasAndRecommendations();
                        if (!incompleteBMC)
                        {
                            await GetCustomerSummary();

                            if (generateVisio) { PopulateVisio(); }

                            if (generateExcel)
                            {
                                PopulateExcel(customerTPID, customerName, salesUnit, country, rowID);
                                rowID++;
                                xlExportWorkbook?.Save();
                            }

                            //Set Completed to true
                            xlCustomerWorksheet.Cells[i, 7].Value = true;
                            xlCustomerWorkbook.Save();
                        }

                        DateTime dtEnd = DateTime.Now;
                        double dbTotalMin = dtEnd.Subtract(dtStart).TotalMinutes;
                        Logging.WriteLog($"Done processing {customerName} ({customerTPID}) in {dbTotalMin.ToString("0.#")} minutes");
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"Failure processing {customerName} ({customerTPID}): {ex.Message} continuing on to next customer...");
                    Logging.WriteLog($"Failure processing {customerName} ({customerTPID}): {ex.Message} continuing on to next customer...");
                }
            }
        }
        #endregion

        #region LLM Functions
        private async Task GetDomain(string customer, string country)
        {
            string domainQuery = $"What is the lastest most up to date web domain associated with {customer} out of the country of {country}.  Only return the actual web domain.";

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

        private async Task GetCanvasAndRecommendations()
        {
            //Trace.WriteLine($"Sleeping {iSleep} seconds before querying for BMC Canvas details");
            //Logging.WriteLog($"Sleeping {iSleep} seconds before querying for BMC Canvas details");
            //Thread.Sleep(iSleep * 1000);

            string bmcPrompt = txtBMCQuery.Text.Replace("[CUSTOMER]", customerName) + " Format response as JSON object.";
            string azurePrompt = txtAzureQuery.Text.Replace("[CUSTOMER]", customerName) + " Format response as text.";
            string dynamicsPrompt = txtDynamicsQuery.Text.Replace("[CUSTOMER]", customerName) + " Format response as text.";
            string modernWorkPrompt = txtModernWorkQuery.Text.Replace("[CUSTOMER]", customerName) + " Format response as text.";
            string challengesPrompt = txtChallengesQuery.Text.Replace("[CUSTOMER]", customerName) + " Format response as text.";

            IKernelBuilder builder = Kernel.CreateBuilder();
            builder.Services.AddAzureOpenAIChatCompletion(aiModel, url, key);
            builder.Plugins.AddFromType<DatePlugin>();
            builder.Plugins.AddFromType<BingPlugin>();

            Kernel kernel = builder.Build();

            try
            {
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

                openAIPromptExecutionSettings = new()
                {
                    ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions
                };

                result = await chat.GetChatMessageContentAsync(
                    history,
                    executionSettings: openAIPromptExecutionSettings,
                    kernel: kernel);

                modernWorkRecommendations = CleantText(result.ToString(), false);

                // Challenges Recommendations //
                history.AddUserMessage(challengesPrompt);

                Trace.WriteLine($"Querying {aiModel} for Challenges Recommentations");
                Logging.WriteLog($"Querying {aiModel} for Challenges Recommendations");

                openAIPromptExecutionSettings = new()
                {
                    ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions
                };

                result = await chat.GetChatMessageContentAsync(
                    history,
                    executionSettings: openAIPromptExecutionSettings,
                    kernel: kernel);

                challengesRecommendations = CleantText(result.ToString(), false);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Error creating chat service: {ex.Message}");
                Logging.WriteLog($"Error creating chat service: {ex.Message}");
            }
        }
        #endregion

        #region Output Functions
        private void SetupExcel()
        {            
            xlExportApp = new Microsoft.Office.Interop.Excel.Application();
            string exportFile = $"{outputFolderPath}\\{excelExportFileName}";

            if (File.Exists(exportFile))
            {
                xlExportWorkbook = xlExportApp.Workbooks.Open(exportFile);
                xlExportWorksheet = xlExportWorkbook.Worksheets[1];
            }
            else
            {
                xlExportWorkbook = xlExportApp.Workbooks.Add();
                xlExportWorksheet = xlExportWorkbook.Worksheets[1];
                xlExportWorksheet.Name = "BMC Output";
                xlExportWorksheet.Cells[1, 1].Value = "ID";
                xlExportWorksheet.Cells[1, 2].Value = "Customer";
                xlExportWorksheet.Cells[1, 3].Value = "Customer Segments";
                xlExportWorksheet.Cells[1, 4].Value = "Value Propositions";
                xlExportWorksheet.Cells[1, 5].Value = "Channels";
                xlExportWorksheet.Cells[1, 6].Value = "Customer Relationships";
                xlExportWorksheet.Cells[1, 7].Value = "Key Resources";
                xlExportWorksheet.Cells[1, 8].Value = "Key Activities";
                xlExportWorksheet.Cells[1, 9].Value = "Key Partners";
                xlExportWorksheet.Cells[1, 10].Value = "Cost Structure";
                xlExportWorksheet.Cells[1, 11].Value = "Revenue Streams";
                xlExportWorksheet.Cells[1, 12].Value = "Summary";
                xlExportWorksheet.Cells[1, 13].Value = "Azure";
                xlExportWorksheet.Cells[1, 14].Value = "Dynamics";
                xlExportWorksheet.Cells[1, 15].Value = "Modern Work";
                xlExportWorksheet.Cells[1, 16].Value = "Challenges";
                xlExportWorksheet.Cells[1, 17].Value = "Sales Unit";
                xlExportWorksheet.Cells[1, 18].Value = "Country";
                xlExportWorkbook.SaveAs(exportFile);
            }
            rowID = xlExportWorksheet.UsedRange.Rows.Count;
            if (rowID > 1)
            {
                string lastRowID = xlExportWorksheet.Cells[rowID, 1].Value.ToString();
                rowID = int.Parse(lastRowID) + 1;
            }
            xlExportWorkbook?.Save();
        }

        private void PopulateExcel(double tpid, string customer, string salesUnit, string country, int rowID)
        {
            int row = xlExportWorksheet.UsedRange.Rows.Count + 1;

            xlExportWorksheet.Cells[row, 1].Value = tpid;
            xlExportWorksheet.Cells[row, 2].Value = customer;
            if (bmCanvas?.CustomerSegments != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (string item in bmCanvas?.CustomerSegments) { sb.AppendLine(item); }
                xlExportWorksheet.Cells[row, 3].Value = sb.ToString();
            }
            if (bmCanvas?.CustomerSegments != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (string item in bmCanvas?.ValuePropositions) { sb.AppendLine(item); }
                xlExportWorksheet.Cells[row, 4].Value = sb.ToString();
            }
            if (bmCanvas?.CustomerSegments != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (string item in bmCanvas?.Channels) { sb.AppendLine(item); }
                xlExportWorksheet.Cells[row, 5].Value = sb.ToString();
            }
            if (bmCanvas?.CustomerSegments != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (string item in bmCanvas?.CustomerRelationships) { sb.AppendLine(item); }
                xlExportWorksheet.Cells[row, 6].Value = sb.ToString();
            }
            if (bmCanvas?.CustomerSegments != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (string item in bmCanvas?.KeyResources) { sb.AppendLine(item); }
                xlExportWorksheet.Cells[row, 7].Value = sb.ToString();
            }
            if (bmCanvas?.CustomerSegments != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (string item in bmCanvas?.KeyActivities) { sb.AppendLine(item); }
                xlExportWorksheet.Cells[row, 8].Value = sb.ToString();
            }
            if (bmCanvas?.CustomerSegments != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (string item in bmCanvas?.KeyPartners) { sb.AppendLine(item); }
                xlExportWorksheet.Cells[row, 9].Value = sb.ToString();
            }
            if (bmCanvas?.CustomerSegments != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (string item in bmCanvas?.CostStructure) { sb.AppendLine(item); }
                xlExportWorksheet.Cells[row, 10].Value = sb.ToString();
            }
            if (bmCanvas?.CustomerSegments != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (string item in bmCanvas?.RevenueStreams) { sb.AppendLine(item); }
                xlExportWorksheet.Cells[row, 11].Value = sb.ToString();
            }
            xlExportWorksheet.Cells[row, 12].Value = summary;
            xlExportWorksheet.Cells[row, 13].Value = azureRecommendations;
            xlExportWorksheet.Cells[row, 14].Value = dynamicsRecommendations;
            xlExportWorksheet.Cells[row, 15].Value = modernWorkRecommendations;
            xlExportWorksheet.Cells[row, 16].Value = challengesRecommendations;
            xlExportWorksheet.Cells[row, 17].Value = salesUnit;
            xlExportWorksheet.Cells[row, 18].Value = country;

            xlExportWorkbook?.Save();
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
            //PopulateShape(page, "Metrics", metrics);
            PopulateShape(page, "Azure", azureRecommendations);
            PopulateShape(page, "Dynamics", dynamicsRecommendations);
            PopulateShape(page, "ModernWork", modernWorkRecommendations);
            PopulateShape(page, "Challenges", challengesRecommendations);

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

            string folder = $"{outputFolderPath}\\{rootSaveFolderName}\\{pod}";
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

        private void PopulateShape(Page page, string shapeName, string? shapeValue)
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
        #endregion

        #region UI Events
        private async void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            try
            {                
                Logging.LogCleanup();               
                Logging.WriteLog("Starting new run!");                

                generateExcel = Convert.ToBoolean(ckGenerateExcel.IsChecked);
                generateVisio = Convert.ToBoolean(ckGenerateVisio.IsChecked);
                if (generateExcel) { SetupExcel(); }

                await Semantisize();                
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Failure processing {customerName} ({customerTPID}): {ex.Message}");
                Logging.WriteLog($"Failure processing {customerName} ({customerTPID}): {ex.Message}");

                MessageBox.Show(ex.Message);
                Environment.Exit(-1);
            }
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
                txtTPID.IsEnabled = false;
            }

        }
        #endregion

        #region Utility Functions

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

        private string? CleantText(string text, bool isDomain)
        {
            text = text.Replace("### ", "");
            text = text.Replace("**", "");

            if (isDomain)
            {
                text = text.Replace("https://www.", "");
                text = text.Replace("http://www.", "");
                text = text.Replace("www.", "");
                text = text.Replace("/", "");

                if (text.Length > 30) { text = "bing.com"; }
            }

            return text;
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

        private void CloseExcel(Worksheet? xlWorksheet, Workbook? xlWorkbook, Microsoft.Office.Interop.Excel.Application? xlApp)
        {
            xlWorksheet = null;
            xlWorkbook?.Close();
            xlWorkbook = null;
            xlApp?.Quit();
            xlApp = null;
        }
        #endregion
    }
}