using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Data;
using HtmlAgilityPack;
using System.Xml.Linq;
using static Program;
using OfficeOpenXml.Table;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        DataTable dt = new DataTable();
        dt.Clear();
        dt.Columns.Add("License ID");
        dt.Columns.Add("License Type");
        dt.Columns.Add("Legal License Name");
        dt.Columns.Add("Status");

        dt.Columns.Add("Entity Type");
        dt.Columns.Add("Active/Inactive");
        dt.Columns.Add("Original License Date");
        dt.Columns.Add("Expiration Date");
        dt.Columns.Add("Restriction");
        dt.Columns.Add("Trade/Professional Name");
        dt.Columns.Add("Special Privilege");
        dt.Columns.Add("Conditions & Limitations");
        dt.Columns.Add("Class Prefix");
        dt.Columns.Add("Business Code");
        dt.Columns.Add("Educational Code");
        dt.Columns.Add("Business Address");

        dt.Columns.Add("Name");
        dt.Columns.Add("Effective Date");
        dt.Columns.Add("Termination Date");

        dt.Columns.Add("Employee Name");
        dt.Columns.Add("Employee Position");
        dt.Columns.Add("Position Status");
        dt.Columns.Add("Employee License ID");
        dt.Columns.Add("Employee License Status");
        dt.Columns.Add("Employee Effective Date");

        dt.Columns.Add("Employer Name");
        dt.Columns.Add("Employer Position");
        dt.Columns.Add("Employer License ID");
        dt.Columns.Add("Employer License Status");
        dt.Columns.Add("Employer Effective Date");

        dt.Columns.Add("Insurance Type");
        dt.Columns.Add("Insurance Effective");
        dt.Columns.Add("Term Date");
        dt.Columns.Add("Cancel Date");
        dt.Columns.Add("Policy #");
        dt.Columns.Add("Surety");
        dt.Columns.Add("Amount");

        dt.Columns.Add("Class Code");
        dt.Columns.Add("Class Type");
        dt.Columns.Add("Effective");
        dt.Columns.Add("Restricted");
        dt.Columns.Add("License Class Status");
        try
        {


            for (int i = 208; i<=2000; i++)
            {
                string baseUrl = "https://mypvl.dcca.hawaii.gov/public-license-details/?licenseId=CT-"+i+"-0";
                string responseBody = GetBrowserData(baseUrl);
                if (!string.IsNullOrEmpty(responseBody))
                {


                    // Load HTML content into HtmlDocument
                    HtmlDocument htmlDocument = new HtmlDocument();
                    htmlDocument.LoadHtml(responseBody);

                    HtmlNode licenseIdNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='License ID']/following-sibling::p");
                    HtmlNode licenseTypeNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='License Type']/following-sibling::p");

                    HtmlNode LegalLicenseNameNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Legal License Name']/following-sibling::p");
                    HtmlNode StatusNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Status']/following-sibling::p");

                    HtmlNode EntityTypeNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Entity Type']/following-sibling::p");
                    HtmlNode Active_InactiveNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Active/Inactive']/following-sibling::p");


                    HtmlNode OriginalLicenseDateNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Original License Date']/following-sibling::p");
                    HtmlNode ExpirationDateNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Expiration Date']/following-sibling::p");

                    HtmlNode RestrictionNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Restriction']/following-sibling::p");
                    HtmlNode TradeProfessionalNameNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Trade/Professional Name']/following-sibling::p");


                    HtmlNode SpecialPrivilegeNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Special Privilege']/following-sibling::p");
                    HtmlNode ConditionsLimitationsNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Conditions &amp; Limitations']/following-sibling::p");

                    HtmlNode ClassPrefixNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Class Prefix']/following-sibling::p");
                    HtmlNode Business_CodeNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Business Code']/following-sibling::p");

                    HtmlNode EducationalCodeNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Educational Code']/following-sibling::p");
                    HtmlNode BusinessAddressNode = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='font-weight-bold' and text()='Business Address']/following-sibling::p");


                    // Extract License ID and License Type
                    string licenseId = licenseIdNode?.InnerText.Trim();
                    if (!string.IsNullOrEmpty(licenseId))
                    {


                        string licenseType = licenseTypeNode?.InnerText.Trim();
                        string LegalLicenseName = LegalLicenseNameNode?.InnerText.Trim();
                        string Status = StatusNode?.InnerText.Trim();
                        string EntityType = EntityTypeNode?.InnerText.Trim();
                        string Active_Inactive = Active_InactiveNode?.InnerText.Trim();
                        if (Active_Inactive=="Active")
                        {
                            string OriginalLicenseDate = OriginalLicenseDateNode?.InnerText.Trim();
                            string ExpirationDate = ExpirationDateNode?.InnerText.Trim();
                            string Restriction = RestrictionNode?.InnerText.Trim();
                            string TradeProfessionalName = TradeProfessionalNameNode?.InnerText.Trim();
                            string SpecialPrivilege = SpecialPrivilegeNode?.InnerText.Trim();
                            string ConditionsLimitations = ConditionsLimitationsNode?.InnerText.Trim();
                            string ClassPrefix = ClassPrefixNode?.InnerText.Trim();
                            string Business_Code = Business_CodeNode?.InnerText.Trim();
                            string EducationalCode = EducationalCodeNode?.InnerText.Trim();
                            string BusinessAddress = BusinessAddressNode?.InnerText.Trim();

                            List<GridRowData> rowDataList = new List<GridRowData>();
                            List<EmployeeData> employeeDataList = new List<EmployeeData>();
                            List<EmployerData> employerDataList = new List<EmployerData>();
                            List<InsuranceBondData> insuranceBondDataList = new List<InsuranceBondData>();
                            List<LicenseClassData> licenseClassDataList = new List<LicenseClassData>();

                            HtmlNodeCollection gridRows = htmlDocument.DocumentNode.SelectNodes("//table[@id='other-business-persona-names-table']/tbody/tr");

                            if (gridRows != null && gridRows.Count > 0)
                            {
                                // Create a list to store the extracted data
                                // Iterate through each row in the grid
                                foreach (HtmlNode row in gridRows)
                                {
                                    // Extract data from each column in the row
                                    HtmlNodeCollection columns = row.SelectNodes("td");
                                    if (columns != null && columns.Count >= 3) // Ensure at least three columns are present
                                    {
                                        string name = columns[0].InnerText.Trim();
                                        string effectiveDate = columns[1].InnerText.Trim();
                                        string terminationDate = columns[2].InnerText.Trim();

                                        // Create a data object and add it to the list
                                        GridRowData rowData = new GridRowData
                                        {
                                            Name = name,
                                            EffectiveDate = effectiveDate,
                                            TerminationDate = terminationDate
                                        };

                                        rowDataList.Add(rowData);
                                    }
                                }

                            }

                            HtmlNodeCollection gridRowsEmployees = htmlDocument.DocumentNode.SelectNodes("//table[@id='employees-table']/tbody/tr");

                            if (gridRowsEmployees != null && gridRowsEmployees.Count > 0)
                            {
                                // Create a list to store the extracted data


                                // Iterate through each row in the grid
                                foreach (HtmlNode row in gridRowsEmployees)
                                {
                                    // Extract data from each column in the row
                                    HtmlNodeCollection columns = row.SelectNodes("td");
                                    if (columns != null && columns.Count >= 6) // Ensure at least six columns are present
                                    {
                                        string employeeName = columns[0].InnerText.Trim();
                                        string position = columns[1].InnerText.Trim();
                                        string positionStatus = columns[2].InnerText.Trim();
                                        string licenseID = columns[3].InnerText.Trim();
                                        string licenseStatus = columns[4].InnerText.Trim();
                                        string effectiveDate = columns[5].InnerText.Trim();

                                        // Create a data object and add it to the list
                                        EmployeeData employeeData = new EmployeeData
                                        {
                                            EmployeeName = employeeName,
                                            Position = position,
                                            PositionStatus = positionStatus,
                                            LicenseID = licenseID,
                                            LicenseStatus = licenseStatus,
                                            EffectiveDate = effectiveDate
                                        };

                                        employeeDataList.Add(employeeData);
                                    }
                                }
                            }

                            HtmlNodeCollection gridRowsEmployer = htmlDocument.DocumentNode.SelectNodes("//table[@id='employer-table']/tbody/tr");

                            if (gridRowsEmployer != null && gridRowsEmployer.Count > 0)
                            {
                                // Create a list to store the extracted data


                                // Iterate through each row in the grid
                                foreach (HtmlNode row in gridRowsEmployer)
                                {
                                    // Extract data from each column in the row
                                    HtmlNodeCollection columns = row.SelectNodes("td");
                                    if (columns != null && columns.Count >= 5) // Ensure at least five columns are present
                                    {
                                        string employerName = columns[0].InnerText.Trim();
                                        string position = columns[1].InnerText.Trim();
                                        string licenseID = columns[2].InnerText.Trim();
                                        string licenseStatus = columns[3].InnerText.Trim();
                                        string effectiveDate = columns[4].InnerText.Trim();

                                        // Create a data object and add it to the list
                                        EmployerData employerData = new EmployerData
                                        {
                                            EmployerName = employerName,
                                            Position = position,
                                            LicenseID = licenseID,
                                            LicenseStatus = licenseStatus,
                                            EffectiveDate = effectiveDate
                                        };

                                        employerDataList.Add(employerData);
                                    }
                                }
                            }

                            HtmlNodeCollection gridRowsInsurance = htmlDocument.DocumentNode.SelectNodes("//table[@id='insurance-bond-table']/tbody/tr");

                            if (gridRowsInsurance != null && gridRowsInsurance.Count > 0)
                            {
                                // Create a list to store the extracted data


                                // Iterate through each row in the grid
                                foreach (HtmlNode row in gridRowsInsurance)
                                {
                                    // Extract data from each column in the row
                                    HtmlNodeCollection columns = row.SelectNodes("td");
                                    if (columns != null && columns.Count >= 7) // Ensure at least seven columns are present
                                    {
                                        string insuranceType = columns[0].InnerText.Trim();
                                        string effectiveDate = columns[1].InnerText.Trim();
                                        string termDate = columns[2].InnerText.Trim();
                                        string cancelDate = columns[3].InnerText.Trim();
                                        string policyNumber = columns[4].InnerText.Trim();
                                        string surety = columns[5].InnerText.Trim();
                                        string amount = columns[6].InnerText.Trim();

                                        // Create a data object and add it to the list
                                        InsuranceBondData insuranceBondData = new InsuranceBondData
                                        {
                                            InsuranceType = insuranceType,
                                            EffectiveDate = effectiveDate,
                                            TermDate = termDate,
                                            CancelDate = cancelDate,
                                            PolicyNumber = policyNumber,
                                            Surety = surety,
                                            Amount = amount
                                        };

                                        insuranceBondDataList.Add(insuranceBondData);
                                    }
                                }
                            }

                            HtmlNodeCollection gridRowsLicense = htmlDocument.DocumentNode.SelectNodes("//table[@id='license-classes-table']/tbody/tr");

                            if (gridRowsLicense != null && gridRowsLicense.Count > 0)
                            {
                                // Create a list to store the extracted data


                                // Iterate through each row in the grid
                                foreach (HtmlNode row in gridRowsLicense)
                                {
                                    // Extract data from each column in the row
                                    HtmlNodeCollection columns = row.SelectNodes("td");
                                    if (columns != null && columns.Count >= 5) // Ensure at least five columns are present
                                    {
                                        string classCode = columns[0].InnerText.Trim();
                                        string classType = columns[1].InnerText.Trim();
                                        string effectiveDate = columns[2].InnerText.Trim();
                                        string restricted = columns[3].InnerText.Trim();
                                        string status = columns[4].InnerText.Trim();

                                        // Create a data object and add it to the list
                                        LicenseClassData licenseClassData = new LicenseClassData
                                        {
                                            ClassCode = classCode,
                                            ClassType = classType,
                                            EffectiveDate = effectiveDate,
                                            Restricted = restricted,
                                            Status = status
                                        };

                                        licenseClassDataList.Add(licenseClassData);
                                    }
                                }
                            }


                            int maxCount = Math.Max(Math.Max(Math.Max(Math.Max(rowDataList.Count, employeeDataList.Count), employerDataList.Count), insuranceBondDataList.Count), licenseClassDataList.Count);

                            for (int j = 0; j<maxCount; j++)
                            {
                                DataRow _dr = dt.NewRow();
                                _dr["License ID"]=licenseId;
                                _dr["License Type"]= licenseType;
                                _dr["Legal License Name"]=LegalLicenseName;
                                _dr["Status"]=Status;
                                _dr["Active/Inactive"]=Active_Inactive;
                                _dr["Original License Date"]=OriginalLicenseDate;
                                _dr["Expiration Date"]=ExpirationDate;
                                _dr["Restriction"]= Restriction;
                                _dr["Trade/Professional Name"]= TradeProfessionalName;
                                _dr["Special Privilege"]= SpecialPrivilege;
                                _dr["Conditions & Limitations"]= ConditionsLimitations;
                                _dr["Class Prefix"]= ClassPrefix;
                                _dr["Business Code"]= Business_Code;
                                _dr["Educational Code"]= EducationalCode;
                                _dr["Business Address"]= BusinessAddress;

                                if (rowDataList.Count>0 &&rowDataList.Count>j)
                                {
                                    _dr["Name"]=rowDataList[j].Name;
                                    _dr["Effective Date"]=rowDataList[j].EffectiveDate;
                                    _dr["Termination Date"]=rowDataList[j].TerminationDate;
                                }
                                if (employeeDataList.Count>0&&employeeDataList.Count>j)
                                {
                                    _dr["Employee Name"]=employeeDataList[j].EmployeeName;
                                    _dr["Employee Position"]=employeeDataList[j].Position;
                                    _dr["Position Status"]=employeeDataList[j].PositionStatus;
                                    _dr["Employee License ID"]=employeeDataList[j].LicenseID;
                                    _dr["Employee License Status"]=employeeDataList[j].LicenseStatus;
                                    _dr["Employee Effective Date"]=employeeDataList[j].EffectiveDate;
                                }
                                if (employerDataList.Count>0&&employerDataList.Count>j)
                                {
                                    _dr["Employer Name"]=employerDataList[j].EmployerName;
                                    _dr["Employer Position"]=employerDataList[j].Position;
                                    _dr["Employer License ID"]=employerDataList[j].LicenseID;
                                    _dr["Employer License Status"]=employerDataList[j].LicenseStatus;
                                    _dr["Employer Effective Date"]=employerDataList[j].EffectiveDate;
                                }
                                if (insuranceBondDataList.Count>0&&insuranceBondDataList.Count>j)
                                {
                                    _dr["Insurance Type"]=insuranceBondDataList[j].InsuranceType;
                                    _dr["Insurance Effective"]=insuranceBondDataList[j].EffectiveDate;
                                    _dr["Term Date"]=insuranceBondDataList[j].TermDate;
                                    _dr["Cancel Date"]=insuranceBondDataList[j].CancelDate;
                                    _dr["Policy #"]=insuranceBondDataList[j].PolicyNumber;
                                    _dr["Surety"]=insuranceBondDataList[j].Surety;
                                    _dr["Amount"]=insuranceBondDataList[j].Amount;
                                }
                                if (licenseClassDataList.Count>0&&licenseClassDataList.Count>j)
                                {
                                    _dr["Class Code"]=licenseClassDataList[j].ClassCode;
                                    _dr["Class Type"]=licenseClassDataList[j].ClassType;
                                    _dr["Effective"]=licenseClassDataList[j].EffectiveDate;
                                    _dr["Restricted"]=licenseClassDataList[j].Restricted;
                                    _dr["License Class Status"]=licenseClassDataList[j].Status;
                                }

                                dt.Rows.Add(_dr);
                            }
                        }
                        else
                        {
                            DataRow _dr = dt.NewRow();
                            _dr["License ID"]=licenseId;
                            _dr["License Type"]= licenseType;
                            _dr["Legal License Name"]=LegalLicenseName;
                            _dr["Status"]=Status;
                            _dr["Active/Inactive"]=Active_Inactive;
                            dt.Rows.Add(_dr);
                        }
                    }
                }
            }
            
            ExportDataTableToExcel(dt, "Output_License.xlsx");
        }
        catch (Exception ex)
        {

        }
    }

    static void ExportDataTableToExcel(DataTable dataTable, string filePath)
    {
        // Check if the file already exists
        if (File.Exists(@"C:/Users/jeyaganesh.madasamy/Desktop/Output_License.xlsx"))
        {
            // If it exists, delete it
            File.Delete(filePath);
            Console.WriteLine("Existing Excel file deleted.");
        }
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        // Create a new Excel package
        using (var package = new ExcelPackage())
        {
            // Add a new worksheet to the Excel package
            var worksheet = package.Workbook.Worksheets.Add("MyWorksheet");

            // Load data from the DataTable into the worksheet starting from cell A1
            worksheet.Cells.LoadFromDataTable(dataTable, true, TableStyles.Light9);

            // Save the Excel package to a file
            var fileInfo = new FileInfo(@"C:/Users/jeyaganesh.madasamy/Desktop/Output_License.xlsx");
            package.SaveAs(fileInfo);
            //var filePath1 = "MyExcelFile.xlsx";

            //// Return the file as a download
            //var fileBytes = System.IO.File.ReadAllBytes(filePath);
            //return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MyExcelFile1.xlsx");
        }

    }
    public class GridRowData
    {
        public string Name { get; set; }
        public string EffectiveDate { get; set; }
        public string TerminationDate { get; set; }
    }
    public class EmployeeData
    {
        public string EmployeeName { get; set; }
        public string Position { get; set; }
        public string PositionStatus { get; set; }
        public string LicenseID { get; set; }
        public string LicenseStatus { get; set; }
        public string EffectiveDate { get; set; }
    }

    // Define a class to represent the employer data
    public class EmployerData
    {
        public string EmployerName { get; set; }
        public string Position { get; set; }
        public string LicenseID { get; set; }
        public string LicenseStatus { get; set; }
        public string EffectiveDate { get; set; }
    }

    public class InsuranceBondData
    {
        public string InsuranceType { get; set; }
        public string EffectiveDate { get; set; }
        public string TermDate { get; set; }
        public string CancelDate { get; set; }
        public string PolicyNumber { get; set; }
        public string Surety { get; set; }
        public string Amount { get; set; }
    }

    public class LicenseClassData
    {
        public string ClassCode { get; set; }
        public string ClassType { get; set; }
        public string EffectiveDate { get; set; }
        public string Restricted { get; set; }
        public string Status { get; set; }
    }

    public static string GetBrowserData(string url)
    {
        string innerHtml = string.Empty;

        //string edgeDriverPath = @"C:\Users\jeyaganesh.madasamy\Downloads\edgedriver_win64\msedgedriver.exe";
        // Specify the path to the downloaded ChromeDriver executable
        //string chromeDriverPath = @"C:\Users\jeyaganesh.madasamy\Downloads\chromedriver-win64\chromedriver.exe";
        string chromeDriverPath = @"~\chromedriver-win64\chromedriver.exe";

        // Set the path to the ChromeDriver executable
        ChromeOptions options = new ChromeOptions();
        options.BinaryLocation = @"C:\Program Files\Google\Chrome\Application\chrome.exe";


        options.PageLoadStrategy = PageLoadStrategy.Normal; // Default strategy
                                                            //"--headless",
        options.AddArguments(
    "start-maximized",
    "enable-automation",

    "--no-sandbox", //this is the relevant other arguments came from solving other issues
    "--disable-infobars",
    "--disable-dev-shm-usage",
    "--disable-browser-side-navigation",
    "--disable-gpu",
    "--ignore-certificate-errors",
    "--disable-extensions",
    "--dns-prefetch-disable");
        var capability = options.ToCapabilities();
        // Set the page load timeout to a longer duration (e.g., 120 seconds)
        options.AddUserProfilePreference("profile.default_content_setting_values.script_timeout", 360);
        options.AddUserProfilePreference("profile.default_content_setting_values.page_load_timeout", 360);

        //ChromeDriverService service = ChromeDriverService.CreateDefaultService(@"C:\Users\jeyaganesh.madasamy\Downloads\chromedriver-win64", "chromedriver.exe");
        ChromeDriverService service = ChromeDriverService.CreateDefaultService(@"~\chromedriver-win64", "chromedriver.exe");
        service.EnableVerboseLogging = true; // Optional: Enable verbose logging for troubleshooting


        ChromeDriver driver = new ChromeDriver(chromeDriverPath, options, TimeSpan.FromMinutes(6));
        try
        {

            driver.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(360));

            //ChromeDriver driver = new ChromeDriver(chromeDriverPath, options);
            // Set script timeout to a longer duration (e.g., 120 seconds)
            driver.Manage().Timeouts().AsynchronousJavaScript = TimeSpan.FromSeconds(360);

            // Set up WebDriverWait with a timeout of 120 seconds
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(360));

            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(360);
            //// Set a timeout for the GoToUrl operation
            //TimeSpan goToUrlTimeout = TimeSpan.FromSeconds(300);
            //driver.Manage().Timeouts().ImplicitWait = goToUrlTimeout;
            driver.Navigate().GoToUrl(url);

            ////RemoteWebDriver driver1 = new RemoteWebDriver(new Uri(url), capability, TimeSpan.FromMinutes(5));
            ////var remotesource = driver1.PageSource;

            //wait.Until(d => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));


            ////((IJavaScriptExecutor)driver).ExecuteScript($"window.location.href = '{url}';");
            ////driver.Navigate().GoToUrl(url);

            // Wait for some time for the page to load (you might need to adjust this)
            //System.Threading.Thread.Sleep(5000);

            string responseBody = driver.PageSource;
            var element = driver.FindElement(By.TagName("main"));
            innerHtml = element.GetAttribute("innerHTML");
            driver.Quit();
            driver.Dispose();


            // Now you can process responseBody as needed
            return innerHtml;
        }
        catch (Exception ex)
        {
            driver.Quit();
            driver.Dispose();
            return "";
        }
    }

}