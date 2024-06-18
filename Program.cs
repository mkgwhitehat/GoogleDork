using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        // Set the license context for EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Set up the Chrome driver (normal mode, not headless)
        var options = new ChromeOptions();
        options.AddArgument("--start-maximized");
        using (IWebDriver driver = new ChromeDriver(options))
        {
            // Define the base URL and file types to search for
            string baseUrl = "https://www.google.com/search?q=site:example.com+filetype:";
            List<string> fileTypes = new List<string>
            {
                "db", "dbf", "mdb", "sql", "bak", "cfg", "config", "csv",
                "xls", "xlsx", "pdf", "doc", "docx", "ppt", "pptx",
                "log", "passwd", "shadow", "htpasswd", "htaccess", "ini",
                "php", "asp", "aspx", "jsp", "json", "xml", "txt", "env"
            };

            // Create a new Excel package
            using (ExcelPackage package = new ExcelPackage())
            {
                bool hasWorksheets = false; // Flag to check if any worksheets were added

                Random rnd = new Random();

                // Loop through each file type, perform search, and save results in Excel
                foreach (string fileType in fileTypes)
                {
                    string searchUrl = baseUrl + fileType;
                    Console.WriteLine($"Searching for file type: {fileType}");
                    List<string> urls = GetUrls(driver, searchUrl);
                    if (urls.Count > 0)
                    {
                        var worksheet = package.Workbook.Worksheets.Add(fileType);
                        worksheet.Cells[1, 1].Value = "URL";

                        for (int i = 0; i < urls.Count; i++)
                        {
                            worksheet.Cells[i + 2, 1].Value = urls[i];
                        }

                        hasWorksheets = true;
                    }
                    else
                    {
                        Console.WriteLine($"No URLs found for file type: {fileType}");
                    }

                    // Introduce a random delay between 5 and 10 seconds
                    Thread.Sleep(rnd.Next(5000, 10000));
                }

                if (hasWorksheets)
                {
                    // Save the workbook only if there are worksheets
                    string outputPath = "provide full file path here for excel file";
                    FileInfo fileInfo = new FileInfo(outputPath);
                    package.SaveAs(fileInfo);

                    Console.WriteLine($"Search results saved to {outputPath}");
                }
                else
                {
                    Console.WriteLine("No URLs found for any file type.");
                }
            }
        }
    }

    static List<string> GetUrls(IWebDriver driver, string searchUrl)
    {
        driver.Navigate().GoToUrl(searchUrl);
        Thread.Sleep(5000); // Wait for page to load

        // Scroll down the page to load more search results
        for (int i = 0; i < 5; i++)
        {
            ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, document.body.scrollHeight)");
            Thread.Sleep(2000); // Wait for the page to load
        }

        List<string> urls = new List<string>();
        try
        {
            var results = driver.FindElements(By.CssSelector("div.g")); // Selector for entire search result blocks

            Console.WriteLine($"Found {results.Count} results for search URL: {searchUrl}");

            foreach (var result in results)
            {
                try
                {
                    var linkElement = result.FindElement(By.TagName("a"));
                    string url = linkElement.GetAttribute("href");
                    if (!string.IsNullOrEmpty(url) && url.StartsWith("http"))
                    {
                        urls.Add(url);
                    }
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine("No link found in this search result.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while extracting URLs: {ex.Message}");
        }

        Console.WriteLine($"Found {urls.Count} URLs for search URL: {searchUrl}");

        return urls;
    }


}
