using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System.Data;
using System.IO;
using OfficeOpenXml;

public class NeobrokerPortfolioImporter
{
    private IWebDriver driver;

    public IWebDriver SeleniumWebDriver(string webBrowser = "chrome")
    {
        if (webBrowser == "chrome")
        {
            var options = new ChromeOptions();
            options.PageLoadStrategy = PageLoadStrategy.Eager;
            options.AddArgument("--disable-search-engine-choice-screen");
            options.AddArgument("--disable-javascript");
            options.AddUserProfilePreference("intl.accept_languages", "en_us");
            options.AddUserProfilePreference("enable_do_not_track", true);
            options.AddUserProfilePreference("download.prompt_for_download", false);
            options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", true);

            driver = new ChromeDriver(options);
        }
        else if (webBrowser == "firefox")
        {
            var options = new FirefoxOptions();
            options.PageLoadStrategy = PageLoadStrategy.Eager;
            options.SetPreference("javascript.enabled", false);
            options.SetPreference("intl.accept_languages", "en_us");
            options.SetPreference("privacy.donottrackheader.enabled", true);
            options.SetPreference("browser.download.manager.showWhenStarting", false);
            options.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream");
            options.SetPreference("browser.download.folderList", 2);

            driver = new FirefoxDriver(options);
        }

        return driver;
    }

    public void SeleniumWebDriverQuit()
    {
        driver.Quit();
    }

    public DataTable ScalableCapitalPortfolioImport(string login = null, string password = null, string fileType = ".xlsx", string outputPath = null, bool returnDataTable = false)
    {
        if (driver == null)
        {
            driver = SeleniumWebDriver("chrome");
        }

        driver.Navigate().GoToUrl("https://de.scalable.capital/en/secure-login");

        if (!string.IsNullOrEmpty(login) && !string.IsNullOrEmpty(password))
        {
            driver.FindElement(By.Id("username")).SendKeys(login);
            driver.FindElement(By.Id("password")).SendKeys(password);
            System.Threading.Thread.Sleep(2000);
            driver.FindElement(By.XPath(".//*[@type='submit']")).Submit();
        }
        else
        {
            while (true)
            {
                try
                {
                    driver.FindElement(By.XPath(".//div[@data-testid='greeting-text']"));
                    break;
                }
                catch (NoSuchElementException)
                {
                    System.Threading.Thread.Sleep(2000);
                }
            }
        }

        System.Threading.Thread.Sleep(3000);
        try
        {
            driver.ExecuteScript("return document.querySelector(\"#usercentrics-root\").shadowRoot.querySelector(\"button[data-testid='uc-deny-all-button']\").click();");
        }
        catch (Exception) { }

        driver.Navigate().GoToUrl("https://de.scalable.capital/broker/");
        System.Threading.Thread.Sleep(5000);

        try
        {
            driver.FindElement(By.XPath(".//button[contains(text(), 'Close')]")).Click();
        }
        catch (Exception) { }

        System.Threading.Thread.Sleep(3000);
        var parentSection = driver.FindElement(By.XPath(".//section[@aria-label='Security list']"));

        var elements = parentSection.FindElements(By.XPath(".//div[@aria-label='grid']//div[@role='rowgroup']//div[@role='row']//div[@role='table']"));

        var assetNames = new List<string>();
        var currentValues = new List<string>();

        foreach (var element in elements)
        {
            var text = element.Text.Split('\n');
            assetNames.Add(text[0]);
            currentValues.Add(text[1]);
        }

        var isinElements = parentSection.FindElements(By.XPath(".//div[@aria-label='grid']//div[@role='rowgroup']//div[@role='row']//a"));
        var isinCodes = new List<string>();

        foreach (var element in isinElements)
        {
            isinCodes.Add(element.GetAttribute("href"));
        }

        isinCodes = isinCodes.Select(isin => Regex.Replace(isin, @"https://de.scalable.capital/broker/security\?isin=", "")).ToList();

        var assetsTable = new DataTable();
        assetsTable.Columns.Add("asset_name");
        assetsTable.Columns.Add("isin_code");
        assetsTable.Columns.Add("current_value");

        for (int i = 0; i < assetNames.Count; i++)
        {
            var row = assetsTable.NewRow();
            row["asset_name"] = assetNames[i];
            row["isin_code"] = isinCodes[i];
            row["current_value"] = Regex.Replace(currentValues[i], @"(^.*\u20ac)([0-9]+,[0-9]+\.[0-9]+|[0-9]+\.[0-9]+)(.*)?", "$2").Replace(",", "");
            assetsTable.Rows.Add(row);
        }

        var shares = new List<Dictionary<string, object>>();

        foreach (var isinCode in isinCodes)
        {
            driver.Navigate().GoToUrl($"https://de.scalable.capital/broker/security?isin={isinCode}");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div[contains(text(), 'Shares')]//..//span")));

            var shareValue = int.Parse(driver.FindElement(By.XPath("//div[contains(text(), 'Shares')]//..//span")).Text);
            shares.Add(new Dictionary<string, object> { { "isin_code", isinCode }, { "shares", shareValue } });
        }

        var sharesTable = new DataTable();
        sharesTable.Columns.Add("isin_code");
        sharesTable.Columns.Add("shares");

        foreach (var share in shares)
        {
            var row = sharesTable.NewRow();
            row["isin_code"] = share["isin_code"];
            row["shares"] = share["shares"];
            sharesTable.Rows.Add(row);
        }

        var mergedTable = from asset in assetsTable.AsEnumerable()
                          join share in sharesTable.AsEnumerable()
                          on asset["isin_code"] equals share["isin_code"]
                          select new
                          {
                              asset_name = asset["asset_name"],
                              isin_code = asset["isin_code"],
                              shares = share["shares"],
                              current_value = asset["current_value"]
                          };

        var finalTable = new DataTable();
        finalTable.Columns.Add("date");
        finalTable.Columns.Add("type");
        finalTable.Columns.Add("financial_institution");
        finalTable.Columns.Add("asset_name");
        finalTable.Columns.Add("isin_code");
        finalTable.Columns.Add("shares");
        finalTable.Columns.Add("current_value");

        foreach (var row in mergedTable)
        {
            var newRow = finalTable.NewRow();
            newRow["date"] = DateTime.Now.Date;
            newRow["type"] = "Investments";
            newRow["financial_institution"] = "Scalable Capital";
            newRow["asset_name"] = row.asset_name;
            newRow["isin_code"] = row.isin_code;
            newRow["shares"] = row.shares;
            newRow["current_value"] = row.current_value;
            finalTable.Rows.Add(newRow);
        }

        if (fileType == ".xlsx" && outputPath != null)
        {
            using (var package = new ExcelPackage(new FileInfo(outputPath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Portfolio");
                worksheet.Cells["A1"].LoadFromDataTable(finalTable, true);
                package.Save();
            }
        }
        else if (fileType == ".csv" && outputPath != null)
        {
            using (var writer = new StreamWriter(outputPath))
            {
                foreach (DataRow row in finalTable.Rows)
                {
                    writer.WriteLine(string.Join(",", row.ItemArray));
                }
            }
        }
        else
        {
            // Copy to clipboard logic can be implemented here if needed
        }

        if (returnDataTable)
        {
            return finalTable;
        }

        return null;
    }

    public DataTable TradeRepublicPortfolioImport(string login = null, string password = null, string fileType = ".xlsx", string outputPath = null, bool returnDataTable = false)
    {
        if (driver == null)
        {
            driver = SeleniumWebDriver("chrome");
        }

        driver.Navigate().GoToUrl("https://app.traderepublic.com");

        driver.FindElement(By.XPath(".//form[@class='consentCard__form']//span[@class='buttonBase__title']")).Click();

        if (!string.IsNullOrEmpty(login) && !string.IsNullOrEmpty(password))
        {
            driver.FindElement(By.Id("loginPhoneNumber__input")).SendKeys(login);
            System.Threading.Thread.Sleep(1000);
            driver.FindElement(By.XPath(".//span[@class='buttonBase__titleWrapper']")).Click();

            var pinsInput = driver.FindElements(By.XPath(".//input[@type='password']"));
            var pins = password.ToCharArray();

            for (int i = 0; i < pinsInput.Count; i++)
            {
                pinsInput[i].SendKeys(pins[i].ToString());
            }
        }

        while (true)
        {
            try
            {
                driver.FindElement(By.XPath(".//span[@class='portfolio__pageTitle']"));
                break;
            }
            catch (NoSuchElementException)
            {
                System.Threading.Thread.Sleep(2000);
            }
        }

        driver.Navigate().GoToUrl("https://app.traderepublic.com/portfolio");
        System.Threading.Thread.Sleep(5000);

        try
        {
            driver.FindElement(By.XPath(".//div[@class='focusManager__content']//button")).Click();
            System.Threading.Thread.Sleep(2000);
        }
        catch (Exception) { }

        driver.FindElement(By.XPath("//div[@class='dropdownList']")).Click();
        driver.FindElement(By.XPath("//div[@class='dropdownList']//li[@id='investments-sinceBuyabs']")).Click();

        var portfolioList = driver.FindElements(By.XPath("//ul[@class='portfolioInstrumentList']//li"));

        var data = new List<Dictionary<string, object>>();

        foreach (var portfolio in portfolioList)
        {
            var d = new Dictionary<string, object>
            {
                { "asset_name", portfolio.FindElement(By.XPath(".//span[@class='instrumentListItem__name']")).Text },
                { "isin_code", portfolio.GetAttribute("id") },
                { "shares", portfolio.FindElement(By.XPath(".//span[@class='instrumentListItem__priceRow']//span")).Text },
                { "current_value", float.Parse(Regex.Replace(portfolio.FindElement(By.XPath(".//span[@class='instrumentListItem__priceRow']//span[@class='instrumentListItem__currentPrice']")).Text, @" \u20ac", "")) }
            };

            data.Add(d);
        }

        var assetsTable = new DataTable();
        assetsTable.Columns.Add("asset_name");
        assetsTable.Columns.Add("isin_code");
        assetsTable.Columns.Add("shares");
        assetsTable.Columns.Add("current_value");

        foreach (var item in data)
        {
            var row = assetsTable.NewRow();
            row["asset_name"] = item["asset_name"];
            row["isin_code"] = item["isin_code"];
            row["shares"] = item["shares"];
            row["current_value"] = item["current_value"];
            assetsTable.Rows.Add(row);
        }

        var finalTable = new DataTable();
        finalTable.Columns.Add("date");
        finalTable.Columns.Add("type");
        finalTable.Columns.Add("financial_institution");
        finalTable.Columns.Add("asset_name");
        finalTable.Columns.Add("isin_code");
        finalTable.Columns.Add("shares");
        finalTable.Columns.Add("current_value");

        foreach (DataRow row in assetsTable.Rows)
        {
            var newRow = finalTable.NewRow();
            newRow["date"] = DateTime.Now.Date;
            newRow["type"] = "Investments";
            newRow["financial_institution"] = "Trade Republic";
            newRow["asset_name"] = row["asset_name"];
            newRow["isin_code"] = row["isin_code"];
            newRow["shares"] = row["shares"];
            newRow["current_value"] = row["current_value"];
            finalTable.Rows.Add(newRow);
        }

        if (fileType == ".xlsx" && outputPath != null)
        {
            using (var package = new ExcelPackage(new FileInfo(outputPath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Portfolio");
                worksheet.Cells["A1"].LoadFromDataTable(finalTable, true);
                package.Save();
            }
        }
        else if (fileType == ".csv" && outputPath != null)
        {
            using (var writer = new StreamWriter(outputPath))
            {
                foreach (DataRow row in finalTable.Rows)
                {
                    writer.WriteLine(string.Join(",", row.ItemArray));
                }
            }
        }
        else
        {
            // Copy to clipboard logic can be implemented here if needed
        }

        if (returnDataTable)
        {
            return finalTable;
        }

        return null;
    }
}
