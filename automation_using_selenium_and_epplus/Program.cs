using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome; 

ChromeOptions options = new();
options.AddArgument("--start-maximized");

ChromeDriver driver = new(options);

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

FileInfo file = new("C:\\Users\\victo\\www\\work\\learn\\automation_using_selenium_and_epplus\\automation_using_selenium_and_epplus\\Data\\Emit.xlsx");

using (ExcelPackage package = new(file))
{
    ExcelWorksheet? worksheet = package.Workbook.Worksheets.FirstOrDefault();

    if(worksheet != null)
    {
        int rowCount = worksheet.Dimension.Rows;

        for (int row = 2; row <= rowCount; row++)
        {
            string? cpf = worksheet.GetValue(row, 1).ToString();
            string? email = worksheet.GetValue(row, 2).ToString();
            string? description = worksheet.GetValue(row, 3).ToString();
            string? value = worksheet.GetValue(row, 4).ToString();

            Reply(driver, cpf, email, description, value);
        }
    }else
    {
        Console.WriteLine("\n\n[ERROR] Impossível ler arquivo excel!\n\n");
    }
   
}

driver.Quit();

static void Reply(IWebDriver driver, string? cpf, string? email, string? description, string? value)
{
    driver.Navigate().GoToUrl("https://forms.gle/4N5s1BXLQCdFFJ896");

    IWebElement cpfInput = driver.FindElement(By.XPath("//*[@id='mG61Hd']/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input"));
    cpfInput.SendKeys(cpf);

    IWebElement emailInput = driver.FindElement(By.XPath("//*[@id='mG61Hd']/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input"));
    emailInput.SendKeys(email);

    IWebElement descriptionInput = driver.FindElement(By.XPath("//*[@id='mG61Hd']/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input"));
    descriptionInput.SendKeys(description);

    IWebElement valueInput = driver.FindElement(By.XPath("//*[@id='mG61Hd']/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input"));
    valueInput.SendKeys(value);

    IWebElement sendBtn = driver.FindElement(By.XPath("//*[@id='mG61Hd']/div[2]/div/div[3]/div[1]/div[1]/div"));
    sendBtn.SendKeys(Keys.Return);
}
