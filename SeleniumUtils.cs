using Google.Apis.Gmail.v1;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace Tool;
public static class SeleniumUtils
{
    public static void RunAutomation(GmailService service, string gmail)
    {
        try
        {
            // Khởi tạo ChromeDriver
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless");
            using (var driver = new ChromeDriver(options))
            {
                // Mở trang web
                driver.Navigate().GoToUrl("https://vnba.com.vn/net-dep-banker/bai-du-thi/271");

                // Tìm button và click vào nó
                var voteBtn = By.XPath("/HTML/body/main/div[3]/div[1]/div[2]/div/button");
                var indexFail = 0;
                isClickable(driver, voteBtn, ref indexFail);

                IWebElement inputEmail = IsDisplay(driver, By.XPath("//*[@id=\"email\"]"));
                inputEmail.SendKeys(gmail);

                IWebElement submitBtn = driver.FindElement(By.XPath("/html/body/main/div[3]/div/div/div/form/div[2]/button"));
                submitBtn.Click();

                // start get code
                var (msgId, code) = GmailUtils.GetCodeInMail(service);
                if (string.IsNullOrEmpty(msgId))
                {
                    throw new Exception($"---------------------> Cannot get msgId of email: {gmail} or mail sending has been limited <---------------------");
                }
                GmailUtils.MsgMarkAsRead(service, msgId);

                IWebElement inputCode = IsDisplay(driver, By.XPath("/html/body/main/div[3]/div/div/div[1]/form/div[1]/div/input"));
                inputCode.SendKeys(code);

                var submitCodeBtn = By.XPath("/html/body/main/div[3]/div/div/div[1]/form/div[2]/button");
                indexFail = 0;
                isClickable(driver, submitCodeBtn, ref indexFail);

                var vote1Btn = By.XPath("/HTML/body/main/div[3]/div[1]/div[2]/div/button");
                indexFail = 0;
                isClickable(driver, vote1Btn, ref indexFail);

                driver.Quit();
            }
        }
        catch (Exception ex)
        {
            throw new Exception("----------> Error when running automation <----------", ex);
        }
    }

    public static IWebElement IsDisplay(ChromeDriver driver, By element)
    {
        try
        {
            IWebElement input = driver.FindElement(element);
            return input;
        }
        catch
        {
            System.Threading.Thread.Sleep(1000);
        }
        return IsDisplay(driver, element);
    }

    public static void isClickable(ChromeDriver driver, By element, ref int index)
    {
        try
        {
            IWebElement btn = driver.FindElement(element);
            btn.Click();
        }
        catch
        {
            if (index == 5)
            {
                throw new Exception("cannot found element");
            }
            index++;
            System.Threading.Thread.Sleep(1000);
            isClickable(driver, element, ref index);
        }
    }
}
