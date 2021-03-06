// Generated by Selenium IDE
using System;
using System.Collections.Generic;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
//using OpenQA.Selenium.Firefox;

using OpenQA.Selenium.Support.UI;

using System.IO;
using System.Reflection;


public class MyTestCaseTest
{
    private IWebDriver driver;
    public IDictionary<string, object> vars { get; private set; }
    private IJavaScriptExecutor js;
    // グーグルクロームを起動します
    public void SetUp()
    {
        var options = new ChromeOptions();
        options.AddArguments("--lang=en");
        options.AddArguments("--headless");
        driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),options);
        js = (IJavaScriptExecutor)driver;
        vars = new Dictionary<string, object>();
    }
    // プログラムを終了する
    public void TearDown()
    {
        driver.Quit();
    }
    // Azureを起動します
    public void StartAzure(string user, string password)
    {
        try
        {         
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));     
            driver.Navigate().GoToUrl("https://portal.azure.com/#home");
            driver.Manage().Window.Maximize();
            //ユーザー名を追加
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("i0116")));
            driver.FindElement(By.Id("i0116")).SendKeys(user);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("idSIButton9")).Click();

            // パスワードを追加
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("i0118")));
            driver.FindElement(By.Id("i0118")).SendKeys(password);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("idSIButton9")).Click();


            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("_weave_e_24")));
            driver.FindElement(By.Id("_weave_e_24")).Click();

            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("mectrl_switchdirectory")));
            driver.FindElement(By.Id("mectrl_switchdirectory")).Click();

            wait.Until(ExpectedConditions.ElementToBeClickable((By.CssSelector(".fxc-listView-item:nth-child(2) .fxs-directorypane-gridcellcontentdetails:nth-child(2)"))));
            driver.FindElement(By.CssSelector(".fxc-listView-item:nth-child(2) .fxs-directorypane-gridcellcontentdetails:nth-child(2)")).Click();

            Thread.Sleep(3000);

        }
        catch (Exception e)
        {
            Console.WriteLine(e);
        }
    }
    
    // クラウドサービスの開始/停止
    public bool TestCases(string listCloudService, string stateStoping)
    {
        try
        { 
            int countAttend = 0;
            int countSuccess = 0;
            WebDriverWait wait_90 = new WebDriverWait(driver, TimeSpan.FromSeconds(90));
            WebDriverWait wait_30 = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

            wait_30.Until(ExpectedConditions.ElementToBeClickable((By.Id("_weave_e_2"))));
            driver.FindElement(By.Id("_weave_e_2")).Click();

            wait_30.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//a[contains(@href, 'microsoft.classicCompute%2FdomainNames')]")));
            driver.FindElement(By.XPath("//a[contains(@href, 'microsoft.classicCompute%2FdomainNames')]")).Click();

            Thread.Sleep(3000);

            wait_30.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//input[starts-with(@name, '__azc-textBox-tsx') and starts-with(@id, 'form-label-id')]")));
            driver.FindElement(By.XPath("//input[starts-with(@name, '__azc-textBox-tsx') and starts-with(@id, 'form-label-id')]")).SendKeys(listCloudService);

            Thread.Sleep(3000);

            if (IsElementPresent((By.LinkText(listCloudService))) == false)
            {
                wait_30.Until(ExpectedConditions.ElementToBeClickable(By.Id("_weave_e_4")));
                driver.FindElement(By.Id("_weave_e_4")).Click();
                return false;
            }
            else
            {
                wait_30.Until(ExpectedConditions.ElementToBeClickable((By.LinkText(listCloudService))));
                driver.FindElement(By.LinkText(listCloudService)).Click();
            }

            wait_30.Until(ExpectedConditions.ElementToBeClickable((By.XPath("//div[@title=\'Stop\'  or @title=\'Start\'  or @title=\'開始\' or @title=\'停止\' ]"))));

            if (IsElementPresent(By.XPath("//span[contains(text(),'No deployment') or contains(text(),'デプロイなし')]")) == false)
            {
                var curState = driver.FindElement(By.XPath("//div[@title=\'Stop\'  or @title=\'Start\' or @title=\'開始\' or @title=\'停止\' ]")).GetAttribute("title");
                if (String.Compare("stop", stateStoping, true) == 0 || String.Compare("start", stateStoping, true) == 0 || String.Compare("停止", stateStoping, true) == 0 || String.Compare("開始", stateStoping, true) == 0)
                {
                    if (String.Compare(curState, "stop", true) == 0 || String.Compare(curState, "停止", true) == 0)
                    {
                        countAttend++;
                        driver.FindElement(By.XPath("//div[@title=\'Stop\'  or @title=\'Start\' or @title=\'開始\' or @title=\'停止\' ]")).Click();

                        wait_30.Until(ExpectedConditions.ElementToBeClickable((By.XPath("//div[@title=\'Yes\' or @title=\'はい\' ]"))));
                        driver.FindElement(By.XPath("//div[@title=\'Yes\' or @title=\'はい\' ]")).Click();

                        wait_90.Until(ExpectedConditions.ElementExists((By.XPath("//li[contains(@aria-label, 'New notification with title Successfully')]"))));

                        if (IsElementPresent(By.XPath("//li[contains(@aria-label, 'New notification with title Successfully')]")) == true)
                        {
                            countSuccess++;
                        }
                        Thread.Sleep(3000);
                    }
                }
                else
                {
                    return false;
                }

            }

            wait_30.Until(ExpectedConditions.ElementToBeClickable((By.XPath("//div[@title=\'Production\' or @title=\'本番環境\']"))));
            driver.FindElement(By.XPath("//div[@title=\'Production\' or @title=\'本番環境\']")).Click();

            wait_30.Until(ExpectedConditions.ElementToBeClickable((By.XPath("//div[@title=\'Production\' or @title=\'本番環境\' ]"))));
            driver.FindElement(By.XPath("//div[@title=\'Production\' or @title=\'本番環境\']")).Click();

            wait_30.Until(ExpectedConditions.ElementToBeClickable((By.XPath("//li[@title=\'Staging\' or @title=\'ステージング環境\']"))));
            driver.FindElement(By.XPath("//li[@title=\'Staging\' or @title=\'ステージング環境\' ]")).Click();

            Thread.Sleep(3000);

            wait_30.Until(ExpectedConditions.ElementToBeClickable((By.XPath("//div[@title=\'Stop\'  or @title=\'Start\'  or @title=\'開始\' or @title=\'停止\' ]"))));

            if (IsElementPresent(By.XPath("//span[contains(text(),'No deployment') or contains(text(),'デプロイなし')]")) == false)
            {

                var curState = driver.FindElement(By.XPath("//div[@title=\'Stop\'  or @title=\'Start\' or @title=\'開始\' or @title=\'停止\' ]")).GetAttribute("title");
                if (String.Compare("stop", stateStoping, true) == 0 || String.Compare("start", stateStoping, true) == 0 || String.Compare("停止", stateStoping, true) == 0 || String.Compare("開始", stateStoping, true) == 0)
                {
                    if (String.Compare(curState, "stop", true) == 0 || String.Compare(curState, "停止", true) == 0)
                    {
                        countAttend++;
                        driver.FindElement(By.XPath("//div[@title=\'Stop\'  or @title=\'Start\' or @title=\'開始\' or @title=\'停止\' ]")).Click();

                        wait_30.Until(ExpectedConditions.ElementToBeClickable((By.XPath("//div[@title=\'Yes\' or @title=\'はい\' ]"))));
                        driver.FindElement(By.XPath("//div[@title=\'Yes\' or @title=\'はい\' ]")).Click();

                        wait_90.Until(ExpectedConditions.ElementExists((By.XPath("//li[contains(@aria-label, 'New notification with title Successfully')]"))));
                        if (IsElementPresent(By.XPath("//li[contains(@aria-label, 'New notification with title Successfully')]")) == true)
                        {
                            countSuccess++;
                        }
                        Thread.Sleep(3000);
                    }
                }
                else
                {
                    return false;
                }

            }

            if (countSuccess == countAttend)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        catch
        {
            return false;
        }
    }

    // 表示されている要素を確認してください
    public bool IsElementPresent(By by)
    {
        try
        {
            driver.FindElement(by);
            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            return false;
        }
    }
}
