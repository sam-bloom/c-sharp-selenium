using System;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;


namespace TestAutomation.FrameworkSupports
{
    public class DriverScript
    {
        public IWebDriver driver;
        /// <summary>
        /// Initialize web driver
        /// </summary>
        public void InitializeDriver()
        {
            var ieOptions = new InternetExplorerOptions();
            ieOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            ieOptions.RequireWindowFocus = true;

            switch (BaseUtilities.browser.ToLower())
            {
                case "chrome":

                    var chromeOpt = new ChromeOptions();
                    chromeOpt.AddArgument("--disable-extensions");
                    chromeOpt.AddAdditionalCapability("useAutomationExtension", false);
                    chromeOpt.Proxy = null;
                    chromeOpt.AddArgument("no-sandbox");
                    driver = new ChromeDriver(Path.Combine(BaseUtilities.GetFolderPath(), "Drivers"), chromeOpt);

                    break;
                case "ie":
                    driver = new InternetExplorerDriver(Path.Combine(BaseUtilities.GetFolderPath(), "Drivers"), ieOptions);
                    break;

                default:

                    driver = new InternetExplorerDriver(Path.Combine(BaseUtilities.GetFolderPath(), "Drivers"), ieOptions);
                    break;

            }
            PropertiesCollection.driver = driver;
            PropertiesCollection.wait = new WebDriverWait(driver, TimeSpan.FromSeconds(Convert.ToInt16(BaseUtilities.objectIdentificationTimeOut)));
            PropertiesCollection.jsexecutor = (IJavaScriptExecutor)driver;
            PropertiesCollection.actions = new Actions(driver);
            driver.Manage().Window.Maximize();
            driver.Manage().Timeouts().ImplicitWait = (TimeSpan.FromSeconds(Convert.ToInt16(BaseUtilities.objectIdentificationTimeOut)));
            driver.Manage().Cookies.DeleteAllCookies();

        }


    }
}
