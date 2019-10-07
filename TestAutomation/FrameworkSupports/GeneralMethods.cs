using AventStack.ExtentReports.Model;
using TestAutomation.FrameworkSupports;
using TestAutomation.FrameworkSupports.Reporting;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace TestAutomation.FrameworkSupports
{
    
    public class GeneralMethods : PropertiesCollection
    {
        public SelectElement selectElement;
        public IWebElement element;

        /// <summary>
        /// Thread wait 
        /// </summary>
        /// <param name="timeinmillisec"></param>
        public static void WaitImplicit(int timeinmillisec)
        {
            new ManualResetEvent(false).WaitOne(timeinmillisec);
        }
        public static string Encrypt(string plainText)
        {
            var encryptInBytes = Encoding.UTF8.GetBytes(plainText.Trim());
            return Convert.ToBase64String(encryptInBytes);
        }

        public static string Decrypt(string encryptedText)
        {
            var decryptInBytes = Convert.FromBase64String(encryptedText.Trim());
            return Encoding.UTF8.GetString(decryptInBytes);
        }

        /// <summary>
        /// Click element if clickable
        /// </summary>
        /// <param name="by"></param>
        public void ClickByElement(By by)
        {
            try
            {
                WaitElementToBeClickable(by);
                driver.FindElement(by).Click();
            }

                    catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// Click element if clickable
        /// </summary>
        /// <param name="by"></param>
        public void ClickByElement(IWebElement ele)
        {
            try
            {
                WaitElementToBeClickable(ele);
                ele.Click();
            }

            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// wait till element to be clickable
        /// </summary>
        /// <param name="by"></param>
        public void WaitElementToBeClickable(By by)
        {
            try
            {
                wait.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(by));
            }

            catch (NoSuchElementException e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
                Assert.Fail();
            }
        }
        /// <summary>
        /// wait till element to be clickable
        /// </summary>
        /// <param name="by"></param>
        public void WaitElementToBeClickable(IWebElement ele)
        {
            try
            {
                wait.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(ele));
            }

            catch (NoSuchElementException e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
                Assert.Fail();
            }
        }


        /// <summary>
        /// wait till element to be exists
        /// </summary>
        /// <param name="by"></param>
        public void WaitElementIsExists(By by)
        {
            try
            {
                wait.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(by));
            }

            catch (NoSuchElementException e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
                Assert.Fail();
            }
        }
        /// <summary>
        /// wait till element text to be exists
        /// </summary>
        /// <param name="by"></param>
        public void WaitElementTextIsExists(By by, String text)
        {
            try
            {
                wait.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.TextToBePresentInElementLocated(by, text));
            }

            catch (NoSuchElementException e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
                Assert.Fail();
            }
        }
        /// <summary>
        /// wait till element to be visible
        /// </summary>
        /// <param name="by"></param>
        public void WaitElementToBeVisible(By by)
        {
            try
            {
                wait.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
            }
            catch (NoSuchElementException e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
                Assert.Fail();
            }
        }
        /// <summary>
        /// wait till page load complete
        /// </summary>
        /// <param name="by"></param>
        public void JSWaitTillPageLoads()
        {
            try
            {
                jsexecutor.ExecuteScript("return document.readyState").Equals("complete");
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// wait till element to be Invisible
        /// </summary>
        /// <param name="by"></param>
        public void WaitElementToBeInVisible(By by)
        {
            try
            {
                wait.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(by));
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// wait till element text to be Invisible
        /// </summary>
        /// <param name="by"></param>
        public void WaitElementWithTextInvisible(By by, string text)
        {
            try
            {
                wait.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementWithText(by, text));
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// retrun true if element is exists
        /// </summary>
        /// <param name="by"></param>
        /// <returns></returns>
        public bool IsElementExist(By by)
        {

            IList<IWebElement> elecount = driver.FindElements(by);
            if (elecount.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Wait element not exists
        /// </summary>
        /// <param name="by"></param>
        public void WaitForElementToNotExist(By by)
        {
            try
            {
                element = driver.FindElement(by);
                if (element.Displayed)
                {
                    do
                    {
                        WaitImplicit(500);
                    }
                    while (element.Displayed);
                }

            }
            catch (Exception e)
            {

                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// Wait till elemenat Enabled
        /// </summary>
        /// <param name="by"></param>
        public void WaitForElementToBeEnabled(By by)
        {
            try
            {
                element = driver.FindElement(by);
                if (!element.Enabled)
                {
                    wait.PollingInterval = TimeSpan.FromMilliseconds(500);
                    wait.Until<bool>((d) =>
                    {
                        return d.FindElement(by).Enabled;
                    });
                }
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="by"></param>
        /// <param name="inputvalue"></param>
        public void SendAttachment(By by, String inputvalue)
        {
            try
            {
                driver.FindElement(by).SendKeys(inputvalue);

            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// Send value to element
        /// </summary>
        /// <param name="by"></param>
        /// <param name="inputvalue"></param>
        public void SendValue(By by, String inputvalue)
        {
            try
            {
                driver.FindElement(by).Click();
                driver.FindElement(by).Clear();
                driver.FindElement(by).SendKeys(inputvalue);

            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// Send values and Enter
        /// </summary>
        /// <param name="by"></param>
        /// <param name="value"></param>
        public void SendValueAndPressEnter(By by, string value)
        {
            driver.FindElement(by).Click();
            driver.FindElement(by).Clear();
            driver.FindElement(by).SendKeys(value);
            driver.FindElement(by).SendKeys(Keys.Enter);
        }
        /// <summary>
        /// Send values and press tab
        /// </summary>
        /// <param name="by"></param>
        /// <param name="value"></param>
        public void SendValueAndPressTab(By by, string value)
        {
            driver.FindElement(by).Click();
            driver.FindElement(by).Clear();
            driver.FindElement(by).SendKeys(value);
            driver.FindElement(by).SendKeys(Keys.Tab);
        }
        /// <summary>
        /// Get element text
        /// </summary>
        /// <param name="by"></param>
        /// <returns></returns>
        public static String GetElementText(By by)
        {
            return driver.FindElement(by).Text;
        }
        /// <summary>
        /// Get element attribute value
        /// </summary>
        /// <param name="by"></param>
        /// <param name="AttributeValue"></param>
        /// <returns></returns>
        public String GetElementAttribute(By by, String AttributeValue)
        {
            return driver.FindElement(by).GetAttribute(AttributeValue);
        }
        /// <summary>
        /// Navigate to page
        /// </summary>
        /// <param name="PageURL"></param>
        public static void NavigateToURL(String PageURL)
        {
            driver.Navigate().GoToUrl(PageURL);
        }

        /// <summary>
        /// Get title
        /// </summary>
        /// <returns></returns>
        public string GetTitle()
        {
            return driver.Title;
        }

        /// <summary>
        /// Verify Text 
        /// </summary>
        /// <param name="Actual"></param>
        /// <param name="Expected"></param>
        /// <returns></returns>
        public string VerifyText(string Actual,string Expected)
        {
            if (Actual.ToLower()== Expected.ToLower())
            {
                return "'"+Actual+"' text matching with '"+ Expected + "'";
            }
            else
            {
                return "'" + Actual + "' text not matching with '" + Expected + "'";
            }
        }
        /// <summary>
        /// Click element using javascript
        /// </summary>
        /// <param name="element"></param>
        public void JSClickElement(By by)
        {
            try
            {
                //WaitElementToBeClickable(by);
                element = driver.FindElement(by);
                jsexecutor.ExecuteScript("arguments[0].click();", element);
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// get hidden element text
        /// </summary>
        /// <param name="element"></param>
        public static string GetHiddenElementText(IWebElement ele)
        {
            return ele.GetAttribute("innerHTML");

        }
        /// <summary>
        /// Send values to element using javascript
        /// </summary>
        /// <param name="element"></param>
        public void JSSendValue(By by, string value)
        {
            try
            {
                element = driver.FindElement(by);
                jsexecutor.ExecuteScript("arguments[0].value='" + value + "';", element);
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// Set attribute to element using javascript
        /// </summary>
        /// <param name="element"></param>
        public void JSSetAttributeValue(By by, string sttribute, string value)
        {
            try
            {
                element = driver.FindElement(by);
                jsexecutor.ExecuteScript("arguments[0].SetSttribute('" + sttribute + "','" + value + "')", element);
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);
            }
        }
        /// <summary>
        /// Switch to alert
        /// </summary>
        public void SwitchToAlert()
        {
            driver.SwitchTo().Alert();
        }
        /// <summary>
        /// Get Alert Text
        /// </summary>
        /// <returns></returns>
        public String GetAlertText()
        {
            return driver.SwitchTo().Alert().Text;
        }
        /// <summary>
        /// Accept Alert
        /// </summary>
        public static void AcceptAlert()
        {
            driver.SwitchTo().Alert().Accept();
        }
        /// <summary>
        /// Dismiss Alert
        /// </summary>
        public void DissmissAlert()
        {
            driver.SwitchTo().Alert().Dismiss();
        }
        /// <summary>
        /// Send Value to Alert
        /// </summary>
        /// <param name="Value"></param>
        public void SendValueToAlert(String Value)
        {
            driver.SwitchTo().Alert().SendKeys(Value);
        }
        /// <summary>
        /// Return tru eif alert is present
        /// </summary>
        /// <returns></returns>
        public Boolean IsAlertExist()
        {
            try
            {
                driver.SwitchTo().Alert();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        /// <summary>
        /// Generate Random String
        /// </summary>
        /// <returns></returns>
        public string GenerateRandomString()
        {
            Random rnd = new Random();
            string txtRand = string.Empty;
            for (int i = 0; i < 8; i++)
            {
                txtRand += ((char)rnd.Next(97, 122)).ToString();
            }
            return txtRand;
        }
        /// <summary>
        /// Close Browser using System Process
        /// </summary>
        /// <param name="strBrowserProcess"></param>
        public void CloseBrowser(string strBrowserProcess)
        {
            Process[] allProcesses = Process.GetProcessesByName(strBrowserProcess);
            foreach (var process in allProcesses)
            {
                if (process.MainWindowTitle != "")
                {
                    process.CloseMainWindow();
                }
            }
        }
        /// <summary>
        /// Kill Process
        /// </summary>
        /// <param name="strBrowserProcess"></param>
        public void KillProcess(string browser)
        {
            Process[] allProcesses = null;
            switch (browser.ToLower())
            {
                case "chrome":
                    allProcesses = Process.GetProcessesByName("chromedriver");
                    break;
                case "ie":
                    allProcesses = Process.GetProcessesByName("IEDriverServer");
                    break;
                default:
                    allProcesses = Process.GetProcessesByName("IEDriverServer");
                    break;
            }
            foreach (var process in allProcesses)
            {
                process.Kill();

            }
        }
        /// <summary>
        /// return to element from by locatore
        /// </summary>
        /// <param name="by"></param>
        /// <returns></returns>
        public static IWebElement ReturnElement(By by)
        {
            return driver.FindElement(by);
        }
        /// <summary>
        /// Wait till page Loads
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="intTimeout"></param>
        public void WaitForPageLoad(int intTimeout)
        {
            DateTime timeOnFunctionCall = DateTime.Now;
            string strState;

            while ((DateTime.Now - timeOnFunctionCall).TotalSeconds < intTimeout)
            {

                try
                {
                    strState = (string)jsexecutor.ExecuteScript("return document.readyState");
                    if (strState == "complete")
                        break;
                    WaitImplicit(1);
                }
                catch (Exception e)
                {
                    Reports.errorMessage = e.Message + e.StackTrace;
                    Reports.ReportEvent("Exception Error", Reporting.Status.Fail, true);

                }

            }
        }
        /// <summary>
        /// Return element if element is selected.
        /// </summary>
        /// <param name="by"></param>
        /// <returns></returns>
        public Boolean IsElementSelected(By by)
        {
            Boolean IsSelected;
            if (driver.FindElement(by).Selected == true)
            {
                IsSelected = true;
            }
            else
            {
                IsSelected = false;
            }
            return IsSelected;
        }
        /// <summary>
        /// Select VisibleText
        /// </summary>
        /// <param name="by"></param>
        /// <param name="visibleText"></param>
        public void SelectByVisibleText(By by, String visibleText)
        {
            selectElement = new SelectElement(driver.FindElement(by));
            selectElement.SelectByText(visibleText);
        }

        /// <summary>
        /// Select by Index
        /// </summary>
        /// <param name="by"></param>
        /// <param name="visibleText"></param>
        public void SelectByIndex(By by, int index)
        {
            selectElement = new SelectElement(driver.FindElement(by));
            selectElement.SelectByIndex(index);
        }
        /// <summary>
        /// Select by value
        /// </summary>
        /// <param name="by"></param>
        /// <param name="visibleText"></param>
        public void SelectByValue(By by, String value)
        {
            selectElement = new SelectElement(driver.FindElement(by));
            selectElement.SelectByValue(value);
        }
        /// <summary>
        /// Switch To Frame ByName
        /// </summary>
        /// <param name="FrameName"></param>
        public void SwitchToFrameByName(String FrameName)
        {
            driver.SwitchTo().Frame(FrameName);
        }
        /// <summary>
        /// Switch To Frame ById
        /// </summary>
        /// <param name="Frameid"></param>
        public void SwitchToFrameById(int Frameid)
        {
            driver.SwitchTo().Frame(Frameid);
        }
        /// <summary>
        /// Switch To Frame By Element
        /// </summary>
        /// <param name="by"></param>
        public void SwitchToFrameByElement(By by)
        {
            driver.SwitchTo().Frame(driver.FindElement(by));
        }
        /// <summary>
        /// Get By object
        /// </summary>
        /// <param name="ElementTypeAndValue"></param>
        /// <returns></returns>
        public static By GetByObject(string ElementTypeAndValue)
        {
            string[] element= Regex.Split(ElementTypeAndValue, "~");
            string IdentifierType = element[0];
            string IdentifierValue= element[1];
            By byLocator;
            try
            {
                switch (IdentifierType.ToLower())
                {
                    case "id":
                        byLocator = By.Id(IdentifierValue);
                        return byLocator;
                    case "class":
                        byLocator = By.ClassName(IdentifierValue);
                        return byLocator;
                    case "name":
                        byLocator = By.Name(IdentifierValue);
                        return byLocator;
                    case "xpath":
                        byLocator = By.XPath(IdentifierValue);
                        return byLocator;
                    case "linktext":
                        byLocator = By.LinkText(IdentifierValue);
                        return byLocator;
                    case "partiallinktext":
                        byLocator = By.PartialLinkText(IdentifierValue);
                        return byLocator;
                    case "tagname":
                        byLocator = By.TagName(IdentifierValue);
                        return byLocator;
                    case "cssselector":
                        byLocator = By.CssSelector(IdentifierValue);
                        return byLocator;

                    default:
                        return null;
                }
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception error", Status.Fail,false);
                return null;
            }
        }
        /// <summary>
        /// Swicth To Default content
        /// </summary>
        public void SwicthToDefault()
        {
            driver.SwitchTo().DefaultContent();
        }
        /// <summary>
        /// Switch To Window
        /// </summary>
        public void SwitchToWindow(String WindowHandle)
        {
            driver.SwitchTo().Window(WindowHandle);
        }
        /// <summary>
        /// Switch To Child Window
        /// </summary>
        public void SwitchToChildWindow(String WindowHandle)
        {
            List<string> lstWindow = driver.WindowHandles.ToList();
            foreach (var handle in lstWindow)
            {
                if (handle != WindowHandle)
                {
                    SwitchToWindow(handle);
                }
            }

        }

        /// <summary>
        /// Switch To Last Window
        /// </summary>
        public void SwitchToLastWindow(String WindowHandle)
        {
            List<string> lstWindow = driver.WindowHandles.ToList();
            String LastWindow = "";
            foreach (var handle in lstWindow)
            {
                if (handle != WindowHandle)
                {
                    LastWindow = handle;

                }
            }
            SwitchToWindow(LastWindow);

        }

        /// <summary>
        /// Switch To Window by int
        /// </summary>
        public void SwitchToWindowByIndex(int index)
        {
            List<string> lstWindow = driver.WindowHandles.ToList();
            List<String> windows = new List<String>();
            foreach (var handle in lstWindow)
            {
                windows.Add(handle);
            }

            SwitchToWindow(windows[index]);

        }
        /// <summary>
        /// Close window
        /// </summary>
        public void Close()
        {
            driver.Close();
        }
        /// <summary>
        /// Return if browser is active
        /// </summary>
        /// <returns></returns>
        public static Boolean IsDriverActive()
        {
            try
            {
                String CurrentWindow = driver.CurrentWindowHandle;
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }

        /// <summary>
        /// Return if browser is active
        /// </summary>
        /// <returns></returns>
        public Boolean IsElementExists(By by)
        {
            try
            {
                WaitElementIsExists(by);
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }

        /// <summary>
        /// Return if browser is active
        /// </summary>
        /// <returns></returns>
        public Boolean IsElementVisible(By by)
        {
            try
            {
                WaitElementToBeVisible(by);
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }
        /// <summary>
        /// Close browser and dispose
        /// </summary>
        public static void CloseBrowserAndDispose()
        {
            if (IsDriverActive())
            {
                driver.Close();
                driver.Dispose();
            }


        }

        /// <summary>
        /// Refresh browser
        /// </summary>
        public void RefreshBrowser()
        {
            driver.Navigate().Refresh();
        }
        /// <summary>
        /// goback browser
        /// </summary>
        public void GoBackBrowser()
        {
            driver.Navigate().Back();
        }
        /// <summary>
        /// Forward browser
        /// </summary>
        public void ForwardBrowser()
        {
            driver.Navigate().Forward();
        }
        /// <summary>
        /// Take screenshot
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        public static void ScreenCapture(string path, string fileName)
        {

            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            ss.SaveAsFile(path + "\\" + fileName + ".png");

        }
        /// <summary>
        /// Move to Element
        /// </summary>
        /// <param name="by"></param>
        public void MoveToElement(By by)
        {
            actions.MoveToElement(ReturnElement(by)).Build().Perform();
        }
        /// <summary>
        /// Move to Element then click
        /// </summary>
        /// <param name="by"></param>
        public void MoveToElementThenClick(By by)
        {
            actions.MoveToElement(ReturnElement(by)).Click().Build().Perform();
        }
        /// <summary>
        /// Click Element and hold
        /// </summary>
        /// <param name="by"></param>
        public void ClickAndHold(By by)
        {
            actions.ClickAndHold(ReturnElement(by)).Build().Perform();
        }

        /// <summary>
        /// Double Click
        /// </summary>
        /// <param name="by"></param>
        public void DoubleClick()
        {
            actions.DoubleClick().Build().Perform();

        }

        /// <summary>
        /// Darg and drop
        /// </summary>
        /// <param name="by"></param>
        public void DragAndDrop(By source, By target)
        {
            actions.DragAndDrop(ReturnElement(source), ReturnElement(target)).Build().Perform();

        }

        /// <summary>
        /// delete cookies
        /// </summary>
        /// <param name="by"></param>
        public void DeleteCookies()
        {
            driver.Manage().Cookies.DeleteAllCookies();

        }

        /// <summary>
        /// get current date by given fromat
        /// </summary>
        /// <param name="by"></param>
        public string GetCurrentDate(string format)
        {
            return DateTime.Now.ToString(format);

        }

    }
}
