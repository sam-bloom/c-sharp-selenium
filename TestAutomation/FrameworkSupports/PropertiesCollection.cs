using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;


namespace TestAutomation.FrameworkSupports
{
    /// <summary>
    /// Global properties collection for framewrok 
    /// </summary>
   public class PropertiesCollection
    {
        public static IWebDriver driver { get; set; }
        public static WebDriverWait wait { get; set; }
        public static IJavaScriptExecutor jsexecutor { get; set; }
        public static Actions actions { get; set; }
    }
    
}
