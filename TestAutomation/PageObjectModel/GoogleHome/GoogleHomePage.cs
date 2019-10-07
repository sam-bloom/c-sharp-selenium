using TestAutomation.FrameworkSupports;
using TestAutomation.FrameworkSupports.Reporting;
using TestAutomation.PageObjectModel;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using TestAutomation.FrameworkSupports.DataAccess;

namespace TestAutomation.PageObjectModel.GoogleHome
{
    public class GoogleHomePage : GeneralMethods
    {
        public string ApplicationURL = GetXMLData.ConfigXMLData["ApplicationURL"];



        public void launchApplication()
        {
            try
            {
                NavigateToURL(ApplicationURL);
                Reports.ReportEvent("Launching application", Status.Pass, true);
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Launching application", Status.Fail, true);
            }
        }
 
    }
}
