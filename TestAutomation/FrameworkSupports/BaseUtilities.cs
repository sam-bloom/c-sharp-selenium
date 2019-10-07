using NPOI.SS.UserModel;
using System;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Reflection;
using NUnit.Framework;
using OpenQA.Selenium;
using TestAutomation.FrameworkSupports;
using TechTalk.SpecFlow;
using OpenQA.Selenium.Support.UI;
using NUnit.Framework.Internal;
using TestAutomation.FrameworkSupports.Reporting;
using System.Globalization;
using TestAutomation.FrameworkSupports.DataAccess;

namespace TestAutomation.FrameworkSupports
{
    public class BaseUtilities
    {

        public static Int16 objectIdentificationTimeOut;
        public static Int16 objectSyncTimeout;
        public static string testCaseName;
        public static string pdfDefaultDirectory;
        public static string currentTestCase;
        public static string scenarioName;
        public static string featureFileName;
        public static string[] scenarioTag;
        public static string testSetName;
        public static string scenarioStatus;
        public static string inputDataFilePath = GetFolderPath() + @"\InputData";
        public static string inputDataFileName;
        public static string testcaseFileName = "TestCases.xlsx";
        public static string webObjectRepFilePath = GetFolderPath() + @"\ObjectRepository\Web";
        public static string webObjectRepFileName = "WebElementLocators.xlsx";
        public static string mobileObjectRepFilePath = GetFolderPath() + @"\ObjectRepository\Mobile";
        public static string mobileObjectRepFileName = "MobileElementLocators.xlsx";
        private static string propertiesFilePath = GetFolderPath() + @"\Config\GlobalSettings.properties";
        public static string applicationType;
        public static string dateFormat = "dd-MM-yyyy_HH-mm-ss_tt";
        private static string PropertiesFilePath = GetFolderPath() + @"\Config\GlobalSettings.properties";
        public static string XMLFilePath = GetFolderPath() + @"\Config\GlobalSettings.xml";


        /// <summary>
        /// The Dictionary variable holding all properties adn corresponding values
        /// </summary>

        public static string browser;
        public static Dictionary<string, int> scenariosList;

        /// <summary>
        ///  The separator string to be used for directories and files based on the current OS
        /// </summary>
        public static String GetFileSeparator()
        {
            return Path.DirectorySeparatorChar.ToString();
        }

        /// <summary>
        ///   Function to return the current time
        /// </summary>
        /// <returns>The current time</returns>
        /// <see cref="#getCurrentFormattedTime(String)"/>
        public static String GetCurrentTime()
        {
            DateTime currentTime = DateTime.Now;
            return String.Format("{0:t}", currentTime);
        }

        /// <summary>
        /// Function to return the current time, formatted as per the DateFormatString setting
        /// </summary>
        /// <param name="dateFormatString">The date format string to be applied</param>
        /// <returns> The current time, formatted as per the date format string specified</returns>
        /// <see cref="#GetCurrentTime()"/>
        /// <see cref="#GetFormattedTime(Date, String)"/>
        public static String GetCurrentFormattedTime(String dateFormatString)
        {
            DateTime currentTime = DateTime.Now;
            return String.Format("{0:" + dateFormatString + "}", currentTime);
        }

        /// <summary>
        /// Get current time with dateFormat  
        /// </summary>
        /// <returns></returns>
        public static String GetCurrentTimeWithFormat()
        {
            return GetCurrentFormattedTime(dateFormat);
        }

        /// <summary>
        ///  Function to format the given time variable as specified by the DateFormatString setting
        /// </summary>
        /// <param name="time">The date/time variable to be formatted</param>
        /// <param name="dateFormatString">The date format string to be applied</param>
        /// <returns>The specified date/time, formatted as per the date format string specified</returns>
        /// <see cref="#getCurrentFormattedTime(String)"/>
        public static String GetFormattedTime(DateTime time, String dateFormatString)
        {
            return String.Format(dateFormatString, time);
        }

        public static String GetFormattedTime(String dateString)
        {
            DateTime result = DateTime.ParseExact(dateString, "yyyy-MM-dd HH:mm tt", null);

            return String.Format("M-dd-yyyy_h:mm:ss_tt", result);
        }

        /// <summary>
        ///   Function to get the time difference between 2 Date variables in minutes/seconds format
        /// </summary>
        /// <param name="startTime">The start time</param>
        /// <param name="endTime"> The end time</param>
        /// <returns>Time difference in minutes/seconds format</returns>
        public static String GetTimeDifference(DateTime startTime, DateTime endTime)
        {
            // TODO: Search for better method for date difference than using '-' operator
            return endTime.Subtract(startTime).TotalMinutes.ToString();
        }


        /// <summary>
        /// Get Current Execution Folder Path
        /// </summary>
        /// <returns>Resturn Folder Path</returns>
        public static string GetFolderPath()
        {
            return Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));
        }
        /// <summary>
        /// Setup initial data Before execution start
        /// </summary>
        public static void FrameworkInitSetup()
        {
            //Load initial data to framework
            GetInitialConfigData();
            scenariosList = new Dictionary<string, int>();
            //Create Test Run result folder for current run
            Reports.CreateTestsRunDierectory();


        }

        public static void GetInitialConfigData()
        {
            //Load initial data to framework
            //GetProperty.properties = GetProperty.GetInstance(PropertiesFilePath);
            //GetProperty.GetConfigDataFromProperty();
            GetXMLData.ConfigXMLData = GetXMLData.GetXMLConfigData(XMLFilePath);
            GetXMLData.GetConfigDataFromXML();

        }
        /// <summary>
        /// Initial report setup for current running  scenario 
        /// </summary>
        public void InitializeDataAndReport()
        {

            //Setup testCaseName, scenarioName and featureFileName
            testCaseName = TestContext.CurrentContext.Test.Name;
            scenarioName = ScenarioContext.Current.ScenarioInfo.Title;
            
            if (scenariosList.Count == 0)
            {
                scenariosList.Add(scenarioName, 1);
            }
            else
            {
                foreach (KeyValuePair<string, int> list in scenariosList)
                {
                    if (list.Key.ToLower() == scenarioName.ToLower())
                    {

                        int newValue = list.Value + 1;
                        scenariosList.Remove(scenarioName);
                        scenariosList.Add(scenarioName, newValue);
                        scenarioName = scenarioName + "_Iteration" + newValue;
                        break;

                    }
                    else
                    {
                        scenariosList.Add(scenarioName, 1);
                        break;
                    }
                }
            }
            
            featureFileName = FeatureContext.Current.FeatureInfo.Title;
            scenarioTag = ScenarioContext.Current.ScenarioInfo.Tags;
            if (scenarioTag.Length > 0)
            {
                testSetName = scenarioTag[0];
            }

            //Initialize report setup for current running  scenario
            if (Reports.IsRequiredReport) {
                Reports.InitializeReports();
            }
        }

        /// <summary>
        /// Initial report setup for current running  scenario 
        /// </summary>
        public static void InitializeErrorReport(Exception ex)
        {

            //Setup testCaseName, scenarioName and featureFileName
            testCaseName = "Pre execution Exception error";
            scenarioName = "Runtime exception";
            featureFileName = "";
            Reports.errorMessage = ex.Message + ex.StackTrace;
            //Initialize report setup for current running  scenario
            Reports.InitializeReports();
            Reports.ReportEvent("Error", Reporting.Status.Fail, true);
            Reports.SaveReports();
            Reports.SaveCustomizedHTMLReport();
            // Assert.Fail();
        }

    }
}

