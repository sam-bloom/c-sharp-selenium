using TestAutomation.FrameworkSupports.Reporting;
using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TechTalk.SpecFlow;

namespace TestAutomation.FrameworkSupports
{

    [Binding]
    public class SpecflowHookSetup
    {


        BaseUtilities baseutilities = new BaseUtilities();
        Reports reports = new Reports();
        GeneralMethods gm = new GeneralMethods();

        /// <summary>
        /// Run before test run and setup framework initial data/inputs
        /// </summary>
        [BeforeTestRun]
        public static void BeforeTestRun()
        {
            try
            {
                BaseUtilities.FrameworkInitSetup();

            }
            catch (Exception ex)
            {
                BaseUtilities.InitializeErrorReport(ex);
                //Reports.SetupErrorLog(ex);
            }
        }

        /// <summary>
        /// run before executing each scenario to create/insert report  values
        /// </summary>
        [BeforeScenario]
        public void BeforeScenarioWeb()
        {
            try
            {
                baseutilities.InitializeDataAndReport();

            }
            catch (Exception ex)
            {
                GeneralMethods.CloseBrowserAndDispose();
                gm.KillProcess(BaseUtilities.browser);
                Reports.SetupErrorLog(ex);
                Assert.Fail();
            }
        }

        /// <summary>
        /// It will run after executing each scenario to save reports
        /// </summary>
        [AfterScenario]
        public void AfterScenario()
        {
            try
            {
                Reports.SaveHTMLReport();
                ReportConversionNew Convert = new ReportConversionNew(HTMLExtentReport.createhtmlfile);
                Convert.ConvertHTMLReportAndALMUpdates();


            }
            finally
            {
                BaseUtilities.testCaseName = null;
                BaseUtilities.scenarioName = null;
                BaseUtilities.featureFileName = null;
                BaseUtilities.currentTestCase = null;
                GeneralMethods.CloseBrowserAndDispose();
                gm.KillProcess(BaseUtilities.browser);
            }
        }
        /// <summary>
        /// it will run after executing complete test to flush the customized HTML report
        /// </summary>
        [AfterTestRun]
        public static void AfterTestRun()
        {
            Reports.SaveCustomizedHTMLReport();
            ReportConversionNew Convert = new ReportConversionNew(HTMLExtentReport.createhtmlfile);
            Convert.ConvertHTMLSummaryReport();


        }

    }
}
