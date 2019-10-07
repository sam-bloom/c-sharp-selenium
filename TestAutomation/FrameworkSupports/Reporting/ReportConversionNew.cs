using HtmlAgilityPack;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace TestAutomation.FrameworkSupports.Reporting
{
    class ReportConversionNew : GeneralMethods
    {


        public string ReportName { get; set; }
        public string TotalTestsPassed { get; set; }
        public string TotalTestsFailed { get; set; }
        public string TestName { get; set; }
        public string TestStatus { get; set; }
        public string TestDescription { get; set; }
        public string TestStartTime { get; set; }
        public string TestEndTime { get; set; }
        public string TestTimeTaken { get; set; }
        public string TestSummaryStartTime { get; set; }
        public string TestSummaryEndTime { get; set; }
        public string TestSummaryTimeTaken { get; set; }
        public string StepStatus { get; set; }
        public string StepDetails { get; set; }
        public string StepImage { get; set; }


        HtmlDocument doc = new HtmlDocument();
        public ReportConversionNew(string HTMLPath)
        {
            doc.Load(HTMLPath);
        }
        private int GetTestResultCount()
        {
            int count = 0;
            count = doc.DocumentNode.SelectNodes("//ul[@id='test-collection']/li").Count;
            return count;
        }

        public void ConvertHTMLReportAndALMUpdates()
        {
            if (Reports.IsRequiredReport && (Reports.isWordReport || Reports.isPdfReport))
            {
                Reports.errorMessage = null;
                Reports.runtimescnshtpath = null;
                int count = GetTestResultCount();
                if (count > 0)
                {
                    for (int i = 1; i <= count; i++)
                    {
                        String rootPath = "//li[@test-id='" + i + "']";
                        TestName = GetTestName(rootPath);
                        if (TestName.Trim() == BaseUtilities.scenarioName.Trim())
                        {
                            BaseUtilities.scenarioName = TestName;
                            BaseUtilities.scenarioStatus = GetTestStatus(rootPath);
                            Reports.starttime = GetTestStartTime(rootPath);
                            Reports.endtime = GetTestEndTime(rootPath);
                            Reports.timetaken = GetTestTimeTaken(rootPath);
                            string[] TestDesc = Regex.Split(GetTestDescription(rootPath).Trim(), "</b>");
                            BaseUtilities.featureFileName = Regex.Split(TestDesc[1].Trim(), "<b>")[0];
                            BaseUtilities.testSetName = TestDesc[2].Trim();
                            if (Reports.isWordReport)
                            {
                                WordReport.InitSetupWord();

                            }
                            if (Reports.isPdfReport)
                            {
                                PDFReport.InitSetupPDF();


                            }
                            Dictionary<string, string> Steps = GetStepDetails(rootPath);
                            if (Steps.Count > 0)
                            {
                                foreach (KeyValuePair<string, string> list in Steps)
                                {
                                    var html = @"<!DOCTYPE html><html><body>" + list.Key + "</body></html> ";
                                    var htmlDoc = new HtmlDocument();
                                    htmlDoc.LoadHtml(html);
                                    string[] StepDetailsArray = Regex.Split(htmlDoc.DocumentNode.SelectSingleNode("//body").InnerText.Replace("Screenshot:", ""), "Error Message:");
                                    try
                                    {
                                        Reports.runtimescnshtpath = htmlDoc.DocumentNode.SelectSingleNode("//body//img").GetAttributeValue("src", "");
                                    }
                                    catch (Exception)
                                    {
                                        Reports.runtimescnshtpath = "";
                                    }
                                    if (StepDetailsArray.Length > 1)
                                    {
                                        Reports.errorMessage = StepDetailsArray[1];
                                    }
                                    if (Reports.isWordReport)
                                    {
                                        WordReport.InsertWordResults(StepDetailsArray[0], Reports.GetEnumStatus(list.Value));
                                    }
                                    if (Reports.isPdfReport)
                                    {
                                        PDFReport.InsertPDFResults(StepDetailsArray[0], Reports.GetEnumStatus(list.Value));
                                    }

                                    Reports.errorMessage = null;
                                    Reports.runtimescnshtpath = null;

                                }
                            }
                            if (Reports.isPdfReport)
                            {
                                PDFReport.PDFdocument.Close();
                            }
                            if (Reports.isReuiredALMUpdates)
                            {
                                ALMQCIntegration.UpdateTestRunsInALMQC();
                            }
                            BaseUtilities.featureFileName = null;
                            BaseUtilities.scenarioStatus = null;
                            BaseUtilities.testSetName = null;
                            Reports.starttime = null;
                            Reports.endtime = null;
                            Reports.timetaken = null;
                            BaseUtilities.scenarioName = null;
                        }



                    }
                }
            }
        }
        private string GetReportName()
        {
            string ReportName = "";
            ReportName = doc.DocumentNode.SelectSingleNode("//span[@class='report-name']").InnerText;
            return ReportName;
        }
        private string GetTotalTestsPassed()
        {
            string GetTotalTestsPassed = "";
            GetTotalTestsPassed = doc.DocumentNode.SelectSingleNode("(//div[@class='block text-small'])[1]").InnerText.Replace(" test(s) passed", "");
            return GetTotalTestsPassed;
        }
        private string GetTotalTestsFailed()
        {
            string GetTotalTestsPassed = "";
            GetTotalTestsPassed = doc.DocumentNode.SelectSingleNode("(//div[@class='block text-small'])[2]").InnerText.Replace(" test(s) failed, 0 others", "");
            return GetTotalTestsPassed;
        }
        private string GetTestName(string rootPath)
        {
            string TestName = "";
            TestName = doc.DocumentNode.SelectSingleNode(rootPath + "//span[@class='test-name']").InnerText;
            return TestName;
        }
        private string GetTestStatus(string rootPath)
        {
            string TestStatus = "";
            TestStatus = doc.DocumentNode.SelectSingleNode(rootPath).GetAttributeValue("status", "");
            return TestStatus;
        }
        private string GetTestDescription(string rootPath)
        {
            string TestDescription = "";
            TestDescription = doc.DocumentNode.SelectSingleNode(rootPath + "//div[@class='test-desc']").InnerHtml;
            return TestDescription;
        }
        private string GetTestStartTime(string rootPath)
        {
            string TestStartTime = "";
            TestStartTime = doc.DocumentNode.SelectSingleNode(rootPath + "//div[@class='test-time-info']/span[1]").InnerText;
            return TestStartTime;
        }
        private string GetTestEndTime(string rootPath)
        {
            string TestEndTime = "";
            TestEndTime = doc.DocumentNode.SelectSingleNode(rootPath + "//div[@class='test-time-info']/span[2]").InnerText;
            return TestEndTime;
        }
        private string GetTestTimeTaken(string rootPath)
        {
            string TestTimeTaken = "";
            TestTimeTaken = doc.DocumentNode.SelectSingleNode(rootPath + "//div[@class='test-time-info']/span[3]").InnerText;
            return TestTimeTaken;
        }
        private string GetTestSummaryStartTime()
        {
            string TestSummaryStartTime = "";
            TestSummaryStartTime = doc.DocumentNode.SelectSingleNode("(//div[@class='panel-lead'])[3]").InnerText;
            return TestSummaryStartTime;
        }
        private string GetTestSummaryEndTime()
        {
            string TestSummaryEndTime = "";
            TestSummaryEndTime = doc.DocumentNode.SelectSingleNode("(//div[@class='panel-lead'])[4]").InnerText;
            return TestSummaryEndTime;
        }
        private string GetTestSummaryTimeTaken()
        {
            string TestSummaryTimeTaken = "";
            TestSummaryTimeTaken = doc.DocumentNode.SelectSingleNode("(//div[@class='panel-lead'])[5]").InnerText;
            return TestSummaryTimeTaken;
        }
        private int TestStepsCount(string rootPath)
        {
            int count = 0;
            try
            {
                count = doc.DocumentNode.SelectNodes(rootPath + "//div[@class='test-steps']//tbody/tr").Count;
            }
            catch (Exception)
            {
                count = 0;
            }
            return count;
        }

        private Dictionary<string, string> GetStepDetails(string rootPath)
        {
            Dictionary<string, string> results = new Dictionary<string, string>();
            int trCount = TestStepsCount(rootPath);
            if (trCount > 0)
            {
                for (int i = 1; i <= trCount; i++)
                {
                    StepStatus = doc.DocumentNode.SelectSingleNode(rootPath + "//div[@class='test-steps']//tbody/tr[" + i + "]").GetAttributeValue("status", "");
                    StepDetails = doc.DocumentNode.SelectSingleNode(rootPath + "//div[@class='test-steps']//tbody/tr[" + i + "]/td[@class='step-details']").InnerHtml;
                    results.Add(StepDetails, StepStatus);

                }
            }
            return results;
        }

        public void ConvertHTMLSummaryReport()
        {
            if (Reports.IsRequiredReport && (Reports.isWordReport || Reports.isPdfReport))
            {

                Reports.reportname = GetReportName().Trim();
                Reports.ttltcspass = GetTotalTestsPassed().Trim();
                Reports.ttltcsfail = GetTotalTestsFailed().Trim();
                if (Reports.isWordReport)
                {
                    WordReport.InitWordSummarySetup();
                }
                if (Reports.isPdfReport)
                {
                    PDFReport.InitSummarySetupPDF();
                }
                int count = GetTestResultCount();
                if (count > 0)
                {
                    for (int i = 1; i <= count; i++)
                    {
                        String rootPath = "//li[@test-id='" + i + "']";
                        BaseUtilities.scenarioName = GetTestName(rootPath);
                        BaseUtilities.scenarioStatus = GetTestStatus(rootPath);
                        if (BaseUtilities.scenarioStatus != null && BaseUtilities.scenarioName != null)
                        {
                            if (Reports.isWordReport)
                            {
                                WordReport.SetReportpath();
                                WordReport.InsertSummaryResults();
                            }
                            if (Reports.isPdfReport)
                            {
                                PDFReport.SetReportpath();
                                PDFReport.InsertSummaryResultPDF();

                            }
                            BaseUtilities.scenarioStatus = null;
                            BaseUtilities.scenarioName = null;

                        }
                    }
                }
                Reports.smrystarttime = GetTestSummaryStartTime();
                Reports.smryendtime = GetTestSummaryEndTime();
                Reports.smrytimetaken = GetTestSummaryTimeTaken();
                if (Reports.isWordReport)
                {
                    WordReport.SummarizeTimeWord();
                }
                if (Reports.isPdfReport)
                {
                    PDFReport.SummarizeTimeResultPDF();
                    PDFReport.SummaryPDFdocument.Close();
                }

                Reports.smrystarttime = null;
                Reports.smryendtime = null;
                Reports.smrytimetaken = null;


            }
        }
    }
}
