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
    class ReportConversion : GeneralMethods
    {
        /// <summary>
        /// Element Locators
        /// </summary>
        private static By reportName = By.ClassName("report-name");
        private static By ttlPass = By.CssSelector("#charts-row > div:nth-child(1) > div > div:nth-child(3) > span > span");
        private static By ttlFail = By.CssSelector("#charts-row > div:nth-child(1) > div > div:nth-child(4) > span:nth-child(1)");
        private static By testCollection = By.Id("test-collection");
        private static By liTag = By.TagName("li");
        private static By testName = By.CssSelector("div.test-heading > span.test-name");
        private static By testStatus = By.CssSelector("div.test-heading > span:nth-child(3)");
        private static By testdesc = By.CssSelector("div.test-content.hide > div.test-desc");
        private static By resultimg = By.CssSelector("td.step-details img");
        private static By testStartTime = By.CssSelector("div.test-content.hide > div.test-time-info > span.label.start-time");
        private static By testEndTime = By.CssSelector("div.test-content.hide > div.test-time-info > span.label.end-time");
        private static By testTotalTime = By.CssSelector("div.test-content.hide > div.test-time-info > span:nth-child(3)");
        private static By testSteps = By.CssSelector("div.test-content.hide > div.test-steps");
        private static By trTag = By.TagName("tr");
        private static By testStepStatus = By.CssSelector("td.status.pass");
        private static By testStepDetails = By.CssSelector("td.step-details");
        private static By rptStartTime = By.CssSelector("#dashboard-view > div > div > div:nth-child(3) > div > div");
        private static By rptEndTime = By.CssSelector("#dashboard-view > div > div > div:nth-child(4) > div > div");
        private static By rptTtlExeTime = By.CssSelector("#dashboard-view > div > div > div:nth-child(5) > div > div");

        public static void ConvertToWordAndPdf()
        {
            try
            {
                string path = HTMLExtentReport.createhtmlfile;

                var driverhelper = new DriverScript();
                driverhelper.InitializeDriver();
                NavigateToURL("file:///" + HTMLExtentReport.createhtmlfile);
                jsexecutor.ExecuteScript("alert('Processing Report Conversion and ALM result updates. Please wait....');");
                WaitImplicit(3000);
                AcceptAlert();
                Reports.reportname = GetElementText(reportName);

                Reports.ttltcspass = GetElementText(ttlPass);
                Reports.ttltcsfail = GetElementText(ttlFail);
                if (Reports.isWordReport)
                {
                    WordReport.InitWordSummarySetup();
                }
                if (Reports.isPdfReport)
                {
                    PDFReport.InitSummarySetupPDF();
                }

                IList<IWebElement> liele = ReturnElement(testCollection).FindElements(liTag);

                foreach (IWebElement li in liele)
                {
                    BaseUtilities.scenarioName = li.FindElement(testName).Text;
                    BaseUtilities.scenarioStatus = li.FindElement(testStatus).Text;
                    Reports.starttime = GetHiddenElementText(li.FindElement(testStartTime));
                    Reports.endtime = GetHiddenElementText(li.FindElement(testEndTime));
                    Reports.timetaken = GetHiddenElementText(li.FindElement(testTotalTime));
                    string testDescription = GetHiddenElementText(li.FindElement(testdesc)).Trim();
                    string[] ftr = Regex.Split(testDescription, "</b>");
                    BaseUtilities.featureFileName = ftr[1].Trim().Substring(0, ftr[1].Trim().Length - 21);
                    BaseUtilities.testSetName = ftr[2].Trim();
                    if (Reports.isWordReport)
                    {
                        WordReport.InitSetupWord();
                        WordReport.InsertSummaryResults();
                    }
                    if (Reports.isPdfReport)
                    {
                        PDFReport.InitSetupPDF();
                        PDFReport.InsertSummaryResultPDF();

                    }

                    IList<IWebElement> trele = li.FindElement(testSteps).FindElement(By.TagName("tbody")).FindElements(trTag);
                    int i = 1;
                    foreach (IWebElement tr in trele)
                    {
                        Status stepStatus = Reports.GetEnumStatus(tr.GetAttribute("status"));

                        Reports.runtimescnshtpath = tr.FindElement(resultimg).GetAttribute("data-src");
                        string[] teststepDtl = Regex.Split(GetHiddenElementText(tr.FindElement(testStepDetails)).Trim(), "<br><b>");
                        string stepDetails = teststepDtl[0];
                        if (teststepDtl.Length >= 3)
                        {
                            Reports.errorMessage = teststepDtl[1].Trim().Substring(19);

                        }
                        if (Reports.isWordReport)
                        {
                            WordReport.InsertWordResults(stepDetails, stepStatus);
                        }
                        if (Reports.isPdfReport)
                        {
                            PDFReport.InsertPDFResults(stepDetails, stepStatus);
                        }

                        Reports.runtimescnshtpath = null;
                        Reports.errorMessage = null;

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
                Reports.smrystarttime = GetHiddenElementText(ReturnElement(rptStartTime));
                Reports.smryendtime = GetHiddenElementText(ReturnElement(rptEndTime));
                Reports.smrytimetaken = GetHiddenElementText(ReturnElement(rptTtlExeTime));
                if (Reports.isWordReport)
                {
                    WordReport.SummarizeTimeWord();
                }
                if (Reports.isPdfReport)
                {
                    PDFReport.SummarizeTimeResultPDF();
                    PDFReport.SummaryPDFdocument.Close();
                }

                jsexecutor.ExecuteScript("alert('Processing completed for Report Conversion and ALM result updates.');");
                WaitImplicit(2000);
                AcceptAlert();
                CloseBrowserAndDispose();
            }
            finally
            {
                if (Reports.isPdfReport)
                {
                    PDFReport.PDFdocument.Close();
                    PDFReport.SummaryPDFdocument.Close();
                }
                CloseBrowserAndDispose();
            }

        }
        public static void ConvertHTMLReportAndALMUpdates()
        {

            string path = HTMLExtentReport.createhtmlfile;
            List<String> list = new List<String>();
            List<String> wordlist = new List<String>();
            var filestream = new FileStream(path, FileMode.Open, FileAccess.Read);
            using (var StreamReader = new StreamReader(filestream, Encoding.UTF8))
            {
                String ln;
                while ((ln = StreamReader.ReadLine()) != null)
                {
                    if (ln.Length > 0)
                    {
                        list.Add(ln);
                    }
                }

            }
            int rr = 0;


            for (int i = rr; i < list.Count; i++)
            {
                if (list[i].Trim().Equals("<ul id='test-collection' class='test-collection'>"))
                {
                    int js = i;
                    nextline:
                    for (int j = js; j < list.Count; j++)
                    {
                        if (list[j].Trim().StartsWith("<li"))
                        {
                            Boolean IsWordSet = true;
                            Boolean IsCurrentScenario = false;
                            for (int jj = j; jj < list.Count; jj++)
                            {

                                if (list[jj].Trim().Equals("</ul>"))
                                {
                                    break;
                                }
                                else
                                {
                                    if (IsWordSet)
                                    {

                                        if (list[jj].Trim().Contains("<span class='test-name'>"))
                                        {
                                            string scenario = list[jj].Trim().Substring(24, list[jj].Trim().Length - 31);
                                            if (scenario != BaseUtilities.scenarioName.Trim())
                                            {
                                                js = j + 1;
                                                goto nextline;
                                            }
                                            else
                                            {
                                                IsCurrentScenario = true;
                                            }

                                        }
                                        if (list[jj].Trim().Contains("<span class='test-status"))
                                        {
                                            BaseUtilities.scenarioStatus = list[jj].Trim().Substring(list[jj].Trim().Length - 11, 4);
                                        }
                                        if (list[jj].Trim().Contains("<span class='label start-time'>"))
                                        {
                                            string[] start = Regex.Split(list[jj].Trim(), "'>");
                                            Reports.starttime = start[1].Trim().Replace("</span>", "");

                                        }
                                        if (list[jj].Trim().Contains("<span class='label end-time'>"))
                                        {
                                            string[] end = Regex.Split(list[jj].Trim(), "'>");
                                            Reports.endtime = end[1].Trim().Replace("</span>", "");
                                        }
                                        if (list[jj].Trim().Contains("<span class='label time-taken"))
                                        {
                                            string[] ttl = Regex.Split(list[jj].Trim(), "'>");
                                            Reports.timetaken = ttl[1].Trim().Replace("</span>", "");
                                        }
                                        if (list[jj].Trim().Contains("<div class='test-desc'>"))
                                        {
                                            string[] ftr = Regex.Split(list[jj].Trim().Substring(44, list[jj].Trim().Length - 50).Trim(), "</b>");
                                            BaseUtilities.featureFileName = ftr[0].Trim().Substring(0, ftr[0].Trim().Length - 21);
                                            BaseUtilities.testSetName = ftr[1].Trim();

                                        }

                                        if (IsCurrentScenario && BaseUtilities.featureFileName != null && BaseUtilities.scenarioStatus != null && BaseUtilities.scenarioName != null && Reports.starttime != null && Reports.endtime != null && Reports.timetaken != null)
                                        {
                                            if (Reports.isWordReport)
                                            {
                                                WordReport.InitSetupWord();

                                            }
                                            if (Reports.isPdfReport)
                                            {
                                                PDFReport.InitSetupPDF();


                                            }
                                            IsWordSet = false;

                                        }
                                    }

                                    if (list[jj].Trim().Equals("<tbody>"))
                                    {
                                        int trr = jj;
                                        nexttr:
                                        for (int k = trr; k < list.Count; k++)
                                        {

                                            if (list[k].Trim().Contains("<tr class='log'"))
                                            {
                                                string stepStatus = list[k + 1].Trim().Substring(18, 4);
                                                string sstatus = list[k + 3].Trim().Substring(25);
                                                string[] stepName = Regex.Split(sstatus, "</br><b>");
                                                if (stepStatus == "fail")
                                                {
                                                    Boolean istrend = false;
                                                    string error = stepName[1].Trim().Substring(19);
                                                    if (stepName.Length > 2)
                                                    {
                                                        string[] errimasrc = Regex.Split(stepName[2], "data-src=");
                                                        Reports.runtimescnshtpath = errimasrc[1].Trim().Substring(1, errimasrc[1].Trim().Length - 9);
                                                    }

                                                    for (int s = k + 4; s < list.Count; s++)
                                                    {
                                                        if (list[s].Trim().Contains("</br><b>"))
                                                        {
                                                            string[] errorimg = Regex.Split(list[s].Trim(), "</br><b>");
                                                            error += " " + errorimg[0].Trim();
                                                            string[] errimasrc = Regex.Split(errorimg[1], "data-src=");
                                                            Reports.runtimescnshtpath = errimasrc[1].Trim().Substring(1, errimasrc[1].Trim().Length - 9);
                                                        }
                                                        else if (list[s].Trim().Equals("</tr>"))
                                                        {
                                                            trr = s;
                                                            istrend = true;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            error += " " + list[s].Trim();
                                                        }

                                                    }
                                                    if (istrend)
                                                    {
                                                        Reports.errorMessage = error;
                                                        if (Reports.isWordReport)
                                                        {
                                                            WordReport.InsertWordResults(stepName[0], Reports.GetEnumStatus(stepStatus));
                                                        }
                                                        if (Reports.isPdfReport)
                                                        {
                                                            PDFReport.InsertPDFResults(stepName[0], Reports.GetEnumStatus(stepStatus));
                                                        }

                                                        Reports.errorMessage = null;
                                                        Reports.runtimescnshtpath = null;
                                                        trr = k + 1;
                                                        goto nexttr;
                                                    }

                                                }
                                                else
                                                {
                                                    if (stepName.Length > 1)
                                                    {
                                                        string[] imasrc = Regex.Split(stepName[1], "data-src=");
                                                        Reports.runtimescnshtpath = imasrc[1].Trim().Substring(1, imasrc[1].Trim().Length - 9);
                                                        if (Reports.isWordReport)
                                                        {
                                                            WordReport.InsertWordResults(stepName[0], Reports.GetEnumStatus(stepStatus));
                                                        }
                                                        if (Reports.isPdfReport)
                                                        {
                                                            PDFReport.InsertPDFResults(stepName[0], Reports.GetEnumStatus(stepStatus));
                                                        }

                                                        Reports.runtimescnshtpath = null;
                                                        trr = k + 1;
                                                        goto nexttr;
                                                    }

                                                    else
                                                    {
                                                        string addstatus = list[k + 4].Trim();
                                                        string[] addstepName = Regex.Split(addstatus, "</br><b>");
                                                        string[] imasrc = Regex.Split(addstepName[1], "data-src=");
                                                        Reports.runtimescnshtpath = imasrc[1].Trim().Substring(1, imasrc[1].Trim().Length - 9);
                                                        if (Reports.isWordReport)
                                                        {
                                                            WordReport.InsertWordResults(stepName[0] + addstepName[0], Reports.GetEnumStatus(stepStatus));
                                                        }
                                                        if (Reports.isPdfReport)
                                                        {
                                                            PDFReport.InsertPDFResults(stepName[0] + addstepName[0], Reports.GetEnumStatus(stepStatus));
                                                        }

                                                        Reports.runtimescnshtpath = null;
                                                        trr = k + 1;
                                                        goto nexttr;
                                                    }
                                                }
                                            }
                                            if (list[k].Trim().Equals("</tbody>"))
                                            {
                                                js = k + 1;
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
                                                break;
                                            }
                                        }
                                    }

                                }

                            }
                        }


                    }
                }

            }
            if (Reports.isPdfReport)
            {
                PDFReport.PDFdocument.Close();
            }

        }

        public static void ConvertHTMLSummaryReport()
        {

            string path = HTMLExtentReport.createhtmlfile;
            List<String> list = new List<String>();
            List<String> wordlist = new List<String>();
            var filestream = new FileStream(path, FileMode.Open, FileAccess.Read);
            using (var StreamReader = new StreamReader(filestream, Encoding.UTF8))
            {
                String ln;
                while ((ln = StreamReader.ReadLine()) != null)
                {
                    if (ln.Length > 0)
                    {
                        list.Add(ln);
                    }
                }

            }
            int rr = 0;

            next:
            for (int i = rr; i < list.Count; i++)
            {
                if (list[i].Trim().Contains("<span class='report-name'>"))
                {
                    string[] reportname = Regex.Split(list[i].Trim(), "'>");
                    Reports.reportname = reportname[1].Trim().Replace("</span>", "");

                }
                if (list[i].Trim().Equals("<div id='test-view-charts' class='subview-full'>"))
                {
                    string[] passcount = Regex.Split(list[i + 9].Trim(), "'>");
                    Reports.ttltcspass = passcount[1].Trim().Replace("</span> test(s) passed</span>", "");
                    string[] failcount = Regex.Split(list[i + 12].Trim(), "'>");
                    Reports.ttltcsfail = failcount[1].Trim().Replace("</span> test(s) failed, <span class='strong", "");
                    if (Reports.isWordReport)
                    {
                        WordReport.InitWordSummarySetup();
                    }
                    if (Reports.isPdfReport)
                    {
                        PDFReport.InitSummarySetupPDF();
                    }

                    rr = i + 13;
                    goto next;
                }
                if (list[i].Trim().Equals("<ul id='test-collection' class='test-collection'>"))
                {
                    int js = i;
                    nextline:
                    for (int j = js; j < list.Count; j++)
                    {
                        if (list[j].Trim().StartsWith("<li"))
                        {
                            Boolean IsWordSet = true;
                            for (int jj = j; jj < list.Count; jj++)
                            {

                                if (list[jj].Trim().Equals("</ul>"))
                                {
                                    break;
                                }
                                else
                                {
                                    if (IsWordSet)
                                    {

                                        if (list[jj].Trim().Contains("<span class='test-name'>"))
                                        {
                                            BaseUtilities.scenarioName = list[jj].Trim().Substring(24, list[jj].Trim().Length - 31);
                                        }
                                        if (list[jj].Trim().Contains("<span class='test-status"))
                                        {
                                            BaseUtilities.scenarioStatus = list[jj].Trim().Substring(list[jj].Trim().Length - 11, 4);
                                        }


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
                                            IsWordSet = false;
                                            BaseUtilities.scenarioStatus = null;
                                            BaseUtilities.scenarioName = null;
                                            js = j + 1;
                                            goto nextline;
                                        }
                                    }



                                }

                            }
                        }


                    }
                }
                if (list[i].Trim().StartsWith("<div class='panel-lead'>"))
                {
                    string[] reportsttime = Regex.Split(list[i].Trim(), "'>");
                    Reports.smrystarttime = reportsttime[1].Trim().Replace("</div>", "");
                    string[] reportendtime = Regex.Split(list[i + 6].Trim(), "'>");
                    Reports.smryendtime = reportendtime[1].Trim().Replace("</div>", "");
                    string[] reporttimetaken = Regex.Split(list[i + 12].Trim(), "'>");
                    Reports.smrytimetaken = reporttimetaken[1].Trim().Replace("</div>", "");
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
                    break;
                }
            }
            if (Reports.isPdfReport)
            {
                PDFReport.SummaryPDFdocument.Close();
            }

        }
    }
}
