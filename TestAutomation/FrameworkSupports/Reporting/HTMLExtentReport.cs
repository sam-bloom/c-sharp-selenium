using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
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
    public static class HTMLExtentReport
    {
        public static ExtentTest test;
        public static ExtentReports extent;
        public static String createhtmlfile;
        public static Boolean isAdded = true;
        //public String categories, author;
        /// <summary>
        /// Initiate HTML initial setup
        /// </summary>
        public static void StartResult()
        {

            String htmlfile = "Summary Report" + ".html";
            //FileStream fs = new FileStream(htmlfile, FileMode.Create);
            //fs.Close();
            createhtmlfile = Reports.testRunResultHTMLFolder + "\\" + htmlfile;
            var resulthtmlfile = new FileStream(createhtmlfile, FileMode.Create);

            var htmlReporter = new ExtentHtmlReporter(createhtmlfile);
            htmlReporter.LoadConfig(BaseUtilities.GetFolderPath() + @"\Config\extent-config.xml");

            // create ExtentReports and attach reporter(s)
            extent = new ExtentReports();
            extent.AttachReporter(htmlReporter);
            extent.AddSystemInfo("Environment", "SysTest");



        }
        /// <summary>
        /// Insert Status and description
        /// </summary>
        /// <param name="Description"></param>
        /// <param name="Status"></param>
        public static void InsertResultStep(String description, Status status,Boolean isScreenshot)
        {
            if (isScreenshot) {
                string[] ImageCustom = Regex.Split(Reports.runtimescnshtpath, "Screenshots");
                string ImgPath = "..\\Screenshots" + ImageCustom[1];
                String screenshotImage = "<img data-featherlight='" + ImgPath + "' class='step-img' src='" + ImgPath + "' data-src='" + ImgPath + "'/>";
                switch (status)
                {
                    case Status.Pass:
                        test.Pass(description + "</br><b>Screenshot:</b>" + screenshotImage);
                        break;
                    case Status.Fail:
                        test.Fail(description + "</br><b>Error Message: </b>" + Reports.errorMessage + "</br><b>Screenshot:</b>" + screenshotImage);
                        break;
                    case Status.Warning:
                        test.Warning(description + "</br><b>Screenshot:</b>" + screenshotImage);
                        break;
                    case Status.Information:
                        test.Info(description + "</br><b>Screenshot:</b>" + screenshotImage);
                        break;

                    case Status.Skip:
                        test.Skip(description + "</br><b>Screenshot:</b>" + screenshotImage);
                        break;

                    case Status.Error:
                        test.Error(description + "</br><b>Screenshot:</b>" + screenshotImage);
                        break;


                }
            }
            else
            {
                switch (status)
                {
                    case Status.Pass:
                        test.Pass(description );
                        break;
                    case Status.Fail:
                        test.Fail(description + "</br><b>Error Message: </b>" + Reports.errorMessage );
                        break;
                    case Status.Warning:
                        test.Warning(description);
                        break;
                    case Status.Information:
                        test.Info(description);
                        break;
                    case Status.Skip:
                        test.Skip(description);
                        break;
                    case Status.Error:
                        test.Error(description);
                        break;
                }
            }
        }
        /// <summary>
        /// Set TestCase Name and Description
        /// </summary>
        /// <param name="TestcaseName"></param>
        /// <param name="TestDescription"></param>
        public static void StartTestCases(String TestcaseName, String TestDescription)
        {
            test = extent.CreateTest(TestcaseName, TestDescription);
           
        }
        /// <summary>
        /// Flush the HTML report
        /// </summary>
        public static void EndResult()
        {
            extent.Flush();
        }

        /// <summary>
        /// Customize html final report
        /// </summary>
        public static void CustomizeHTMLReport()
        {

            var list = new List<String>();
            var filestream = new FileStream(HTMLExtentReport.createhtmlfile, FileMode.Open, FileAccess.Read);
            using (var StreamReader = new StreamReader(filestream, Encoding.UTF8))
            {
                String ln;
                while ((ln = StreamReader.ReadLine()) != null)
                {
                    //customie logo and other stuff
                    if (ln.Trim().Equals("<a href=\"http://extentreports.relevantcodes.com\" class=\"brand-logo blue darken-3\">Extent</a>"))
                    {

                        list.Add("        <a target='_blank' href=\"https://one.cba\" ><img src='" + Reports.htmlreportheaderimage + "'  height='95%'/></a>");
                    }
                    else if (ln.Trim().Contains("<span class=\"label blue darken-3\">"))
                    {
                    }

                    else
                    {
                        list.Add(ln);

                    }
                }
            }

            //Rewrite html report with customized content
            using (FileStream fs = new FileStream(HTMLExtentReport.createhtmlfile, FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    int sr = 0;
                next:
                    for (int i = sr; i < list.Count; i++)
                    {

                        if (list[i].Trim().Equals("<!-- search -->") && isAdded)
                        {
                            if (Reports.isWordReport || Reports.isPdfReport)
                            {
                                sw.WriteLine("<!-- Other Report types -->");
                                sw.WriteLine("<div class=\"chip transparent\">");
                                sw.WriteLine("<a class=\"dropdown-button tests-toggle\" href=\"#\" data-hover=\"false\" data-beloworigin=\"true\" data-constrainwidth=\"true\" data-activates=\"tests-toggleCus\">Other Report Types</a>");
                                sw.WriteLine("<ul id=\"tests-toggleCus\" class=\"dropdown-content\">");
                                if (Reports.isWordReport)
                                {
                                    sw.WriteLine("<li><a href=\"" + Reports.testRunResultWordFolder + "\\Summary Report.docx\">Word Report </a></li>");
                                }
                                if (Reports.isPdfReport)
                                {
                                    sw.WriteLine("<li><a href=\"" + Reports.testRunResultPDFFolder + "\\Summary Report.pdf\" target=\"_blank\">PDF Report </a></li>");
                                }
                                sw.WriteLine(" </ul>");
                                sw.WriteLine("</div>");
                                sw.WriteLine("<!-- Other Report types -->");
                            }
                            sw.WriteLine(list[i]);
                            sr = i + 1;
                            isAdded = false;
                            goto next;

                        }
                        else
                        {
                            sw.WriteLine(list[i]);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Customize html final report
        /// </summary>
        public static void CustomizeHTMLReportOld()
        {
            var list = new List<String>();
            var filestream = new FileStream(HTMLExtentReport.createhtmlfile, FileMode.Open, FileAccess.Read);
            using (var StreamReader = new StreamReader(filestream, Encoding.UTF8))
            {
                String ln;
                while ((ln = StreamReader.ReadLine()) != null)
                {
                    //customie logo and other stuff
                    if (ln.Trim().Equals("<a href=\"http://extentreports.relevantcodes.com\" class=\"brand-logo blue darken-3\">Extent</a>"))
                    {

                        list.Add("        <a target='_blank' href=\"https://one.cba\" ><img src='" + BaseUtilities.GetFolderPath() + "\\Framework_Supports\\Reporting\\Images\\HTMLlogo.png'  height='95%'/></a>");
                    }
                    else if (ln.Trim().Contains("<span class=\"label blue darken-3\">"))
                    {
                    }
                    else
                    {
                        list.Add(ln);

                    }
                }
            }

            //Rewrite html report with customized content
            using (FileStream fs = new FileStream(HTMLExtentReport.createhtmlfile, FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    int sr = 0;
                next:
                    for (int i = sr; i < list.Count; i++)
                    {
                        if (list[i].Trim().Equals("<!-- search -->") && isAdded)
                        {
                            if (Reports.isWordReport || Reports.isPdfReport)
                            {
                                sw.WriteLine("<!-- Custom Report Convertion -->");
                                sw.WriteLine("<div class=\"chip transparent\">");
                                sw.WriteLine("<a class=\"dropdown-button tests-toggle\" href=\"#\" data-hover=\"false\" data-beloworigin=\"true\" data-constrainwidth=\"true\" data-activates=\"tests-toggleCus\">Report Convertion</a>");
                                sw.WriteLine("<ul id=\"tests-toggleCus\" class=\"dropdown-content\">");
                                if (Reports.isWordReport)
                                {
                                    sw.WriteLine("<li><a data-toggle=\"modal\" data-target=\"#wordModal\">Convert to Word </a></li>");
                                }
                                if (Reports.isPdfReport)
                                {
                                    sw.WriteLine("<li><a data-toggle=\"modal\" data-target=\"#pdfModal\">Convert to Pdf </a></li>");
                                }
                                sw.WriteLine(" </ul>");
                                sw.WriteLine("<script src=\"https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js\"></script>");
                                sw.WriteLine("<script src=\"https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js\"></script>");
                                if (Reports.isWordReport)
                                {
                                    sw.WriteLine("<div class=\"modal fade\" id=\"wordModal\" role=\"dialog\" >");
                                    sw.WriteLine("<div class=\"modal-dialog\">");
                                    sw.WriteLine(" <div class=\"modal-content\"/>");
                                    sw.WriteLine("<div class=\"modal-header\">");
                                    sw.WriteLine("<h4 class=\"modal-title\">Report Convertion to Word</h4>");
                                    sw.WriteLine("</div>");
                                    sw.WriteLine("<div class=\"modal-body\">");
                                    sw.WriteLine("<input type=\"radio\"/>Convert report into Default Location </br>");
                                    sw.WriteLine("<button type=\"button\" class=\"btn btn-info btn-lg\" >Convert</button>");
                                    sw.WriteLine("<button type=\"button\" class=\"btn btn-info btn-lg\" data-dismiss=\"modal\">Close</button>");
                                    sw.WriteLine("</div>");
                                    sw.WriteLine("</div>");
                                    sw.WriteLine("</div>");
                                    sw.WriteLine("</div>");
                                }
                                if (Reports.isPdfReport)
                                {
                                    sw.WriteLine("<div class=\"modal fade\" id=\"pdfModal\" role=\"dialog\" >");
                                    sw.WriteLine("<div class=\"modal-dialog\">");
                                    sw.WriteLine("<div class=\"modal-content\"/>");
                                    sw.WriteLine("<div class=\"modal-header\">");
                                    sw.WriteLine("<h4 class=\"modal-title\">Report Convertion to PDF</h4>");
                                    sw.WriteLine("</div>");
                                    sw.WriteLine("<div class=\"modal-body\">");
                                    sw.WriteLine("<input type=\"radio\"/>Convert report into Default Location </br>");
                                    sw.WriteLine("<button type=\"button\" class=\"btn btn-info btn-lg\" >Convert</button>");
                                    sw.WriteLine("<button type=\"button\" class=\"btn btn-info btn-lg\" data-dismiss=\"modal\">Close</button>");
                                    sw.WriteLine("</div>");
                                    sw.WriteLine("</div>");
                                    sw.WriteLine("</div>");
                                    sw.WriteLine(" </div>");
                                }
                                sw.WriteLine("</div>");
                                sw.WriteLine("<!-- Custom Report Convertion -->");
                            }
                            sw.WriteLine(list[i]);
                            sr = i + 1;
                            isAdded = false;
                            goto next;

                        }
                        else
                        {
                            sw.WriteLine(list[i]);
                        }
                    }
                }
            }
        }
    }
}
