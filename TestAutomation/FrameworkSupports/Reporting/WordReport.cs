using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace TestAutomation.FrameworkSupports.Reporting
{
    public static class WordReport
    {
        public static DocX resultdocument;
        public static DocX resultSummarydocument;
        public static int temp = 1;
        public static int tempsum = 1;
        public static string reportpath;
        public static int smrw = 1;

        /// <summary>
        /// Set report path
        /// </summary>
        public static void SetReportpath()
        {
            if (BaseUtilities.scenarioName.Length <= 115)
            {
                reportpath = Reports.testRunResultWordFolder + "\\" + BaseUtilities.scenarioName + ".docx";
            }
            else
            {
                reportpath = Reports.testRunResultWordFolder + "\\" + BaseUtilities.scenarioName.Substring(0, 115) + ".docx";
            }
        }
        /// <summary>
        /// Initialize Word Report --> Support Method
        /// </summary>
        public static void InitWordSummarySetup()
        {
            try
            {
                string directory = BaseUtilities.GetFolderPath();

                resultSummarydocument = DocX.Create(Reports.testRunResultWordFolder + "\\" + "Summary Report.docx");
                //Resultdocument = DocX.Create(directory + @"Results\WordResults\Testrr.docx");

                resultSummarydocument.AddHeaders();
                Header ResultHeader = resultSummarydocument.Headers.Odd;

                Image Headerimage = resultSummarydocument.AddImage(Reports.reportheaderimage);
                Picture picture = Headerimage.CreatePicture();
                picture.Width = 150;
                picture.Height = 28;
                Paragraph Header = ResultHeader.InsertParagraph();
                Header.Alignment = Alignment.right;
                Header.AppendPicture(picture);

                Paragraph Date = resultSummarydocument.InsertParagraph();
                String StrDate = System.DateTime.Today.ToString("dd-MM-yyyy");
                Paragraph date = Date.Append("Date: " + StrDate).FontSize(8).Font("Calibri").Color(System.Drawing.Color.FromArgb(1, 74, 73, 71));
                date.Alignment = Alignment.left;

                Paragraph title = resultSummarydocument.InsertParagraph().Append(Reports.reportname).FontSize(12).Font("Century Gothic").Color(System.Drawing.Color.FromArgb(1, 91, 168));
                title.Alignment = Alignment.center;
                title.AppendLine();
                Paragraph tcscountpass = resultSummarydocument.InsertParagraph().Append("Total Test(s) Passed: " + Reports.ttltcspass).Color(System.Drawing.Color.FromArgb(52, 168, 83));
                tcscountpass.Alignment = Alignment.left;
                Paragraph tcscountfail = resultSummarydocument.InsertParagraph().Append("Total Test(s) Failed: " + Reports.ttltcsfail).Color(System.Drawing.Color.FromArgb(234, 67, 53));
                tcscountfail.Alignment = Alignment.left;
                tcscountfail.AppendLine();

                String[] Tableheading = { "S.No", "Scenario Name", "Test Status", "For More Details" };
                String[] Tabledata = { BaseUtilities.scenarioName, BaseUtilities.scenarioStatus };

                Table table = resultSummarydocument.AddTable(1, 4);
                // Specify some properties for this Table.
                table.Alignment = Alignment.left;
                table.AutoFit = AutoFit.Contents;
                table.Design = TableDesign.TableGrid;
                table.SetColumnWidth(0, 667.87);
                table.SetColumnWidth(1, 5347.87);
                table.SetColumnWidth(2, 1255.87);
                table.SetColumnWidth(3, 1825.82);

                table.Rows[0].Cells[0].Paragraphs.First().Append(Tableheading[0]).Bold();
                table.Rows[0].Cells[1].Paragraphs.First().Append(Tableheading[1]).Bold();
                table.Rows[0].Cells[2].Paragraphs.First().Append(Tableheading[2]).Bold();
                table.Rows[0].Cells[3].Paragraphs.First().Append(Tableheading[3]).Bold();


                resultSummarydocument.InsertTable(table);

                resultSummarydocument.AddFooters();
                Footer footer_default = resultSummarydocument.Footers.Odd;
                Paragraph footer = footer_default.InsertParagraph();
                footer.Append(Reports.reportfooterName);

                resultSummarydocument.Save();
            }
            catch (Exception e)
            {
                Reports.SetupErrorLog(e);
            }
        }

        /// <summary>
        /// Initialize Word Report --> Support Method
        /// </summary>
        public static void InsertSummaryResults()
        {
            try
            {
                Table table = resultSummarydocument.AddTable(1, 4);
                // Specify some properties for this Table.
                table.Alignment = Alignment.left;
                table.AutoFit = AutoFit.Contents;
                table.Design = TableDesign.TableGrid;
                table.SetColumnWidth(0, 667.87);
                table.SetColumnWidth(1, 5347.87);
                table.SetColumnWidth(2, 1255.87);
                table.SetColumnWidth(3, 1825.82);

                table.Rows[0].Cells[0].Paragraphs.First().Append(smrw.ToString() + ".");
                smrw += 1;
                table.Rows[0].Cells[1].Paragraphs.First().Append(BaseUtilities.scenarioName);
                switch (BaseUtilities.scenarioStatus.ToLower())
                {
                    case "pass":

                        table.Rows[0].Cells[2].Paragraphs.First().Append("Pass").Color(System.Drawing.Color.FromArgb(52, 168, 83));
                        break;
                    case "fail":

                        table.Rows[0].Cells[2].Paragraphs.First().Append("Fail").Color(System.Drawing.Color.FromArgb(234, 67, 53));
                        break;
                    case "skip":

                        table.Rows[0].Cells[2].Paragraphs.First().Append("Skip").Color(System.Drawing.Color.FromArgb(234, 67, 53));
                        break;
                }


                Hyperlink link = resultSummarydocument.AddHyperlink("Click here", new Uri(reportpath));
                table.Rows[0].Cells[3].Paragraphs.First().AppendHyperlink(link).Color(System.Drawing.Color.FromArgb(1, 91, 168));

                resultSummarydocument.InsertTable(table);

                resultSummarydocument.Save();
            }
            catch (Exception e)
            {
                Reports.SetupErrorLog(e);
            }
        }

        /// <summary>
        /// Initialize Word Report --> Support Method
        /// </summary>
        public static void SummarizeTimeWord()
        {
            try
            {
                Paragraph last = resultSummarydocument.InsertParagraph();
                last.AppendLine();
                last.AppendLine();

                String[] Tableheading = { "Total Count", "Start Time", "End Time", "Total Time Taken" };
                String[] Tabledata = { (smrw-1).ToString(), Reports.smrystarttime, Reports.smryendtime, Reports.smrytimetaken };

                Table table = resultSummarydocument.AddTable(2, 4);
                // Specify some properties for this Table.
                table.Alignment = Alignment.left;
                table.AutoFit = AutoFit.Contents;
                table.Design = TableDesign.TableGrid;
                table.SetColumnWidth(0, 1335.74);
                table.SetColumnWidth(1, 2732.83);
                table.SetColumnWidth(2, 2732.84);
                table.SetColumnWidth(3, 2323.84);

                table.Rows[0].Cells[0].Paragraphs.First().Append(Tableheading[0]).Bold();
                table.Rows[0].Cells[1].Paragraphs.First().Append(Tableheading[1]).Bold();
                table.Rows[0].Cells[2].Paragraphs.First().Append(Tableheading[2]).Bold();
                table.Rows[0].Cells[3].Paragraphs.First().Append(Tableheading[3]).Bold();
                table.Rows[1].Cells[0].Paragraphs.First().Append(Tabledata[0]);
                table.Rows[1].Cells[1].Paragraphs.First().Append(Tabledata[1]);
                table.Rows[1].Cells[2].Paragraphs.First().Append(Tabledata[2]);
                table.Rows[1].Cells[3].Paragraphs.First().Append(Tabledata[3]);

                resultSummarydocument.InsertTable(table);


                resultSummarydocument.Save();
            }
            catch (Exception e)
            {
                Reports.SetupErrorLog(e);
            }
        }
        /// <summary>
        /// Initialize Word Report --> Support Method
        /// </summary>
        public static void InitSetupWord()
        {
            try
            {
                string directory = BaseUtilities.GetFolderPath();
                SetReportpath();
                resultdocument = DocX.Create(reportpath);
                //Resultdocument = DocX.Create(directory + @"Results\WordResults\Testrr.docx");

                resultdocument.AddHeaders();
                Header ResultHeader = resultdocument.Headers.Odd;

                Image Headerimage = resultdocument.AddImage(Reports.reportheaderimage);
                Picture picture = Headerimage.CreatePicture();
                picture.Width = 150;
                picture.Height = 28;
                Paragraph Header = ResultHeader.InsertParagraph();
                Header.Alignment = Alignment.right;
                Header.AppendPicture(picture);

                Paragraph Date = resultdocument.InsertParagraph();
                String StrDate = System.DateTime.Today.ToString("dd-MM-yyyy");
                Paragraph date = Date.Append("Date: " + StrDate).FontSize(8).Font("Calibri").Color(System.Drawing.Color.FromArgb(1, 74, 73, 71));
                date.Alignment = Alignment.left;

               
                
               
                Paragraph title = resultdocument.InsertParagraph().Append("Test Case Results").FontSize(20).Font("Century Gothic").Color(System.Drawing.Color.FromArgb(1, 91, 168));
                title.Alignment = Alignment.center;
                title.AppendLine();
               
                String[] Tableheading = { "Feature Name","ALM TestSet Name", "Scenario Name", "Start Time", "End Time", "Time Taken", "Test Status" };
                String[] Tabledata = { BaseUtilities.featureFileName, BaseUtilities.testSetName, BaseUtilities.scenarioName, Reports.starttime, Reports.endtime, Reports.timetaken, BaseUtilities.scenarioStatus };

                Table table = resultdocument.AddTable(Tableheading.Length, 2);
                // Specify some properties for this Table.
                table.Alignment = Alignment.center;
                table.AutoFit = AutoFit.Contents;
                table.Design = TableDesign.TableGrid;


                for (int row = 0; row < Tableheading.Length; row++)
                {

                    if (Tableheading[row].Trim().ToLower() == "test status")
                    {
                        switch (Tabledata[row].ToLower())
                        {
                            case "pass":

                                table.Rows[row].Cells[0].Paragraphs.First().Append(Tableheading[row]).Bold().Color(System.Drawing.Color.FromArgb(52, 168, 83));
                                table.Rows[row].Cells[1].Paragraphs.First().Append(Tabledata[row]).Color(System.Drawing.Color.FromArgb(52, 168, 83));
                                break;
                            case "fail":
                                table.Rows[row].Cells[0].Paragraphs.First().Append(Tableheading[row]).Bold().Color(System.Drawing.Color.FromArgb(234, 67, 53));
                                table.Rows[row].Cells[1].Paragraphs.First().Append(Tabledata[row]).Color(System.Drawing.Color.FromArgb(234, 67, 53));
                                break;
                            case "skip":
                                table.Rows[row].Cells[0].Paragraphs.First().Append(Tableheading[row]).Bold().Color(System.Drawing.Color.FromArgb(234, 67, 53));
                                table.Rows[row].Cells[1].Paragraphs.First().Append(Tabledata[row]).Color(System.Drawing.Color.FromArgb(234, 67, 53));
                                break;
                        }
                    }
                    else
                    {
                        table.Rows[row].Cells[0].Paragraphs.First().Append(Tableheading[row]).Bold();
                        table.Rows[row].Cells[1].Paragraphs.First().Append(Tabledata[row]);
                    }
                }

                resultdocument.InsertTable(table);
                Paragraph p1 = resultdocument.InsertParagraph();
                p1.AppendLine();
                p1.AppendLine();
                resultdocument.AddFooters();
                Footer footer_default = resultdocument.Footers.Odd;
                Paragraph footer = footer_default.InsertParagraph();
                footer.Append(Reports.reportfooterName);

                Paragraph last = resultdocument.InsertParagraph();

                for (int newline = 0; newline <= 36; newline++)
                {
                    last.AppendLine();
                }
                resultdocument.Save();
            }
            catch (Exception e)
            {
                Reports.SetupErrorLog(e);
            }
        }
        /// <summary>
        /// Insert Word Results for each step --> Support Method
        /// </summary>
        /// <param name="StepName">Name of the step</param>
        /// <param name="Stepdescription">Description of the step</param>
        /// <param name="Stepstatus">Status of the step</param>
        /// <param name="img">Step image</param>
        public static void InsertWordResults(string StepName, Status Stepstatus)
        {
            try
            {
                Paragraph title1 = resultdocument.InsertParagraph().Append("Step Name: " + StepName).FontSize(12).Font("Calibri").Color(System.Drawing.Color.FromArgb(0, 0, 0));

                Paragraph title2 = null;
                Paragraph title3 = null;

                switch (Stepstatus)
                {
                    case Status.Pass:
                        title2 = resultdocument.InsertParagraph().Append("Step Status: Passed").FontSize(12).Font("Calibri").Color(System.Drawing.Color.FromArgb(52, 168, 83));
                        break;
                    case Status.Warning:
                        title2 = resultdocument.InsertParagraph().Append("Step Status: Warning").FontSize(12).Font("Calibri").Color(System.Drawing.Color.FromArgb(251, 188, 5));
                        break;
                    case Status.Fail:
                        title2 = resultdocument.InsertParagraph().Append("Step Status: Failed").FontSize(12).Font("Calibri").Color(System.Drawing.Color.FromArgb(234, 67, 53));
                        title3 = resultdocument.InsertParagraph().Append("Error Message: " + Reports.errorMessage).FontSize(12).Font("Calibri").Color(System.Drawing.Color.FromArgb(234, 67, 53));
                        break;
                    case Status.Information:
                        title2 = resultdocument.InsertParagraph().Append("Step Status: Information").FontSize(12).Font("Calibri").Color(System.Drawing.Color.FromArgb(0, 0, 225));
                        break;
                    case Status.Error:
                        title2 = resultdocument.InsertParagraph().Append("Step Status: Error").FontSize(12).Font("Calibri").Color(System.Drawing.Color.FromArgb(234, 67, 53));
                        break;
                    case Status.Skip:
                        title2 = resultdocument.InsertParagraph().Append("Step Status: Skip").FontSize(12).Font("Calibri").Color(System.Drawing.Color.FromArgb(102, 225, 178));
                        break;
                }

                title1.Alignment = Alignment.left;
                Paragraph step = resultdocument.InsertParagraph();
                step.AppendLine();
                if (!string.IsNullOrEmpty(Reports.runtimescnshtpath)) {
                    string WordImg = Reports.getCurrentTestRunPath+Reports.runtimescnshtpath.Replace("..","");
                    Image image = resultdocument.AddImage(WordImg);

                    Picture picture = image.CreatePicture();
                    picture.Width = 600;
                    picture.Height = 330;
                    step.AppendPicture(picture);
                }
                if (temp > 1)
                {
                    for (int i = 0; i <= 4; i++)
                    {
                        step.AppendLine();
                        resultdocument.Save();
                    }
                    temp = 1;
                }
                else
                {
                    step.AppendLine();
                    resultdocument.Save();
                    temp = temp + 1;
                }

            }
            catch (Exception e)
            {
                Reports.SetupErrorLog(e);
            }
        }
    }
}
