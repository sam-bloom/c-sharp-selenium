using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAutomation.FrameworkSupports.Reporting
{
    public static class PDFReport
    {
        public static Document PDFdocument;
        public static Document SummaryPDFdocument;
        public static int Temp = 1;
        public static int smrw = 1;
        public static string reportpath;

        /// <summary>
        /// Set report path
        /// </summary>
        public static void SetReportpath()
        {
            if (BaseUtilities.scenarioName.Length <= 115)
            {
                reportpath = Reports.testRunResultPDFFolder + "\\" + BaseUtilities.scenarioName + ".pdf";
            }
            else
            {
                reportpath = Reports.testRunResultPDFFolder + "\\" + BaseUtilities.scenarioName.Substring(0, 115) + ".pdf";
            }
        }
        /// <summary>
        /// Initialize PDF Report --> Support Method
        /// </summary>
        public static void InitSummarySetupPDF()
        {
            try
            {
                string directory = BaseUtilities.GetFolderPath();
                SummaryPDFdocument = new Document(PageSize.A4, 50, 50, 25, 25);

                var output = new FileStream(Reports.testRunResultPDFFolder + "\\Summary Report.pdf", FileMode.Create);
                var writer = PdfWriter.GetInstance(SummaryPDFdocument, output);
                SummaryPDFdocument.Open();

                Font date = new Font(FontFactory.GetFont(BaseFont.COURIER, 8, new BaseColor(1, 74, 73)));
                Font heading = new Font(FontFactory.GetFont(BaseFont.COURIER, 15, new BaseColor(1, 91, 168)));
                Font subTitleFont = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, Font.BOLD));
                Font tabledata = new Font(FontFactory.GetFont(BaseFont.COURIER, 11));
                Font passstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(52, 168, 83)));
                Font warningstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(251, 188, 5)));
                Font failstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(234, 67, 53)));
                Font subpassstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(52, 168, 83)));
                Font subwarningstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(251, 188, 5)));
                Font subfailstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(234, 67, 53)));


                var logo = iTextSharp.text.Image.GetInstance(Reports.reportheaderimage);
                logo.ScaleToFit(150, 28);
                logo.Alignment = Element.ALIGN_RIGHT;

                SummaryPDFdocument.Add(logo);

                String StrDate = System.DateTime.Today.ToString("dd-MM-yyyy");
                Paragraph emptyline = new Paragraph(" ", tabledata);
                emptyline.Alignment = Element.ALIGN_LEFT;
                Paragraph Pdate = new Paragraph("Date: " + StrDate, date);
                SummaryPDFdocument.Add(Pdate);

                Paragraph heading1 = new Paragraph(Reports.reportname, heading);
                heading1.Alignment = Element.ALIGN_CENTER;

                SummaryPDFdocument.Add(heading1);
                //SummaryPDFdocument.Add(Chunk.NEWLINE);
                Paragraph ttlctcspass = new Paragraph("Total Test(s) Passed: " + Reports.ttltcspass, passstep);
                ttlctcspass.Alignment = Element.ALIGN_LEFT;
                SummaryPDFdocument.Add(ttlctcspass);
                Paragraph ttlctcsfail = new Paragraph("Total Test(s) Failed: " + Reports.ttltcsfail, failstep);
                ttlctcsfail.Alignment = Element.ALIGN_LEFT;
                SummaryPDFdocument.Add(ttlctcsfail);
                SummaryPDFdocument.Add(emptyline);
                PdfPTable table = new PdfPTable(4);
                table.WidthPercentage = 100;
                int[] tblWidth = { 3, 18, 6, 8 };
                table.SetWidths(tblWidth);
                //table.HorizontalAlignment = Element.ALIGN_LEFT;
                Paragraph Sno = new Paragraph("S.No", subTitleFont);
                Paragraph ScenarioName = new Paragraph("Scenario Name", subTitleFont);
                Paragraph Teststatus = new Paragraph("Test Status", subTitleFont);
                Paragraph Fmi = new Paragraph("For More Details", subTitleFont);
                table.AddCell(Sno);

                table.AddCell(ScenarioName);
                table.AddCell(Teststatus);
                table.AddCell(Fmi);


                SummaryPDFdocument.Add(table);

                //SummaryPDFdocument.NewPage();
                //PDFReport.PDFdocument.Close();
            }
            catch (Exception e)
            {

                Reports.SetupErrorLog(e);
            }

        }

        /// <summary>
        /// Initialize PDF Report --> Support Method
        /// </summary>
        public static void InsertSummaryResultPDF()
        {
            try
            {

                Font date = new Font(FontFactory.GetFont(BaseFont.COURIER, 8, new BaseColor(1, 74, 73)));
                Font heading = new Font(FontFactory.GetFont(BaseFont.COURIER, 15, new BaseColor(1, 91, 168)));
                Font link = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(1, 91, 168)));
                Font subTitleFont = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, Font.BOLD));
                Font tabledata = new Font(FontFactory.GetFont(BaseFont.COURIER, 11));
                Font passstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(52, 168, 83)));
                Font warningstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(251, 188, 5)));
                Font failstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(234, 67, 53)));
                Font subpassstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(52, 168, 83)));
                Font subwarningstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(251, 188, 5)));
                Font subfailstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(234, 67, 53)));

                PdfPTable table = new PdfPTable(4);
                table.WidthPercentage = 100;
                int[] tblWidth = { 3, 18, 6, 8 };
                table.SetWidths(tblWidth);

                Paragraph Sno = new Paragraph(smrw.ToString(), tabledata);
                Paragraph ScenarioName = new Paragraph(BaseUtilities.scenarioName, tabledata);
                Anchor anchor = new Anchor("Click here", link);
                anchor.Reference = "file:///" + reportpath;
                //Paragraph Fmi = new Paragraph("Click here", link);

                table.AddCell(Sno);
                smrw += 1;
                table.AddCell(ScenarioName);
                switch (BaseUtilities.scenarioStatus.ToLower())
                {
                    case "pass":
                        Paragraph _Teststatus = new Paragraph("Pass", passstep);
                        table.AddCell(_Teststatus);
                        break;
                    case "fail":
                        Paragraph _Teststatusfl = new Paragraph("Fail", failstep);

                        table.AddCell(_Teststatusfl);
                        break;
                }

                table.AddCell(anchor);


                SummaryPDFdocument.Add(table);

                //SummaryPDFdocument.NewPage();

            }
            catch (Exception e)
            {
                Reports.SetupErrorLog(e);
            }

        }
        /// <summary>
        /// Initialize PDF Report --> Support Method
        /// </summary>
        public static void SummarizeTimeResultPDF()
        {
            try
            {

                Font date = new Font(FontFactory.GetFont(BaseFont.COURIER, 8, new BaseColor(1, 74, 73)));
                Font heading = new Font(FontFactory.GetFont(BaseFont.COURIER, 15, new BaseColor(1, 91, 168)));
                Font link = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(1, 91, 168)));
                Font subTitleFont = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, Font.BOLD));
                Font tabledata = new Font(FontFactory.GetFont(BaseFont.COURIER, 11));
                Font passstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(52, 168, 83)));
                Font warningstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(251, 188, 5)));
                Font failstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(234, 67, 53)));
                Font subpassstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(52, 168, 83)));
                Font subwarningstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(251, 188, 5)));
                Font subfailstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(234, 67, 53)));

                SummaryPDFdocument.Add(Chunk.NEWLINE);
                PdfPTable table = new PdfPTable(4);
                table.WidthPercentage = 100;
                int[] tblWidth = { 5, 9, 9, 7 };
                table.SetWidths(tblWidth);

                Paragraph ttl = new Paragraph("Total Count", subTitleFont);
                Paragraph sttime = new Paragraph("Start Time", subTitleFont);
                Paragraph endtime = new Paragraph("End Time", subTitleFont);
                Paragraph ttltime = new Paragraph("Time Taken", subTitleFont);
                table.AddCell(ttl);

                table.AddCell(sttime);
                table.AddCell(endtime);
                table.AddCell(ttltime);
                table.AddCell(new Paragraph((smrw - 1).ToString(), tabledata));
                table.AddCell(new Paragraph(Reports.smrystarttime, tabledata));
                table.AddCell(new Paragraph(Reports.smryendtime, tabledata));
                table.AddCell(new Paragraph(Reports.smrytimetaken, tabledata));

                SummaryPDFdocument.Add(table);


                //SummaryPDFdocument.NewPage();

            }
            catch (Exception e)
            {

                Reports.SetupErrorLog(e);

            }

        }
        /// <summary>
        /// Initialize PDF Report --> Support Method
        /// </summary>
        public static void InitSetupPDF()
        {
            try
            {
                string directory = BaseUtilities.GetFolderPath();
                PDFdocument = new Document(PageSize.A4, 50, 50, 25, 25);
                SetReportpath();
                var output = new FileStream(reportpath, FileMode.Create);
                var writer = PdfWriter.GetInstance(PDFdocument, output);
                PDFdocument.Open();

                Font date = new Font(FontFactory.GetFont(BaseFont.COURIER, 8, new BaseColor(1, 74, 73)));
                Font heading = new Font(FontFactory.GetFont(BaseFont.COURIER, 20, new BaseColor(1, 91, 168)));
                Font subTitleFont = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, Font.BOLD));
                Font tabledata = new Font(FontFactory.GetFont(BaseFont.COURIER, 11));
                Font passstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(52, 168, 83)));
                Font warningstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(251, 188, 5)));
                Font failstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(234, 67, 53)));
                Font subpassstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(52, 168, 83)));
                Font subwarningstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(251, 188, 5)));
                Font subfailstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 11, new BaseColor(234, 67, 53)));

                var logo = iTextSharp.text.Image.GetInstance(Reports.reportheaderimage);
                logo.ScaleToFit(150, 28);
                logo.Alignment = Element.ALIGN_RIGHT;

                PDFdocument.Add(logo);

                String StrDate = System.DateTime.Today.ToString("dd-MM-yyyy");

                Paragraph Pdate = new Paragraph("Date: " + StrDate, date);
                PDFdocument.Add(Pdate);
                Paragraph emptyline = new Paragraph("  ", tabledata);
                emptyline.Alignment = Element.ALIGN_CENTER;
                Paragraph heading1 = new Paragraph("Test Case Results", heading);
                heading1.Alignment = Element.ALIGN_CENTER;
                PDFdocument.Add(Chunk.NEWLINE);
                PDFdocument.Add(heading1);
                PDFdocument.Add(emptyline);
                PdfPTable table = new PdfPTable(2);
                int[] tblWidth = { 9, 20 };
                table.SetWidths(tblWidth);
                Paragraph FeatureName = new Paragraph("Feature Name", subTitleFont);
                Paragraph TestSetName = new Paragraph("ALM TestSet Name", subTitleFont);
                Paragraph ScenarioName = new Paragraph("Scenario Name", subTitleFont);
                Paragraph Starttime = new Paragraph("Start Time", subTitleFont);
                Paragraph Endtime = new Paragraph("End Time", subTitleFont);
                Paragraph Timetaken = new Paragraph("Time Taken", subTitleFont);



                Paragraph _FeatureName = new Paragraph(BaseUtilities.featureFileName, tabledata);
                Paragraph _TestSetName = new Paragraph(BaseUtilities.testSetName, tabledata);
                Paragraph _ScenarioName = new Paragraph(BaseUtilities.scenarioName, tabledata);
                Paragraph _Starttime = new Paragraph(Reports.starttime, tabledata); ;
                Paragraph _Endtime = new Paragraph(Reports.endtime, tabledata);
                Paragraph _Timetaken = new Paragraph(Reports.timetaken, tabledata);



                table.AddCell(FeatureName);
                table.AddCell(_FeatureName);
                table.AddCell(TestSetName);
                table.AddCell(_TestSetName);

                table.AddCell(ScenarioName);
                table.AddCell(_ScenarioName);

                table.AddCell(Starttime);
                table.AddCell(_Starttime);

                table.AddCell(Endtime);
                table.AddCell(_Endtime);

                table.AddCell(Timetaken);
                table.AddCell(_Timetaken);
                switch (BaseUtilities.scenarioStatus.ToLower())
                {
                    case "pass":
                        Paragraph Teststatus = new Paragraph("Test Status", passstep);
                        Paragraph _Teststatus = new Paragraph("Pass", passstep);
                        table.AddCell(Teststatus);
                        table.AddCell(_Teststatus);
                        break;
                    case "fail":
                        Paragraph Teststatusfl = new Paragraph("Test Status", failstep);
                        Paragraph _Teststatusfl = new Paragraph("Fail", failstep);
                        table.AddCell(Teststatusfl);
                        table.AddCell(_Teststatusfl);
                        break;
                }



                PDFdocument.Add(table);

                PDFdocument.NewPage();
                //PDFReport.PDFdocument.Close();
            }
            catch (Exception e)
            {

                Reports.SetupErrorLog(e);
            }

        }

        /// <summary>
        /// Insert PDF Results for each step --> Support Method
        /// </summary>
        /// <param name="StepName">Name of the step</param>
        /// <param name="Stepdescription">Description of the step</param>
        /// <param name="Stepstatus">Status of the step</param>
        /// <param name="img">Step image</param>
        public static void InsertPDFResults(string StepName, Status Stepstatus)
        {
            try
            {


                Font passstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 9, new BaseColor(52, 168, 83)));
                Font warningstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 9, new BaseColor(251, 188, 5)));
                Font failstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 9, new BaseColor(234, 67, 53)));
                Font step = new Font(FontFactory.GetFont(BaseFont.COURIER, 9, new BaseColor(0, 0, 0)));
                Font infostep = new Font(FontFactory.GetFont(BaseFont.COURIER, 9, new BaseColor(0, 0, 225)));
                Font skipstep = new Font(FontFactory.GetFont(BaseFont.COURIER, 9, new BaseColor(102, 225, 178)));

                Paragraph Stepname = new Paragraph("Step Name: " + StepName, step);
                Stepname.Alignment = Element.ALIGN_LEFT;
                PDFdocument.Add(Stepname);

                Paragraph PStepStatus = null;
                Paragraph errormsg = null;


                switch (Stepstatus)
                {
                    case Status.Pass:
                        PStepStatus = new Paragraph("Step Status: Passed", passstep);
                        PStepStatus.Alignment = Element.ALIGN_LEFT;
                        PDFdocument.Add(PStepStatus);
                        PDFdocument.Add(Chunk.NEWLINE);
                        break;
                    case Status.Warning:
                        PStepStatus = new Paragraph("Step Status: Warning", warningstep);
                        PStepStatus.Alignment = Element.ALIGN_LEFT;
                        PDFdocument.Add(PStepStatus);
                        PDFdocument.Add(Chunk.NEWLINE);
                        break;
                    case Status.Fail:
                        PStepStatus = new Paragraph("Step Status: Failed", failstep);
                        PStepStatus.Alignment = Element.ALIGN_LEFT;
                        PDFdocument.Add(PStepStatus);
                        errormsg = new Paragraph("Error Message: " + Reports.errorMessage, failstep);
                        errormsg.Alignment = Element.ALIGN_LEFT;
                        PDFdocument.Add(errormsg);
                        PDFdocument.NewPage();
                        break;
                    case Status.Information:
                        PStepStatus = new Paragraph("Step Status: Information", infostep);
                        PStepStatus.Alignment = Element.ALIGN_LEFT;
                        PDFdocument.Add(PStepStatus);
                        PDFdocument.Add(Chunk.NEWLINE);
                        break;
                    case Status.Skip:
                        PStepStatus = new Paragraph("Step Status: Information", skipstep);
                        PStepStatus.Alignment = Element.ALIGN_LEFT;
                        PDFdocument.Add(PStepStatus);
                        PDFdocument.Add(Chunk.NEWLINE);
                        break;
                    case Status.Error:
                        PStepStatus = new Paragraph("Step Status: Information", failstep);
                        PStepStatus.Alignment = Element.ALIGN_LEFT;
                        PDFdocument.Add(PStepStatus);
                        PDFdocument.Add(Chunk.NEWLINE);
                        break;
                }


                //var images = iTextSharp.text.Image.GetInstance(BaseUtilities.scnshtpath + BaseUtilities.scnshtdatetime + ".png");
                if (!string.IsNullOrEmpty(Reports.runtimescnshtpath))
                {
                    string PDFImg = Reports.getCurrentTestRunPath + Reports.runtimescnshtpath.Replace("..", "");
                    var images = iTextSharp.text.Image.GetInstance(PDFImg);
                    images.ScaleToFit(590, 300);
                    images.Alignment = Element.ALIGN_CENTER;
                    PDFdocument.Add(images);
                    PDFdocument.Add(Chunk.NEWLINE);
                }
                PDFdocument.Add(Chunk.NEWLINE);
                if (Temp > 1)
                {
                    for (int i = 0; i <= 5; i++)
                    {
                        PDFdocument.Add(Chunk.NEWLINE);
                    }
                    Temp = 1;
                }

                else
                {
                    PDFdocument.Add(Chunk.NEWLINE);
                    Temp = Temp + 1;
                }


            }
            catch (Exception e)
            {

                Reports.SetupErrorLog(e);
            }

        }

    }
}
