using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestAutomation.FrameworkSupports.Reporting
{
    class Reports : PropertiesCollection
    {

        public static string resultImgFolder;
        public static string testRunResultWordFolder;
        public static string testRunResultPDFFolder;
        public static string testRunResultHTMLFolder;
        public static string testRunResultImgFolder;
        public static string testRunResultImgFolderCustom;
        public static string scnshtpath;
        public static string runtimescnshtpath;
        public static string errorMessage;
        public static string getCurrentTestRunPath;
        public static string testRunResultOutputFolder;
        public static string testRunOutputFileName;
        public static string reportfooterName = "Automation Team";
        public static string reportname;
        public static string starttime;
        public static string endtime;
        public static string timetaken;
        public static string smrystarttime;
        public static string smryendtime;
        public static string smrytimetaken;
        public static Boolean isWordReport = false;
        public static Boolean isPdfReport = false;
        public static Boolean isHTMLReport = false;
        public static Boolean isReuiredALMUpdates = false;
        public static Boolean IsRequiredReport = false;
        public static string ttltcspass;
        public static string ttltcsfail;
        public static string reportheaderimage = BaseUtilities.GetFolderPath() + @"\FrameworkSupports\Reporting\Images\logo.png";
        public static string htmlreportheaderimage = BaseUtilities.GetFolderPath() + @"\FrameworkSupports\Reporting\Images\HTMLlogo.png";
        public static string emptyimage = BaseUtilities.GetFolderPath() + @"\FrameworkSupports\Reporting\Images\EmptyImage.png";
        /// <summary>
        /// Save screenshots into result Directory --> Support Method
        /// </summary>
        /// <param name="img">Image</param>
        public static void SaveScreenshot(System.Drawing.Image img)
        {
            scnshtpath = resultImgFolder + "\\";
            runtimescnshtpath = scnshtpath + BaseUtilities.GetCurrentTimeWithFormat() + ".png";
            img.Save(runtimescnshtpath);
        }

        public static void SetCurrentTestRunPath()
        {
            string directory = BaseUtilities.GetFolderPath();
            string datetime = BaseUtilities.GetCurrentTimeWithFormat();
            getCurrentTestRunPath = directory + @"\Results\Run_" + datetime;
        }
        /// <summary>
        /// Save web screenshots into result Directory --> Support Method
        /// </summary>
        /// <param name="img">ScreenShot</param>
        public static void SaveWebScreenshot(Screenshot img)
        {
            scnshtpath = resultImgFolder + "\\";
            runtimescnshtpath = scnshtpath + BaseUtilities.GetCurrentTimeWithFormat() + ".png";
            img.SaveAsFile(runtimescnshtpath);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        public static void TakeScreenshot(String path)
        {
            var bmpScreenShot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format32bppArgb);
            var gfxScreenshot = Graphics.FromImage(bmpScreenShot);
            gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
            bmpScreenShot.Save(path, ImageFormat.Png);
            /*
             //Another way to take screenshot
            Rectangle bounds = Screen.GetBounds(Point.Empty);
            using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height))
            {
                using (Graphics gp = Graphics.FromImage(bitmap))
                {
                    gp.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                }
                bitmap.Save(path, ImageFormat.Png);
            }
             */
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

        public static void Wrapup()
        {
            if (BaseUtilities.applicationType.ToLower() == "web")
            {
                if (IsDriverActive())
                {
                    driver.Quit();
                }

            }
            if (BaseUtilities.applicationType.ToLower() == "mobile")
            {

            }
            if (BaseUtilities.applicationType.ToLower() == "windows")
            {

            }
        }
        /// <summary>
        /// Create Result Directory for storing images --> Support Method
        /// </summary>
        public static void CreateTestRunImageFolder()
        {


            if (BaseUtilities.scenarioName.Length > 115)
            {
                BaseUtilities.scenarioName = BaseUtilities.scenarioName.Substring(0, 115);
            }
            DirectoryInfo di = Directory.CreateDirectory(resultImgFolder = testRunResultImgFolder + "\\" + BaseUtilities.scenarioName);
        }
        /// <summary>
        /// Initialize test reports
        /// </summary>
        public static void InitializeReports()
        {
            //Initial setup for Word and PDF Results

            if (isHTMLReport)
            {
                HTMLExtentReport.StartTestCases(BaseUtilities.scenarioName, "<b>Feature Name: </b> " + BaseUtilities.featureFileName + "<b> ALM TestSet Name: </b> " + BaseUtilities.testSetName);
            }
            //create image folder
            CreateTestRunImageFolder();
        }
        /// <summary>
        /// Save reports 
        /// </summary>
        public static void SaveReports()
        {
            SaveHTMLReport();
            SaveWordReport();
            SavePdfReport();
        }
        /// <summary>
        /// Save html report
        /// </summary>
        public static void SaveHTMLReport()
        {
            if (isHTMLReport)
            {
                HTMLExtentReport.EndResult();

            }
        }
        /// <summary>
        /// Save word report
        /// </summary>
        public static void SaveWordReport()
        {
            if (Reports.isWordReport)
            {

                WordReport.resultdocument.Save();

            }
        }
        /// <summary>
        /// Save pdf report
        /// </summary>
        public static void SavePdfReport()
        {
            if (Reports.isPdfReport)
            {

                PDFReport.PDFdocument.Close();
            }
        }
        /// <summary>
        /// Customize html report
        /// </summary>
        public static void SaveCustomizedHTMLReport()
        {
            if (Reports.isHTMLReport)
            {
                HTMLExtentReport.CustomizeHTMLReport();

            }
        }
        /// <summary>
        /// Create Result Directory for storing images --> Support Method
        /// </summary>
        public static void CreateTestsRunDierectory()
        {
            SetCurrentTestRunPath();
            testRunResultWordFolder = getCurrentTestRunPath + "\\WordReport";
            testRunResultPDFFolder = getCurrentTestRunPath + "\\PDFReport";
            testRunResultHTMLFolder = getCurrentTestRunPath + "\\HTMLReport";
            testRunResultImgFolder = getCurrentTestRunPath + "\\Screenshots";
            testRunResultOutputFolder = getCurrentTestRunPath + "\\OutputData";
            if (Reports.IsRequiredReport) {
                DirectoryInfo dirimg = Directory.CreateDirectory(testRunResultImgFolder);
            }

            if (Reports.IsRequiredReport && isWordReport)
            {

                DirectoryInfo dirword = Directory.CreateDirectory(testRunResultWordFolder);


            }
            if (Reports.IsRequiredReport && isPdfReport)
            {

                DirectoryInfo dirpdf = Directory.CreateDirectory(testRunResultPDFFolder);

            }
            if (isHTMLReport)
            {
                DirectoryInfo dirhtml = Directory.CreateDirectory(testRunResultHTMLFolder);
                HTMLExtentReport.StartResult();
            }

        }
        /// <summary>
        /// 
        /// </summary>
        public static void CreateOutputData()
        {
            //Copy user imput file and past it into runtime outputfile
            DirectoryInfo dirop = Directory.CreateDirectory(testRunResultOutputFolder);
            // Use Path class to manipulate file and directory paths.
            string sourceFile = System.IO.Path.Combine(BaseUtilities.inputDataFilePath, BaseUtilities.inputDataFileName);
            string destFile = System.IO.Path.Combine(testRunResultOutputFolder, BaseUtilities.inputDataFileName);

            // To copy a file to another location and 
            // overwrite the destination file if it already exists.
            System.IO.File.Copy(sourceFile, destFile, true);

            String fileExtn = BaseUtilities.inputDataFileName.Substring(BaseUtilities.inputDataFileName.IndexOf("."));
            if (fileExtn != null)
            {
                if (fileExtn.ToLower() == ".xlsx")
                {
                    testRunOutputFileName = "OutputData.xlsx";
                }
                else if (fileExtn.ToLower() == ".xls")
                {
                    testRunOutputFileName = "OutputData.xls";

                }
            }

            if (testRunOutputFileName != null)
            {
                System.IO.File.Move(destFile, testRunResultOutputFolder + "\\" + testRunOutputFileName);
            }

        }

        /// <summary>
        /// Report Event --> Main Method
        /// </summary>
        /// <param name="stepName">Name of the step</param>
        /// <param name="stepDescription">Description of the step</param>
        /// <param name="stepStatus">Status of the step</param>
        public static void ReportEvent(string stepName, Status stepStatus, Boolean isScreenshot)
        {

            //System.Drawing.Image image = null;
            if (IsRequiredReport && isScreenshot)
            {
                if (BaseUtilities.applicationType != null)
                {
                    if (BaseUtilities.applicationType.ToLower().Contains("web"))
                    {
                        if (IsDriverActive())
                        {
                            Screenshot screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                            scnshtpath = resultImgFolder + "\\";
                            runtimescnshtpath = scnshtpath + BaseUtilities.GetCurrentTimeWithFormat() + ".png";
                            screenshot.SaveAsFile(runtimescnshtpath);

                        }

                    }
                    if (BaseUtilities.applicationType.ToLower().Contains("windows"))
                    {
                        scnshtpath = resultImgFolder + "\\";
                        runtimescnshtpath = scnshtpath + BaseUtilities.GetCurrentTimeWithFormat() + ".png";
                        TakeScreenshot(runtimescnshtpath);

                    }
                }
            }
            if (isHTMLReport)
            {
                HTMLExtentReport.InsertResultStep(stepName, stepStatus, isScreenshot);
            }
        }
        /// <summary>
        /// Setup error report log 
        /// </summary>
        public static void ExceptionHandler(Exception ex)
        {
            Reports.errorMessage = ex.Message + ex.StackTrace;
            Reports.ReportEvent("Error", Status.Fail, true);
        }
        /// <summary>
        /// Setup error log 
        /// </summary>
        public static void SetupErrorLog(Exception ex)
        {
            if (getCurrentTestRunPath == null)
            {
                SetCurrentTestRunPath();
            }

            String errorLogFile = getCurrentTestRunPath + "\\ErrorLog.txt";

            if (!File.Exists(errorLogFile))
            {
                File.Create(errorLogFile).Dispose();
            }

            StreamWriter sw = File.AppendText(errorLogFile);
            sw.WriteLine("=========================Error Logging========================");
            sw.WriteLine("===============Start===============");
            sw.WriteLine("Error Message: " + ex.Message);
            sw.WriteLine("Stack Trace: " + ex.StackTrace);
            sw.WriteLine("===============End===============");

        }

        public static Status GetEnumStatus(string status)
        {

            if (status.ToLower() == "pass")
            {
                return Status.Pass;

            }
            if (status.ToLower() == "fail")
            {
                return Status.Fail;

            }
            if (status.ToLower() == "error")
            {
                return Status.Error;

            }
            if (status.ToLower() == "skip")
            {
                return Status.Skip;

            }
            if (status.ToLower() == "warning")
            {
                return Status.Warning;

            }
            if (status.ToLower() == "debug")
            {
                return Status.Debug;

            }
            if (status.ToLower() == "information")
            {
                return Status.Information;

            }
            if (status.ToLower() == "done")
            {
                return Status.Done;

            }

            return Status.Null;
        }
    }
}
