using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SecureConnect;
using TDAPIOLELib;
using System.IO.Compression;
using System.IO;
using TestAutomation.FrameworkSupports.Reporting;
using TechTalk.SpecFlow;
using TestAutomation.FrameworkSupports.DataAccess;

namespace TestAutomation.FrameworkSupports
{
    public class ALMQCIntegration
    {
        public static string almFolderPath;
        public static string almDomain;
        public static string almProjectName;
        public static string almUserName;
        public static string almUserDomain;
        public static string almPassword;
        public static string almTestStatus;
        public static string almTestSetName;
        public static string almReportAttachmentType;
        public static void UpdateTestRunsInALMQC()
        {
            string testType = GetProperty.properties["TestType"];
            string testName;
            
            if (testType == null || !testType.Equals("E2E"))
            {
                testName = "[1]" + BaseUtilities.scenarioName;

            }
            else
            {
                testName = "[1]" + BaseUtilities.scenarioName;
                almTestSetName = BaseUtilities.testSetName;
            }
            if (BaseUtilities.scenarioStatus.ToLower() == "pass")
            {
                almTestStatus = "Passed";
            }
            if (BaseUtilities.scenarioStatus.ToLower() == "fail")
            {
                almTestStatus = "Failed";
            }
            ConnectALMQCAndUpdateTestRuns(almFolderPath, almTestSetName, testName);
        }

        private static  void ConnectALMQCAndUpdateTestRuns(string testFolderPath, string testSetName, string Test_Name)
        {
            TDConnection qcConn;
            string url = "https://abc.com";

            try
            {

                qcConn = new TDConnection();
                qcConn.InitConnection(url, almDomain, GeneralMethods.Decrypt(almPassword));
                qcConn.ConnectProject(almProjectName, almUserName, GeneralMethods.Decrypt(almPassword));
                //qcConn.InitConnectionEx(url);
                //qcConn.ConnectProjectEx(domain, projectName, userName, Password.Decrypt(KeyFunctions.GetConfigData("LAN.EncPassword"), userDomain + "-" + userName));
                ////qcConn.ConnectProjectEx(domain, projectName, userName, password, userDomain + "-" + userName));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return;
            }


            if (qcConn.ProjectConnected)
            {
                var testInfo = new List<string>();
                var tsTreeMgr = (TestSetTreeManager)qcConn.TestSetTreeManager;
                var tsFolder = (TestSetFolder)tsTreeMgr.get_NodeByPath(almFolderPath);
                //var tsFolder = (TestSetFolder)tsTreeMgr.get_NodeById(22703);
                //var tsFolder = (TestSetFolder)tsTreeMgr.get_NodeByPath(testFolderPath);
                var attchFactory = (AttachmentFactory)tsFolder.Attachments;
                //string testname1 = tsFolder.Name;
                //var set1 = tsFolder.get_Child(1);
                int nodeid = tsFolder.NodeID;
                var tsList = tsFolder.FindTestSets(testSetName, false, null);
                //var tsList = tsFolder.FindTestSets(testSetName, true,null);
                foreach (TestSet ts in tsList)
                {
                    var tstFolder = (TestSetFolder)ts.TestSetFolder;
                    var tsTestFactory = (TSTestFactory)ts.TSTestFactory;
                    var mylist = tsTestFactory.NewList("");
                    foreach (TSTest tsTest in mylist)
                    {
                        if (tsTest.Name.Equals(Test_Name))
                        {
                            var runFactory = (RunFactory)tsTest.RunFactory;

                            var run = (Run)runFactory.AddItem("AutRun_" + Reports.starttime);
                            run.CopyDesignSteps();

                            //runResult just tells me if overall my test run passes or fails - it's not built in. It was my way of tracking things though the code.
                            run.Status = almTestStatus;
                            run.Post();
                            AttachmentFactory attachmentFactory = (AttachmentFactory)tsTest.Attachments;
                            TDAPIOLELib.Attachment attachment = (TDAPIOLELib.Attachment)attachmentFactory.AddItem(System.DBNull.Value);
                            attachment.Description = "Attached Automation execution result";
                            attachment.Type = 1;
                            if (almReportAttachmentType.ToLower() == "word" && Reports.isWordReport) {
                                attachment.FileName = WordReport.reportpath;
                            }
                            if (almReportAttachmentType.ToLower() == "pdf" && Reports.isPdfReport)
                            {
                                attachment.FileName = PDFReport.reportpath;
                            }
                            attachment.Post();

                            StepFactory rsFactory = (StepFactory)run.StepFactory;
                            dynamic rdata_stepList = rsFactory.NewList("");
                            var rstepList = (TDAPIOLELib.List)rdata_stepList;
                            foreach (dynamic rstep in rstepList)
                            {
                                rstep.Status = almTestStatus;
                                rstep.Post();

                            }
                        }
                    }
                }

            }
        }
    }
}
