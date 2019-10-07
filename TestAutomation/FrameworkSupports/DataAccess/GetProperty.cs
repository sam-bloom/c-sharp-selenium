using TestAutomation.FrameworkSupports.Reporting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAutomation.FrameworkSupports.DataAccess
{
    class GetProperty
    {
        public static Dictionary<string, string> properties;
        /// <summary>
        /// The private object used to lock on 
        /// </summary>
        private static object _syncRoot = new Object();

        /// <summary>
        /// Function to return the singleton instance of the Properties object
        /// </summary>
        /// <returns> Instance of the  Properties object</returns>
        public static Dictionary<string, string> GetInstance(String propertiesFilePath)
        {

            if (properties == null)
            {
                lock (_syncRoot)
                {
                    if (properties == null)
                        LoadFromPropertiesFile(propertiesFilePath);
                }
            }

            return properties;

        }

        /// <summary>
        /// Function to read and load properties from file to dictionary variable
        /// </summary>
        private static void LoadFromPropertiesFile(String propertiesFilePath)
        {

            try
            {
                properties = new Dictionary<string, string>();
                foreach (var row in File.ReadAllLines(propertiesFilePath))
                    properties.Add(row.Split('=')[0], row.Split('=')[1]);
            }
            catch (Exception e)
            {
                throw new FrameworkException(e.Message + e.StackTrace);
            }
        }

        /// <summary>
        /// Set all Configuration Data from properties file
        /// </summary>
        public static void GetConfigDataFromProperty()
        {
            BaseUtilities.objectIdentificationTimeOut = Convert.ToInt16(properties["ObjectIdentificationTimeOut"]);

            BaseUtilities.objectSyncTimeout = Convert.ToInt16(properties["ObjectSyncTimeout"]);
            BaseUtilities.browser = properties["Browser"];

            BaseUtilities.inputDataFileName = properties["InputDataFileName"];
            if(properties["WordReport"].ToLower()=="true"){
                Reports.isWordReport = true; 
            }
            if (properties["PDFReport"].ToLower() == "true")
            {
                Reports.isPdfReport = true;
            }
            if (properties["IsRequiredReport"].ToLower() == "true")
            {
                Reports.isHTMLReport = true;
            }
            //inputDataFileName = KeyFunctions.GetConfigData("InputDataFileName").Trim();
            ALMQCIntegration.almFolderPath = properties["ALM.TestFolder"].Trim();
            ALMQCIntegration.almDomain = properties["ALM.Domain"].Trim();
            ALMQCIntegration.almProjectName = properties["ALM.Project"].Trim();
            ALMQCIntegration.almUserName = properties["LAN.UserName"].Trim();
            ALMQCIntegration.almUserDomain = properties["LAN.Domain"].Trim();
            ALMQCIntegration.almPassword = properties["LAN.EncPassword"].Trim();
            ALMQCIntegration.almReportAttachmentType = properties["ALMReportAttachmentType"].Trim();

            if (properties["IsReuiredALMUpdates"].Trim().ToLower() == "true")
            {
                Reports.isReuiredALMUpdates = true;
            }


        }
    }
}
