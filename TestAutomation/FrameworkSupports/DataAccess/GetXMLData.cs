using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using System.Xml;
using TestAutomation.FrameworkSupports.Reporting;
using System.Xml.Linq;

namespace TestAutomation.FrameworkSupports.DataAccess
{
    public class GetXMLData
    {
        public static Dictionary<string, string> ConfigXMLData;
        /// <summary>
        /// The private object used to lock on 
        /// </summary>
        private static object _syncRoot = new Object();
        /// <summary>
        /// Function to return the singleton instance of the xml object
        /// </summary>
        /// <returns> Instance of the  Properties object</returns>
        public static Dictionary<string, string> GetXMLConfigData(String XMLFilePath)
        {

            if (ConfigXMLData == null)
            {
                lock (_syncRoot)
                {
                    if (ConfigXMLData == null)
                        LoadFromXMLFile(XMLFilePath);
                }
            }

            return ConfigXMLData;

        }

        /// <summary>
        /// Function to read and load xml from file to dictionary variable
        /// </summary>
        private static void LoadFromXMLFile(String XMLFilePath)
        {

            try
            {
                XDocument xmldoc = new XDocument();
                XmlTextReader read = new XmlTextReader(XMLFilePath);
                FileStream fs = new FileStream(XMLFilePath, FileMode.Open, FileAccess.Read);
                ConfigXMLData = new Dictionary<string, string>();
                read.Read();
                while (!read.EOF)
                {
                    if (read.Name=="provider")
                    {
                        ConfigXMLData.Add(read.GetAttribute("name"), read.GetAttribute("value"));
                    }
                    read.Read();
                }
                read.Close();
            }
            catch (Exception e)
            {
                throw new FrameworkException(e.Message + e.StackTrace);
            }
        }

        /// <summary>
        /// Set all Configuration Data from properties file
        /// </summary>
        public static void GetConfigDataFromXML()
        {
            BaseUtilities.objectIdentificationTimeOut = Convert.ToInt16(ConfigXMLData["ObjectIdentificationTimeOut"]);

            BaseUtilities.objectSyncTimeout = Convert.ToInt16(ConfigXMLData["ObjectSyncTimeout"]);
            BaseUtilities.browser = ConfigXMLData["Browser"];

            BaseUtilities.inputDataFileName = ConfigXMLData["InputDataFileName"];
            if (ConfigXMLData["WordReport"].ToLower() == "true")
            {
                Reports.isWordReport = true;
            }
            if (ConfigXMLData["PDFReport"].ToLower() == "true")
            {
                Reports.isPdfReport = true;
            }
            if (ConfigXMLData["IsRequiredReport"].ToLower() == "true")
            {
                Reports.isHTMLReport = true;
                Reports.IsRequiredReport = true;
            }
            //inputDataFileName = KeyFunctions.GetConfigData("InputDataFileName").Trim();
            ALMQCIntegration.almFolderPath = ConfigXMLData["ALM.TestFolder"].Trim();
            ALMQCIntegration.almDomain = ConfigXMLData["ALM.Domain"].Trim();
            ALMQCIntegration.almProjectName = ConfigXMLData["ALM.Project"].Trim();
            ALMQCIntegration.almUserName = ConfigXMLData["LAN.UserName"].Trim();
            ALMQCIntegration.almUserDomain = ConfigXMLData["LAN.Domain"].Trim();
            ALMQCIntegration.almPassword = ConfigXMLData["LAN.EncPassword"].Trim();
            ALMQCIntegration.almReportAttachmentType = ConfigXMLData["ALMReportAttachmentType"].Trim();

            if (ConfigXMLData["IsReuiredALMUpdates"].Trim().ToLower() == "true")
            {
                Reports.isReuiredALMUpdates = true;
            }
        }
    }
}
