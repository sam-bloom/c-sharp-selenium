using TestAutomation.FrameworkSupports.Reporting;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using TestAutomation.FrameworkSupports.DataAccess;

namespace TestAutomation.FrameworkSupports.DataAccess
{
    public class ExcelDataAccess
    {

        public static int dataSheetRowCnt;
        /// <summary>
        /// The Dictionary variable holding all properties adn corresponding values
        /// </summary>
        public static Dictionary<string, DataTable> inputDataDict;
        /// <summary>
        /// The Dictionary variable holding all properties adn corresponding values
        /// </summary>
        public static Dictionary<string, DataTable> webObjRepDict;
        public static Dictionary<string, DataTable> mobileObjRepDict;
        public static string uniqueFieldName = "TestCaseName";
        public static string objRepColumnName = "ObjectRepKey";

        /// <summary>
        /// Return excel work book based on ecel file type
        /// </summary>
        /// <param name="path">file location</param>
        /// <param name="excelfileName">filename </param>
        /// <returns></returns>
        public static IWorkbook GetWorkBook(String path, String excelfileName)
        {
            IWorkbook workbook = null;
            try
            {
                using (FileStream stream = new FileStream(path + "\\" + excelfileName, FileMode.Open, FileAccess.Read))
                {
                    String fileExtn = excelfileName.Substring(excelfileName.IndexOf("."));
                    if (fileExtn != null)
                    {
                        if (fileExtn.ToLower() == ".xlsx")
                        {
                            workbook = new XSSFWorkbook(stream);
                        }
                        else if (fileExtn.ToLower() == ".xls")
                        {
                            workbook = new HSSFWorkbook(stream);
                        }
                        else
                        {
                            throw new FrameworkException("Error ", "Not a valid excel format either .xlx or .xlsx");
                        }
                    }
                    else
                    {
                        throw new FrameworkException("Error ", "Not a valid excel file name");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new FrameworkException(ex.Message + ex.StackTrace);
            }
            return workbook;

        }
        /// <summary>
        /// Function to return the singleton instance of the inputdata dictionary
        /// </summary>
        /// <returns> Instance of the  inputdata dictionary</returns>
        public static Dictionary<string, DataTable> GetDataInstance(String path, String excelfileName)
        {
            Dictionary<string, DataTable> excelDataDict = new Dictionary<string, DataTable>();

            IWorkbook workbooks = GetWorkBook(path, excelfileName);


            for (int i = 0; i < workbooks.NumberOfSheets; i++)
            {
                String ShtName = workbooks.GetSheetName(i);
                DataTable dataTable = RetrieveDataFromExcel(path, excelfileName, ShtName);
                excelDataDict.Add(ShtName, dataTable);
            }


            return excelDataDict;

        }

        public static Dictionary<string, List<string>> MyWebSSOInputData(String path, String excelfileName)
        {
            Dictionary<string, List<string>> excelDataDict = new Dictionary<string, List<string>>();

            XLSExcelDataAccess xlsfile = new XLSExcelDataAccess(path, excelfileName);
            xlsfile.DatasheetName = xlsfile.GetSheetName(0);
            int totalrow = xlsfile.GetLastRowNum();
            for (int i = 1; i <= totalrow; i++)
            {
                List<string> list = new List<string>();
                list.Add(xlsfile.GetValue(i, 0).Trim());
                list.Add(xlsfile.GetValue(i, 1).Trim());
                list.Add(xlsfile.GetValue(i, 2).Trim());
                string TenantURL = xlsfile.GetValue(i, 3).Trim();
                list.Add(TenantURL);
                list.Add(xlsfile.GetValue(i, 4).Trim());
                list.Add(xlsfile.GetValue(i, 5).Trim());
                list.Add(xlsfile.GetValue(i, 6).Trim());
                if (!string.IsNullOrEmpty(TenantURL)) {
                    excelDataDict.Add("User" + i, list);
                }
            }
            return excelDataDict;

        }

        /// <summary>
        /// Write data to Excel Sheet --> Main Method
        /// </summary>
        /// <param name="uniqueFieldName">Unique Field Name</param>
        /// <param name="uniqueFieldValue">Unique Field Value: Probably the Test Case Name</param>
        /// <param name="targetFieldName">Target Field Name: header name of the data</param>
        /// <param name="targetFieldValue">Target Field Value: data which we want to write into excel</param>
        public static void WriteToDataSheet(String path, String excelfile, String sheetName, string targetFieldName, string targetFieldValue)
        {

            string InputDataSheet = path + "\\" + excelfile;
            string uniquefieldname = "false";
            string targetfieldname = "false";
            int uniquefieldcellnum = -1;
            int targetfieldcellnum = -1;

            IWorkbook workbook = GetWorkBook(path, excelfile);
            ISheet sheet = workbook.GetSheet(sheetName);

            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (row == 0)
                {
                    IRow headerRow = sheet.GetRow(row);
                    foreach (ICell headerCell in headerRow)
                    {
                        if (headerCell.ToString() == uniqueFieldName)
                        {
                            uniquefieldname = "true";
                            uniquefieldcellnum = headerCell.ColumnIndex;
                        }
                        if (headerCell.ToString() == targetFieldName)
                        {
                            targetfieldname = "true";
                            targetfieldcellnum = headerCell.ColumnIndex;
                        }
                        if (uniquefieldname == "true" & targetfieldname == "true")
                        {
                            break;
                        }

                    }

                    if (uniquefieldname == "true" & targetfieldname == "false")
                    {
                        targetfieldcellnum = headerRow.LastCellNum;
                        String fileExtn = excelfile.Substring(excelfile.IndexOf("."));
                        if (fileExtn != null)
                        {
                            if (fileExtn.ToLower() == ".xls")
                            {
                                XLSExcelDataAccess xlsfile = new XLSExcelDataAccess(path, excelfile);
                                xlsfile.DatasheetName = sheetName;
                                xlsfile.AddColumn(targetFieldName);
                            }
                            else
                            {
                                throw new FrameworkException("Error ", "Not a valid excel format either .xls file");
                            }
                        }



                    }


                }

                if (sheet.GetRow(row).GetCell(uniquefieldcellnum).ToString() == BaseUtilities.scenarioName)
                {
                    String fileExtn = excelfile.Substring(excelfile.IndexOf("."));
                    if (fileExtn != null)
                    {
                        if (fileExtn.ToLower() == ".xls")
                        {
                            XLSExcelDataAccess xlsfile = new XLSExcelDataAccess(path, excelfile);
                            xlsfile.DatasheetName = sheetName;
                            xlsfile.SetValue(row, targetfieldcellnum, targetFieldValue);
                        }
                        else
                        {
                            throw new FrameworkException("Error ", "Not a valid excel format either .xls");
                        }
                    }

                    break;
                }

            }
        }
        /// <summary>
        /// Store output data into current running scenarion
        /// </summary>
        public static void WriteData(String sheetName, string targetFieldName, string fieldvalue)
        {
            WriteToDataSheet(Reports.testRunResultOutputFolder, Reports.testRunOutputFileName, sheetName, targetFieldName, fieldvalue);
        }

        public void XSSFWriteData(string InputDataSheet, string sheetName)
        {
            FileStream fs = new FileStream(InputDataSheet, FileMode.Create, FileAccess.Write);
            XSSFWorkbook workbook = new XSSFWorkbook(fs);
            //XSSFSheet sheet = workbook.GetSheetName(sheetName);

        }
        /// <summary>
        /// Take data from excel --> Support Method
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        public static DataTable RetrieveDataFromExcel(String path, string excelFilename, string sheetName)
        {
            try
            {

                IWorkbook workbooks = GetWorkBook(path, excelFilename);
                ISheet sheet = workbooks.GetSheet(sheetName);

                OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + "\\" + excelFilename + ";Extended Properties=Excel 12.0;");

                StringBuilder stbQuery = new StringBuilder();

                stbQuery.Append("SELECT * FROM [" + sheet.SheetName + "$]");

                OleDbDataAdapter adp = new OleDbDataAdapter(stbQuery.ToString(), con);

                DataSet dsXLS = new DataSet();

                adp.Fill(dsXLS);

                dataSheetRowCnt = dsXLS.Tables[0].Rows.Count;
                return (dsXLS.Tables[0]);

            }
            catch (Exception e)
            {
                throw new FrameworkException(e.Message + e.StackTrace);

            }

        }


        /// <summary>
        /// Get Input Data from the InputData Datatable --> Main Method
        /// </summary>
        /// <param name="uniqueFieldName">Unique Field Name</param>
        /// <param name="uniqueFieldValue">Unique Field Value: Probably the Test Case Name</param>
        /// <param name="targetFieldName">Target Field Name: header name of the data</param>
        /// <returns></returns>
        public static string GetexcelData(String path, string excelName, string sheetName, string UniqueColumnName, string uniqueFieldValue, string targetFieldName)
        {
            string data = null;
            try
            {

                DataTable inputData = RetrieveDataFromExcel(path, excelName, sheetName);
                DataRow dr = FilterDataTable(inputData, UniqueColumnName + "='" + uniqueFieldValue + "'");
                data = dr[targetFieldName].ToString();
                return (data);
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception error", Reporting.Status.Fail,true);

            }
            return (data);
        }
        /// <summary>
        /// Get user data from input data dictionary
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="uniqueFieldValue"></param>
        /// <param name="targetFieldName"></param>
        /// <returns></returns>

        public static string GetUserInput(String sheetName, string targetFieldName)
        {
            string data = null;
            try
            {

                foreach (KeyValuePair<string, DataTable> list in inputDataDict)
                {

                    if (list.Key.ToLower() == sheetName.ToLower())
                    {

                        DataTable input = list.Value;
                        DataRow dr = FilterDataTable(input, uniqueFieldName + "='" + BaseUtilities.scenarioName + "'");
                        data = dr[targetFieldName].ToString();
                    }
                }


                return (data);
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception error", Reporting.Status.Fail,true);

            }
            return (data);
        }

        /// <summary>
        /// Get element data from object data dictionary
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="uniqueFieldValue"></param>
        /// <param name="targetFieldName"></param>
        /// <returns></returns>
        public static string GetWebObjRepData(String sheetName, string uniqueFieldValue, string targetFieldName)
        {
            string data = null;
            try
            {

                foreach (KeyValuePair<string, DataTable> list in webObjRepDict)
                {

                    if (list.Key.ToLower() == sheetName.ToLower())
                    {

                        DataTable input = list.Value;
                        DataRow dr = FilterDataTable(input, objRepColumnName + "='" + uniqueFieldValue + "'");
                        data = dr[targetFieldName].ToString();
                    }
                }


                return (data);
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception error", Reporting.Status.Fail, true);

            }
            return (data);
        }

        /// <summary>
        /// Get element data from object data dictionary
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="uniqueFieldValue"></param>
        /// <param name="targetFieldName"></param>
        /// <returns></returns>
        public static string GetMobileObjRepData(String sheetName, string uniqueFieldValue, string targetFieldName)
        {
            string data = null;
            try
            {

                foreach (KeyValuePair<string, DataTable> list in mobileObjRepDict)
                {

                    if (list.Key.ToLower() == sheetName.ToLower())
                    {

                        DataTable input = list.Value;
                        DataRow dr = FilterDataTable(input, objRepColumnName + "='" + uniqueFieldValue + "'");
                        data = dr[targetFieldName].ToString();
                    }
                }


                return (data);
            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception error", Reporting.Status.Fail, true);

            }
            return (data);
        }
        /// <summary>
        /// Take values from the Data Table --> Support Method
        /// </summary>
        /// <param name="dataTab">Data Table in which take data from</param>
        /// <param name="FilterQuery">Filter Query from data table</param>
        /// <returns>Return data based on query</returns>
        public static DataRow FilterDataTable(DataTable dataTab, String FilterQuery)
        {
            try
            {
                DataRow datRow = null;
                DataRow[] dataRow = dataTab.Select(FilterQuery);
                datRow = dataRow[0];
                //FilterdTestDataRow = datRow;
                return (datRow);

            }
            catch (Exception e)
            {
                Reports.errorMessage = e.Message + e.StackTrace;
                Reports.ReportEvent("Exception error", Reporting.Status.Fail, true);
                return (null);
            }

        }

    }
}
