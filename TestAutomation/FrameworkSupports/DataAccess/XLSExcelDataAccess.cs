using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using TestAutomation.FrameworkSupports;
using TestAutomation.FrameworkSupports.DataAccess;


namespace TestAutomation.FrameworkSupports.DataAccess
{
    /// <summary>
    /// Class to encapsulate the excel data access layer of the framework
    /// </summary>
    public class XLSExcelDataAccess
    {
        private String _filePath, _fileName;

        /// <summary>
        ///  The Excel sheet name
        /// </summary>
        public String DatasheetName { get; set; }

        /// <summary>
        /// Constructor to initialize the excel data filepath and filename
        /// </summary>
        /// <param name="filePath">The absolute path where the excel data file is stored</param>
        /// <param name="fileName">The name of the excel data file</param>
        public XLSExcelDataAccess(String filePath, String fileName)
        {
            this._filePath = filePath;
            this._fileName = fileName;
        }

        private void CheckPreRequisites()
        {
            if (DatasheetName == null)
            {
                throw new FrameworkException("ExcelDataAccess.datasheetName is not set!");
            }
        }

        private HSSFWorkbook OpenFileForReading()
        {
            String absoluteFilePath = _filePath + "\\" + _fileName+".xls";

            FileStream fileStreamObject;
            try
            {
                fileStreamObject = new FileStream(absoluteFilePath, FileMode.OpenOrCreate, FileAccess.Read);
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine(e.StackTrace);
                throw new FrameworkException("The specified file \"" + absoluteFilePath + "\" does not exist!");
            }

            HSSFWorkbook workbook;
            try
            {
                workbook = new HSSFWorkbook(fileStreamObject);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                throw new FrameworkException("Error while opening the specified Excel workbook \"" + absoluteFilePath + "\"");
            }

            return workbook;
        }

        private void WriteIntoFile(HSSFWorkbook workbook)
        {
            String absoluteFilePath = _filePath + "\\" + _fileName;

            FileStream fileStreamObject;
            try
            {
                fileStreamObject = new FileStream(absoluteFilePath, FileMode.OpenOrCreate, FileAccess.Write);
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine(e.StackTrace);
                throw new FrameworkException("The specified file \"" + absoluteFilePath + "\" does not exist!");
            }

            try
            {
                workbook.Write(fileStreamObject);
                fileStreamObject.Close();
            }
            catch (IOException e)
            {
                Console.WriteLine(e.StackTrace);
                throw new FrameworkException("Error while writing into the specified Excel workbook \"" + absoluteFilePath + "\"");
            }
        }

        private HSSFSheet GetWorkSheet(HSSFWorkbook workbook)
        {
            HSSFSheet worksheet = (HSSFSheet)workbook.GetSheet(DatasheetName);
            if (worksheet == null)
            {
                throw new FrameworkException("The specified sheet \"" + DatasheetName + "\" does not exist within the workbook \"" + _fileName + ".xls\"");
            }

            return worksheet;
        }

        private String GetCellValue(HSSFSheet worksheet, int rowNum, int columnNum)
        {
            HSSFRow row = (HSSFRow)worksheet.GetRow(rowNum);
            HSSFCell cell = (HSSFCell)row.GetCell(columnNum);

            String cellValue;
            if (cell == null)
            {
                cellValue = "";
            }
            else
            {
                cellValue = cell.StringCellValue.Trim();	// Assumption: All cells are formatted as Text
            }
            return cellValue;
        }

        private String GetCellValue(HSSFSheet worksheet, HSSFRow row, int columnNum)
        {
            HSSFCell cell = (HSSFCell)row.GetCell(columnNum);

            String cellValue;
            if (cell == null)
            {
                cellValue = "";
            }
            else
            {
                cellValue = cell.StringCellValue.Trim();	// Assumption: All cells are formatted as Text
            }
            return cellValue;
        }
        public string GetSheetName(int index)
        {
            HSSFWorkbook workbook = OpenFileForReading();

            return workbook.GetSheetName(index);
        }
        /// <summary>
        ///  Function to search for a specified key within a column, and return the corresponding row number
        /// </summary>
        /// <param name="key">The value being searched for</param>
        /// <param name="columnNum">The column number in which the key should be searched</param>
        /// <param name="startRowNum">The row number from which the search should start</param>
        /// <returns>The row number in which the specified key is found (-1 if the key is not found)</returns>
        public int GetRowNum(String key, int columnNum, int startRowNum)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            String currentValue;
            for (int currentRowNum = startRowNum;
                    currentRowNum <= worksheet.LastRowNum; currentRowNum++)
            {

                currentValue = GetCellValue(worksheet, currentRowNum, columnNum);

                if (currentValue.Equals(key))
                {
                    return currentRowNum;
                }
            }

            return -1;
        }

        /// <summary>
        /// Function to search for a specified key within a column, and return the corresponding row number
        /// </summary>
        /// <param name="key">The value being searched for</param>
        /// <param name="columnNum">The column number in which the key should be searched</param>
        /// <returns>The row number in which the specified key is found (-1 if the key is not found)</returns>
        public int GetRowNum(String key, int columnNum)
        {
            return GetRowNum(key, columnNum, 0);
        }

        /// <summary>
        /// Function to get the last row number within the worksheet
        /// </summary>
        /// <returns>The last row number within the worksheet</returns>
        public int GetLastRowNum()
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            return worksheet.LastRowNum;
        }

        /// <summary>
        ///Function to search for a specified key within a column, and return the corresponding occurence count
        /// </summary>
        /// <param name="key">The value being searched for</param>
        /// <param name="columnNum">The column number in which the key should be searched</param>
        /// <param name="startRowNum">The row number from which the search should start</param>
        /// <returns>The occurence count of the specified key</returns>
        public int GetRowCount(String key, int columnNum, int startRowNum)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            int rowCount = 0;
            Boolean keyFound = false;

            String currentValue;
            for (int currentRowNum = startRowNum;
                    currentRowNum <= worksheet.LastRowNum; currentRowNum++)
            {

                currentValue = GetCellValue(worksheet, currentRowNum, columnNum);

                if (currentValue.Equals(key))
                {
                    rowCount++;
                    keyFound = true;
                }
                else
                {
                    if (keyFound)
                    {
                        break;	// Assumption: Keys always appear contiguously
                    }
                }
            }

            return rowCount;
        }

        /// <summary>
        ///  Function to search for a specified key within a column, and return the corresponding occurence count
        /// </summary>
        /// <param name="key">The value being searched for</param>
        /// <param name="columnNum">The column number in which the key should be searched</param>
        /// <returns>The occurence count of the specified key</returns>
        public int GetRowCount(String key, int columnNum)
        {
            return GetRowCount(key, columnNum, 0);
        }

        /// <summary>
        /// Function to search for a specified key within a row, and return the corresponding column number
        /// </summary>
        /// <param name="key">The value being searched for</param>
        /// <param name="rowNum">The row number in which the key should be searched</param>
        /// <returns>The column number in which the specified key is found (-1 if the key is not found)</returns>
        public int GetColumnNum(String key, int rowNum)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            HSSFRow row = (HSSFRow)worksheet.GetRow(rowNum);
            String currentValue;
            for (int currentColumnNum = 0;
                    currentColumnNum < row.LastCellNum; currentColumnNum++)
            {

                currentValue = GetCellValue(worksheet, row, currentColumnNum);

                if (currentValue.Equals(key))
                {
                    return currentColumnNum;
                }
            }

            return -1;
        }

        /// <summary>
        ///  Function to get the value in the cell identified by the specified row and column numbers
        /// </summary>
        /// <param name="rowNum"> The row number of the cell</param>
        /// <param name="columnNum">The column number of the cell</param>
        /// <returns>The value present in the cell</returns>
        public String GetValue(int rowNum, int columnNum)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            String cellValue = GetCellValue(worksheet, rowNum, columnNum);

            return cellValue;
        }
        /// <summary>
        ///  Function to get the value in the cell identified by the specified row number and column header
        /// </summary>
        /// <param name="rowNum">The row number of the cell</param>
        /// <param name="columnHeader">The column header of the cell</param>
        /// <returns>The column header of the cell</returns>
        public String GetValue(int rowNum, String columnHeader)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            HSSFRow row = (HSSFRow)worksheet.GetRow(0);	//0 because header is always in the first row
            int columnNum = -1;
            String currentValue;
            for (int currentColumnNum = 0;
                    currentColumnNum < row.LastCellNum; currentColumnNum++)
            {

                currentValue = GetCellValue(worksheet, row, currentColumnNum);

                if (currentValue.Equals(columnHeader))
                {
                    columnNum = currentColumnNum;
                    break;
                }
            }

            if (columnNum == -1)
            {
                throw new FrameworkException("The specified column header \"" + columnHeader + "\" is not found in the sheet \"" + DatasheetName + "\"!");
            }
            else
            {
                String cellValue = GetCellValue(worksheet, rowNum, columnNum);
                return cellValue;
            }
        }

        private HSSFCellStyle ApplyCellStyle(HSSFWorkbook workbook,
                                                ExcelCellFormatting cellFormatting)
        {
            HSSFCellStyle cellStyle = (HSSFCellStyle)workbook.CreateCellStyle();
            if (cellFormatting.Centred)
            {
                cellStyle.Alignment = HorizontalAlignment.Center;
            }
            cellStyle.FillForegroundColor = cellFormatting.BackColorIndex;
            cellStyle.FillPattern = FillPattern.SolidForeground;

            HSSFFont font = (HSSFFont)workbook.CreateFont();
            font.FontName = cellFormatting.FontName;
            font.FontHeightInPoints = cellFormatting.FontSize;
            if (cellFormatting.Bold)
            {
                font.Boldweight = (short)FontBoldWeight.Bold;
            }
            font.Color = cellFormatting.ForeColorIndex;
            cellStyle.SetFont(font);

            return cellStyle;
        }

        /// <summary>
        /// Function to set the specified value in the cell identified by the specified row and column numbers
        /// </summary>
        /// <param name="rowNum">The row number of the cell</param>
        /// <param name="columnNum">The column number of the cell</param>
        /// <param name="value">The value to be set in the cell</param>
        public void SetValue(int rowNum, int columnNum, String value)
        {
            SetValue(rowNum, columnNum, value, null);
        }

        /// <summary>
        ///  Function to set the specified value in the cell identified by the specified row and column numbers
        /// </summary>
        /// <param name="rowNum">The row number of the cell</param>
        /// <param name="columnNum">The column number of the cell</param>
        /// <param name="value">The value to be set in the cell</param>
        /// <param name="cellFormatting">The ExcelCellFormatting to be applied to the cell</param>
        public void SetValue(int rowNum, int columnNum, String value,
                                                ExcelCellFormatting cellFormatting)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            HSSFRow row = (HSSFRow)worksheet.GetRow(rowNum);
            HSSFCell cell = (HSSFCell)row.CreateCell(columnNum);
            cell.SetCellType(CellType.String);
            cell.SetCellValue(value);

            if (cellFormatting != null)
            {
                HSSFCellStyle cellStyle = ApplyCellStyle(workbook, cellFormatting);
                cell.CellStyle = cellStyle;
            }

            WriteIntoFile(workbook);
        }

        /// <summary>
        ///   Function to set the specified value in the cell identified by the specified row number and column header
        /// </summary>
        /// <param name="rowNum">The row number of the cell</param>
        /// <param name="columnHeader">The column header of the cell</param>
        /// <param name="value">The value to be set in the cell</param>
        public void SetValue(int rowNum, String columnHeader, String value)
        {
            SetValue(rowNum, columnHeader, value, null);
        }

        /// <summary>
        ///  Function to set the specified value in the cell identified by the specified row number and column header
        /// </summary>
        /// <param name="rowNum">The row number of the cell</param>
        /// <param name="columnHeader">The column header of the cell</param>
        /// <param name="value">The value to be set in the cell</param>
        /// <param name="cellFormatting">The ExcelCellFormatting to be applied to the cell</param>
        public void SetValue(int rowNum, String columnHeader, String value,
                                                    ExcelCellFormatting cellFormatting)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            HSSFRow row = (HSSFRow)worksheet.GetRow(0);	//0 because header is always in the first row
            int columnNum = -1;
            String currentValue;
            for (int currentColumnNum = 0;
                    currentColumnNum < row.LastCellNum; currentColumnNum++)
            {

                currentValue = GetCellValue(worksheet, row, currentColumnNum);

                if (currentValue.Equals(columnHeader))
                {
                    columnNum = currentColumnNum;
                    break;
                }
            }

            if (columnNum == -1)
            {
                throw new FrameworkException("The specified column header " + columnHeader + " is not found in the sheet \"" + DatasheetName + "\"!");
            }
            else
            {
                row = (HSSFRow)worksheet.GetRow(rowNum);
                HSSFCell cell = (HSSFCell)row.CreateCell(columnNum);
                cell.SetCellType(CellType.String);
                cell.SetCellValue(value);

                if (cellFormatting != null)
                {
                    HSSFCellStyle cellStyle = ApplyCellStyle(workbook, cellFormatting);
                    cell.CellStyle = cellStyle;
                }

                WriteIntoFile(workbook);
            }
        }

        /// <summary>
        ///  Function to set a hyperlink in the cell identified by the specified row and column numbers
        /// </summary>
        /// <param name="rowNum">The row number of the cell</param>
        /// <param name="columnNum">The column number of the cell</param>
        /// <param name="linkAddress">The link address to be set
        public void SetHyperlink(int rowNum, int columnNum, String linkAddress)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            HSSFRow row = (HSSFRow)worksheet.GetRow(rowNum);
            HSSFCell cell = (HSSFCell)row.GetCell(columnNum);
            if (cell == null)
            {
                throw new FrameworkException("Specified cell is empty! Please set a value before including a hyperlink...");
            }

            SetCellHyperlink(workbook, cell, linkAddress);

            WriteIntoFile(workbook);
        }

        private void SetCellHyperlink(HSSFWorkbook workbook, HSSFCell cell, String linkAddress)
        {
            HSSFCellStyle cellStyle = (HSSFCellStyle)cell.CellStyle;
            HSSFFont font = (HSSFFont)cellStyle.GetFont(workbook);
            font.Underline = FontUnderlineType.Single;
            cellStyle.SetFont(font);
            HSSFHyperlink hyperlink = new HSSFHyperlink(HyperlinkType.Url);
            hyperlink.Address = linkAddress;
            cell.Hyperlink = hyperlink;
            cell.CellStyle = cellStyle;

        }

        /// <summary>
        ///  Function to set a hyperlink in the cell identified by the specified row number and column header
        /// </summary>
        /// <param name="rowNum"> The row number of the cell</param>
        /// <param name="columnHeader">The column header of the cell</param>
        /// <param name="linkAddress">The link address to be set</param>
        public void SetHyperlink(int rowNum, String columnHeader, String linkAddress)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            HSSFRow row = (HSSFRow)worksheet.GetRow(0);	//0 because header is always in the first row
            int columnNum = -1;
            String currentValue;
            for (int currentColumnNum = 0;
                    currentColumnNum < row.LastCellNum; currentColumnNum++)
            {

                currentValue = GetCellValue(worksheet, row, currentColumnNum);

                if (currentValue.Equals(columnHeader))
                {
                    columnNum = currentColumnNum;
                    break;
                }
            }

            if (columnNum == -1)
            {
                throw new FrameworkException("The specified column header " + columnHeader + " is not found in the sheet \"" + DatasheetName + "\"!");
            }
            else
            {
                row = (HSSFRow)worksheet.GetRow(rowNum);
                HSSFCell cell = (HSSFCell)row.GetCell(columnNum);
                if (cell == null)
                {
                    throw new FrameworkException("Specified cell is empty! Please set a value before including a hyperlink...");
                }

                SetCellHyperlink(workbook, cell, linkAddress);

                WriteIntoFile(workbook);
            }
        }

        /// <summary>
        ///   Function to create a new Excel workbook
        /// </summary>
        public void CreateWorkbook()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();

            WriteIntoFile(workbook);
        }

        /// <summary>
        ///   Function to add a sheet to the Excel workbook
        /// </summary>
        /// <param name="sheetName">The sheet name to be added</param>
        public void AddSheet(String sheetName)
        {
            HSSFWorkbook workbook = OpenFileForReading();

            HSSFSheet worksheet = (HSSFSheet)workbook.CreateSheet(sheetName);
            worksheet.CreateRow(0);	//include a blank row in the sheet created

            WriteIntoFile(workbook);

            this.DatasheetName = sheetName;
        }

        /// <summary>
        ///  Function to add a new row to the Excel worksheet
        /// </summary>
        /// <returns>The row number of the newly added row</returns>
        public int AddRow()
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            int newRowNum = worksheet.LastRowNum + 1;
            worksheet.CreateRow(newRowNum);

            WriteIntoFile(workbook);

            return newRowNum;
        }

        /// <summary>
        /// Function to add a new column to the Excel worksheet
        /// </summary>
        /// <param name="columnHeader">The column header to be added</param>
        public void AddColumn(String columnHeader)
        {
            AddColumn(columnHeader, null);
        }

        /// <summary>
        ///   Function to add a new column to the Excel worksheet
        /// </summary>
        /// <param name="columnHeader">The column header to be added</param>
        /// <param name="cellFormatting">The ExcelCellFormatting to be applied to the column header</param>
        public void AddColumn(String columnHeader, ExcelCellFormatting cellFormatting)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            HSSFRow row = (HSSFRow)worksheet.GetRow(0);	//0 because header is always in the first row
            int lastCellNum = row.LastCellNum;
            if (lastCellNum == -1)
            {
                lastCellNum = 0;
            }

            HSSFCell cell = (HSSFCell)row.CreateCell(lastCellNum);
            cell.SetCellType(CellType.String);
            cell.SetCellValue(columnHeader);

            if (cellFormatting != null)
            {
                HSSFCellStyle cellStyle = ApplyCellStyle(workbook, cellFormatting);
                cell.CellStyle = cellStyle;
            }

            WriteIntoFile(workbook);
        }


        /// <summary>
        ///  Function to outline (i.e., group together) the specified rows
        /// </summary>
        /// <param name="firstRow">The first row</param>
        /// <param name="lastRow">The last row</param>
        public void GroupRows(int firstRow, int lastRow)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            worksheet.GroupRow(firstRow, lastRow);

            WriteIntoFile(workbook);
        }

        /// <summary>
        ///  Function to automatically adjust the column width to fit the contents for the specified range of columns (all inputs are 0-based)
        /// </summary>
        /// <param name="firstCol">The first column</param>
        /// <param name="lastCol">The last column</param>
        public void AutoFitContents(int firstCol, int lastCol)
        {
            CheckPreRequisites();

            HSSFWorkbook workbook = OpenFileForReading();
            HSSFSheet worksheet = GetWorkSheet(workbook);

            if (firstCol < 0)
            {
                firstCol = 0;
            }

            if (firstCol > lastCol)
            {
                throw new FrameworkException("First column cannot be greater than last column!");
            }

            for (int currentColumn = firstCol;
                        currentColumn <= lastCol; currentColumn++)
            {
                worksheet.AutoSizeColumn(currentColumn);
            }

            WriteIntoFile(workbook);
        }



    }
}

