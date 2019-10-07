using TestAutomation.FrameworkSupports;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestAutomation.FrameworkSupports.DataAccess
{
    /// <summary>
    /// Class to encapsulate the cell formatting settings for the Excel spreadsheet
    /// </summary>
    public class ExcelCellFormatting
    {
        /// <summary>
        ///  Function to get the name of the font to be used in the cell
        /// </summary>
        public String FontName { get; set; }

        /// <summary>
        /// The font size to be used in the cell
        /// </summary>
        public short FontSize { get; set; }

        private short _backColorIndex;
        /// <summary>
        ///The index of the background color for the cell
        /// </summary>
        public short BackColorIndex { get { return _backColorIndex; } set { if (value < 0x8 || value > 0x40) { throw new FrameworkException("Valid indexes for the Excel custom palette are from 0x8 to 0x40 (inclusive)!"); } _backColorIndex = value; } }

        private short _foreColorIndex;
        /// <summary>
        ///  The index of the foreground color (i.e., font color) for the cell
        /// </summary>
        public short ForeColorIndex { get { return _foreColorIndex; } set { if (value < 0x8 || value > 0x40) { throw new FrameworkException("Valid indexes for the Excel custom palette are from 0x8 to 0x40 (inclusive)!"); } _foreColorIndex = value; } }

        /// <summary>
        /// Boolean variable to control whether the cell contents are in bold
        /// </summary>
        public Boolean Bold = false;


        /// <summary>
        /// Boolean variable to control whether the cell contents are in italics
        /// </summary>

        public Boolean Italics = false;


        /// <summary>
        /// Boolean variable to control whether the cell contents are centred
        /// </summary>
        public Boolean Centred = false;
    }
}