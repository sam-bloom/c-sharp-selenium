using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAutomation.FrameworkSupports.Reporting
{
    /// <summary>
    ///  Eumeration to represent the status of the current test step
    /// </summary>
    public enum Status
    {
        /// <summary>
        /// Indicates that the outcome of a verification was not successful
        /// </summary>
        Fail,
        /// <summary>
        ///  Indicates a warning message
        /// </summary>
        Warning,
        /// <summary>
        /// Indicates a infiormation message
        /// </summary>
        Information,
        /// <summary>
        /// Indicates that the outcome of a verification was successful
        /// </summary>
        Pass,

        /// <summary>
        ///Indicates a message that is logged into the results for informational purposes
        /// </summary>
        Done,
        /// <summary>
        /// Indicates a debug-level message, typically used by automation developers
        /// </summary>
        Debug,
         /// <summary>
        /// Indicates test cases skiped 
        /// </summary>
        Skip,
        /// <summary>
        /// Indicates to get error while running
        /// </summary>
        Error,
        /// <summary>
        /// null value
        /// </summary>
        Null

    }

    
}
