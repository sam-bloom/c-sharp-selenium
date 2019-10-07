using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace TestAutomation.FrameworkSupports
{



    /// <summary>
    ///  Exception class for the framework
    /// </summary>
    //@SuppressWarnings("serial")
    //# pragma warning disable 0618
    //[CLSCompliant(false)]
    public class FrameworkException : Exception
    {
        /// <summary>
        ///   The step name to be specified for the exception
        /// </summary>
        public String errorName = "Error";


        /// <summary>
        ///   Constructor to initialize the exception from the framework
        /// </summary>
        /// <param name="errorDescription">The Exception message to be thrown</param>
        public FrameworkException(String errorDescription)
            : base(errorDescription)
        {

        }

        /// <summary>
        ///   Constructor to initialize the exception from the framework
        /// </summary>
        /// <param name="errorName">The step name for the error</param>
        /// <param name="errorDescription">The Exception message to be thrown</param>
        public FrameworkException(String errorName, String errorDescription)
            : base(errorDescription)
        {
            this.errorName = errorName;
            throw new Exception(errorDescription);
        }
    }
}
