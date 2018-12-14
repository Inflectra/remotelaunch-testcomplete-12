using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartBear.TestComplete12;

namespace Inflectra.RemoteLaunch.Engines.TestCompleteEngine.TestComplete12
{
    /// <summary>
    /// Represents a single action in the TC test log
    /// </summary>
    public class TestResultStep
    {
        public TestResult.TestResultEntryType Type { get; set; }
        
        public string Message { get; set; }

        public string AdditionalInfo { get; set; }

        public string ImageFilename { get; set; }

        public byte[] Image { get; set; }
    }
}
