using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartBear.TestComplete12;
using System.Xml;
using System.Xml.XPath;
using System.Collections;
using System.IO;

namespace Inflectra.RemoteLaunch.Engines.TestCompleteEngine.TestComplete12
{
    /// <summary>
    /// Represents a single test result
    /// </summary>
    public class TestResult
    {
        #region Enumerations
        
        /// <summary>
        /// http://stackoverflow.com/questions/10137334/why-does-testcomplete-keep-changing-status-image-name
        /// </summary>
        public enum TestResultEntryType
        {
            Message = 0,
            Event = 1,
            Warning = 2,
            Error = 3,
            Checkpoint = 7
        }

        #endregion

        protected string messageLog;
        protected List<TestResultStep> steps = new List<TestResultStep>();

        /// <summary>
        /// Contains the last execution status
        /// </summary>
        public TC_LOG_STATUS ExecutionStatus { get; set; }

        /// <summary>
        /// The start time of the execution
        /// </summary>
        public Nullable<DateTime> StartTime { get; set; }

        /// <summary>
        /// The end time of the execution
        /// </summary>
        public Nullable<DateTime> EndTime { get; set; }

        /// <summary>
        /// Is the test completed
        /// </summary>
        public bool IsTestCompleted { get; set; }

        /// <summary>
        /// The error count
        /// </summary>
        public int ErrorCount { get; set; }

        /// <summary>
        /// The warning count
        /// </summary>
        public int WarningCount { get; set; }

        /// <summary>
        /// The type of test
        /// </summary>
        public string TestType { get; set; }

        /// <summary>
        /// The detailed message log
        /// </summary>
        public string MessageLog
        {
            get
            {
                return this.messageLog;
            }
        }

        public List<TestResultStep> Steps
        {
            get
            {
                return this.steps;
            }
        }

        /// <summary>
        /// Extracts the XML TestComplete logfile into a string format that can be reported back
        /// </summary>
        /// <param name="filename">The full path of the RootLogData.dat file</param>
        /// <param name="testItemName">The project test item name</param>
        /// <param name="scriptLogFile">Is this the logfile for a script/routine vs. a project or project test item</param>
        public void ExtractLog(string filename, string testItemName, bool scriptLogFile)
        {
            //Use a stringbuilder to create the message log
            StringBuilder logMessages = new StringBuilder();

            //The provided filename is for Description.tcLog, whereas we need RootLogData.dat
            filename = filename.Replace("Description.tcLog", "RootLogData.dat");

            //Open up the RootLogData.dat XML Log file
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(filename);

            //Extract the list of individual test files from this top-level logfile (RootLogData.dat)
            XmlNodeList xmlNodes = xmlDoc.SelectNodes("/Nodes/Node/Node");
            if (xmlNodes == null)
            {
                throw new ApplicationException("Unable to retrieve node list");
            }

            //See if we have a script or a test project / project item
            if (scriptLogFile)
            {
                //Iterate through each of the nodes
                foreach (XmlNode xmlNode in xmlNodes)
                {
                    //Access each node in turn and get the script filename
                    XmlNode xmlPrpNode = xmlNode.SelectSingleNode("./Prp[@name='filename']");
                    if (xmlPrpNode != null)
                    {
                        //Extract the filename attribute from the node
                        if (xmlPrpNode.Attributes["value"] != null)
                        {
                            string detailedLogFilename = xmlPrpNode.Attributes["value"].Value;

                            //Since this filename does not include the path, need to get that from the root path.
                            string pathname = filename.Replace("RootLogData.dat", detailedLogFilename);
                            ExtractLogMessages(pathname, logMessages);
                        }
                    }
                }
            }
            else
            {
                //Iterate through each of the nodes
                foreach (XmlNode xmlNode in xmlNodes)
                {
                    //Access each node in turn and get the test name
                    //It needs to match the projectkey node
                    string projectKey = "";
                    XmlNode xmlProjectKeyNode = xmlNode.SelectSingleNode("./Prp[(@name='projectkey')]");
                    if (xmlProjectKeyNode != null)
                    {
                        projectKey = xmlProjectKeyNode.Attributes["value"].Value;
                    }
                    XmlNode xmlPrpNode = xmlNode.SelectSingleNode("./Prp[(@name='test' and @value='" + projectKey + "')]");
                    if (xmlPrpNode != null)
                    {
                        //Since we might have multiple matching test items under this test, need to make sure that the test item name matches
                        if (String.IsNullOrEmpty(testItemName))
                        {
                            //We are not given a specific test item, so all results need to be appended to the message
                            //Now get the actual child node that points to the log file
                            XmlNode xmlPrpNode2 = xmlPrpNode.SelectSingleNode("../Node[@name='children']/Prp[@name='child 0']");
                            if (xmlPrpNode2 != null)
                            {
                                //Extract the filename attribute from the child-node
                                if (xmlPrpNode2.Attributes["value"] != null)
                                {
                                    string detailedLogFilename = xmlPrpNode2.Attributes["value"].Value;

                                    //Since this filename does not include the path, need to get that from the root path.
                                    string pathname = filename.Replace("RootLogData.dat", detailedLogFilename);
                                    ExtractLogMessages(pathname, logMessages);
                                }
                            }
                        }
                        else
                        {
                            XmlNode xmlPrpTestItemNode = xmlPrpNode.SelectSingleNode("../Prp[(@name='name' and @value='" + testItemName + "')]");
                            if (xmlPrpTestItemNode != null)
                            {
                                //Now get the actual child node that points to the log file
                                XmlNode xmlPrpNode2 = xmlPrpNode.SelectSingleNode("../Node[@name='children']/Prp[@name='child 0']");
                                if (xmlPrpNode2 != null)
                                {
                                    //Extract the filename attribute from the child-node
                                    if (xmlPrpNode2.Attributes["value"] != null)
                                    {
                                        string detailedLogFilename = xmlPrpNode2.Attributes["value"].Value;

                                        //Since this filename does not include the path, need to get that from the root path.
                                        string pathname = filename.Replace("RootLogData.dat", detailedLogFilename);
                                        ExtractLogMessages(pathname, logMessages);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //Store the results
            this.messageLog = logMessages.ToString();
            if (String.IsNullOrWhiteSpace(this.messageLog))
            {
                this.messageLog = "No messages found in logfile: " + filename;
            }
        }

        /// <summary>
        /// Extracts the log for a specific detailed log file
        /// </summary>
        /// <param name="pathname"></param>
        /// <param name="logMessages"></param>
        public void ExtractLogMessages(string pathname, StringBuilder logMessages)
        {
            //Open up the Log file
            logMessages.Append("> Opening up logfile: " + pathname + "\n");
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(pathname);

            //Extract the list of individual test test messages from this logfile
            //Need to sort on the 
            XPathNavigator navigator = xmlDoc.CreateNavigator();
            XPathExpression expression = navigator.Compile("/Nodes/Node/Node");
            expression.AddSort("@name", new NodeNameComparer());
            XPathNodeIterator iterator = navigator.Select(expression);
            if (iterator == null)
            {
                throw new ApplicationException("Unable to retrieve test nodes from logfile");
            }

            //Iterate through each of the nodes and get the message and results. Then add to string
            //(need to iterate in reverse order so that they appear in chronological order)
            foreach (XPathNavigator xmlNavNode in iterator)
            {
                //Access each node in turn
                XmlNode xmlNode = ((System.Xml.IHasXmlNode)xmlNavNode).GetNode();

                //Now select the Node/Prp element that has the message and add to the string
                XmlNode xmlPrpNode = xmlNode.SelectSingleNode("./Prp[@name='message']");
                if (xmlPrpNode != null)
                {
                    //Extract the value attribute from this node
                    XmlAttribute xmlAttribute = xmlPrpNode.Attributes["value"];
                    if (xmlAttribute != null)
                    {
                        string message = xmlAttribute.Value;

                        //Now select the Node/Prp element that has the type and add the message to the string
                        XmlNode xmlPrpNode2 = xmlNode.SelectSingleNode("./Prp[@name='type']");
                        if (xmlPrpNode2 != null)
                        {
                            //Now get the type and add the appropriate message prefix
                            XmlAttribute xmlAttribute2 = xmlPrpNode2.Attributes["value"];
                            if (xmlAttribute2 != null)
                            {
                                string typeString = xmlAttribute2.Value;
                                int type;
                                if (Int32.TryParse(typeString, out type))
                                {
                                    //There can be a lot of 'Event' messages so check if we want to include them
                                    if (type != (int)TestResultEntryType.Event || Properties.Settings.Default.IncludeEventMessages)
                                    {
                                        if (type == (int)TestResultEntryType.Error)
                                        {
                                            //Error
                                            logMessages.Append("*ERROR*: ");
                                        }
                                        else if (type == (int)TestResultEntryType.Message)
                                        {
                                            //Action/Message
                                            logMessages.Append("Action: ");
                                        }
                                        else if (type == (int)TestResultEntryType.Event)
                                        {
                                            //Event
                                            logMessages.Append("Event: ");
                                        }
                                        else if (type == (int)TestResultEntryType.Warning)
                                        {
                                            //Warning
                                            logMessages.Append("!Warning!: ");
                                        }
                                        else
                                        {
                                            //Other
                                            logMessages.Append("Other: ");
                                        }

                                        //Now add to the log message
                                        logMessages.Append(message + "\n");

                                        //See if we also have a value for the 'remarks' node
                                        string remarks = "";
                                        XmlNode xmlRemarksPrpNode = xmlNode.SelectSingleNode("./Prp[@name='remarks']");
                                        if (xmlRemarksPrpNode != null)
                                        {
                                            XmlAttribute xmlRemarksValueAttribute = xmlRemarksPrpNode.Attributes["value"];
                                            if (xmlRemarksValueAttribute != null)
                                            {
                                                remarks = xmlRemarksValueAttribute.Value;
                                            }
                                        }

                                        //Also add to the list of 'test steps'
                                        TestResultStep testResultStep = new TestResultStep();
                                        testResultStep.Type = (TestResult.TestResultEntryType)type;
                                        testResultStep.Message = message;
                                        testResultStep.AdditionalInfo = remarks;
                                        this.steps.Add(testResultStep);

                                        //See if we have an image to attach
                                        XmlNode xmlImagePrpNode = xmlNode.SelectSingleNode("./Node[@name='visualizer']/Node[@name='current']/Prp[@name='image file name']");
                                        if (xmlImagePrpNode != null)
                                        {
                                            XmlAttribute xmlImageValueAttribute = xmlImagePrpNode.Attributes["value"];
                                            if (xmlImageValueAttribute != null)
                                            {
                                                string imageFilename = xmlImageValueAttribute.Value;
                                                if (!String.IsNullOrWhiteSpace(imageFilename))
                                                {
                                                    string imageDir = Path.GetDirectoryName(pathname);
                                                    string imagePath = Path.Combine(imageDir, imageFilename);
                                                    try
                                                    {
                                                        testResultStep.Image = File.ReadAllBytes(imagePath);
                                                        testResultStep.ImageFilename = imageFilename;
                                                    }
                                                    catch (Exception)
                                                    {
                                                        //Fail quietly for now
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Compares the order of two node names by their name="message X" value where X is a numeric order parameter
    /// </summary>
    public class NodeNameComparer : IComparer
    {
        public int Compare(object x, object y)
        {
            string message1 = x.ToString();
            string message2 = y.ToString();

            //Strip of the word "message ";
            try
            {
                int messageNumber1 = Int32.Parse(message1.Replace("message ", ""));
                int messageNumber2 = Int32.Parse(message2.Replace("message ", ""));
                return messageNumber1.CompareTo(messageNumber2);
            }
            catch (FormatException)
            {
                //Cannot parse, so just do a string comparison
                return message1.CompareTo(message2);
            }
        }
    } 

}
