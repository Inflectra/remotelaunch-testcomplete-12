using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Inflectra.RemoteLaunch.Interfaces;
using System.Diagnostics;
using SmartBear.TestComplete12;
using Inflectra.RemoteLaunch.Interfaces.DataObjects;
using System.Threading;

namespace Inflectra.RemoteLaunch.Engines.TestCompleteEngine.TestComplete12
{
    /// <summary>
    /// Implements the IAutomationEngine class for integration with TestComplete 12.x by SmartBear
    /// This class is instantiated by the RemoteLaunch application
    /// </summary>
    /// <remarks>
    /// The AutomationEngine class provides some of the generic functionality
    /// </remarks>
    public class TestCompleteEngine : AutomationEngine, IAutomationEngine4
    {
        private const string CLASS_NAME = "TestCompleteEngine";

        private const string AUTOMATION_ENGINE_TOKEN = "TestComplete12";
        private const string AUTOMATION_ENGINE_VERSION = "4.0.0";

        /// <summary>
        /// Constructor
        /// </summary>
        public TestCompleteEngine()
        {
            //Set status to OK
            base.status = EngineStatus.OK;
        }

        /// <summary>
        /// Returns the author of the test automation engine
        /// </summary>
        public override string ExtensionAuthor
        {
            get
            {
                return "Inflectra Corporation";
            }
        }

        /// <summary>
        /// The unique GUID that defines this automation engine
        /// </summary>
        public override Guid ExtensionID
        {
            get
            {
                return new Guid("{84C9E71B-B6FC-4E3F-8A20-261F6B28D885}");
            }
        }

        /// <summary>
        /// Returns the display name of the automation engine
        /// </summary>
        public override string ExtensionName
        {
            get
            {
                return "TestComplete 12 Automation Engine";
            }
        }

        /// <summary>
        /// Returns the unique token that identifies this automation engine to SpiraTest
        /// </summary>
        public override string ExtensionToken
        {
            get
            {
                return AUTOMATION_ENGINE_TOKEN;
            }
        }

        /// <summary>
        /// Returns the version number of this extension
        /// </summary>
        public override string ExtensionVersion
        {
            get
            {
                return AUTOMATION_ENGINE_VERSION;
            }
        }

        /// <summary>
        /// Adds a custom settings panel for configuring the TestComplete engine
        /// </summary>
        public override System.Windows.UIElement SettingsPanel
        {
            get
            {
                return new EngineSettings();
            }
            set
            {
                EngineSettings engineSettings= (EngineSettings)value;
                engineSettings.SaveSettings();
            }
        }

        public override AutomatedTestRun StartExecution(AutomatedTestRun automatedTestRun)
        {
            //Not used since we implement the V4 API instead
            throw new NotImplementedException();
        }

        /// <summary>
        /// This is the main method that is used to start automated test execution
        /// </summary>
        /// <param name="automatedTestRun">The automated test run object</param>
        /// <param name="projectId">The id of the project</param>
        /// <returns>Either the populated test run or an exception</returns>
        public AutomatedTestRun4 StartExecution(AutomatedTestRun4 automatedTestRun, int projectId)
        {
            //Set status to OK
            base.status = EngineStatus.OK;

            try
            {
                //Instantiate the TestComplete wrapper class
                using (TestComplete testComplete = new TestComplete())
                {
                    //See if we have an attached or linked test script
                    if (automatedTestRun.Type == AutomatedTestRun4.AttachmentType.URL)
                    {
                        //The "URL" of the test is actually three pipe-separated components:
                        //Suite Filename|Project Name|Project Item Name

                        //Or if we want to use parameters we have to pass the actual unit and routine name
                        //Suite Filename|Project Name|Unit Name|Routine Name

                        //Extract upto three components from the URL
                        //Project Suite|Project Name|Project Item Name
                        //only the first one is mandatory
                        string[] components = automatedTestRun.FilenameOrUrl.Split('|');

                        /* For executing a project item */
                        //string suiteFilename = @"[CommonDocuments]\TestComplete 7 Samples\Open Apps\OrdersDemo\C#\TCProject\Orders.pjs";
                        //string projectName = "Orders_C#_C#Script";
                        //string projectItemName = "ProjectTestItem1";

                        /* For executing a specific routine */
                        //string suiteFilename = @"[CommonDocuments]\TestComplete 7 Samples\Scripts\Hello\Hello.pjs";
                        //string projectName = "Hello_C#Script";
                        //string unitName = "hello_cs";
                        //string routineName = "Hello";

                        if (components.Length >= 4)
                        {
                            string suiteFilename = components[0];
                            string projectName = components[1];
                            string unitName = components[2];
                            string routineName = components[3];

                            //Get any parameters
                            //TC uses ordered parameters rather than name parameters, so we need to sort the list by key
                            SortedList<string, string> parameters = new SortedList<string, string>();
                            if (automatedTestRun.Parameters != null)
                            {
                                foreach (TestRunParameter testRunParameter in automatedTestRun.Parameters)
                                {
                                    if (!parameters.ContainsKey(testRunParameter.Name))
                                    {
                                        parameters.Add(testRunParameter.Name, testRunParameter.Value);
                                    }
                                }
                            }

                            //Run the TestComplete test
                            testComplete.ExecuteEx(suiteFilename, projectName, unitName, routineName, parameters);

                            //Extract the project name
                            automatedTestRun.RunnerTestName = projectName + " > " + unitName + " > " + routineName;
                        }
                        else if (components.Length >= 3)
                        {
                            string suiteFilename = components[0];
                            string projectName = components[1];
                            string projectItemName = components[2];

                            //Run the TestComplete test
                            testComplete.Execute(suiteFilename, projectName, projectItemName);

                            //Extract the project name
                            automatedTestRun.RunnerTestName = projectName + " > " + projectItemName;
                        }
                        else if (components.Length >= 2)
                        {
                            string suiteFilename = components[0];
                            string projectName = components[1];

                            //Run the TestComplete test
                            testComplete.Execute(suiteFilename, projectName);

                            //Extract the project name
                            automatedTestRun.RunnerTestName = projectName;
                        }

                        else if (components.Length >= 1)
                        {
                            string suiteFilename = components[0];

                            //Run the TestComplete test
                            testComplete.Execute(suiteFilename);

                            //Extract the project name
                            automatedTestRun.RunnerTestName = suiteFilename;
                        }
                    }
                    else
                    {
                        //We have an embedded script which we need to execute directly
                        //This is not currently supported since TC requires a complex set of
                        //project files which are not easily attached
                        throw new InvalidOperationException("The TestComplete automation engine only supports linked test scrips");
                    }

                    //If no result, throw an exception
                    if (testComplete.Result == null)
                    {
                        throw new ApplicationException("No result returned from test execution");
                    }

                    //Now extract the status and populate the test run object
                    if (String.IsNullOrEmpty(automatedTestRun.RunnerName))
                    {
                        automatedTestRun.RunnerName = this.ExtensionName;
                    }
                    automatedTestRun.RunnerAssertCount = testComplete.Result.ErrorCount + testComplete.Result.WarningCount;
                    automatedTestRun.RunnerMessage = testComplete.Result.ErrorCount + " errors and " + testComplete.Result.WarningCount + " warnings.";
                    switch (testComplete.Result.ExecutionStatus)
                    {
                        case TC_LOG_STATUS.lsOk:
                            automatedTestRun.ExecutionStatus = AutomatedTestRun4.TestStatusEnum.Passed;
                            break;

                        case TC_LOG_STATUS.lsError:
                            automatedTestRun.ExecutionStatus = AutomatedTestRun4.TestStatusEnum.Failed;
                            break;

                        case TC_LOG_STATUS.lsWarning:
                            automatedTestRun.ExecutionStatus = AutomatedTestRun4.TestStatusEnum.Caution;
                            break;
                    }
                    if (testComplete.Result.StartTime.HasValue)
                    {
                        automatedTestRun.StartDate = testComplete.Result.StartTime.Value;
                    }
                    else
                    {
                        automatedTestRun.StartDate = DateTime.Now;
                    }

                    if (testComplete.Result.EndTime.HasValue)
                    {
                        automatedTestRun.EndDate = testComplete.Result.EndTime.Value;
                    }
                    else
                    {
                        automatedTestRun.EndDate = DateTime.Now;
                    }

                    //The result log (we send back the results as Plain Text)
                    automatedTestRun.Format = AutomatedTestRun4.TestRunFormat.PlainText;
                    automatedTestRun.RunnerStackTrace = testComplete.Result.MessageLog;

                    //The detailed 'test steps' and screenshows
                    if (testComplete.Result.Steps != null && testComplete.Result.Steps.Count > 0)
                    {
                        automatedTestRun.TestRunSteps = new List<TestRunStep4>();
                        automatedTestRun.Screenshots = new List<TestRunScreenshot4>();
                        int position = 1;
                        foreach (TestResultStep testResultStep in testComplete.Result.Steps)
                        {
                            TestRunStep4 testRunStep = new TestRunStep4();
                            switch (testResultStep.Type)
                            {
                                case TestResult.TestResultEntryType.Message:
                                case TestResult.TestResultEntryType.Checkpoint:
                                    testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.Passed;
                                    break;

                                case TestResult.TestResultEntryType.Warning:
                                    testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.Caution;
                                    break;

                                case TestResult.TestResultEntryType.Error:
                                    testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.Failed;
                                    break;

                                case TestResult.TestResultEntryType.Event:
                                    testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.NotApplicable;
                                    break;

                                default:
                                    testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.NotApplicable;
                                    break;
                            }
                            testRunStep.Description = testResultStep.Message;
                            testRunStep.ActualResult = testResultStep.AdditionalInfo;
                            testRunStep.Position = position++;
                            automatedTestRun.TestRunSteps.Add(testRunStep);

                            //Check for screenshots
                            if (!String.IsNullOrWhiteSpace(testResultStep.ImageFilename) && testResultStep.Image != null && testResultStep.Image.Length > 0)
                            {
                                TestRunScreenshot4 screenshot = new TestRunScreenshot4();
                                screenshot.Data = testResultStep.Image;
                                screenshot.Filename = testResultStep.ImageFilename;
                                screenshot.Description = testResultStep.Message;
                                automatedTestRun.Screenshots.Add(screenshot);
                            }
                        }
                    }
                }

                //It can take TC several seconds to close down, so we need to build in a pause so that the next test using TC
                //doesn't fail with a TC already running error
                Thread.Sleep(Properties.Settings.Default.WaitTimeMilliseconds);

                //Report as complete               
                base.status = EngineStatus.OK;
                return automatedTestRun;
            }
            catch (Exception exception)
            {
                //Log the error and denote failure
                LogEvent(exception.Message + " (" + exception.StackTrace + ")", EventLogEntryType.Error);

                //Report as completed with error
                base.status = EngineStatus.Error;
                throw exception;
            }
        }
    }
}
