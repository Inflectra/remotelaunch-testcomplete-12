using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using SmartBear.TestComplete12;
using System.Runtime.InteropServices;
using System.Threading;

namespace Inflectra.RemoteLaunch.Engines.TestCompleteEngine.TestComplete12
{
	/// <summary>
	/// The primary class that handles interaction with TestComplete
	/// </summary>
	public class TestComplete : IDisposable
	{
		protected TestCompleteApplication tcApplication;
		protected IaqBaseManager tcBaseManager;
		protected TestResult result;

		/// <summary>
		/// Constructor
		/// </summary>
		public TestComplete()
		{
            //First see if we can access a running instance (better for performance that way)
            try
            {
                //TestComplete.TestCompleteApplication
                //TestExecute.TestExecuteApplication
                this.tcApplication = (TestCompleteApplication)System.Runtime.InteropServices.Marshal.GetActiveObject("TestComplete.TestCompleteApplication");
            }
            catch (COMException)
            {
                //Instantiate the TC API class instead
                this.tcApplication = new TestCompleteApplication();
            }

			this.tcApplication.Visible = Properties.Settings.Default.ApplicationVisible;

			//Get a handle to the base manager
			this.tcBaseManager = this.tcApplication.Manager;
            
            //Run in silent mode
            this.tcBaseManager.RunMode = TaqRunMode.rmSilent;
		}

		/// <summary>
		/// The last test result
		/// </summary>
		public TestResult Result
		{
			get
			{
				return this.result;
			}
		}

		/// <summary>
		/// Opens and executes a specific item in a specific project
		/// </summary>
		/// <param name="suiteFilename">The suite filename</param>
		public void Execute(string suiteFilename)
		{
			Execute(suiteFilename, null, null);
		}

		/// <summary>
		/// Opens and executes a specific item in a specific project
		/// </summary>
		/// <param name="suiteFilename">The suite filename</param>
		/// <param name="projectName">The project name</param>
		public void Execute(string suiteFilename, string projectName)
		{
			Execute(suiteFilename, projectName, null);
		}

		/// <summary>
		/// Opens and executes a specific item in a specific project
		/// </summary>
		/// <param name="suiteFilename">The suite filename</param>
		/// <param name="projectItemName">The project item name</param>
		/// <param name="projectName">The project name</param>
		public void Execute(string suiteFilename, string projectName, string projectItemName)
		{
			//To make it easier, we have certain shortcuts that can be used in the filename
			suiteFilename = suiteFilename.Replace("[MyDocuments]", Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments));
			suiteFilename = suiteFilename.Replace("[CommonDocuments]", Environment.GetFolderPath(System.Environment.SpecialFolder.CommonDocuments));
			suiteFilename = suiteFilename.Replace("[DesktopDirectory]", Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory));

			//Get a handle to the integration object
			ItcIntegration tcIntegration = this.tcApplication.Integration;

			//First open the project suite
			bool success = tcIntegration.OpenProjectSuite(suiteFilename);
			if (!success || !tcIntegration.IsProjectSuiteOpened())
			{
				throw new ApplicationException("Unable to open project suite - '" + suiteFilename + "'");
			}

			//Next try and run the tests
			if (String.IsNullOrEmpty(projectName))
			{
				tcIntegration.RunProjectSuite();
			}
			else if (String.IsNullOrEmpty(projectItemName))
			{
				tcIntegration.RunProject(projectName);
			}
			else
			{
				tcIntegration.RunProjectTestItem(projectName, projectItemName);
			}

			//Wait until it finishes
			bool running = true;
			MessageFilter.Register();
			while (running)
			{
				try
				{
					running = tcIntegration.IsRunning();
				}
				catch (COMException)
				{
					//Ignore and carry on
				}

				System.Threading.Thread.Sleep(1);
			}

			MessageFilter.Revoke();


			//Check the results
			ItcIntegrationResultDescription resultDesc = tcIntegration.GetLastResultDescription();
			if (resultDesc == null)
			{
				throw new ApplicationException("No result description available");
			}

			//Get the execution status and summary data
			this.result = new TestResult();
			this.result.ExecutionStatus = resultDesc.Status;
			if (resultDesc.StartTime is DateTime)
			{
				this.result.StartTime = resultDesc.StartTime;
			}
			if (resultDesc.EndTime is DateTime)
			{
				this.result.EndTime = resultDesc.EndTime;
			}
			this.result.IsTestCompleted = resultDesc.IsTestCompleted;
			this.result.ErrorCount = resultDesc.ErrorCount;
			this.result.WarningCount = resultDesc.WarningCount;
			this.result.TestType = resultDesc.TestType;

			//Now we need to extract the log
            this.result.ExtractLog(resultDesc.LogFileName, projectItemName, false);

			//Close the project suite
			tcIntegration.CloseProjectSuite();
		}

		/// <summary>
		/// Opens and executes a specific script routine in a specific project with parameters
		/// </summary>
		/// <param name="suiteFilename">The suite filename</param>
		/// <param name="unitName">The project script name</param>
		/// <param name="projectName">The project name</param>
		/// <param name="routineName">The name of the actual script routine</param>
		/// <param name="parameters">The parameters to be passed to the script</param>
		public void ExecuteEx(string suiteFilename, string projectName, string unitName, string routineName, SortedList<string,string> parameters)
		{
			//To make it easier, we have certain shortcuts that can be used in the filename
			suiteFilename = suiteFilename.Replace("[MyDocuments]", Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments));
			suiteFilename = suiteFilename.Replace("[CommonDocuments]", Environment.GetFolderPath(System.Environment.SpecialFolder.CommonDocuments));
			suiteFilename = suiteFilename.Replace("[DesktopDirectory]", Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory));

			//Get a handle to the integration object
			ItcIntegration tcIntegration = this.tcApplication.Integration;

			//First open the project suite
			bool success = tcIntegration.OpenProjectSuite(suiteFilename);
			if (!success || !tcIntegration.IsProjectSuiteOpened())
			{
				throw new ApplicationException("Unable to open project suite - '" + suiteFilename + "'");
			}

			//Convert the parameters to simple array (sorted by the Spira key)
			List<string> tcParameters = new List<string>();
			if (parameters != null)
			{
				foreach (KeyValuePair<string, string> kvp in parameters)
				{
					tcParameters.Add(kvp.Value);
				}
			}

			//Next try and run the tests
			if (tcParameters == null || tcParameters.Count == 0)
			{
				tcIntegration.RunRoutine(projectName, unitName, routineName);
			}
			else
			{
				tcIntegration.RunRoutineEx(projectName, unitName, routineName, tcParameters.ToArray());
			}

			//Wait until it finishes
			bool running = true;
			MessageFilter.Register();
			while (running)
			{
				try
				{
					running = tcIntegration.IsRunning();
				}
				catch (COMException)
				{
					//Ignore and carry on
				}

				System.Threading.Thread.Sleep(1);
			}

			MessageFilter.Revoke();


			//Check the results
			ItcIntegrationResultDescription resultDesc = tcIntegration.GetLastResultDescription();
			if (resultDesc == null)
			{
				throw new ApplicationException("No result description available");
			}

			//Get the execution status and summary data
			this.result = new TestResult();
			this.result.ExecutionStatus = resultDesc.Status;
			if (resultDesc.StartTime is DateTime)
			{
				this.result.StartTime = resultDesc.StartTime;
			}
			if (resultDesc.EndTime is DateTime)
			{
				this.result.EndTime = resultDesc.EndTime;
			}
			this.result.IsTestCompleted = resultDesc.IsTestCompleted;
			this.result.ErrorCount = resultDesc.ErrorCount;
			this.result.WarningCount = resultDesc.WarningCount;
			this.result.TestType = resultDesc.TestType;

			//Now we need to extract the log
			this.result.ExtractLog(resultDesc.LogFileName, null, true);

			//Close the project suite
			tcIntegration.CloseProjectSuite();
		}

		#region IDisposable Members

		/// <summary>
		/// Called when the class is disposed
		/// </summary>
		public void Dispose()
		{
            try
			{
			    this.tcBaseManager = null;
			    if (this.tcApplication != null)
			    {
                    if (tcApplication.Integration != null && tcApplication.Integration.IsProjectSuiteOpened())
				    {
					    tcApplication.Integration.CloseProjectSuite();
				    }

                    //Close the app if specified
                    if (Properties.Settings.Default.CloseApplicationAfterTest)
                    {
                        this.tcApplication.Quit();
                        this.tcApplication = null;
                        this.tcBaseManager = null;
                    }
			    }
			}
			catch { }
		}

		#endregion
	}
}
