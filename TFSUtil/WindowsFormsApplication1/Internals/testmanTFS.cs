using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using TFSUtil.Internals;

namespace TFSUtil.Internals
{    
    class testmanTFS
    {
        public Dictionary<string, string> getSuiteList = new Dictionary<string, string>();                       
        public Dictionary<string, string> getTCListFromSuite = new Dictionary<string, string>();
        public Dictionary<string, Dictionary<string,string>> getTCInfo = new Dictionary<string, Dictionary<string, string>>();
        List<Dictionary<string, string>> getDicList = new List<Dictionary<string, string>>();        
        public List<string> xmlTestCaseFields = new List<string>();
        public List<string> xmlTestCaseFieldsVal = new List<string>();
        public Dictionary<string, string[]> tcFieldAllowedValues = new Dictionary<string, string[]>();
        static string getSuiteQry = "Select * From TestSuite";
        static string getPlanQuery = "Select * from TestPlan";
        IStaticTestSuite currentTestSuite = null;
        ITestPlan currentTestPlan = null;
        int currentPlanId = 0;
        public void GetTestSuites(string getSuiteNum="")
        {
            if (getSuiteNum.Length > 0)
            {
                getSuiteQry = "Select * From TestSuite where Id=" + getSuiteNum;
            }
            else
            {
                getSuiteQry = "Select * From TestSuite";
            }
            foreach (ITestSuiteBase ts in tsCollection)
            {
                List<string> currentDirectory = new List<string>();
                ITestSuiteBase getParent = null;
                if (!ts.IsRoot)
                {                
                    currentDirectory.Add(ts.Title);                    
                    getParent = ts.Parent;
                    currentDirectory.Add(getParent.Title);
                    while (getParent.Title != ts.Plan.Name)
                    {
                        getParent = getParent.Parent;
                        currentDirectory.Add(getParent.Title);
                    }
                }
                if (ts.TestCases.Count > 0)
                {
                    string suitePath = "";
                    for (int x = currentDirectory.Count-1; x >= 0; x--)
                    {
                        suitePath = suitePath + currentDirectory[x] + @"\";                       
                    }
                    getSuiteList.Add(Convert.ToString(ts.Id), suitePath.Remove(suitePath.LastIndexOf('\\')));
                }
            }
        }

        public void GetTestCaseInformation(string tcID)
        {
            loadXMLTCFields();
            //getTCFields.Clear();            

            ITestCase testCase = connectTFS.tfsTeamProject.TestCases.Find(Convert.ToInt32(tcID));
            Dictionary<string, string> getTCFields = new Dictionary<string, string>();
            var testcaseResult = connectTFS.tfsTeamProject.TestResults.Query("Select * from TestResult where TestCaseID='" + tcID + "'");
            try
            {                
                foreach (string getField in xmlTestCaseFields)
                {
                    if(getField != "Test Plan")
                    {                    
                        if (!getTCFields.ContainsKey(getField))
                        {
                            getTCFields.Add(getField, Convert.ToString(testCase.WorkItem.Fields[getField].Value));
                        }
                    }
                }
            }
            catch(FieldDefinitionNotExistException)
            {
                int stepCtr = 1;
                string fullStep = "";
                string dc="";
                string oc="";
                if (stepCtr == 1)
                {
                    if (testcaseResult.Count > 0)
                    {
                        dc = testcaseResult[testcaseResult.Count - 1].DateCompleted.ToString();
                        oc = testcaseResult[testcaseResult.Count - 1].Outcome.ToString();
                    }
                    getTCFields.Add("Test Outcome", oc);
                    getTCFields.Add("Date Completed", dc);                    
                }
                foreach (ITestAction testStep in testCase.Actions)
                {                   
                    ITestStep ts = (ITestStep)testStep;
                    string getStep = ts.Title.ToString();
                    string getExpRes = ts.ExpectedResult.ToString();
                    if(fullStep.Length==0)
                    {
                        fullStep = stepCtr.ToString() + "<+++>" + getStep + "<--->" + getExpRes;
                    }
                    else
                    {
                        fullStep = fullStep + "<...>" + stepCtr.ToString() + "<+++>" + getStep + "<--->" + getExpRes;
                    }
                    
                    stepCtr++;
                }
                getTCFields.Add("getSteps", fullStep);
            }
            getDicList.Add(getTCFields);
            getTCInfo.Add(tcID, getDicList[getDicList.Count-1]);
        }

        public bool CreateTestCaseExtractFile(string path)
        {
            try
            {
                ExcelProcessing xlProc = new ExcelProcessing();
                string actualFileName = "TestCase" + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".xlsx";
                string fullFilePath = path + "\\" + actualFileName;
                string[] arrSNOID = { "SNo", "ID" };
                xlProc.createNewExcelTemplate("TestCase");
                xlProc.generateTestCaseTemplate();
                xlProc.updateAndCopy(fullFilePath, "TestCase", true, actualFileName);
                //Start Populating excel
                string compDestPath = xlProc.destPath + "TestCase\\" + actualFileName;
                File.Copy(fullFilePath, compDestPath, true);
                xlProc.updateTestCaseExcelData(fullFilePath, "TestCase", getTCInfo);
                xlProc.updateAndCopy(fullFilePath, "TestCase", false, actualFileName);
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        public bool CreateTestCaseExtractFile(string path, bool isSpecial)
        {
            try
            {
                ExcelProcessing xlProc = new ExcelProcessing();
                string actualFileName = "TestCase" + DateTime.Now.ToString("ddMMyyyyhhmmss") + ".xlsx";
                string fullFilePath = path + "\\" + actualFileName;
                string[] arrSNOID = { "SNo", "ID" };
                xlProc.createNewExcelTemplate("TestCase");
                new testmanTFS().getTFSTestCaseFields();
                xlProc.updateAndCopy(fullFilePath, "TestCase", true, actualFileName);
                //Start Populating excel
                string compDestPath = xlProc.destPath + "TestCase\\" + actualFileName;
                File.Copy(fullFilePath, compDestPath, true);
                xlProc.updateTestCaseExcelDataSpecFormat(fullFilePath, "TestCase", getTCInfo);
                xlProc.updateAndCopy(fullFilePath, "TestCase", false, actualFileName);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public void LoadIntoTFS(string getPath)
        {
            ExcelProcessing xlProc = new ExcelProcessing();
            xlProc.readExcelDataForTC(getPath, "Testcase");
            loadXMLTCFieldsForValidation();
            
            int getSteps = 0;
            ITestCase getTc = null;
            bool isNew = false;
            int getExtract = 0;
            if(xlProc.validateFileFormat(xmlTestCaseFieldsVal, getPath))
            {
                foreach (Dictionary<string, string> dicFromListTC in xlProc.extractedTC)
                {
                    List<string> exp = new List<string>();
                    List<string> title = new List<string>();
                    List<string> stpno = new List<string>();
                    foreach (KeyValuePair<string, string> entry in dicFromListTC.Reverse())
                    {
                        switch (Convert.ToString(entry.Key))
                        {
                            case "Test Plan":
                                string getTestPlan = dicFromListTC["Test Plan"].Substring(0, dicFromListTC["Test Plan"].IndexOf(@"\"));
                                string[] arrTestPlanLoc = dicFromListTC["Test Plan"].Split('\\');
                                if (!validateTestPlanExist(getTestPlan))
                                {
                                    createTestPlan(getTestPlan);
                                }
                                createTestSuites(arrTestPlanLoc);
                                break;

                            case "ID":
                                if (string.IsNullOrEmpty(Convert.ToString(entry.Value)))
                                {
                                    ITestCase tc = connectTFS.tfsTeamProject.TestCases.Create();
                                    getTc = tc;
                                    isNew = true;
                                }
                                else
                                {
                                    ITestCase tc = connectTFS.tfsTeamProject.TestCases.Find(Convert.ToInt32(entry.Value));
                                    getTc = tc;
                                    getSteps = tc.Actions.Count;
                                    isNew = false;
                                }
                                break;
                            case "Step No":
                            case "Step Title":
                            case "Step Expected Result":
                                if (Convert.ToString(entry.Key).Contains("No"))
                                {
                                    stpno = getSplitVals(entry.Value);
                                }
                                if (Convert.ToString(entry.Key).Contains("Title"))
                                {
                                    title = getSplitVals(entry.Value);
                                }
                                if (Convert.ToString(entry.Key).Contains("Expected Result"))
                                {
                                    exp = getSplitVals(entry.Value);
                                }
                                if (title.Count > 0 && exp.Count > 0)
                                {
                                    if (getSteps > stpno.Count)
                                    {
                                        int getStepDif = getSteps - stpno.Count;
                                        for (int dif = 1; dif <= getStepDif; dif++)
                                        {
                                            getTc.Actions.RemoveAt(dif);                        
                                        }
                                    }
                                    else if (getSteps < stpno.Count)
                                    {
                                        int getStepDif = stpno.Count - getSteps;
                                        for (int dif = 1; dif <= getStepDif; dif++)
                                        {
                                            ITestStep newStep = getTc.CreateTestStep();
                                            getTc.Actions.Add(newStep);
                                        }
                                    }
                                    int stind = 0;
                                    foreach (ITestAction testStep in getTc.Actions)
                                    {
                                        ITestStep ts = (ITestStep)testStep;
                                        ts.Title = title[stind];
                                        ts.ExpectedResult = exp[stind];
                                        stind++;
                                    }
                                }
                                break;
                            case "SNo":
                            case "Test Outcome":
                            case "Date Completed":
                                break;
                            default:
                                getTc.WorkItem.Fields[Convert.ToString(entry.Key)].Value = Convert.ToString(entry.Value);
                                break;
                        }
                    }
                    if (isNew)
                    {
                        try
                        {
                            IdAndName defaultConfigIdAndName = new IdAndName(defConfig.Id, defConfig.Name);
                            currentTestSuite.SetDefaultConfigurations(new IdAndName[] { defaultConfigIdAndName });
                            getTc.Save();
                            currentTestSuite.Entries.Add(getTc);
                            currentTestPlan.Save();
                            xlProc.updateTestCaseTestID(getPath, getTc.Id, xlProc.rowTestId[getExtract]);
                        }
                        catch (TestManagementValidationException tme)
                        {
                            Console.WriteLine(tme.ToString());
                            Console.WriteLine(getTc.Title);
                        }
                    }
                    else
                    {
                        getTc.Save();
                    }
                    getExtract++;
                }
            }
        }

        private List<string> getSplitVals(string inputVal)
        {
            string[] arrGetVals = inputVal.Split(new string[] { "<...>" }, StringSplitOptions.None);
            return arrGetVals.ToList();
        }

        private IStaticTestSuite FindSuite(ITestSuiteEntryCollection collection, string title)
        {
            foreach (ITestSuiteEntry entry in collection)
            {
                IStaticTestSuite suite = entry.TestSuite as IStaticTestSuite;

                if (suite != null)
                {
                    if (suite.Title == title)
                        return suite;
                    else if (suite.Entries.Count > 0)
                        FindSuite(suite.Entries, title);
                }
            }
            return connectTFS.tfsTeamProject.TestSuites.CreateStatic();
        }

        private void createTestSuites(string[] arrTestPlanLoc)
        {
            ITestPlan plan = connectTFS.tfsTeamProject.TestPlans.Find(currentPlanId);
            currentTestPlan = plan;
            IStaticTestSuite newSuite = connectTFS.tfsTeamProject.TestSuites.CreateStatic();
            for (int x = 1; x <= arrTestPlanLoc.Length - 1; x++)
            {
                if (x == 1)
                {
                    ITestSuiteEntryCollection collection = plan.RootSuite.Entries;
                    newSuite = FindSuite(collection, arrTestPlanLoc[x]);
                    newSuite.Title = arrTestPlanLoc[x];
                    plan.RootSuite.Entries.Add(newSuite);
                    plan.Save();
                }
                else
                {
                    ITestSuiteEntryCollection collection = newSuite.Entries;
                    IStaticTestSuite subSuite = FindSuite(collection, arrTestPlanLoc[x]);
                    if (subSuite.Id == 0)
                    {
                        subSuite.Title = arrTestPlanLoc[x];
                        newSuite.Entries.Add(subSuite);
                        newSuite = subSuite;
                        currentTestSuite = subSuite;
                        plan.Save();
                    }
                    else
                    {
                        newSuite = subSuite;
                        currentTestSuite = subSuite;
                    }
                }
            }
        }
        
        private void createTestPlan(string testPlanName)
        {
            ITestPlan plan = connectTFS.tfsTeamProject.TestPlans.Create();
            plan.Name = testPlanName;
            plan.StartDate = DateTime.Now;
            plan.EndDate = DateTime.Now.AddMonths(2);
            plan.Save();
            currentPlanId = plan.Id;
        }

        private bool validateTestPlanExist(string getTestPlan)
        {
            bool found = false;
            if(getTestPlans.Count>0)
            {
                foreach(ITestPlan getPlan in getTestPlans)
                {
                    if(getPlan.Name== getTestPlan)
                    {
                        found = true;
                        currentPlanId = getPlan.Id;
                        return true;
                    }
                }
                if (!found)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return false;
            }
        }
        private bool validateTestSuiteExist(string tsuite)
        {
            getSuiteQry = "Select * from TestSuite where Title='" + tsuite + "'";
            if (tsCollection.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void GetTestCases(string suitWkItmId)
        {
            ITestSuiteEntryCollection allTestCase = connectTFS.tfsTeamProject.TestSuites.Find(Convert.ToInt32(suitWkItmId)).TestCases;
            foreach(ITestSuiteEntry getTC in allTestCase)
            {
                getTCListFromSuite.Add(Convert.ToString(getTC.Id), Convert.ToString(getTC.Id) + " - " + Convert.ToString(getTC.Title));
            }
        }

        public void getTFSTestCaseFields()
        {
            WorkItem workItem = new WorkItem(testCaseWorkItemType);
            List<string> allowedVals = new List<string>();

            loadXMLTCFields();

            foreach (Field dField in workItem.Fields)
            {
                Console.WriteLine(dField.Name);
                allowedVals.Clear();
                if (xmlTestCaseFields.Contains(dField.Name))
                {

                    if (dField.AllowedValues.Count > 0 || dField.Name == "Area Path" || dField.Name == "Iteration Path")
                    {
                        for (int av = 0; av <= dField.AllowedValues.Count - 1; av++)
                        {
                            allowedVals.Add(dField.AllowedValues[av].ToString());
                        }
                        processCustomListItems(dField, allowedVals);
                        tcFieldAllowedValues[dField.Name] = allowedVals.ToArray();
                    }
                }
            }
        }

        private void processCustomListItems(Field dField, List<string> vals)
        {

            if (dField.Name == "State")
            {
                vals.Clear();
                if (vals.Count < 6)
                {
                    vals.Add("New");
                    vals.Add("Assigned");
                    vals.Add("OnHold");
                    vals.Add("Rejected");
                    vals.Add("Resolved");
                    vals.Add("Closed");
                    vals.Add("Reopened");
                    vals.Add("Draft");
                    vals.Add("Design");
                }
            }

            if (dField.Name == "Area Path")
            {
                vals.Clear();
                foreach (Node getArea in Globals.getProjectCol.AreaRootNodes)
                {
                    if (getArea.HasChildNodes)
                    {
                        foreach (Node children in getArea.ChildNodes)
                        {
                            vals.Add(children.Path.ToString());
                        }
                    }
                    else
                    {
                        vals.Add(getArea.Path.ToString());
                    }
                }
            }

            if (dField.Name == "Iteration Path")
            {
                vals.Clear();
                foreach (Node getIteration in Globals.getProjectCol.IterationRootNodes)
                {
                    vals.Add(getIteration.Path.ToString());
                }
            }
        }

        public void loadXMLTCFields()
        {
            XDocument xdoc = XDocument.Load(@"References\" + Globals.getTestCaseFieldsFromSetting + ".xml");
            //XDocument xdoc = XDocument.Load(@"References\TestCaseFields.xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;

            foreach (XElement r in xRows)
            {
                if (!xmlTestCaseFields.Contains(r.Value))
                {
                    xmlTestCaseFields.Add(r.Value);
                }
            }
        }

        public void loadXMLTCFields(string getFile)
        {
            XDocument xdoc = XDocument.Load(@"References\" + getFile);
            //XDocument xdoc = XDocument.Load(@"References\TestCaseFields.xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;

            foreach (XElement r in xRows)
            {
                if (!xmlTestCaseFields.Contains(r.Value))
                {
                    xmlTestCaseFields.Add(r.Value);
                }
            }
        }

        public void loadXMLTCFieldsForValidation()
        {
            XDocument xdoc = XDocument.Load(@"References\" + Globals.getTestCaseFieldsFromSetting + ".xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;
            xmlTestCaseFieldsVal.Add("SNo");
            xmlTestCaseFieldsVal.Add("ID");
            foreach (XElement r in xRows)
            {
                if (!xmlTestCaseFieldsVal.Contains(r.Value))
                {
                    xmlTestCaseFieldsVal.Add(r.Value);
                }
            }
        }

        public static WorkItemType testCaseWorkItemType
        {
            get
            {
                return Globals.workItemTypes["Test Case"];
            }
        }

        public static ITestSuiteCollection tsCollection
        {
            get
            {
                return getTestSuites.Query(getSuiteQry);
            }
        }
        //public static ITestSuiteEntryCollection tsEntryCollection
        //{
        //    get
        //    {
        //        return 
        //    }
        //}
        public static ITestSuiteHelper getTestSuites
        {
            get
            {
                return connectTFS.tfsTeamProject.TestSuites;
            }
        }

        public static ITestPlanCollection getTestPlans
        {
            get
            {
                return connectTFS.tfsTeamProject.TestPlans.Query(getPlanQuery);
            }
        }

        public static ITestConfiguration defConfig
        {
            get
            {
                foreach (ITestConfiguration config in connectTFS.tfsTeamProject.TestConfigurations.Query(
                    "Select * from TestConfiguration"))
                {
                    return config;                    
                }
                return defConfig;
            }
        }
    }
}
