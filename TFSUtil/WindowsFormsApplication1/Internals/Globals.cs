using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Xml.Linq;

namespace TFSUtil.Internals
{
    class Globals
    {
        public static string getExtractDrive { get; set; }
        public static string getFileName { get; set; }
        public static string testCaseFileName { get; set; }

        public static string getTemplateDrive { get; set; }
        public static string getRTMDrive { get; set; }

        public static string getTCTemplateDrive { get; set; }

        public static string getTCExtractDrive { get; set; }
        public static string getReportPath { get; set; }

        public static string getTestPlan { get; set; }
        public static string getTestCaseFieldsFromSetting { get; set; }
        public static string getDefectFieldsFromSetting { get; set; }

        public static string getRtmDestinationFile { get; set; }
        public static Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore workItemStore
        {
            get
            {
                return connectTFS.myTfsTeamProjectCollection.GetService<WorkItemStore>();
            }
        }

        public static Project getProjectCol
        {
            get
            {
                return workItemStore.Projects[connectTFS.tfsTeamProject.TeamProjectName];
            }
        }

        public static WorkItemTypeCollection workItemTypes
        {
            get
            {
                return getProjectCol.WorkItemTypes;
            }
        }

        public static void AddToDictionary(Dictionary<string, string[]> getDic, string getKey, string getValue, List<string> getList)
        {
            getList.Add(getValue);
            if (getDic.ContainsKey(getKey))
            {
                getDic[getKey] = getList.ToArray();
            }
            else
            {
                getDic.Add(getKey, getList.ToArray());
            }
        }
        public static void loadSettings()
        {
            XElement xdoc = XElement.Load(@"References\ProgramSettings.xml");
            IEnumerable<XElement> xRows = xdoc.Elements();
            // Read the entire XML
            
            foreach (XElement r in xRows)
            {
                getTestCaseFieldsFromSetting = r.Element("TestCaseRef").Value;
                getDefectFieldsFromSetting = r.Element("DefectRef").Value;            
            }
        }
    }
}
