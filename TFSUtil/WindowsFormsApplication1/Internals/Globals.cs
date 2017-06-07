using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Xml.Linq;
using System.Runtime.CompilerServices;
using System.Diagnostics;
using System.Windows.Forms;

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
        public static bool isConnected { get; set; }
        public static bool isTestCase { get; set; }

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
        [MethodImpl(MethodImplOptions.NoInlining)]
        public static string GetCurrentMethod()
        {
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }

        /// <summary>
        /// Displays error messages
        /// typeMsg Parameter Definition:
        /// 1 - Message box with OK Button and Error Icon
        /// 2 - Message box with OK and Information Icon
        /// 3 - Message box with Ok and Warning Icon
        /// 4 - Message Box with OK and Cancel and Question Icon
        /// 5 - Message Box with Yes and No and Question Icon
        /// 6 - Message Box with Yes, No and Cancel and Question Icon
        /// </summary>
        /// <param name="errorMsg"></param>
        /// <param name="getCaption"></param>
        /// <param name="typeMsg"></param>
        public static void DisplayErrorMessage(string errorMsg, string getCaption, int typeMsg)
        {
            switch(typeMsg)
            {
                case 1:
                    MessageBox.Show(errorMsg, getCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case 2:
                    MessageBox.Show(errorMsg, getCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                case 3:
                    MessageBox.Show(errorMsg, getCaption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
                case 4:
                    MessageBox.Show(errorMsg, getCaption, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    break;
                case 5:
                    MessageBox.Show(errorMsg, getCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    break;
                case 6:
                    MessageBox.Show(errorMsg, getCaption, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    break;
            }
            
        }
    }
}
