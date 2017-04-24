using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace TFSUtil.Internals
{
    class connectTFS
    {
        public static TeamProjectPicker tfsPP = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false, new UICredentialsProvider());
        public static void connectToTFS()
        {
            Uri tfsUri = null;
            try
            {
                tfsPP.ShowDialog();                
                if (tfsPP.SelectedTeamProjectCollection != null)
                {
                    tfsUri = tfsPP.SelectedTeamProjectCollection.Uri;
                }                
            }
            catch { }            
        }

        public static TfsTeamProjectCollection myTfsTeamProjectCollection
        {
            get
            {
                return tfsPP.SelectedTeamProjectCollection;
            }
        }

        public static ITestManagementService tfsService
        {
            get
            {
                return myTfsTeamProjectCollection.GetService<ITestManagementService>();
            }
        }

        public static ITestManagementTeamProject tfsTeamProject
        {
            get
            {
                return tfsService.GetTeamProject(tfsPP.SelectedProjects[0].Name);
            }
        }
    }
}
