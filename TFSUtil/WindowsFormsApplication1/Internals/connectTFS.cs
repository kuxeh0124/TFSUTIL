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
        /// <summary>
        /// Initiate the connection to the TFS server using the TeamProjectPicker built in library.
        /// </summary>
        public static TeamProjectPicker tfsPP = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false, new UICredentialsProvider());

        /// <summary>
        /// Connects to the serves
        /// </summary>
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
        
        /// <summary>
        /// Returns the team project collection
        /// </summary>
        public static TfsTeamProjectCollection myTfsTeamProjectCollection
        {
            get
            {
                return tfsPP.SelectedTeamProjectCollection;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public static ITestManagementService tfsService
        {
            get
            {
                try
                {
                    return myTfsTeamProjectCollection.GetService<ITestManagementService>();
                }
                catch
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// returns the name of the connected team projecs
        /// </summary>
        public static ITestManagementTeamProject tfsTeamProject
        {
            get
            {
                try
                {
                    return tfsService.GetTeamProject(tfsPP.SelectedProjects[0].Name);
                }
                catch
                {
                    return null;
                }
            }
        }
    }
}
