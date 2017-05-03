using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace TFSUtil
{
    public partial class CustomFieldConfig : Form
    {
        public CustomFieldConfig()
        {
            InitializeComponent();
            loadValues();
        }

        private void loadValues()
        {
            //CustomFieldConfig cfc = new CustomFieldConfig();
            XDocument xdoc = XDocument.Load(@"References\CustomFields.xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;

            foreach (XElement r in xRows)
            {
                if(Convert.ToString(r.Name).Contains("TextField"))
                {
                    this.groupBox1.Controls["txt_" + r.Name].Text = r.Value;
                }           
                else
                {
                    this.groupBox2.Controls["txt_" + r.Name].Text = r.Value;
                }
            }            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //var textBoxes = groupBox1.Controls.OfType<TextBox>();
            XDocument xdoc = XDocument.Load(@"References\CustomFields.xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;
            foreach (XElement r in xRows)
            {
                if (Convert.ToString(r.Name).Contains("TextField"))
                {
                    r.SetValue(this.groupBox1.Controls["txt_" + r.Name].Text);
                }                
            }
            xdoc.Save(@"References\CustomFields.xml");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //var textBoxes = groupBox1.Controls.OfType<TextBox>();
            XDocument xdoc = XDocument.Load(@"References\CustomFields.xml");
            var xRows = from xRow in xdoc.Descendants("Row") select xRow.FirstNode;
            foreach (XElement r in xRows)
            {
                if (Convert.ToString(r.Name).Contains("Remarks"))
                {                    
                    r.SetValue(this.groupBox2.Controls["txt_" + r.Name].Text);
                }
            }
            xdoc.Save(@"References\CustomFields.xml");
        }
    }
}
