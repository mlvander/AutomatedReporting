using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PropelProfileGenerator
{
    public partial class ProjectForm : Form
    {
        
        public string iniFilePath;

        public ProjectForm(string iniFilePath)
        {
            InitializeComponent();
        }

        private void loadData()
        {
            /*this.active.Text = project.active.ToString();
            this.projectName.Text = project.projectName;
            this.projectAbbr.Text = project.projectAbbr;
            this.templateFolder.Text = project.templateFolder;
            this.excelFolder.Text = project.excelFolder;
            this.profileFolder.Text = project.profileFolder;
            this.database.Text = project.database;
            this.username.Text = project.username;
            this.password.Text = project.password;
            this.table.Text = project.table;
            this.tableID.Text = project.tableID;
            this.tableTitle.Text = project.tableTitle;
            this.dataReady.Text = project.dataReady;
            this.profileGenerated.Text = project.profileGenerated;
            this.profilePosted.Text = project.profilePosted;
            */

        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
