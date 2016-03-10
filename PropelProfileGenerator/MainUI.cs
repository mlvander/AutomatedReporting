using PropelProfileGenerator.Properties;
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
    public partial class MainUI : Form
    {
        string iniFilePath = @"\\ahsfile\Propel$\Active Projects\PropelProfileGenerator_V3";
        List<Project> projects = new List<Project>();
        Project selectedProject;
        Organization selectedOrganization;

        public MainUI()
        {
            InitializeComponent();
            loadProjects();
            
        }
        private void loadProjects()
        {
            projectSelectionPanel.Controls.Clear();

            Projects projectList = new Projects(iniFilePath, this);
            projects = projectList.getProjects();

            int yPos = 10;
            foreach (Project thisProject in projects)
            {
                Button newButton = new Button();

                newButton.Location = new System.Drawing.Point(10, yPos);
                newButton.Name = "button" + thisProject.projectAbbr;
                newButton.Size = new System.Drawing.Size(146, 23);
                newButton.TabIndex = 0;
                newButton.Text = thisProject.projectName;
                newButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
                newButton.UseVisualStyleBackColor = true;
                newButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
                newButton.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
                newButton.Click += new System.EventHandler(this.projectButton_Click);

                yPos += 25;

                projectSelectionPanel.Controls.Add(newButton);
            }
        }
        private void setSelectedProject(Project project)
        {
            editProjectButton.Enabled = true;
            // Everytime the selected project changes several UI changes are required.
            selectedProject = project;
            //Start by clearing the panel
            profileTypePanel.Controls.Clear();
            int yPos = 10;
            // Display the different types of profiles that could be run
            foreach (Profile profile in selectedProject.profiles)
            {
                if (profile.active)
                {
                    CheckBox newCheckBox = new CheckBox();
                    newCheckBox.AutoSize = true;
                    newCheckBox.Location = new System.Drawing.Point(10, yPos);
                    newCheckBox.Name = "checkBox_" + profile.profileID;
                    newCheckBox.Size = new System.Drawing.Size(80, 17);
                    newCheckBox.TabIndex = 6;
                    newCheckBox.Text = profile.profileName;
                    newCheckBox.UseVisualStyleBackColor = true;
                    newCheckBox.Checked = true;
                    newCheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    newCheckBox.Click += new System.EventHandler(this.updateSelectedProfiles);
                                        
                    profileTypePanel.Controls.Add(newCheckBox);
                    yPos += 25;
                }
            }
            // get list of orgs with profile data ready
            this.organizationListReady.Items.Clear(); 
            Organizations orgListReady = new Organizations(selectedProject, this, false, false);
            List<Organization> orgDetailsReady = orgListReady.getOrgDetails();
            this.profileDataReady.Text = "Schools Data Ready (" + orgDetailsReady.Count.ToString() + ")";
            foreach (Organization org in orgDetailsReady)
            {
                string[] row = { org.ID, org.name, org.dataReady };

                ListViewItem item = new ListViewItem(row);
                this.organizationListReady.Items.Add(item);
                this.organizationListReady.View = View.Details; 
            }
            
            // get list of orgs with profiles generated
            this.organizationListGenerated.Items.Clear();
            Organizations orgListGenerated = new Organizations(selectedProject, this, true, false);
            List<Organization> orgDetailsGenerated = orgListGenerated.getOrgDetails();
            this.profileBuilt.Text = "Schools Profiles Built (" + orgDetailsGenerated.Count.ToString() + ")";
            foreach (Organization org in orgDetailsGenerated)
            {
                string[] row = { org.ID, org.name, org.profileGenerated };

                ListViewItem item = new ListViewItem(row);
                this.organizationListGenerated.Items.Add(item);
                this.organizationListGenerated.View = View.Details;
            }
            // get list of orgs with profiles posted
            this.organizationListPosted.Items.Clear();
            Organizations orgListPosted = new Organizations(selectedProject, this, false, true);
            List<Organization> orgDetailsPosted = orgListPosted.getOrgDetails();
            this.profilePosted.Text = "Schools Profiles Posted (" + orgDetailsPosted.Count.ToString() + ")";
            foreach (Organization org in orgDetailsPosted)
            {
                string[] row = { org.ID, org.name, org.profilePosted };

                ListViewItem item = new ListViewItem(row);
                this.organizationListPosted.Items.Add(item);
                this.organizationListPosted.View = View.Details;
            }
        }
        private void setSelectedOrganization(string id)
        {
            Organizations getSelectedOrg = new Organizations(selectedProject, id);
            selectedOrganization = getSelectedOrg.getOrgDetails()[0];
                
            this.log(selectedOrganization.ID + " - " + selectedOrganization.name + " selected.");
        }   
        private void projectButton_Click(object sender, EventArgs e)
        {

            foreach (Button thisButton in projectSelectionPanel.Controls)
            {
                thisButton.Image = null;
                thisButton.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            }
            foreach (Project thisProject in projects)
            {
                if(thisProject.projectName == ((Button)sender).Text)
                {
                    this.setSelectedProject(thisProject);
                }
                ((Button)sender).Image = ((System.Drawing.Image)(Resources.ResourceManager.GetObject("checkmark_16x16")));
                ((Button)sender).Padding = new System.Windows.Forms.Padding(0, 0, 0, 0);
                
            }
        }
        private void updateSelectedProfiles(object sender, EventArgs e)
        {
            CheckBox thisCheckbox = ((CheckBox)sender);
            for (int i = 0; i < selectedProject.profiles.Count; i++)
            {
                if(thisCheckbox.Text == selectedProject.profiles[i].profileName)
                {
                    selectedProject.profiles[i].selected = thisCheckbox.Checked;
                }
            }
        }
        private void selectedIndexChanged(object sender, EventArgs e)
        {
            ListView thisList = ((ListView)sender);

            for (int x = 0; x < thisList.Items.Count; x++)
            {
                if (thisList.Items[x].Selected)
                {
                    this.setSelectedOrganization(thisList.Items[x].SubItems[0].Text);
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string templateName;
            string excelFile;
            string saveAs;
            if (selectedOrganization != null)
            {
                bool profileGenerated = false;

                for (int i = 0; i < selectedProject.profiles.Count; i++)
                {
                    if (selectedProject.profiles[i].active && selectedProject.profiles[i].selected)
                    {
                        log("Build Started..");
                        log("Gathering details for " + selectedProject.profiles[i].profileName + "...");

                        templateName = selectedProject.templateFolder + "\\" + selectedProject.profiles[i].profileTemplate;
                        excelFile = selectedProject.excelFolder + "\\" + selectedProject.projectAbbr + "_" + selectedOrganization.ID + ".xlsx";
                        saveAs = selectedProject.profileFolder + "\\" + selectedProject.projectAbbr + "_" + selectedOrganization.ID + "_" + selectedProject.profiles[i].profileName.Replace(" ", "") + ".docx";

                        ProfileDocument newProfile = new ProfileDocument(templateName, excelFile, saveAs, this);
                        if (!profileGenerated)
                        {
                            profileGenerated = newProfile.getStatus();
                        }
                    }
                }
                if (profileGenerated)
                {
                    selectedOrganization.setProfileGenerated(selectedProject);
                    setSelectedProject(selectedProject);
                }
            }
        }
        public void log(string text)
        {
            this.processingLog.Text += text + System.Environment.NewLine;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.processingLog.Text = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ProfileUpload uploadProfile = new ProfileUpload(selectedProject, selectedOrganization, this);
            setSelectedProject(selectedProject);
        }

        private void toggle(object sender, EventArgs e)
        {
            if(((TextBox)sender).Text.ToLower() == "true")
            {
                ((TextBox)sender).Text = "False";
            }
            else
            {
                ((TextBox)sender).Text = "True";
            }
        }
        private void addProjectButton_Click(object sender, EventArgs e)
        {
            mainPanel.Visible = false;
            project_projectPanel.Visible = true;

            project_button1.Visible = true;
            project_button2.Visible = false;

            project_active.Text = "True";
            project_projectName.Text = null;
            project_projectAbbr.Text = null;
            project_templateFolder.Text = null;
            project_excelFolder.Text = null;
            project_profileFolder.Text = null;
            project_database.Text = null;
            project_username.Text = null;
            project_password.Text = null;
            project_table.Text = null;
            project_tableID.Text = null;
            project_tableTitle.Text = null;
            project_dataReady.Text = null;
            project_profileGenerated.Text = null;
            project_profilePosted.Text = null;
        }

        private void editProjectButton_Click(object sender, EventArgs e)
        {
            mainPanel.Visible = false;
            project_projectPanel.Visible = true;

            project_active.Text = selectedProject.active.ToString();
            project_projectName.Text = selectedProject.projectName;
            project_projectAbbr.Text = selectedProject.projectAbbr;
            project_templateFolder.Text = selectedProject.templateFolder;
            project_excelFolder.Text = selectedProject.excelFolder;
            project_profileFolder.Text = selectedProject.profileFolder;
            project_database.Text = selectedProject.database;
            project_username.Text = selectedProject.username;
            project_password.Text = selectedProject.password;
            project_table.Text = selectedProject.table;
            project_tableID.Text = selectedProject.tableID;
            project_tableTitle.Text = selectedProject.tableTitle;
            project_dataReady.Text = selectedProject.dataReady;
            project_profileGenerated.Text = selectedProject.profileGenerated;
            project_profilePosted.Text = selectedProject.profilePosted;

            // Profile Settings
            project_profile_active.Text = selectedProject.profiles[0].active.ToString();
            project_profile_profileID.Text = selectedProject.profiles[0].profileID;
            //project_profile_profileAbbr.Text = selectedProject.profiles[0].profileAbbr;
            project_profile_profileTemplate.Text = selectedProject.profiles[0].profileTemplate;
            project_profile_profileName.Text = selectedProject.profiles[0].profileName;

            // Parent Summary
            project_psum_active.Text = selectedProject.profiles[1].active.ToString();
            project_psum_profileID.Text = selectedProject.profiles[1].profileID;
            //project_psum_profileAbbr.Text = selectedProject.profiles[1].profileAbbr;
            project_psum_profileTemplate.Text = selectedProject.profiles[1].profileTemplate;
            project_psum_profileName.Text = selectedProject.profiles[1].profileName;

            // School Summary
            project_ssum_active.Text = selectedProject.profiles[2].active.ToString();
            project_ssum_profileID.Text = selectedProject.profiles[2].profileID;
            //project_psum_profileAbbr.Text = selectedProject.profiles[2].profileAbbr;
            project_ssum_profileTemplate.Text = selectedProject.profiles[2].profileTemplate;
            project_ssum_profileName.Text = selectedProject.profiles[2].profileName;

            // General Summary
            project_sum_active.Text = selectedProject.profiles[3].active.ToString();
            project_sum_profileID.Text = selectedProject.profiles[3].profileID;
            //project_psum_profileAbbr.Text = selectedProject.profiles[3].profileAbbr;
            project_sum_profileTemplate.Text = selectedProject.profiles[3].profileTemplate;
            project_sum_profileName.Text = selectedProject.profiles[3].profileName;

            project_button1.Visible = false;
            project_button2.Visible = true;
        }

        private void project_button1_Click(object sender, EventArgs e)
        {
            // adding new project
            string newName = iniFilePath + "\\" + project_projectAbbr.Text.Replace(" ","_") + ".ini";
            Project newProject = new Project(newName);
            saveProjectData(newProject);

            mainPanel.Visible = true;
            project_projectPanel.Visible = false;
            // reset the list of project buttons
            loadProjects();
        }

        private void saveProjectData (Project thisProject)
        {
            if (project_active.Text == "True")
            {
                thisProject.active = true;
            }
            else
            {
                thisProject.active = false;
            }
            thisProject.projectName = project_projectName.Text;
            thisProject.projectAbbr = project_projectAbbr.Text;
            thisProject.templateFolder = project_templateFolder.Text;
            thisProject.excelFolder = project_excelFolder.Text;
            thisProject.profileFolder = project_profileFolder.Text;
            thisProject.database = project_database.Text;
            thisProject.username = project_username.Text;
            thisProject.password = project_password.Text;
            thisProject.table = project_table.Text;
            thisProject.tableID = project_tableID.Text;
            thisProject.tableTitle = project_tableTitle.Text;
            thisProject.dataReady = project_dataReady.Text;
            thisProject.profileGenerated = project_profileGenerated.Text;
            thisProject.profilePosted = project_profilePosted.Text;

            // Profile Settings
            if (project_profile_active.Text == "True")
            {
                thisProject.profiles[0].active = true;
            }
            else
            {
                thisProject.profiles[0].active = false;
            }
            thisProject.profiles[0].profileID = project_profile_profileID.Text;
            //thisProject.profiles[0].profileAbbr = project_profile_profileAbbr.Text;
            thisProject.profiles[0].profileTemplate = project_profile_profileTemplate.Text;
            thisProject.profiles[0].profileName = project_profile_profileName.Text;

            // Parent Summary
            if (project_psum_active.Text == "True")
            {
                thisProject.profiles[1].active = true;
            }
            else
            {
                thisProject.profiles[1].active = false;
            }
            thisProject.profiles[1].profileID = project_psum_profileID.Text;
            //thisProject.profiles[1].profileAbbr = project_psum_profileAbbr.Text;
            thisProject.profiles[1].profileTemplate = project_psum_profileTemplate.Text;
            thisProject.profiles[1].profileName = project_psum_profileName.Text;

            // School Summary
            if (project_ssum_active.Text == "True")
            {
                thisProject.profiles[2].active = true;
            }
            else
            {
                thisProject.profiles[2].active = false;
            }
            thisProject.profiles[2].profileID = project_ssum_profileID.Text;
            //thisProject.profiles[2].profileAbbr = project_ssum_profileAbbr.Text;
            thisProject.profiles[2].profileTemplate = project_ssum_profileTemplate.Text;
            thisProject.profiles[2].profileName = project_ssum_profileName.Text;

            // General Summary
            if (project_sum_active.Text == "True")
            {
                thisProject.profiles[3].active = true;
            }
            else
            {
                thisProject.profiles[3].active = false;
            }
            thisProject.profiles[3].profileID = project_sum_profileID.Text;
            //thisProject.profiles[3].profileAbbr = project_sum_profileAbbr.Text;
            thisProject.profiles[3].profileTemplate = project_sum_profileTemplate.Text;
            thisProject.profiles[3].profileName = project_sum_profileName.Text;

            thisProject.save();

        }
        private void project_button2_Click(object sender, EventArgs e)
        {
            // editing existsing project
            saveProjectData(selectedProject);

            mainPanel.Visible = true;
            project_projectPanel.Visible = false;
            // reset the list of project buttons
            loadProjects();
            // reset the selected project to refresh the UI
            setSelectedProject(selectedProject);
            
        }
        private void closeForm(object sender, EventArgs e)
        {
            mainPanel.Visible = true;
            project_projectPanel.Visible = false;

        }
    }
}
