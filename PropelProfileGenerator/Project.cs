using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PropelProfileGenerator
{
    class Project
    {
        public bool active { get; set; }
        public string projectName { get; set; }
        public string projectAbbr { get; set; }
        public string templateFolder { get; set; }
        public string excelFolder { get; set; }
        public string profileFolder { get; set; }
        public string database { get; set; }
        public string username { get; set; }
        public string password { get; set; }
        public string table { get; set; }
        public string tableID { get; set; }
        public string tableTitle { get; set; }
        public string dataReady { get; set; }
        public string profileGenerated { get; set; }
        public string profilePosted { get; set; }

        IniFile MyIni;

        public string[] profileTypeList = new string[4] { "Profile", "Parent Summary", "School Summary", "Summary" };

        public List<Profile> profiles = new List<Profile>();

        public Project(string iniFile)
        {
            MyIni = new IniFile(iniFile);

            active = false;
            var isActive = MyIni.Read("Active","Main");

            if (MyIni.Read("Active", "Main").ToLower() == "true")
            {
                active = true;
            }

            projectName = MyIni.Read("Project Name", "Main");
            projectAbbr = MyIni.Read("Project Abbr", "Main");
            templateFolder = MyIni.Read("Template folder", "Main");
            excelFolder = MyIni.Read("Excel Folder", "Main");
            profileFolder = MyIni.Read("Profile Folder", "Main");
            database = MyIni.Read("Database", "Main");
            username = MyIni.Read("Username", "Main");
            password = MyIni.Read("Password", "Main");
            table = MyIni.Read("Table", "Main");
            tableID = MyIni.Read("TableID", "Main");
            tableTitle = MyIni.Read("TableTitle", "Main");
            dataReady = MyIni.Read("DataReady", "Main");
            profileGenerated = MyIni.Read("ProfileGenerated", "Main");
            profilePosted = MyIni.Read("ProfilePosted", "Main");

            foreach (string profileType in profileTypeList)
            {
                Profile thisProfile = new Profile();
                if (MyIni.Read("Active", profileType) == "true")
                {
                    thisProfile.active = true;
                }
                else
                {
                    thisProfile.active = false;
                }

                thisProfile.profileType = profileType;
                thisProfile.profileAbbr = MyIni.Read("Profile Abbr", profileType);
                thisProfile.profileName = MyIni.Read("Profile Name", profileType);
                thisProfile.profileTemplate = MyIni.Read("Profile Template", profileType);
                thisProfile.profileID = MyIni.Read("Profile ID", profileType);
                thisProfile.selected = true;
                profiles.Add(thisProfile);
            }
            
        }

        public void save()
        {
            MyIni.Write("Active", active.ToString(), "Main");
            MyIni.Write("Project Name", projectName, "Main");
            MyIni.Write("Project Abbr", projectAbbr, "Main");
            MyIni.Write("Template folder", templateFolder, "Main");
            MyIni.Write("Excel Folder", excelFolder, "Main");
            MyIni.Write("Profile Folder", profileFolder, "Main");
            MyIni.Write("Database", database, "Main");
            MyIni.Write("Username", username, "Main");
            MyIni.Write("Password", password, "Main");
            MyIni.Write("Table", table, "Main");
            MyIni.Write("TableID", tableID, "Main");
            MyIni.Write("TableTitle", tableTitle, "Main");
            MyIni.Write("DataReady", dataReady, "Main");
            MyIni.Write("ProfileGenerated", profileGenerated, "Main");
            MyIni.Write("ProfilePosted", profilePosted, "Main");

            foreach (Profile p in profiles)
            {
                MyIni.Write("Active", p.active.ToString().ToLower(), p.profileType);
                MyIni.Write("Profile Abbr", p.profileAbbr, p.profileType);
                MyIni.Write("Profile Name", p.profileName, p.profileType);
                MyIni.Write("Profile Template", p.profileTemplate, p.profileType);
                MyIni.Write("Profile ID", p.profileID, p.profileType);
            }
        }
    }

    class Profile
    {
        public bool active { get; set; }
        public string profileType { get; set; }
        public string profileAbbr { get; set; }
        public string profileName { get; set; }
        public string profileTemplate { get; set; }
        public string profileID { get; set; }
        public bool selected { get; set; }

    }
}
