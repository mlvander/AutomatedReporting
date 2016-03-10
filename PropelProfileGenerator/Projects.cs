using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PropelProfileGenerator
{
    class Projects
    {
        List<Project> projects = new List<Project>();

        public Projects(string path, MainUI mainDisplay) {
            if (Directory.Exists(path))
            {
                string[] fileEntries = Directory.GetFiles(path);
                foreach (string file in fileEntries)
                {
                    string[] thisFile = file.Split('.');
                    if (thisFile[1] == "ini")
                    {
                        Project thisProject = new Project(file);
                        if (thisProject.active)
                        {
                            projects.Add(thisProject);
                        }
                    }
                }
            }
            else
            {
                mainDisplay.log("Critical Error: Path to ini project files not valid.");
            }

            if (projects.Count == 0)
            {
                mainDisplay.log("Critical Error: There are no active project INI files.");
            }
        }

        public List<Project> getProjects()
        {
            return projects;
        }
    }
}
