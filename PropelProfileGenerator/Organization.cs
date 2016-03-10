using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MySql.Data.MySqlClient;

namespace PropelProfileGenerator
{
    class Organization
    {
        public string name {get; set;}
        public string ID { get; set; }
        public string dataReady { get; set; }
        public string profileGenerated { get; set; }
        public string profilePosted { get; set; }

        public void setProfileGenerated(Project thisProject)
        {
            Project project = thisProject;

            MySqlConnection dataConnect = null;
            // parameters for the connection string should be entered into the ini file for this project
            string connectionString = "server=propelwebserv.uwaterloo.ca;database=" + project.database + ";uid=" + project.username + ";pwd=" + project.password + ";";
            dataConnect = new MySqlConnection(connectionString);

            dataConnect.Open();
            MySqlCommand dataCommand1 = new MySqlCommand();
            dataCommand1.Connection = dataConnect;
            dataCommand1.CommandText = "UPDATE " + project.table + " SET " + project.profileGenerated + " = NOW() WHERE "  + project.tableID + " = " + ID;
            dataCommand1.ExecuteNonQuery();
            dataConnect.Close();
            
        }
    }
}
