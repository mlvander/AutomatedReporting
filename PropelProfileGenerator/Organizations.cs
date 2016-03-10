using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MySql.Data.MySqlClient;

namespace PropelProfileGenerator
{
    class Organizations
    {
        Project project = null;
        MySqlConnection dataConnect = null;
        MainUI mainDisplay = null;
        bool profileGenerated;
        bool profilePosted;
        string orgID = null;

        public Organizations(Project thisProject, MainUI thisMainDisplay, bool thisProfileGenerated, bool thisProfilePosted)
        {
            mainDisplay = thisMainDisplay;
            profileGenerated = thisProfileGenerated;
            profilePosted = thisProfilePosted;
            project = thisProject;

            setUpMySQL();
        }
        public Organizations(Project thisProject, string thisOrgID)
        {
            project = thisProject;
            orgID = thisOrgID;
            setUpMySQL();
        }

        private void setUpMySQL()
        {
            string connectionString = null;

            // parameters for the connection string should be entered into the ini file for this project
            connectionString = "server=propelwebserv.uwaterloo.ca;database=" + project.database + ";uid=" + project.username + ";pwd=" + project.password + ";";
            dataConnect = new MySqlConnection(connectionString);

            try
            {
                dataConnect.Open();
                dataConnect.Close();
            }
            catch (Exception ex)
            {
                mainDisplay.log("Cannot open connection to the selected project database (" + project.database + ")! ");
                return;
            }
        }

        public List<Organization> getOrgDetails()
        {
            List<Organization> orgDetail = new List<Organization>();

            try
            {
                dataConnect.Open();
            }
            
            catch (MySqlException e)
            {
                
                return orgDetail;    
            }
            
            string getColNames = "SHOW COLUMNS FROM " + project.table;
            MySqlCommand dataCommand1 = new MySqlCommand();
            dataCommand1.Connection = dataConnect;
            dataCommand1.CommandText = getColNames;
            MySqlDataReader dataResult1;

            try
            {
                dataResult1 = dataCommand1.ExecuteReader();
            }
            catch (MySqlException e)
            {
                mainDisplay.log(e.Message);
                return orgDetail;
            }

            int[] colIndex = new int[5] { 1, 2, 3, 4, 5 };

            int indexCounter = 0;
            while (dataResult1.Read())
            {
                string currentData = dataResult1.GetString(0);

                if (currentData == project.tableID)
                {
                    colIndex[0] = indexCounter;
                }
                if (currentData == project.tableTitle)
                {
                    colIndex[1] = indexCounter;
                }
                if (currentData == project.dataReady)
                {
                    colIndex[2] = indexCounter;
                }
                if (currentData == project.profileGenerated)
                {
                    colIndex[3] = indexCounter;
                }
                if (currentData == project.profilePosted)
                {
                    colIndex[4] = indexCounter;
                }
                indexCounter++;
            }
            dataConnect.Close();
            dataConnect.Open();

            string sqlString;
            if (orgID != null)
            {
                // looking up a specific organization
                sqlString = "SELECT * from " + project.table + " where " + project.tableID + " = " + orgID;
            }
            else if (profilePosted)
            {
                // want to find all the organizations that have indicated the profile has already been generated
                sqlString = "SELECT * from " + project.table + " where " + project.profilePosted + " is not null";
            }
            else if (profileGenerated)
            {
                // want to find all the organizations that have indicated the profile has already been generated
                sqlString = "SELECT * from " + project.table + " where " + project.profileGenerated + " is not null AND " + project.profilePosted + " is null";
            }
            else
            {
                // want to find all the organizations that have indicated the data is ready for profiles
                sqlString = "SELECT * from " + project.table + " where " + project.dataReady + " is not null AND " + project.profileGenerated + " is null AND " + project.profilePosted + " is null";
            }
            MySqlCommand dataCommand2 = new MySqlCommand();
            dataCommand2.Connection = dataConnect;
            dataCommand2.CommandText = sqlString;
            MySqlDataReader dataResult2;

            try
            {
                dataResult2 = dataCommand2.ExecuteReader();
            }
            catch (MySqlException e)
            {
                mainDisplay.log(e.Message);
                return orgDetail;
            }
            
            while (dataResult2.Read())
            {
                // if the profile has not been generated then this value will be null, programming here to prevent null errors.
                string thisProfileGenerated = null;
                if (profileGenerated)
                {
                    thisProfileGenerated = dataResult2.GetString(colIndex[3]);
                }
                string thisProfilePosted = null;
                if (profilePosted)
                {
                    thisProfilePosted = dataResult2.GetString(colIndex[4]);
                }

                orgDetail.Add(new Organization
                {
                    ID = dataResult2.GetString(colIndex[0])
                    ,
                    name = dataResult2.GetString(colIndex[1])
                    ,
                    dataReady = dataResult2.GetString(colIndex[2])
                    ,
                    profileGenerated = thisProfileGenerated
                    ,
                    profilePosted = thisProfilePosted
                });
            }
            dataConnect.Close();
            return orgDetail;
        }
    }
}
