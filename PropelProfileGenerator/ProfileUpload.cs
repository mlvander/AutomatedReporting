using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MySql.Data.MySqlClient;
using System.IO;

namespace PropelProfileGenerator
{
    class ProfileUpload
    {
        string webServerLocation = @"\\propelwebserv\secure files\\shapesreports";

        public ProfileUpload(Project project, Organization organization, MainUI mainDisplay)
        {
            MySqlConnection dataConnect = null;
            // parameters for the connection string should be entered into the ini file for this project
            string connectionString = "server=propelwebserv.uwaterloo.ca;database=" + project.database + ";uid=" + project.username + ";pwd=" + project.password + ";";
            dataConnect = new MySqlConnection(connectionString);

            for (int i = 0; i < project.profiles.Count; i++)
            {
                // 1. check if the profile file exists in the profile folder
                string filename = project.projectAbbr + "_" + organization.ID + "_" + project.profiles[i].profileName.Replace(" ", "") + ".pdf";

                string sourceFilename = project.profileFolder + "\\" + filename;
                string destinationFilename = webServerLocation + "\\" + filename;

                
                if (File.Exists(sourceFilename))
                {
                    mainDisplay.log("Generated file: " + filename + " found.");

                    // 2. Check the database to see if this file as already been inserted
                    bool dataRecordExists = false;
                    try
                    {
                        dataConnect.Open();
                        MySqlCommand dataCommand1 = new MySqlCommand();
                        dataCommand1.Connection = dataConnect;
                        dataCommand1.CommandText = "SELECT * FROM tbl_profile_file WHERE profile_file = '" + filename + "'";
                        MySqlDataReader dataReader = dataCommand1.ExecuteReader();

                        dataRecordExists = dataReader.HasRows;

                        dataConnect.Close();

                        if (dataRecordExists)
                        {
                            mainDisplay.log("     Database: '" + project.database + "' has already been updated, no database changes made.");
                        }
                    }
                    catch (Exception e)
                    {
                        mainDisplay.log("     Access to:" + project.database + " has been denied.");
                        return;
                    }
                    
                    // 3. copy the file to the web server location

                    try
                    {
                        File.Copy(sourceFilename, destinationFilename, true);
                        if (dataRecordExists)
                        {
                            mainDisplay.log("     File replaced on web server.");
                        }
                        else
                        {
                            mainDisplay.log("     File uploaded to web server.");
                        }
                    }
                    catch (Exception e)
                    {
                        mainDisplay.log("     ERROR: Access to web server denied, file can't be uploaded.");
                        return;
                    }

                    
                    // 4. update the database tables to indicate it has been uploaded
                    if (!dataRecordExists)
                    {
                        try
                        {
                            dataConnect.Open();
                            MySqlCommand dataCommand1 = new MySqlCommand();
                            dataCommand1.Connection = dataConnect;

                            dataCommand1.CommandText = "INSERT INTO tbl_profile_file ( profile_file, profileID_fk, schoolID_fk, profile_upload_dt )";
                            dataCommand1.CommandText += " VALUES ( '" + filename + "', '" + project.profiles[i].profileID + "', " + organization.ID + ", NOW())";

                            dataCommand1.ExecuteNonQuery();
                            dataConnect.Close();

                            mainDisplay.log("     Database: '" + project.database + "' has been updated.");
                        }
                        catch (Exception e)
                        {
                            mainDisplay.log("     Access to:" + project.database + " has been denied, can't update.");
                            mainDisplay.log("     " + e.Message);
                            return;
                        }
                    }

                    // 5. Update the main record to indicate the profile has been posted.
                    dataConnect.Open();
                    MySqlCommand dataCommand2 = new MySqlCommand();
                    dataCommand2.Connection = dataConnect;
                    dataCommand2.CommandText = "UPDATE " + project.table + " SET " + project.profilePosted + " = NOW() WHERE " + project.tableID + " = " + organization.ID;
                    dataCommand2.ExecuteNonQuery();
                    dataConnect.Close();
                }
                
            }
        }        
    }
}
