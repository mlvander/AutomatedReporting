using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PropelProfileGenerator
{
    class ProfileDocument
    {
        // Excel does not always quit cleanly, so start by storing the Excel processes before starting.
        Process[] processesBefore = Process.GetProcessesByName("excel"); 
        // Prepare Word application
        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        Microsoft.Office.Interop.Word.Document templateDoc = null;
        // Prepare Excel application
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook sourceDataFile = null;
        // Reference to the Main Display UI
        MainUI mainDisplay;

        bool profileGenerated = false;

        object missing = System.Reflection.Missing.Value;

        public ProfileDocument(object fileName, object dataFile, object saveAs, MainUI mainControl)
        {
            mainDisplay = mainControl;
            if (!File.Exists((string)dataFile))
            {
                mainDisplay.log("ERROR: Excel Datafile " + (string)dataFile + " does not exist!");
                // Close background Excel instance
                ((Microsoft.Office.Interop.Excel._Application)excelApp).Quit();
                // Close background Word instance
                ((Microsoft.Office.Interop.Word._Application)wordApp).Quit();
                return;
            }
            else if (!File.Exists((string)fileName))
            {
                mainDisplay.log("ERROR: Word Template " + (string)fileName + " does not exist!");
                // Close background Excel instance
                ((Microsoft.Office.Interop.Excel._Application)excelApp).Quit();
                // Close background Word instance
                ((Microsoft.Office.Interop.Word._Application)wordApp).Quit();
                return;
            }
            else
            {
                wordApp.Visible = false;
                object readOnly = false;
                object isVisible = false;

                // Open the word template
                templateDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref isVisible, ref missing, ref missing, ref missing, ref missing);
                // Activate the template
                templateDoc.Activate();

                try
                {
                    // Save the template with the new filename
                    templateDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing);

                    mainDisplay.log("New profile created.");
                }
                catch (Exception e)
                {
                    mainDisplay.log("ERROR: Cannot access profile final. Is it open? Please close file and retry.");
                    return;
                }             
                
                //Open Excel data file
                sourceDataFile = excelApp.Workbooks.Open((string)dataFile,missing,true);

                mainDisplay.log("Updating text values ....");
                updateText();

                mainDisplay.log("Updating graph data ....");              
                updateGraphs();

                // Save the document
                ((_Document)templateDoc).Save();
                mainDisplay.log("Document saved");

                // Close excel data
                ((_Workbook)sourceDataFile).Close();
                // Close background Excel instance
                ((Microsoft.Office.Interop.Excel._Application)excelApp).Quit();

                // Close the document
                ((_Document)templateDoc).Close(ref missing, ref missing, ref missing);
                // Close background Word instance
                ((Microsoft.Office.Interop.Word._Application)wordApp).Quit();
                
                mainDisplay.log("Profile created: " + (string)saveAs);
                profileGenerated = true;
            }
            // Get Excel processes still running after program has finished.
            Process[] processesAfter = Process.GetProcessesByName("excel");
            // Now find the process id that was created, and kill it.
            foreach (Process process in processesAfter)
            {
                if (!processesBefore.Select(p => p.Id).Contains(process.Id))
                {
                   process.Kill();
                }
            }
            mainDisplay.log("");
        }

        private bool FindAndReplace(Microsoft.Office.Interop.Word.Application WordApp, object findText, object replaceText)
        {
        // this is to find the tags listed in sheet - S00 - and replace them with the corresponding data
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object nmatchAllWordsForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = true;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            bool result = WordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards, 
                ref matchSoundsLike, ref nmatchAllWordsForms, ref forward, ref wrap, ref format, ref replaceText, 
                ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);

            return result;
        }

        private List<string> FindMissedTags(Microsoft.Office.Interop.Word.Application WordApp)
        {
            // This is to find any tags in the word template that have not been replaced
            List<string> missedTags = new List<string>();

            object findText = "|*|";
            object replaceText = false;
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = true;
            object matchSoundsLike = false;
            object nmatchAllWordsForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = true;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            Microsoft.Office.Interop.Word.Range range = WordApp.ActiveDocument.Content;
            Microsoft.Office.Interop.Word.Find find = range.Find;
            find.Text = "|*|";
            find.MatchWildcards = true;
            find.ClearFormatting();

            range.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

            while (range.Find.Found)
            {
                // Get the range of the result of the find and then select the text in that range
                // the search string is |*| but we want to capture <|*|> so adding a character to each side of range
                WordApp.ActiveDocument.Range(range.Start - 1, range.End + 1).Select();
                //collecting all missing tags in this list
                missedTags.Add(WordApp.Selection.Text);

                //Find the next instance of search string
                range.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);
            }
            return missedTags;
        }

        private void updateText()
        {
            //Worksheet S00 will contain a list of tags and data
            Worksheet textData = sourceDataFile.Sheets["S00"];
            string thisFieldName = " ";
            string thisFieldValue = " ";
            int loopCounter = 1;

            while (thisFieldName != null)
            {
                // convert the cell reference to ranges to get the values
                Microsoft.Office.Interop.Excel.Range fieldName = textData.Cells[loopCounter, 1];
                Microsoft.Office.Interop.Excel.Range fieldValue = textData.Cells[loopCounter, 2];

                thisFieldName = fieldName.Value;

                if (thisFieldName != null)
                {
                    thisFieldValue = fieldValue.Value.ToString();
                    // fields in the document are formatted "<|fieldname|>"
                    bool result = this.FindAndReplace(wordApp, "<|" + thisFieldName + "|>", thisFieldValue);
                    if (result)
                    {
                        mainDisplay.log("     <|" + thisFieldName + "|> replaced.");
                    }
                    else
                    {
                        mainDisplay.log("     ERROR: <|" + thisFieldName + "|> not found in template.");
                    }
                }
                loopCounter++;
            }

            // Search the template for any tags that have not been replaced
            List<string> missedTags = this.FindMissedTags(wordApp);
            // Write list of missed tags in the log window
            foreach(string tag in missedTags)
            {
                mainDisplay.log("     WARNING: " + tag + " has not been replaced in the template.");
            }
        }

        private void updateGraphs()
        {
            // Loop through all the inline shapes in this document.
            foreach (InlineShape thisShape in templateDoc.InlineShapes)
            {
                // If the shape is a chart then continue to update the data
                if (thisShape.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    //Excel sheet names in source datafile will match template chart titles but will not have spaces
                    string sheetname = thisShape.Title.Replace(" ", "_");
                    //Need to activate the chart in order for the data to be editable in Word after the profile is saved
                    thisShape.Chart.ChartData.Activate();
                    Worksheet templateChart = thisShape.Chart.ChartData.Workbook.Worksheets[1];

                    try
                    {   
                        // This is the source data for the graph.
                        Worksheet sourceData = sourceDataFile.Sheets[sheetname];

                        // Arbitrarily copying 26 x 26 block of cells and assigning them to embeded chart data, assuming this will be more than sufficient for every chart
                        for (int column = 1; column <= 26; column++)
                        {
                            bool unreportable = false;
                            for (int row = 1; row <= 26; row++)
                            {
                                double currentValue = 0;

                                try
                                {
                                    currentValue = (double)(sourceData.Cells[column, row] as Microsoft.Office.Interop.Excel.Range).Value;
                                }
                                catch(Exception error)
                                {
                                    currentValue = 0;
                                }

                                if ( currentValue == -1)
                                {
                                    unreportable = true;
                                    templateChart.Cells[column, row] = null;
                                }
                                else
                                {
                                    templateChart.Cells[column, row] = sourceData.Cells[column, row];
                                }
                            }
                            // if there is an unreportable value then we want to add a * to the row title
                            if (unreportable)
                            {
                                (templateChart.Cells[column, 1] as Microsoft.Office.Interop.Excel.Range).Value += "*";
                            }
                        }
                        // posting to the log window on the UI
                        mainDisplay.log("     " + thisShape.Title + " graph updated.");
                    }
                    catch(Exception e)
                    {
                        // posting to the log window on the UI
                        mainDisplay.log("     ERROR: Cannot access '" + thisShape.Title + "' data sheet in excel file.");
                        mainDisplay.log(e.Message);
                    }

                    thisShape.Chart.ChartData.Workbook.Close();
                }
            }
        }
        public bool getStatus()
        {
            return profileGenerated;
        }
    }
}
