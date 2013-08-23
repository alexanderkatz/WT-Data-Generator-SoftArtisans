using System;
using System.Collections.Generic;
using System.Text;
using SoftArtisans.OfficeWriter.WordWriter;
using System.Data;
using System.Collections;

namespace WordTemplateDataGenerator
{
    /// <summary>
    /// Data Generation Class
    /// </summary>
    class WordTemplateDataGenerator
    {
        String[] allMergefields; /* Collection of Mergefields */
        String[] allBookmarks;   /* Collection of Bookmarks */
        WordTemplate wt;         /*Template */
        private delegate string DataGeneratorFunction();

        /* Merge Fields Patterns */
        string colNamePattern = @"^[a-zA-z][\s\w0-9]*[a-zA-Z0-9]$";   // «ColumnName» or «Column Name»
        string sourceColPattern = @"^[a-zA-Z][\s\w]*[a-zA-Z0-9]\.[a-zA-Z][\s\w]*[a-zA-Z0-9]$"; //«DataSource.ColumnName» or «[DataSource].[ColumnName]» or «[Data Source].[Column Name]»

        /// <summary>
        /// GenerateBind()
        /// This method is the only method that needs to be called. It calls all the other methods necessary for generating data.
        /// </summary>
        /// <param name="wt">Template object with an open excel file</param>
        public void GenerateBind(WordTemplate wt)
        {
            // Stores the template for use in other methods
            this.wt = wt;

            // Gets the the merge fields and bookmarks
            allMergefields = wt.FieldMarkers;
            allBookmarks = wt.Bookmarks;

            // Parse Fields
            parseFieldsDocLevel();
            parseFieldsInBookmarks();
        }

        /*Parses fields that are outside of bookmarks*/
        private void parseFieldsDocLevel()
        {
            /*ArrayLists store long and short mergefields not contained within bookmarks*/
            ArrayList shortMerge = new ArrayList();
            ArrayList longMerge = new ArrayList();

            /* Dictionary: Keys = DataTable names, Values = DataTable*/
            Dictionary<String, DataTable> dFields = new Dictionary<string, DataTable>();

            // Loops through all the mergefields and generates data dependent on mergefield type
            for (int i = 0; i < allMergefields.Length; i++)
            {
                //If mergefield has brackets strip them away
                allMergefields[i] = removeBrackets(allMergefields[i]);
                // Check to see if mergefield matches colNamePattern
                if (CheckColNamePattern(allMergefields[i]))
                {
                    //Get the name of the short merge fields. The names will later be used to generate data
                    shortMerge.Add(allMergefields[i]);
                }
                // Check to see if field matches «DataSource.ColumnName»
                else if (CheckSourceColPattern(allMergefields[i]))
                {
                    longMerge.Add(allMergefields[i]);                    
                    fillDictionary(dFields, allMergefields[i], false);
                }
            }
            //Set the data source
            SetShortDataSources(shortMerge.ToArray());
            SetDataSourceDictionary(dFields, false);
            
            //for debugging
            //printFieldCats(shortMerge, longMerge);
        }

        private void parseFieldsInBookmarks()
        {
            /* Dictionary: Keys = DataTable names, Values = DataTable*/
            Dictionary<String, DataTable> dBookmarkFields = new Dictionary<string, DataTable>();/*Mergefields inside of bookmarks*/

            // Loops through all the bookmarks and generates data dependent on mergefield type
            // Handle grouping here
            for (int i = 0; i < allBookmarks.Length; i++)
            {
                //Check to see if bookmark is a group
                if (!allBookmarks[i].StartsWith("group"))
                {
                    //Get all the mergefields in the particular bookmark
                    string[] fieldsInBookmark = wt.BookmarkFieldMarkers(allBookmarks[i]);
                    for (int j = 0; j < fieldsInBookmark.Length; j++)
                    {
                        fieldsInBookmark[j] = removeBrackets(fieldsInBookmark[j]);
                        if (CheckColNamePattern(fieldsInBookmark[j]))
                        {
                            // Create a temp field that includes the data source
                            string field = allBookmarks[i] + "." + fieldsInBookmark[j];
                            fillDictionary(dBookmarkFields, field, true);
                        }
                        // Check to see if long mergefield
                        else if (CheckSourceColPattern(fieldsInBookmark[j]))
                        {
                            fillDictionary(dBookmarkFields, fieldsInBookmark[j], true);
                        }
                    }
                }
            }
            //Set the data source
            SetDataSourceDictionary(dBookmarkFields, true);
        }

        /*Checks to see if string matches «ColumnName»*/
        private bool CheckColNamePattern(string colName)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(colName, colNamePattern);
        }

        /*Checks to see if string matches «DataSource.ColumnName»*/
        private bool CheckSourceColPattern(string colName)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(colName, sourceColPattern);
        }


        /// <summary>
        /// fillDictionary
        /// Adds keys (Table Name) to the dictionary and generate data
        /// </summary>
        private void fillDictionary(Dictionary<String, DataTable> dictionary, String field, Boolean bookmarked)
        {
            // Isolate the key and colName from the mergefield. The key is the DataSource or the substring preceeding the "."            
            int charIndex = field.IndexOf(".");
            string key = field.Substring(0, charIndex);
            string colName = field.Substring(charIndex + 1);

            // If key is not contained within dictionary, add it
            DataTable dt;
            dictionary.TryGetValue(key, out dt);
            if (dt == null)
            {
                dictionary.Add(key, dt = new DataTable());
                dt.TableName = key;
                dt.Rows.Add(dt.NewRow());
                if (bookmarked)
                {
                    for (int i = 0; i < 10; i++)
                    {
                        dt.Rows.Add(dt.NewRow());
                    }
                }
            }
            // add column to dt if it has not been added
            if (!dt.Columns.Contains(colName))
                dt.Columns.Add(colName);

            // Attempt to identify data type for more realistic data values
            DataGeneratorFunction function = identifyData(colName);

            // If the type of data cannot be identified generate generic data points
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var value = string.Empty;
                if (function == null)
                {
                    value = dt.TableName + "." + colName + "_value";
                    if (dt.Rows.Count > 1)
                        value += "_"+(i+1);
                }
                else
                {
                    value = function();
                }
                //dt.Rows[i].SetField<string>(colName, value);
                dt.Rows[i][dt.Columns.IndexOf(colName)] = value;
            }
        }

        /// <summary>
        /// SetShortDataSources()
        /// Sets the data source for all short mergefields
        /// This is the templates first data source
        /// </summary>
        private void SetShortDataSources(object[] shortMerge)
        {
            if (shortMerge.Length > 0)
            {
                string[] fieldNames = new string[shortMerge.Length];
                for (int i = 0; i < shortMerge.Length; i++)
                {
                    fieldNames[i] = (string)shortMerge[i];
                    shortMerge[i] = shortMerge[i] + "_Value";
                }
                wt.SetDataSource(shortMerge, fieldNames);
            }
        }

        /// <summary>
        /// SetDataSourceDictionary
        /// Sets the data source for all Data Tables contained in dictionary
        /// </summary>
        private void SetDataSourceDictionary(Dictionary<string, DataTable> dictionary, Boolean bookmarked)
        {
            foreach (var pair in dictionary)
            {
                if (bookmarked)
                {
                    wt.SetRepeatBlock(pair.Value, pair.Key);
                }
                else
                {
                    wt.SetDataSource(pair.Value, pair.Key);
                }
            }
        }


        /// <summary>
        /// removeBrackets()
        /// Removes the brackets from mergefields
        /// This allows SetDataSource and SetRepeatBlock to be performed successfully
        /// </summary>
        private string removeBrackets(string field)
        {
            field = field.Replace("[", "");
            field = field.Replace("]", "");
            return field;
        }

/* Methods for Specific Data Generation ________________________________________________________
 * If methods for generatic specific types of data are desired add them in this section
 * Useful methods could be a date generator, a number generator, etc.
 */

        /// <summary>
        /// identifyData()
        /// Identifies specific data type
        /// </summary>
        private DataGeneratorFunction identifyData(string colName)
        {
            if (colName.Contains("phone"))
                return new DataGeneratorFunction(getPhoneNums);
            else
                return null;
        }

        /// <summary>
        /// Generates a random phone number
        /// </summary>
        /// <returns></returns>
        private string getPhoneNums()
        {
            Random rand = new Random();
            string num = "";
            for (int j = 0; j < 10; j++)
            {
                int rNum = DateTime.Now.GetHashCode() % rand.Next();
                num += rNum.ToString().Substring(1, 1);
                if (j == 2 || j == 5)
                    num += "-";

            }
            //Console.WriteLine(num);
            return num;
        }

/* Methods for debugging _____________________________________________________________________*/

        /// <summary>
        /// Print the short and long mergefields for debugging
        /// </summary>
        /// <param name="shortF">Short Mergefields</param>
        /// <param name="longF">Long Mergefields</param>
        private void printFieldCats(ArrayList shortF, ArrayList longF)
        {
            Console.WriteLine("Short Mergefields");
            foreach (string s in shortF)
            { Console.Write(s + ", "); }

            Console.WriteLine("\nLong Mergefields");
            foreach (string st in longF)
            { Console.Write(st + ", "); }
        }
        /// <summary>
        /// printData
        /// Prints all the mergefields at the document level and within bookmarks
        /// </summary>
        private void printData(string[] allMergefields, string[] allBookmarks)
        {
            Console.WriteLine("Mergefields");
            foreach (string s in allMergefields)
            {
                Console.Write(s + ", ");
            }

            Console.WriteLine("\nBookmarks");
            foreach (string st in allBookmarks)
            {
                Console.WriteLine(st + ", ");
                // Get all the mergefields within the bookmar
                String[] someFields = wt.BookmarkFieldMarkers(st);
                foreach (string str in someFields)
                {
                    Console.Write(str + ", ");
                }
            }
        }
    }
}
