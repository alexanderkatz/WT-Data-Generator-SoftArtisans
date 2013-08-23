//--------------------------------------------------------------
//--- SoftArtisans OfficeWriter WordTemplate Dummy Data Generator
//--- TS-Tool for Testing
//--- Opens a WordTemplate and creates and sets data sources for the merge fields
//---   Sets repeat blocks for merge fields contained within a bookmark
//--------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using SoftArtisans.OfficeWriter.WordWriter;
using System.Data;
using System.Web;
using System.Text.RegularExpressions;
using System.Collections;
using WordTemplateDataGenerator;

namespace WordTemplateDataGenerator
{
    public class GeneratorTest
    {
        public static void Main(String[] args)
        {
            // Create an instance of WordTemplate and open the template file
            WordTemplate wt = new WordTemplate();
            string file = "Bookmark.docx";
            wt.Open(@"..\..\TestFiles\" + file);

            // Crate an instance of the DataGenerator
            WordTemplateDataGenerator dg = new WordTemplateDataGenerator();
            dg.GenerateBind(wt);

            wt.Process();
            wt.Save(@"../../Output/Output_" + file);
        }
    }
}
    
   

