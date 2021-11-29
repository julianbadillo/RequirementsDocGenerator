using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace RequirementsDocGenerator
{
    /// <summary>
    /// Generates a Requirements document from an excel spreadsheet.
    /// </summary>
    class DocGenerator
    {
        const string SHEET_NAME = "Functional Requirements";
        const string PROJECT_NAME = "ACORN";
        const string TITLE = "Functional Requirements";

        /// <summary>
        /// Abstraction of a requirement
        /// </summary>
        class Requirement
        {
            public string ReqId { get; set; }
            public string UCId { get; set; }
            public string Text { get; set; }
            public string Category { get; set; }
            public ISet<string> Tags { get; set; }
            public int? Level { get; set; }
            public string Notes { get; set; }
            public string Experts { get; set; }
        }

        /// <summary>
        /// Produces a requirements Word document from the table of requirements in Excel
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFile"></param>
        public void Generate(string inputFile, string outputFile)
        {
            // Requirements from the excel table
            var reqList = ReadRequirements(inputFile);
            WriteRequirementsDocument(outputFile, reqList);
        }

        /// <summary>
        /// Writes a document from a list of requriements
        /// </summary>
        /// <param name="outputFile"></param>
        private void WriteRequirementsDocument(string outputFile, IList<Requirement> reqList)
        {
            using (var writer = new WordWriter())
            {

                writer.StartDocument(outputFile);
                // document title
                writer.WriteParagraph();
                writer.WriteTitle(PROJECT_NAME);
                writer.WriteParagraph();
                writer.WriteTitle(TITLE);
                writer.WritePageBreak();

                // Revision history and Table of content
                writer.WriteHeading0("Revision History");

                writer.StartTable();
                writer.AddTableHeader("Revision", "Responsible Person", "Date", "Description of Changes");
                writer.WriteParagraph();
                writer.WritePageBreak();

                // insert a table of content here
                //writer.WriteTOC();
                //writer.WritePageBreak();


                writer.WriteParagraph();
                writer.WriteParagraph();

                // Beginning
                writer.WriteHeading1("Purpose of this Document");
                writer.WriteParagraph("<write the purpose of this document>");

                // Sort by category
                var byCategory = reqList.GroupBy(req => req.Category);
                foreach(var group in byCategory)
                {
                    // skip these categories
                    if (group.Key == "Out_Of_Scope"
                        || group.Key == "Top_Level")
                        continue;
                    // category title
                    writer.WriteHeading1(group.Key);
                    // list of requirements - sort by level
                    foreach(var req in group.OrderBy(r => r.Level ?? 0))
                    {
                        writer.WriteHeading2(req.ReqId);
                        writer.WriteParagraph(req.Text);
                        writer.WriteParagraph($"Level: {req.Level}");
                        writer.WriteParagraph($"Tags: {string.Join("; ",req.Tags)}");
                    }
                }
            }
        }

        /// <summary>
        /// Reads requirements from Excel table, generating a list of
        /// Requirement objects
        /// </summary>
        /// <param name="inputFile"></param>
        /// <returns></returns>
        private IList<Requirement> ReadRequirements(string inputFile)
        {
            using (var input = File.OpenRead(inputFile))
            using (var excel = new ExcelReader())
            {
                var data = excel.ReadFromExcel(input, SHEET_NAME);
                int rowNumber = 0;
                var columns = new Dictionary<string, int>();
                // read the list of requirements and create it
                var reqsList = new List<Requirement>();
                foreach (var row in data)
                {
                    if (rowNumber == 0)
                    {
                        // Read column names
                        int colNumber = 0;
                        foreach (var col in row)
                        {
                            columns[col] = colNumber++;
                        }
                        rowNumber++;
                        continue;
                    }
                    // Build a requirement object from row's data
                    var req = new Requirement()
                    {
                        ReqId = columns["Req ID"] < row.Count ? row[columns["Req ID"]]:"",
                        UCId = columns["UC_ID"] < row.Count? row[columns["UC_ID"]]: "",
                        Category = columns["Category"] < row.Count ? row[columns["Category"]] : "",
                        Text = columns["Requirement"] < row.Count ? row[columns["Requirement"]]: "" ,
                        Tags = columns["Metadata (Additional Tags)"] < row.Count 
                                        ? new HashSet<string>(row[columns["Metadata (Additional Tags)"]]
                                                    .Split(new char[] { ' ', ',', ';' }, 
                                                    StringSplitOptions.RemoveEmptyEntries))
                                        : new HashSet<string>(),
                        Experts = columns["People/Experts"] < row.Count ? row[columns["People/Experts"]]: "",
                    };
                    // make sure is non-blank an in-bounds
                    if (columns["Level"] < row.Count && int.TryParse(row[columns["Level"]], out int level))
                        req.Level = level;

                    reqsList.Add(req);
                    rowNumber++;
                }
                return reqsList;
            }
        }
    }
}
