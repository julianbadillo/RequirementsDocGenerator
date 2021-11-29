using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RequirementsDocGenerator
{
    /// <summary>
    /// To read an excel spreadsheet, using OpenXML
    /// </summary>
    public class ExcelReader : IDisposable
    {
        /// <summary>
        /// If it has still some more data to read.
        /// ReadFromExcel should be called again.
        /// </summary>
        public bool HasMoreData { get; private set; } = false;

        /// <summary>
        /// If finished reading.
        /// </summary>
        public bool FinishedReading { get; private set; } = false;

        /// <summary>
        /// Rows that have been read so far, can be used to calculate progress.
        /// </summary>
        public int RowsRead { get; private set; }

        /// <summary>
        /// Total rows to be read, can be used to calculate progress.
        /// </summary>
        public int TotalRows { get; private set; }

        /// <summary>
        /// Max number of rows read in one call
        /// To avoid loading a too-big list in memory.
        /// </summary>
        public const int MAX_ROWS_ON_ONE_READ = 10000;

        private SharedStringTablePart shareStringPart = null;
        private SpreadsheetDocument doc = null;
        OpenXmlReader DataReader = null;

        /// <summary>
        /// parses an excel file and a sheet name into list of rows of strings.
        /// One should call this method again if HasMoreData is true.
        /// </summary>
        /// <param name="input"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ICollection<List<string>> ReadFromExcel(Stream input, string sheetName)
        {
            var result = new List<List<string>>();

            if (!HasMoreData)
                PrepareRead(input, sheetName);

            //ProcessCaDataPostSheetDOM(result, shareStringPart, part);
            ProcessCaDataPostSheetSAX(result);

            return result;
        }

        private void PrepareRead(Stream input, string sheetName)
        {
            doc = SpreadsheetDocument.Open(input, isEditable: false);

            // Get the SharedStringTablePart
            if (doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                shareStringPart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();

            //foreach (Sheet sheet in doc.WorkbookPart.Workbook.Sheets)
            //    Debug.WriteLine("sheet:" + sheet.Name + " ID:" + sheet.Id);

            foreach (WorksheetPart part in doc.WorkbookPart.WorksheetParts)
            {
                var id = doc.WorkbookPart.GetIdOfPart(part);

                // Get the name from the sheets
                var name = doc.WorkbookPart.Workbook.Sheets.Cast<Sheet>()
                                .First(s => s.Id == id).Name;
                //Debug.WriteLine("Sheet Name: " + name);

                // check only sheet with the given name
                if (name == sheetName || string.IsNullOrEmpty(sheetName))
                {
                    CountRows(part);
                    DataReader = OpenXmlReader.Create(part);
                }
            }
            if (DataReader == null)
                throw new FileFormatException("The sheet name '" + sheetName + "' was not found, make sure the file uploaded has the correct format.");

        }

        /// <summary>
        /// Gets the row count of a single sheet and set TotalRows property and reset RowsRead
        /// </summary>
        /// <param name="part"></param>
        private void CountRows(WorksheetPart part)
        {
            using (var countReader = OpenXmlReader.Create(part))
            {
                RowsRead = 0;
                TotalRows = 0;
                while (countReader.Read())
                {
                    if (countReader.IsStartElement && countReader.ElementType == typeof(Row))
                        ;
                    // reached a row end
                    else if (countReader.IsEndElement && countReader.ElementType == typeof(Row))
                        TotalRows++;
                    // reached a cell start
                    else if (countReader.IsStartElement && countReader.ElementType == typeof(Cell))
                        ;
                    // value start
                    else if (countReader.IsStartElement && countReader.ElementType == typeof(CellValue))
                    {
                        string s = countReader.GetText();
                        if (string.IsNullOrEmpty(s))
                            continue;
                    }
                    // cell end
                    else if (countReader.IsEndElement && countReader.ElementType == typeof(Cell))
                        ;
                }
                //Debug.WriteLine("Total Rows: " + TotalRows);
            }
        }

        // By pieces
        private void ProcessCaDataPostSheetSAX(List<List<string>> result)
        {
            // read lines
            int rowCount = 0;
            List<string> rowList = null;
            string valueType = null;
            string value = null;
            string reference = null;
            string prevReference = null;
            if (DataReader == null)
                throw new FileFormatException("The sheet name was not found, make sure the file uploaded has the correct format.");
            while (DataReader.Read())
            {
                //System.Diagnostics.Debug.WriteLine(reader.IsStartElement+"-"+reader.ElementType +":"+reader.GetText()+", LocalName="+ reader.LocalName
                //    +", Attributes={"+string.Join(",",(from a in reader.Attributes select a.LocalName+":"+a.Value))+"}");
                // reached a row start
                if (DataReader.IsStartElement && DataReader.ElementType == typeof(Row))
                {
                    rowList = new List<string>();
                }
                // reached a row end
                else if (DataReader.IsEndElement && DataReader.ElementType == typeof(Row))
                {
                    // add the row and break.
                    if (rowList != null && rowList.Count > 0)
                        result.Add(rowList);
                    rowCount++;
                    RowsRead++;
                    if (rowCount >= MAX_ROWS_ON_ONE_READ)
                        break;
                }
                // reached a cell start
                else if (DataReader.IsStartElement && DataReader.ElementType == typeof(Cell))
                {
                    valueType = "str";
                    value = "";
                    foreach (var a in DataReader.Attributes)
                        if (a.LocalName == "t")
                            valueType = a.Value;
                        else if (a.LocalName == "r")
                            reference = a.Value;

                    // see if there is any gap
                    if (prevReference != null && reference != null)
                    {
                        long gap = CalculateGap(prevReference, reference);
                        // fill the missing cells with empty strings
                        if (gap > 1)
                            for (int i = 1; i < gap; i++) rowList.Add("");
                    }
                    prevReference = reference;
                }
                // value start
                else if (DataReader.IsStartElement && DataReader.ElementType == typeof(CellValue))
                {
                    // shared string
                    if (valueType == "s")
                    {
                        string s = DataReader.GetText();
                        if (string.IsNullOrEmpty(s))
                            continue;

                        int index = int.Parse(s);
                        SharedStringItem item = (SharedStringItem)shareStringPart.SharedStringTable.ElementAt(index);
                        value = item.Text == null ? "" : item.Text.Text;
                    }
                    else
                        value = DataReader.GetText();
                }
                // cell end
                else if (DataReader.IsEndElement && DataReader.ElementType == typeof(Cell))
                {
                    // remove hidden chars
                    rowList.Add(value.Replace("_x000D_", ""));
                }
            }

            //Debug.WriteLine("Rows Fed: " + RowsFed);

            // we break so we dont overload the list of strings
            if (rowCount >= MAX_ROWS_ON_ONE_READ)
                HasMoreData = true;
            else
            {
                HasMoreData = false;
                FinishedReading = true;
            }
        }

        private void ProcessCaDataPostSheetDOM(List<List<string>> result, SharedStringTablePart shareStringPart, WorksheetPart part)
        {
            SheetData sheetData = part.Worksheet.GetFirstChild<SheetData>();

            // needs to load entire sheet data in memory
            foreach (Row row in sheetData.Elements<Row>())
            {
                var rowList = new List<string>();

                string prevReference = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    string reference = cell.CellReference.Value;

                    // see if there is any gap
                    if (prevReference != null)
                    {
                        long gap = CalculateGap(prevReference, reference);
                        // fill the missing cells with empty strings
                        if (gap > 1)
                            for (int i = 1; i < gap; i++) rowList.Add("");
                    }
                    prevReference = reference;

                    // If shared string -- lookup on the shared table
                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        int index = int.Parse(cell.CellValue.Text);
                        SharedStringItem item = (SharedStringItem)shareStringPart.SharedStringTable.ElementAt(index);
                        rowList.Add(item.Text.Text);
                    }
                    else
                    {
                        rowList.Add(cell.InnerText ?? "");
                    }
                }
                if (rowList.Count > 0)
                    result.Add(rowList);
                rowList = null;
                if (result.Count >= MAX_ROWS_ON_ONE_READ)
                {
                    //Debug.WriteLine("MAX Rows Reached: " + result.Count);
                    HasMoreData = true;
                    break;
                }
            }
        }

        public static long CalculateGap(string cell1, string cell2)
        {
            // trim the integer part
            cell1 = Regex.Replace(cell1, @"\d+", "");
            cell2 = Regex.Replace(cell2, @"\d+", "");
            return FromBase26(cell2) - FromBase26(cell1);
        }

        private static long FromBase26(string s)
        {
            char[] c = s.ToCharArray();
            int b = 1;
            long total = 0;
            for (int i = c.Length - 1; i >= 0; i--)
            {
                total += b * (c[i] - 'A' + 1);
                b *= 26;
            }
            return total;
        }

        public void Dispose()
        {
            // To avoid memory leaks
            if (DataReader != null)
            {
                DataReader.Close();
                DataReader.Dispose();
                DataReader = null;
            }

            if (doc != null)
            {
                doc.Close();
                doc.Dispose();
                doc = null;
            }

            if (shareStringPart != null)
                shareStringPart = null;
        }
    }
}
