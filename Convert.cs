using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ConvertExcelToAnotherFormat
{
    public static class Convert
    {

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static DataSet ToDataSet(string fileName)
        {
            try
            {
                Stream file = File.OpenRead(fileName);
                return ToDataSet(file);
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static DataSet ToDataSet(Stream fileStream)
        {
            try
            {
                DataSet ds = new DataSet();
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileStream, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets sheetsCollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    foreach (Sheet thesheet in sheetsCollection.OfType<Sheet>())
                    {
                        DataTable dtTable = new DataTable();
                        Worksheet workSheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;
                        SheetData sheetData = workSheet.GetFirstChild<SheetData>();

                        for (int rCnt = 0; rCnt < sheetData.ChildElements.Count(); rCnt++)
                        {
                            List<string> rowList = new List<string>();
                            for (int rCnt1 = 0; rCnt1
                                < sheetData.ElementAt(rCnt).ChildElements.Count(); rCnt1++)
                            {

                                Cell thecurrentcell = (Cell)sheetData.ElementAt(rCnt).ChildElements.ElementAt(rCnt1);
                                string currentcellvalue = string.Empty;
                                if (thecurrentcell.DataType != null)
                                {
                                    if (thecurrentcell.DataType == CellValues.SharedString)
                                    {
                                        int id;
                                        if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                        {
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            if (item.Text != null)
                                            {
                                                if (rCnt == 0)
                                                {
                                                    dtTable.Columns.Add(item.Text.Text);
                                                }
                                                else
                                                {
                                                    rowList.Add(item.Text.Text);
                                                }
                                            }
                                            else if (item.InnerText != null)
                                            {
                                                currentcellvalue = item.InnerText;
                                            }
                                            else if (item.InnerXml != null)
                                            {
                                                currentcellvalue = item.InnerXml;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (rCnt != 0)
                                    {
                                        rowList.Add(thecurrentcell.InnerText);
                                    }
                                }
                            }
                            if (rCnt != 0)
                                dtTable.Rows.Add(rowList.ToArray());

                        }

                        ds.Tables.Add(dtTable.Copy());
                    }

                    return ds;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="tableNo"></param>
        /// <returns></returns>
        public static string ToJson(string fileName, int tableNo = 0)
        {
            DataSet dataSet = ToDataSet(fileName);
            if (dataSet.Tables.Count >= tableNo)
                return "";
            return JsonConvert.SerializeObject(dataSet.Tables[tableNo]);
        }


        ///
    }
}
