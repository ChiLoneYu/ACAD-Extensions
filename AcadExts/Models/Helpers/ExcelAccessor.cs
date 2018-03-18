using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.EditorInput;
using EXCEL = Microsoft.Office.Interop.Excel;
using System.IO;

namespace AcadExts
{
    internal class ExcelAccessor
    {
        // Belongs to Excel class but should be available to other classes
        public class Row
        {
            public String LegacyKey;
            public String LegacyLayer;
            public String NewLayer;
            public int Priority;

            public Row(String inKey, String inLegacy, String inNew, int inPriority)
            {
                LegacyKey = LegacyLayer = inLegacy;
                NewLayer = inNew;
                Priority = inPriority;
            }
        }

        public static List<Row> GetExcelData(String inXLS)
        {
            Dictionary<String, List<String>> ExcelLayers = new Dictionary<String, List<String>>();
            List<Row> ExcelTable = new List<Row>();

            ExcelLayers.Add("LegacyLayer", new List<String>());
            ExcelLayers.Add("NewLayer", new List<String>());
            ExcelLayers.Add("PriorityLayer", new List<String>());

            //Autodesk.AutoCAD.EditorInput.Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //PromptResult pr = ed.GetString("Enter RPSTL layer name excel file path (no spaces): ");
            //String inXLS = pr.StringResult.Replace(" ", "").Replace("\r\n", "");

            if (!File.Exists(inXLS) ||
               (Path.GetExtension(inXLS).ToLower() != ".xlsx"))
            {
                throw new System.Exception("Invalid Excel file: " + inXLS);
            }

            EXCEL.Application xlApp = new EXCEL.Application();
            EXCEL.Workbook xlWorkbook = xlApp.Workbooks.Open(inXLS);
            EXCEL._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

            EXCEL.Range xlRange = xlWorksheet.UsedRange;

            int numCols = xlRange.Columns.Count;
            int numRows = xlRange.Rows.Count;

            int LegacyCol = 0;
            int NewCol = 0;
            int PriorityCol = 0;

            for (int col = 1; col <= numCols; col++)
            {
                if (xlRange.Cells[2, col] != null && xlRange.Cells[2, col].Value2 != null)
                {
                    String cellString = xlRange.Cells[2, col].Value2.ToString().Trim().ToUpper();

                    if (cellString.Contains("NEW LAYER")) { NewCol = col; }
                    if (cellString.Contains("LEGACY")) { LegacyCol = col; }
                    if (cellString.Contains("PRIORITY")) { PriorityCol = col; }
                }
            }
            try
            {
                for (int i = 3; i <= numRows; i++)
                {
                    if ((xlRange.Cells[i, NewCol] != null && xlRange.Cells[i, NewCol].Value2 != null) &&
                         (xlRange.Cells[i, LegacyCol] != null && xlRange.Cells[i, LegacyCol].Value2 != null) &&
                         (xlRange.Cells[i, PriorityCol] != null && xlRange.Cells[i, PriorityCol].Value2 != null))
                    {
                        Row currentRow = new Row(xlRange.Cells[i, LegacyCol].Value2.ToString(),
                                                  xlRange.Cells[i, LegacyCol].Value2.ToString(),
                                                  xlRange.Cells[i, NewCol].Value2.ToString(),
                                                  Convert.ToInt32(xlRange.Cells[i, PriorityCol].Value2));
                        ExcelTable.Add(currentRow);
                    }
                }

                ExcelTable = ExcelTable.OrderByDescending(inRow => inRow.LegacyKey.Length).ToList<Row>();
                //for (int i = 3; i <= numRows; i++)
                //{
                //    if (xlRange.Cells[i, LegacyCol] != null && xlRange.Cells[i, LegacyCol].Value2 != null)
                //    {
                //        ExcelLayers["LegacyLayer"].Add(xlRange.Cells[i, LegacyCol].Value2.ToString());
                //    }
                //}

                //for (int i = 3; i <= numRows; i++)
                //{
                //    if (xlRange.Cells[i, PriorityCol] != null && xlRange.Cells[i, PriorityCol].Value2 != null)
                //    {
                //        ExcelLayers["PriorityLayer"].Add(xlRange.Cells[i, PriorityCol].Value2.ToString());
                //    }
                //}
            }
            catch { throw; }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return ExcelTable;
        }
    }
}
