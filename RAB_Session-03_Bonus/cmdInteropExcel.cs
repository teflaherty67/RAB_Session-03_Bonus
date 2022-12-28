#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

#endregion

namespace RAB_Session_03_Bonus
{
    [Transaction(TransactionMode.Manual)]
    public class cmdInteropExcel : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // prompt user to select Excel file
            Forms.OpenFileDialog selectFile = new Forms.OpenFileDialog();
            selectFile.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";
            selectFile.InitialDirectory = "C:\\";
            selectFile.Multiselect = false;

            string excelFile = "";

            if (selectFile.ShowDialog() == Forms.DialogResult.OK)
                excelFile = selectFile.FileName;

            if (excelFile == "")
            {
                TaskDialog.Show("Error", "Please select an Excel file.");
                return Result.Failed;
            }

            // open Excel file

            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(excelFile);
            Excel.Worksheet worksheet = workbook.Worksheets[1];
            Excel.Range range = (Excel.Range)worksheet.UsedRange;

            // get row & column count

            int rows = range.Rows.Count;
            int columns = range.Columns.Count;

            // read Excel data into a list

            List<List<string>> excelData = new List<List<string>>();

            for(int i = 1; i <= rows; i++)
            {
                List<string> rowData = new List<string>();
                for (int j = 1; j <= columns; j++)
                {
                    string cellContent = worksheet.Cells[i, j].Value.ToString();
                    rowData.Add(cellContent);
                }
                excelData.Add(rowData);
            }

            return Result.Succeeded;
        }
    }
}
