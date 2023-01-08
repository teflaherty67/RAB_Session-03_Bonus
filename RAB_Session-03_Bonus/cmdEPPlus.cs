#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Forms = System.Windows.Forms;

#endregion

namespace RAB_Session_03_Bonus
{
    [Transaction(TransactionMode.Manual)]
    public class cmdEPPlus : IExternalCommand
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

            // set EPPlus license context

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // open Excel file

            ExcelPackage excel = new ExcelPackage(excelFile);
            ExcelWorkbook workbook = excel.Workbook;
            ExcelWorksheet worksheet = workbook.Worksheets[0];

            // get row & column count

            int rows = worksheet.Dimension.Rows;
            int columns = worksheet.Dimension.Columns;

            // read Excel data into a list

            List<List<string>> excelData = new List<List<string>>();

            for (int i = 1; i <= rows; i++)
            {
                List<string> rowData = new List<string>();
                for (int j = 1; j <= columns; j++)
                {
                    string cellContent = worksheet.Cells[i, j].Value.ToString();
                    rowData.Add(cellContent);
                }
                excelData.Add(rowData);
            }

            // create new worksheet

            ExcelWorksheet newWorkSheet = workbook.Worksheets.Add("Test EPPlus");

            // write data to excel

            for (int k = 1; k <= 10; k++)
            {
                for (int j = 1; j <= 10; j++)
                {
                    newWorkSheet.Cells[k,j].Value = "Row " + k.ToString() + ": Column " + j.ToString();
                }
            }

            // save & clsoe excel file

            excel.Save();
            excel.Dispose();

            return Result.Succeeded;
        }
    }
}
