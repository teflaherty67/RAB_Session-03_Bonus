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
using Excel = OfficeOpenXml;
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
            ExcelWorksheet worksheet = workbook.Worksheets[1];            

            return Result.Succeeded;
        }
    }
}
