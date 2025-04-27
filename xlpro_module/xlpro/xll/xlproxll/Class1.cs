using ExcelDna.Integration;
using System;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;   
using Excel = Microsoft.Office.Interop.Excel;
public static class MyFunctions
{
    [ExcelFunction(Description = "My first .NET function")]
    public static string SayHello(string name)
    {
        return "Hello " + name;
    }

    [ExcelFunction(Description = "A useful test function that adds two numbers, and returns the sum.")]
    public static double AddThem(
        [ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")] 
        double v1,
        [ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]     
        double v2)
    {
        return v1 + v2;
    }
    [ExcelFunction(Description = "func")]
    public static string comtest()
    {
        try
        {
            // Attach to running Excel instance
            // dynamic app = Interaction.GetObject(null, "Excel.Application");
            // dynamic app = Marshal.GetActiveObject("Excel.Application");

            // Excel.Application app = Globals.ThisAddIn.Application;

            // Excel.Application app = null;
            // app = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Application xlapp = new Excel.Application();;

            xlapp.Visible = true;

            string ret = xlapp.ActiveWorkbook.Name;

            Console.WriteLine("Active workbook: " + ret);
            return ret;
        }
        catch (COMException ex)
        {
            Console.WriteLine("Excel not running or not accessible: " + ex.Message);
            return null;
        }
    }
    [ExcelFunction(Name = "GetActiveWorkbookName", Description = "Gets the name of the active workbook")]
    public static string GetActiveWorkbookName()
    {
        string workbookName = null;

        // Run code on the main Excel thread using ExcelAsyncUtil
        ExcelAsyncUtil.Run(() =>
        {
            // Access the Excel application object
            Excel.Application excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

            // Check if there is an active workbook and get its name
            if (excelApp.ActiveWorkbook != null)
            {
                workbookName = excelApp.ActiveWorkbook.Name;
            }
            else
            {
                workbookName = "No active workbook";
            }
        });

        return workbookName;
    }
}