using System;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main()
    {
        Console.WriteLine("Starting report creation...");

        // Creating reports
        CreateWordReport(@"C:\report.docx");
        CreateExcelReport(@"C:\report.xlsx");

        Console.WriteLine("Reports successfully created!");
    }

    static void CreateExcelReport(string filePath)
    {
        Console.WriteLine("Creating Excel report...");

        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlBook = xlApp.Workbooks.Add();

        // Creating the second sheet
        CreateSheet(xlBook, "Second Sheet");

        // Creating the first sheet
        CreateSheet(xlBook, "First Sheet");



        xlBook.SaveAs(filePath);
        xlBook.Close();
        xlApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
        GC.Collect();
        GC.WaitForPendingFinalizers();

        Console.WriteLine("Excel report successfully created!");
    }

    static void CreateSheet(Excel.Workbook xlBook, string sheetName)
    {
        Excel.Worksheet xlSheet = (Excel.Worksheet)xlBook.Worksheets.Add();
        xlSheet.Name = sheetName;
        xlSheet.Cells[1, 1] = "User ID";
        xlSheet.Cells[1, 2] = "User Name";
        xlSheet.Cells[1, 3] = "Second Name";
        xlSheet.Cells[1, 4] = "Deposit Balance";
        xlSheet.Cells[1, 5] = "Deposit Interest";

        Random random = new Random();
        for (int i = 1; i <= 30; i++)
        {
            xlSheet.Cells[i + 1, 1] = i;
            xlSheet.Cells[i + 1, 2] = "UsernameName" + random.Next(1, 100);
            xlSheet.Cells[i + 1, 3] = "UserSecondName" + random.Next(1, 100);
            xlSheet.Cells[i + 1, 4] = "" + random.Next(1, 1000);
            xlSheet.Cells[i + 1, 5] = "" + random.Next(1, 10);
        }
    }
    static void CreateWordReport(string filePath)
    {
        Console.WriteLine("Creating Word report...");

        Word.Application wdApp = new Word.Application();
        Word.Document doc = wdApp.Documents.Add();

        // Creating the first table
        CreateWordTable(doc);

        doc.SaveAs2(filePath);
        doc.Close();
        wdApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wdApp);
        GC.Collect();
        GC.WaitForPendingFinalizers();

        Console.WriteLine("Word report successfully created!");
    }

    static void CreateWordTable(Word.Document doc)
    {
        Word.Range range = doc.Paragraphs.Add().Range;
        Word.Table table = doc.Tables.Add(range, 31, 5);
        table.Range.Font.Size = 14;
        table.Range.Font.Name = "Times New Roman";
        table.Borders.Enable = 1;
        table.Cell(1, 1).Range.Text = "User ID";
        table.Cell(1, 2).Range.Text = "User Name";
        table.Cell(1, 3).Range.Text = "Second Name";
        table.Cell(1, 4).Range.Text = "Deposit Balance";
        table.Cell(1, 5).Range.Text = "Deposit Interest";

        Random random = new Random();
        for (int i = 1; i <= 30; i++)
        {
            table.Cell(i + 1, 1).Range.Text = (i).ToString();
            table.Cell(i + 1, 2).Range.Text = "UserName " + random.Next(1, 100);
            table.Cell(i + 1, 3).Range.Text = "UserSecondName " + random.Next(1, 100);
            table.Cell(i + 1, 4).Range.Text = ""+random.Next(1, 1000);
            table.Cell(i + 1, 5).Range.Text =  ""+random.Next(1, 10);
        }

    }
}

