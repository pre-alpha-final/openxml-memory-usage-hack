using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlMemoryUsageHack;

internal class Program
{
    private const int RowCount = 100000;
    private const int CellCount = 100;

    private static void Main(string[] args)
    {
        //StandardSave();
        HackySave();

        Console.WriteLine("Done");
        Console.ReadKey();
    }

    /*
     * Standard save using example from https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalue?view=openxml-2.8.1
     * Modified just enough to work with provided data
     */
    private static void StandardSave()
    {
        var fileName = @"standard.xlsx";
        var spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

        // Add a WorkbookPart to the document.
        var workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        // Add a WorksheetPart to the WorkbookPart.
        var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Add Sheets to the Workbook.
        var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

        // Append a new worksheet and associate it with the workbook.
        var sheet = new Sheet()
        {
            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "mySheet"
        };

        sheets.Append(sheet);
        var worksheet = new Worksheet();
        var sheetData = new SheetData();

        uint rowIndex = 1;
        foreach (var row in GetData())
        {
            var openXmlRow = new Row() { RowIndex = rowIndex++ };
            for (var i = 0; i < CellCount; i++)
            {
                var openXmlCell = new Cell()
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(row[i])
                };
                openXmlRow.Append(openXmlCell);
            }
            sheetData.Append(openXmlRow);
        }

        worksheet.Append(sheetData);
        worksheetPart.Worksheet = worksheet;
        workbookpart.Workbook.Save();

        // Close the document.
        spreadsheetDocument.Close();
    }

    private static void HackySave()
    {
        var fileName = @"hacky.xlsx";
        var spreadsheetDocument = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

        // Add a WorkbookPart to the document.
        var workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        // Add a WorksheetPart to the WorkbookPart.
        var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Add Sheets to the Workbook.
        var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

        // Append a new worksheet and associate it with the workbook.
        var sheet = new Sheet()
        {
            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "mySheet"
        };

        sheets.Append(sheet);
        var worksheet = new Worksheet();
        var sheetData = new SheetData();

        uint rowIndex = 1;
        foreach (var row in GetData())
        {
            var openXmlRowProxy = new RowProxy() { RowIndex = rowIndex++ };
            for (var i = 0; i < CellCount; i++)
            {
                openXmlRowProxy.AppendCellValue(row[i]);
            }
            sheetData.Append(openXmlRowProxy);
        }

        worksheet.Append(sheetData);
        worksheetPart.Worksheet = worksheet;
        workbookpart.Workbook.Save();

        // Close the document.
        spreadsheetDocument.Close();
    }


    // Return a string x string table filled with guids (to have some non repeating data)
    private static IEnumerable<List<string>> GetData()
    {
        for (int i = 0; i < RowCount; i++)
        {
            var row = new List<string>();
            for (var j = 0; j < CellCount; j++)
            {
                row.Add(Guid.NewGuid().ToString());
            }

            yield return row;
        }
    }
}