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
        // Standard save
        //SaveXml("standard.xlsx", StandardRowGeneration);

        // Hacky save
        SaveXml("hacky.xlsx", HackyRowGeneration);

        Console.WriteLine("Done");
        Console.ReadKey();
    }

    private static Row StandardRowGeneration(uint rowIndex, List<string> row)
    {
        var openXmlRow = new Row() { RowIndex = rowIndex };
        for (var i = 0; i < CellCount; i++)
        {
            var openXmlCell = new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue(row[i])
            };
            openXmlRow.Append(openXmlCell);
        }
        Console.SetCursorPosition(0, 0);
        Console.Write($"Row generation {rowIndex * 100 / RowCount} %");

        return openXmlRow;
    }

    private static Row HackyRowGeneration(uint rowIndex, List<object> row)
    {
        var openXmlRow = new RowProxy() { RowIndex = rowIndex };
        for (var i = 0; i < CellCount; i++)
        {
            openXmlRow.AppendCellValue(row[i]);
        }
        Console.SetCursorPosition(0, 0);
        Console.Write($"Row generation {rowIndex * 100 / RowCount} %");

        return openXmlRow;
    }

    // Modified save using example from https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalue?view=openxml-2.8.1
    private static void SaveXml(string filename, Func<uint, List<object>, OpenXmlElement> generateRow)
    {
        var spreadsheetDocument = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);

        // Add a WorkbookPart to the document.
        var workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        // Add minimalist styles
        var stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet
        {
            Fonts = new Fonts(new Font()),
            Fills = new Fills(new Fill()),
            Borders = new Borders(new Border()),
            CellStyleFormats = new CellStyleFormats(new CellFormat()),
            CellFormats =
                new CellFormats(
                    new CellFormat(),
                    new CellFormat { NumberFormatId = 14, ApplyNumberFormat = true })
        };

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
            var openXmlRow = generateRow(rowIndex++, row);
            sheetData.Append(openXmlRow);
        }

        worksheet.Append(sheetData);
        worksheetPart.Worksheet = worksheet;
        Console.WriteLine("\nSaving file");
        workbookpart.Workbook.Save();

        // Close the document.
        spreadsheetDocument.Close();
    }

    // Return a string table filled with guids (to have some non repeating data)
    private static IEnumerable<List<object>> GetData()
    {
        for (int i = 0; i < RowCount; i++)
        {
            var row = new List<object>();
            for (var j = 0; j < CellCount; j++)
            {
                row.Add(Guid.NewGuid().ToString());
            }

            yield return row;
        }
    }
}
