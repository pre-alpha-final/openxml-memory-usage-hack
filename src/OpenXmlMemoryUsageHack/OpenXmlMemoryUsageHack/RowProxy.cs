using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlMemoryUsageHack;

internal class RowProxy : Row
{
    private readonly List<object> _cellValues = new();

    private List<OpenXmlElement> Children => MapToCell(_cellValues);
    public override OpenXmlElementList ChildElements => new ElementListProxy(() => Children);
    public override OpenXmlElement FirstChild => Children.FirstOrDefault();
    public override bool HasChildren => Children.Any();

    public void AppendCellValue(object cellValue)
    {
        _cellValues.Add(cellValue);
    }

    private static List<OpenXmlElement> MapToCell(IEnumerable<object> cellValues)
    {
        return cellValues.Select(MapCell).ToList();
    }

    private static OpenXmlElement MapCell(object cellValue)
    {
        switch (cellValue?.GetType().Name ?? "")
        {
            case "":
            case "String":
                return new Cell { DataType = CellValues.String, CellValue = new CellValue((string)cellValue) };
            case "Guid":
                return new Cell { DataType = CellValues.String, CellValue = new CellValue(cellValue?.ToString()) };
            case "Boolean":
                return new Cell { DataType = CellValues.Boolean, CellValue = new CellValue((bool)cellValue) };
            case "DateTime":
                return new Cell { DataType = CellValues.Date, CellValue = new CellValue((DateTime)cellValue), StyleIndex = 1 };
            case "DateTimeOffset":
                return new Cell { DataType = CellValues.Date, CellValue = new CellValue(((DateTimeOffset)cellValue).DateTime), StyleIndex = 1 };
            case "Int32":
                return new Cell { DataType = CellValues.Number, CellValue = new CellValue((int)cellValue) };
            case "Int64":
                return new Cell { DataType = CellValues.Number, CellValue = new CellValue(Convert.ToDouble((long)cellValue)) };
            case "Single":
                return new Cell { DataType = CellValues.Number, CellValue = new CellValue((float)cellValue) };
            case "Double":
                return new Cell { DataType = CellValues.Number, CellValue = new CellValue((double)cellValue) };
            case "Decimal":
                return new Cell { DataType = CellValues.Number, CellValue = new CellValue((decimal)cellValue) };
            default:
                //return new Cell { DataType = CellValues.String, CellValue = new CellValue(cellValue?.ToString()) };
                throw new ArgumentOutOfRangeException($"{cellValue.GetType().Name} type not supported");
        }
    }

    public class ElementListProxy : OpenXmlElementList
    {
        private readonly List<OpenXmlElement> _list;

        public ElementListProxy(Func<List<OpenXmlElement>> getListFunc)
        {
            _list = getListFunc();
        }

        public override OpenXmlElement GetItem(int index)
        {
            return _list.Skip(index).FirstOrDefault();
        }

        public override IEnumerator<OpenXmlElement> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        public override int Count => _list.Count;
    }
}
