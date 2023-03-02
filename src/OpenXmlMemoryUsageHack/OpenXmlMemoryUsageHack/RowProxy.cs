using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace OpenXmlMemoryUsageHack;

internal class RowProxy : Row
{
    private readonly List<string> _cellValues = new();

    private List<OpenXmlElement> Children => MapToCell(_cellValues);
    public override OpenXmlElementList ChildElements => new ElementListProxy(() => Children);
    public override OpenXmlElement FirstChild => Children.FirstOrDefault();
    public override bool HasChildren => Children.Any();

    public void AppendCellValue(string cellValue)
    {
        if (cellValue != null)
        {
            _cellValues.Add(cellValue);
        }
    }

    private static List<OpenXmlElement> MapToCell(IEnumerable<string> cellValues)
    {
        return cellValues.Select(e => (OpenXmlElement)new Cell
        {
            DataType = CellValues.String,
            CellValue = new CellValue(e)
        }).ToList();
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