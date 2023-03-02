namespace OpenXmlMemoryUsageHack;

internal class Program
{
    private const int RowCount = 100000;
    private const int CellCount = 10;

    private static void Main(string[] args)
    {
        var data = MockData().ToList();
    }

    private static IEnumerable<List<string>> MockData()
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