using ClosedXML.Excel;

namespace Enriching_lists;

public static class Extensions
{
    public static int IndexOfColumn(this IXLRow row, string text)
    {
        var inc = 1;
        while ((row.Cell(inc).GetValue<string>() != text || row.Worksheet.Column(inc).IsHidden) &&
               (!string.IsNullOrWhiteSpace(row.Cell(inc).GetValue<string>()) || row.Worksheet.Column(inc).IsHidden))
        {
            inc++;
        }

        if (text.Equals("Причина"))
        {
            inc++;
            while ((row.Cell(inc).GetValue<string>() != text || row.Worksheet.Column(inc).IsHidden) &&
                   (!string.IsNullOrWhiteSpace(row.Cell(inc).GetValue<string>()) || row.Worksheet.Column(inc).IsHidden))
            {
                inc++;
            }
        }

        return inc;
    }
}