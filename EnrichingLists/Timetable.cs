using ClosedXML.Excel;

namespace Enriching_lists;

public class Timetable
{
    public class RecordObject
    {
        public string NumberTrain { get; set; }
        public DateTime DateTime { get; set; }
        public string Route { get; set; }
        public string Cause { get; set; }

        public string Description { get; set; }

        public string Key => $"{NumberTrain}_{DateTime}_{Route}_{Cause}";
    }
    public List<RecordObject> Items { get; set; } = new();

    public Timetable(string path)
    {
        this.path = path;
    }

    public void LoadData()
    {
        using var wbook = new XLWorkbook(path);
        var sheet = wbook.Worksheet("Лист1");

        for (var i = 3; i <= sheet.Rows().Count(); i++)
        {
            if (sheet.Row(i).IsHidden)
            {
                continue;
            }

            var row = sheet.Row(i);
            var obj = new RecordObject
            {
                NumberTrain = row.Cell(1).GetValue<string>(),
                DateTime = DateTime.Parse(row.Cell(2).GetValue<string>()),
                Route = row.Cell(3).GetValue<string>(),
                Cause = row.Cell(4).GetValue<string>(),
                Description = row.Cell(11).GetValue<string>()
            };
            Items.Add(obj);
        }
    }

    public void UpdateData()
    {
        using var wbook = new XLWorkbook(path);
        var sheet = wbook.Worksheet("Лист1");

        for (var i = 3; i <= sheet.Rows().Count(); i++)
        {
            if (sheet.Row(i).IsHidden)
            {
                continue;
            }

            var raw = sheet.Row(i);
            var key =
                $"{raw.Cell(1).GetValue<string>()}_" +
                $"{DateTime.Parse(raw.Cell(2).GetValue<string>())}_" +
                $"{raw.Cell(3).GetValue<string>()}_" +
                $"{raw.Cell(4).GetValue<string>()}";
            sheet.Cell(i, 11).Value = Items.FirstOrDefault(e => e.Key == key)?.Description;
        }
        
        wbook.Save();
    }

    private string path { get; set; }
}