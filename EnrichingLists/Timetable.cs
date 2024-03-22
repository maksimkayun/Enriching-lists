using ClosedXML.Excel;

namespace Enriching_lists;

public class Timetable
{
    public class RecordObject
    {
        public string NumberTrain { get; set; }
        public DateTime DateTime { get; set; }
        public string Route { get; set; }

        public string? Description { get; set; }

        public List<string>? AdditionalDescriptions { get; set; }

        public string Key => $"{NumberTrain}_{DateTime}_{Route}";

        public void AddDescription(string message)
        {
            if (Description == default || string.IsNullOrWhiteSpace(Description))
            {
                Description = message;
            }
            else if (message == Description)
            {
                Description = message;
            } 
            else
            {
                AdditionalDescriptions ??= new List<string>();
                if (!AdditionalDescriptions.Exists(e=> e == message))
                {
                    AdditionalDescriptions.Add(message);
                }
                
            }
        }
    }
    public List<RecordObject> Items { get; set; } = new();

    public Timetable(string path)
    {
        this.path = path;
    }

    public void LoadData()
    {
        using var wbook = new XLWorkbook(path);
        var sheet = wbook.Worksheet(1);

        var header = sheet.Row(2);
        
        for (var i = 5; i <= sheet.Rows().Count(); i++)
        {
            if (sheet.Row(i).IsHidden)
            {
                continue;
            }

            var row = sheet.Row(i);
            if (string.IsNullOrWhiteSpace(row.Cell(1).GetValue<string>()) || row.IsMerged())
            {
                break;
            }
            
            var obj = new RecordObject
            {
                NumberTrain = row.Cell(header.IndexOfColumn("Номер\nпоезда")).GetValue<string>(),
                DateTime = DateTime.Parse(row.Cell(header.IndexOfColumn("Дата\nотправления")).GetValue<string>()),
                Route = row.Cell(header.IndexOfColumn("Маршрут")).GetValue<string>(),
                Description = row.Cell(header.IndexOfColumn("Причина")).GetValue<string>()
            };
            Items.Add(obj);
        }
    }

    public void UpdateData()
    {
        using var wbook = new XLWorkbook(path);
        var sheet = wbook.Worksheet(1);
        var header = sheet.Row(2);

        for (var i = 5; i <= sheet.Rows().Count(); i++)
        {
            if (sheet.Row(i).IsHidden)
            {
                continue;
            }

            var row = sheet.Row(i);
            if (string.IsNullOrWhiteSpace(row.Cell(1).GetValue<string>()) || row.IsMerged())
            {
                break;
            }
            var key =
                $"{row.Cell(header.IndexOfColumn("Номер\nпоезда")).GetValue<string>()}_" +
                $"{DateTime.Parse(row.Cell(header.IndexOfColumn("Дата\nотправления")).GetValue<string>())}_" +
                $"{row.Cell(header.IndexOfColumn("Маршрут")).GetValue<string>()}";
            
            var increment = header.IndexOfColumn("Причина");
            var value = sheet.Cell(i, increment).GetValue<string>();
            while (!string.IsNullOrWhiteSpace(value))
            {
                increment++;
                value = sheet.Cell(i, increment).GetValue<string>();
            }

            var item = Items.FirstOrDefault(e => e.Key == key);
            sheet.Cell(i, increment).Value = item?.Description;
            if (item?.AdditionalDescriptions != null && item.AdditionalDescriptions.Any())
            {
                foreach (var desc in item.AdditionalDescriptions)
                {
                    increment++;
                    sheet.Cell(i, increment).Value = desc;
                }
            }
        }
        
        wbook.Save();
    }

    private string path { get; set; }
}