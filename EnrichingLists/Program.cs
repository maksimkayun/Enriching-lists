// See https://aka.ms/new-console-template for more information

using ClosedXML.Excel;

namespace Enriching_lists
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            var tableMain = GetTimeTable("Введите путь до файла с данными (основной) >");
            var tableInstr = GetTimeTableDirectory("Введите путь до ДИРЕКТОРИИ с файлами с данными (инструкторские) >");
            
            FindAndSetDescription(tableMain, tableInstr);
            
            tableMain.UpdateData();
        }

        internal static Timetable GetTimeTable(string message)
        {
            Console.Write(message);
            var path = Console.ReadLine()?.TrimStart('\"').TrimEnd('\"');
            var timetable = new Timetable(path);
            timetable.LoadData();
            return timetable;
        }

        internal static void FindAndSetDescription(Timetable mainTable, List<Timetable> secondaryTables)
        {
            foreach (var secondaryTable in secondaryTables)
            {
                for (var i = 0; i < secondaryTable.Items.Count; i++)
                {
                    var key = secondaryTable.Items[i].Key;
                    var record = mainTable.Items
                        .FirstOrDefault(e => e.Key == key)?
                        .Description ?? default;

                    if (string.IsNullOrWhiteSpace(mainTable.Items[i].Description) && record != default &&
                        !string.IsNullOrWhiteSpace(record))
                    {
                        mainTable.Items[i].Description = record;
                    }
                }
            }
        }

        internal static List<Timetable> GetTimeTableDirectory(string message)
        {
            Console.Write(message);
            var pathDir = Console.ReadLine()?.TrimStart('\"').TrimEnd('\"');
            var files = Directory.GetFiles(pathDir);

            var result = new List<Timetable>();

            foreach (var path in files)
            {
                var timetable = new Timetable(path);
                timetable.LoadData();

                result.Add(timetable);
            }

            return result;
        }
    }
}