using ClosedXML.Excel;
using CsvHelper;
using System.Data;

if (args == null || args.Length == 0)
    throw new ArgumentNullException(nameof(args));

string filePath = args[0];
if (string.IsNullOrEmpty(filePath))
{
    //throw new ArgumentNullException(nameof(filePath));
    Console.WriteLine($"the file doesn't exist: {filePath}");
    return;
}

using (var reader = new StreamReader(filePath))
{
    using (var csv = new CsvReader(reader, System.Globalization.CultureInfo.InvariantCulture))
    {
        // do any configuration to `CsvReader` before creating CsvDataReader
        using (var dr = new CsvDataReader(csv))
        {
            DataTable dt = new DataTable();
            dt.Load(dr);
            if (dt==null || dt.Rows.Count == 0)
            {
                Console.WriteLine("empty csv file, existing program...");
            }
            else
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                var fileFolder = Path.GetDirectoryName(fileName);
                using(XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dt, fileName);
                    wb.SaveAs($"{fileFolder}/{fileName}.xlsx");
                    Console.WriteLine("done.");
                }
            }
        }
        Console.WriteLine("press any key to continue...");
        Console.ReadKey();
    }
}