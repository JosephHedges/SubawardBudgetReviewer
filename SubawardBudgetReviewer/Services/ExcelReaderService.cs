using System.Data;
using OfficeOpenXml;
using SubawardBudgetReviewer.Models;

namespace SubawardBudgetReviewer.Services;

/// <summary>
/// Service holding the necessary logic to read in an array of spreadsheet file paths and determine which subawards are present in each along
/// with their totals.
/// </summary>
public class ExcelReaderService
{
    /// <summary>
    /// Returns an aggregate list of subawards for the provided array of spreadsheet file paths.
    /// </summary>
    /// <param name="budgets">Files to evaluate</param>
    /// <returns>List of subawards</returns>
    public static List<Subaward> GetSubawards(string[] budgets)
    {
        var subawards = new List<Subaward>();

        foreach (var budget in budgets)
        {
            var fileName = budget[(budget.LastIndexOf("\\") + 1)..];

            Console.WriteLine($"\r\n{fileName}");
            Console.WriteLine("---------------------------------------");

            var fileStream = File.OpenRead(budget);
            var budgetSubawards = GetSubawardsFromExcelFile(fileStream, fileName);

            foreach (var subaward in budgetSubawards)
            {
                if (subaward.Name == "Name Missing")
                {
                    subawards.Add(subaward);
                    continue;
                }

                var exists = subawards.FirstOrDefault(s => s.Name == subaward.Name);

                if (exists is not null)
                {
                    exists.Amount += subaward.Amount;
                }
                else
                {
                    subawards.Add(subaward);
                }
            }
        }

        return subawards;
    }

    /// <summary>
    /// Processes the provided file stream to obtain its subawards
    /// </summary>
    /// <param name="fileStream">File stream from disk to evaluate</param>
    /// <param name="fileName">Name of the file</param>
    /// <returns>List of subawards found in the provided file</returns>
    public static List<Subaward> GetSubawardsFromExcelFile(Stream fileStream, string fileName)
    {
        var file = ReadExcelFile(fileStream);

        if (file is null)
        {
            return new();
        }

        var subawards = GetSubawards(file.Columns[1]?.Table?.Rows, fileName);

        return subawards;
    }

    /// <summary>
    /// Parses the Excel file stream to a DataTable for logic processing.
    /// </summary>
    /// <param name="stream"></param>
    /// <returns></returns>
    private static DataTable? ReadExcelFile(Stream? stream)
    {
        if (stream is null)
        {
            return null;
        }

        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets[0];
        var table = new DataTable();
        var columns = worksheet.Dimension.Columns;
        var rows = worksheet.Dimension.Rows;

        for (int i = 1; i <= columns; i++)
        {
            table.Columns.Add(worksheet.Cells[1, i].Value?.ToString() ?? "");
        }

        for (int i = 1; i <= rows; i++)
        {
            var row = table.NewRow();

            for (int j = 1; j <= columns; j++)
            {
                row[j - 1] = worksheet.Cells[i, j].Value;
            }

            table.Rows.Add(row);
        }

        return table;
    }

    /// <summary>
    /// Evaluates the provided rows to determine which are subawards and returns a list of them along with their award totals.<br />
    /// In the event a subaward is missing a name, it will be given the name "Name Missing".
    /// </summary>
    /// <param name="rows"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    private static List<Subaward> GetSubawards(DataRowCollection? rows, string fileName)
    {
        if (rows is null)
        {
            return new();
        }

        var subawards = new List<Subaward>();
        var totalColumn = -1;

        foreach (DataRow row in rows)
        {
            for (var i = 0; i < row.ItemArray.Length; i++)
            {
                var item = row.ItemArray[i];

                if (item is string itemString)
                {
                    if (totalColumn == -1 && itemString.Equals("Total", StringComparison.InvariantCultureIgnoreCase))
                    {
                        totalColumn = i;
                    }
                    else if (itemString.Contains("Subaward:", StringComparison.InvariantCultureIgnoreCase))
                    {
                        var name = string.IsNullOrEmpty(row.ItemArray[2]?.ToString())
                            ? row.ItemArray[1]?.ToString()?.Replace("Subaward:", "").Trim()
                            : row.ItemArray[2]?.ToString()?.Trim();

                        if (string.IsNullOrEmpty(name))
                        {
                            name = "Name Missing";
                        }

                        var subaward = new Subaward
                        {
                            Name = name,
                            Amount = decimal.Parse(row.ItemArray[totalColumn]?.ToString() ?? "0.00"),
                            FileName = fileName
                        };

                        Console.WriteLine($"Subaward: {name}");

                        subawards.Add(subaward);
                        break;
                    }
                }
            }
        }

        return subawards;
    }
}
