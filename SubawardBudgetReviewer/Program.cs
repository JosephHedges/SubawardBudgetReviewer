using OfficeOpenXml;
using SubawardBudgetReviewer.Services;

/*
 * Author: Joseph T. Hedges
 * 2023
 * 
 * This application is designed to extract data from a spreadsheet containing subaward budgets.
 * 
 * Makes use of the EPPlus library under NonCommercial licensing. 
 * Do not use this application in a commercial setting unless you intend to provide a license.
*/

ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // https://polyformproject.org/licenses/noncommercial/1.0.0/

Console.WriteLine("  ***  Welcome to the Subaward Budget Reviewer  ***");

Console.Write("Please enter the path to the folder containing the Subaward Budgets: ");
var budgetPath = Console.ReadLine();

if (string.IsNullOrEmpty(budgetPath))
{
    Console.WriteLine("You must enter a path to the folder containing the Subaward Budgets.");
    return;
}
else if (!Directory.Exists(budgetPath))
{
    Console.WriteLine("The path you entered does not exist.");
    return;
}

var budgets = Directory.GetFiles(budgetPath, "*.xlsx");

Console.WriteLine($"\r\nReading {budgets.Length} spreadsheets...");

var subawards = ExcelReaderService.GetSubawards(budgets);

Console.WriteLine("\r\nTotals:");
var namesWereMissed = false;

foreach (var subaward in subawards)
{
    var line = $"{$"{subaward.Name}:",-14}\t{subaward.Amount:C}";

    if (subaward.Name == "Name Missing")
    {
        line += $"\t({subaward.FileName})";
        namesWereMissed = true;
    }

    Console.WriteLine(line);
}

if (namesWereMissed)
{
    Console.WriteLine("\r\nAt least one provided file did not have a valid Subaward name.\r\nPlease check the file(s) and ensure that the name is provided.");
}

Console.Read();