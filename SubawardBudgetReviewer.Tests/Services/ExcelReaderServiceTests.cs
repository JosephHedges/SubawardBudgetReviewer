using System.Reflection;
using FluentAssertions;

namespace SubawardBudgetReviewer.Services.Tests;

public class ExcelReaderServiceTests
{
    [Theory]
    [InlineData("SubawardBudgetExample1.xlsx")]
    public void GetSubawardsFromExcelFile_For_Example1_Should_Have_Four_Subawards(string fileName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var budgetExample1Spreadsheet = assembly.GetManifestResourceStream($"SubawardBudgetReviewer.Tests.Data.{fileName}");

        if (budgetExample1Spreadsheet is null)
        {
            Assert.Fail($"Could not find the file {fileName} in the test project's resources.");
        }

        var subawards = ExcelReaderService.GetSubawardsFromExcelFile(budgetExample1Spreadsheet, fileName);

        subawards.Should().HaveCount(4);
        subawards.Should().Contain(s => s.Name == "Indiana");
        subawards.Should().Contain(s => s.Name == "Mayo");
        subawards.Should().Contain(s => s.Name == "Purdue");
        subawards.Should().Contain(s => s.Name == "Florida");
    }
}