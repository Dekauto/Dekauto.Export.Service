using System.Reflection;
using Dekauto.Export.Service.Domain.Entities;
using Dekauto.Export.Service.Domain.Entities.DTO;
using Dekauto.Export.Service.Domain.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using OfficeOpenXml;

namespace Dekauto.Export.Tests;

[TestClass]
public class ProgramMasteringGiaBlockExportTests
{
    private static DiplomaSupplementExportService CreateService()
    {
        var config = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["ExportDiploma:SupplementDirPath"] = "Templates/diploma_supplement",
                ["ExportCommentText"] = "test",
                ["ExportCommentAuthor"] = "Dekauto"
            })
            .Build();
        return new DiplomaSupplementExportService(config, NullLogger<DiplomaSupplementExportService>.Instance);
    }

    [TestMethod]
    public void FillProgramMasteringSheet_empty_gia_always_writes_header_and_vtom_block()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage();
        var sheet = package.Workbook.Worksheets.Add("3 Освоение программы");

        var data = new DiplomaSupplementData
        {
            DisciplineResults = new List<StudentDisciplineResult>(),
            TargetGiaCreditsFromPlan = 9,
            DiplomaWithHonors = false
        };

        var service = CreateService();
        var method = typeof(DiplomaSupplementExportService).GetMethod(
            "FillProgramMasteringSheet",
            BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.IsNotNull(method);

        method!.Invoke(service, new object[] { sheet, data });

        var columnBTexts = new List<string>();
        var endRow = sheet.Dimension?.End.Row ?? 0;
        for (int r = 1; r <= endRow; r++)
        {
            var t = sheet.Cells[r, 2].Text?.Trim();
            if (!string.IsNullOrEmpty(t))
                columnBTexts.Add(t);
        }

        Assert.IsTrue(
            columnBTexts.Any(t =>
                t.Contains("Государственная итоговая аттестация", StringComparison.OrdinalIgnoreCase)),
            "Ожидался заголовок блока ГИА");
        Assert.IsTrue(
            columnBTexts.Any(t => string.Equals(t.Trim(), "в том числе:", StringComparison.OrdinalIgnoreCase)),
            "Ожидалась строка «в том числе:»");
    }
}
