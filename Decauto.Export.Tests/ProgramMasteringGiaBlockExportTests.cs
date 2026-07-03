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

        InvokeFillProgramMasteringSheet(sheet, data, "bachelor");

        var columnBTexts = CollectColumnBTexts(sheet);

        Assert.IsTrue(
            columnBTexts.Any(t =>
                t.Contains("Государственная итоговая аттестация", StringComparison.OrdinalIgnoreCase)),
            "Ожидался заголовок блока ГИА");
        Assert.IsTrue(
            columnBTexts.Any(t => string.Equals(t.Trim(), "в том числе:", StringComparison.OrdinalIgnoreCase)),
            "Ожидалась строка «в том числе:»");
    }

    [TestMethod]
    public void FillProgramMasteringSheet_vkr_writes_title_and_topic_on_separate_rows()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage();
        var sheet = package.Workbook.Worksheets.Add("3 Освоение программы");

        var data = new DiplomaSupplementData
        {
            DisciplineResults = new List<StudentDisciplineResult>
            {
                new()
                {
                    DisciplineName = "Разработка веб-приложения",
                    ControlType = "защита вкр",
                    Score = "5",
                    CreditUnits = 0,
                    AudHours = 0
                }
            },
            TargetGiaCreditsFromPlan = 9,
            DiplomaWithHonors = false
        };

        InvokeFillProgramMasteringSheet(sheet, data, "bachelor");

        var columnBTexts = CollectColumnBTexts(sheet);

        Assert.IsTrue(
            columnBTexts.Any(t =>
                t.Contains("бакалаврская работа", StringComparison.OrdinalIgnoreCase)),
            "Ожидался заголовок ВКР для бакалавриата");
        Assert.IsTrue(
            columnBTexts.Any(t => t.Contains("Разработка веб-приложения", StringComparison.OrdinalIgnoreCase)),
            "Ожидалась тема ВКР на отдельной строке");
    }

    [TestMethod]
    public void FillProgramMasteringSheet_card_only_course_work_appears_in_course_block()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage();
        var sheet = package.Workbook.Worksheets.Add("3 Освоение программы");

        var data = new DiplomaSupplementData
        {
            DisciplineResults = new List<StudentDisciplineResult>
            {
                new()
                {
                    DisciplineName = "НАЗВАНИЕ ДИСЦИПЛИНЫ \"Курсовая работа номер 1\"",
                    ControlType = "Курсовая работа",
                    Score = "4",
                    Semester = 3,
                    IsCardOnlyUnmatchedPlan = true
                }
            },
            DiplomaWithHonors = false
        };

        InvokeFillProgramMasteringSheet(sheet, data, "bachelor");

        var columnBTexts = CollectColumnBTexts(sheet);

        Assert.IsTrue(
            columnBTexts.Any(t => t.Contains("Курсовая работа номер 1", StringComparison.OrdinalIgnoreCase)),
            "Курсовая из карточки должна попасть в блок курсовых");
        Assert.IsFalse(
            columnBTexts.Any(t =>
                t.Contains("Неопределенные записи", StringComparison.OrdinalIgnoreCase)),
            "Несопоставленная курсовая не должна уходить только в ручной блок");
    }

    private static void InvokeFillProgramMasteringSheet(
        OfficeOpenXml.ExcelWorksheet sheet,
        DiplomaSupplementData data,
        string? educationLevel)
    {
        var service = CreateService();
        var method = typeof(DiplomaSupplementExportService).GetMethod(
            "FillProgramMasteringSheet",
            BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.IsNotNull(method);
        method!.Invoke(service, new object?[] { sheet, data, educationLevel });
    }

    private static List<string> CollectColumnBTexts(OfficeOpenXml.ExcelWorksheet sheet)
    {
        var columnBTexts = new List<string>();
        var endRow = sheet.Dimension?.End.Row ?? 0;
        for (int r = 1; r <= endRow; r++)
        {
            var t = sheet.Cells[r, 2].Text?.Trim();
            if (!string.IsNullOrEmpty(t))
                columnBTexts.Add(t);
        }
        return columnBTexts;
    }
}
