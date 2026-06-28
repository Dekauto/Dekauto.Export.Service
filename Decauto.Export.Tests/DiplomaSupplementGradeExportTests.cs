using System.Reflection;
using Dekauto.Export.Service.Domain.Entities;
using Dekauto.Export.Service.Domain.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;

namespace Dekauto.Export.Tests;

[TestClass]
public class DiplomaSupplementGradeExportTests
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

    private static string? InvokePrivate(string methodName, params object?[] args)
    {
        var svc = CreateService();
        var method = typeof(DiplomaSupplementExportService).GetMethod(
            methodName,
            BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.IsNotNull(method, $"Method {methodName} not found");
        return (string?)method!.Invoke(svc, args);
    }

    [TestMethod]
    public void FormatCreditsCell_manual_validation_no_tail_in_column_c()
    {
        var item = new StudentDisciplineResult
        {
            CreditUnits = 3,
            RequiresManualValidation = true
        };

        var result = InvokePrivate("FormatCreditsCell", item);

        Assert.AreEqual("3 з.е.", result);
        Assert.IsFalse(result!.Contains("ТРЕБУЕТ ПРОВЕРКИ"));
    }

    [TestMethod]
    public void GetGradeTextCore_bare_zach_numeric_12_maps_to_horosho()
    {
        var item = new StudentDisciplineResult
        {
            ControlType = "зачёт",
            Score = "12"
        };

        var result = InvokePrivate("GetGradeTextCore", item);

        Assert.AreEqual("хорошо", result);
    }

    [TestMethod]
    public void FormatGradeCell_manual_validation_appends_tail_to_column_d()
    {
        var item = new StudentDisciplineResult
        {
            ControlType = "экзамен",
            Score = "12",
            RequiresManualValidation = true
        };

        var result = InvokePrivate("FormatGradeCell", item);

        Assert.AreEqual("хорошо (ТРЕБУЕТ ПРОВЕРКИ)", result);
    }
}
