using Dekauto.Export.Service.Domain.Entities.DTO;
using Dekauto.Export.Service.Domain.Interfaces;
using OfficeOpenXml;
using System.Globalization;

namespace Dekauto.Export.Service.Domain.Services
{
    public class DiplomaSupplementExportService : IDiplomaSupplementExportService
    {
        private IConfiguration _configuration;
        private readonly ILogger<DiplomaSupplementExportService> _logger;
        private readonly string cellComment = "Значение изначально было установлено автоматически (с помощью Dekauto).";
        private readonly string commentAuthor = "Dekauto";
        private ExcelWorksheet? activeWorksheet;

        public DiplomaSupplementExportService(IConfiguration configuration,
            ILogger<DiplomaSupplementExportService> logger)
        {
            _configuration = configuration;
            _configuration.GetRequiredSection("ExportDiploma");
            cellComment = _configuration.GetValue<string>("ExportCommentText");
            commentAuthor = _configuration.GetValue<string>("ExportCommentAuthor");
            _logger = logger;

        }

        private void AddComment(string cell)
        {
            activeWorksheet.Cells[cell].AddComment(cellComment, commentAuthor);
        }

        private void SetCellValue(string cell, object value)
        {
            object cellValue;

            // Проверяем сначала DateOnly (это сработает и для DateOnly? с значением)
            if (value is DateOnly dateOnly)
            {
                cellValue = FormatDate(dateOnly);
            }
            // Затем проверяем на null
            else if (value is null)
            {
                cellValue = null;
                _logger.LogWarning($"{cell} has received a null value.");
            }
            else
            {
                cellValue = value;
            }

            activeWorksheet.Cells[cell].Value = cellValue;
            AddComment(cell);

            _logger.LogTrace($"Для ячейки {cell} установлено значение \"{cellValue}\"");
        }

        private string FormatDate(DateOnly date)
        {
            var culture = new CultureInfo("ru-RU");
            return date.ToString("d MMMM yyyy 'года'", culture);
        }

        public async Task<(MemoryStream, string)> ExportDiplomaSupplement(DiplomaSupplementExportRequest request)
        {
            ArgumentNullException.ThrowIfNull(request);
            activeWorksheet = null;

            var (templatePath, e1, m1) = ChooseTemplateFile(request.manufacturer, request.educationLevel);
            _logger.LogInformation($"Найден файл шаблона: {m1}, {e1}");

            var diplomaFile = await FillDiplomaSupplementAsync(templatePath, request.supplementData);
            _logger.LogInformation($"Приложение диплома сформировано.");

            string fileName = $"Приложение диплома {request.supplementData.Surname} {request.supplementData.Name} {request.supplementData.Patronymic} {m1} {e1}";

            return (diplomaFile, fileName);
        }

        private (string, string, string) ChooseTemplateFile(string manufacturer, string educationLevel)
        {
            var m = _configuration.GetValue<string>($"Manufacturers:{manufacturer}:path");
            var m1 = _configuration.GetValue<string>($"Manufacturers:{manufacturer}:name");
            var e = _configuration.GetValue<string>($"EducationLevels:{educationLevel}:path");
            var e1 = _configuration.GetValue<string>($"EducationLevels:{educationLevel}:name");

            var templatePath = Path.Combine(Directory.GetCurrentDirectory(), m, e); //Путь шаблона

            if (!File.Exists(templatePath))
            {
                _logger.LogError($"File {templatePath} doesn't exist.");
                throw new FileNotFoundException("Файл шаблона не найден. Обратитесь к администратору");
            }


            return (templatePath, e1, m1);

        }

        private async Task<MemoryStream> FillDiplomaSupplementAsync(string templatePath, DiplomaSupplementData supplementData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var diplomaSupplement = new MemoryStream(); //Используем временное хранилище

            using (var package = new ExcelPackage(new FileInfo(templatePath)))
            {
                FillExcel(supplementData, package);
                await package.SaveAsAsync(diplomaSupplement); //Сохраняем файл
            }
            diplomaSupplement.Position = 0; //Сбрасываем позицию

            return diplomaSupplement;
        }

        private void FillExcel(DiplomaSupplementData data, ExcelPackage package)
        {
            if (package.Workbook.Worksheets.Count == 0)
                throw new InvalidOperationException("Файл шаблона не содержит листов");
            activeWorksheet = null;

            var ws = package.Workbook.Worksheets;
            var diplomaSheet = ws.Where(w => w.Name.Trim()
                .Contains("диплом", StringComparison.OrdinalIgnoreCase)) // ищем "Диплом"
                .MinBy(w => w.Name.Length);
            var ownerSheet = ws.Where(w => w.Name.Trim().Contains("обладат", StringComparison.OrdinalIgnoreCase))
                .First(); // ищем "1 Обладатель диплома"
            var programMasteringSheet = ws.Where(w => w.Name.Trim().Contains("програм", StringComparison.OrdinalIgnoreCase))
                .First(); // ищем "3 Освоение программы"
            //var otherDataSheet = ws.Where(w => w.Name.Trim().Contains("сведен", StringComparison.OrdinalIgnoreCase))
            //  .First(); // ищем "4 доп.сведения"

            _logger.LogInformation($"Начинаем заполнение страницы диплома (лист \"{diplomaSheet.Name}\")...");
            FillDiplomaSheet(diplomaSheet, data);
            _logger.LogInformation("Заполнение страницы диплома завершено.");

            _logger.LogInformation($"Начинаем заполнение страницы обладателя (лист \"{ownerSheet.Name}\")...");
            FillOwnerSheet(ownerSheet, data);
            _logger.LogInformation("Заполнение страницы обладателя завершено.");

            _logger.LogInformation($"Начинаем заполнение страницы освоения программы (лист \"{programMasteringSheet.Name}\")...");
            FillProgramMasteringSheet(programMasteringSheet, data);
            _logger.LogInformation("Заполнение страницы освоения программы завершено.");

            //logger.LogInformation($"Начинаем заполнение страницы доп. сведений (лист \"{otherDataSheet.Name}\")...");
            //FillOtherDataSheet(otherDataSheet, data);
            //logger.LogInformation("Заполнение страницы доп. сведений завершено.");
        }

        private void FillDiplomaSheet(ExcelWorksheet sheet, DiplomaSupplementData data)
        {
            activeWorksheet = sheet;

            if (data is null || data.DiplomaWithHonors is null)
            {
                _logger.LogWarning("Статус диплома \"с отличием\" не найден, пропускаем...");
                return;
            }

            SetCellValue("B6", data.DiplomaWithHonors.Value ? "с отличием" : null);

            activeWorksheet = null;
        }

        private void FillOwnerSheet(ExcelWorksheet sheet, DiplomaSupplementData data)
        {
            activeWorksheet = sheet;

            SetCellValue("B3", data.Surname);
            SetCellValue("B4", data.Name);
            SetCellValue("B5", data.Patronymic);
            SetCellValue("B6", data.BirthdayDate);
            SetCellValue("B12", data.DiplomaWithHonors.Value ? "с отличием" : null);

            activeWorksheet = null;
        }

        private void FillProgramMasteringSheet(ExcelWorksheet sheet, DiplomaSupplementData data)
        {
            activeWorksheet = sheet;

            activeWorksheet = null;
        }
    }
}