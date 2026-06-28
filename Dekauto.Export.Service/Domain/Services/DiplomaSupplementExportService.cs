using Dekauto.Export.Service.Domain.Entities;
using Dekauto.Export.Service.Domain.Entities.DTO;
using Dekauto.Export.Service.Domain.Interfaces;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Dekauto.Export.Service.Domain.Services
{
    public class DiplomaSupplementExportService : IDiplomaSupplementExportService
    {
        private IConfiguration _configuration;
        private readonly ILogger<DiplomaSupplementExportService> _logger;
        private readonly string cellComment = "Значение изначально было установлено автоматически (с помощью Dekauto).";
        private readonly string commentAuthor = "Dekauto";
        private ExcelWorksheet? activeWorksheet;

        /// <summary>Плейсхолдер в столбце C/D как в образце РИД (кириллическое «х»).</summary>
        private const string Xmark = "\u0445"; // кирилл. х

        private const string ManualReviewTail = " (ТРЕБУЕТ ПРОВЕРКИ)";

        private static readonly Dictionary<string, string> KnownTextGradeMap = new(StringComparer.OrdinalIgnoreCase)
        {
            { "отлично", "отлично" },
            { "хорошо", "хорошо" },
            { "удовлетворительно", "удовлетворительно" },
            { "неудовлетворительно", "неудовлетворительно" },
            { "зачтено", "зачтено" },
            { "не зачтено", "не зачтено" },
            { "х", "\u0445" },
            { "x", "\u0445" },
        };

        // Поля для отслеживания текущего состояния записи
        private int _currentRow;
        private int _itemIndex;

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
            if (activeWorksheet == null) return;
            var r = activeWorksheet.Cells[cell];
            if (r.Comment != null)
            {
                r.Comment.Text = cellComment;
                r.Comment.Author = commentAuthor;
            }
            else
                r.AddComment(cellComment, commentAuthor);
        }

        private void SetCellValue(int row, int col, object? value, bool addComment = true)
        {
            if (activeWorksheet == null) return;
            var cell = activeWorksheet.Cells[row, col];

            if (value is null)
            {
                cell.Value = null;
            }
            else
            {
                cell.Value = value;
                if (addComment)
                {
                    if (cell.Comment != null)
                    {
                        cell.Comment.Text = cellComment;
                        cell.Comment.Author = commentAuthor;
                    }
                    else
                    {
                        cell.AddComment(cellComment, commentAuthor);
                    }
                }
            }
        }

        private static bool IsProgramMasteringNoiseName(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return true;
            var lower = raw.Trim().ToLowerInvariant();
            if (lower.Contains("итого") && lower.Contains("семестр"))
                return true;
            if (lower.StartsWith("итого", StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(lower, "в том числе:", StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(lower, "в том числе", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        private static bool IsProbableGiaDisciplineName(string nameLower)
        {
            return nameLower.Contains("государственн") ||
                   nameLower.Contains("выпускная") ||
                   nameLower.Contains("квалификационн") ||
                   nameLower.Contains("защита вкр") ||
                   nameLower.Contains("итоговый") ||
                   (nameLower.Contains("защит") && nameLower.Contains("выпускн"));
        }

        /// <summary>Совпадение с блоком практик плана после разбиения длинных имён или без якоря «Блок 2» на листе.</summary>
        private static bool IsLikelyPracticeBlockDisciplineName(string? disciplineName)
        {
            if (string.IsNullOrWhiteSpace(disciplineName))
                return false;
            var n = disciplineName.Trim();
            var nl = n.ToLowerInvariant();
            if (nl.Contains("проектный практикум"))
                return false;
            if (n.StartsWith("Учебная практика", StringComparison.OrdinalIgnoreCase))
                return true;
            if (n.StartsWith("Производственная практика", StringComparison.OrdinalIgnoreCase))
                return true;
            if (nl.Contains("научно-исследовательская работа") ||
                nl.Contains("научно-исследовательской работы"))
                return true;
            if (nl.Contains("практика по получению профессиональных"))
                return true;
            if (nl.StartsWith("навыков ", StringComparison.OrdinalIgnoreCase))
                return true;
            if (nl.Contains("научно-исследовательской работы)"))
                return true;
            if (nl.StartsWith("умений ", StringComparison.OrdinalIgnoreCase) && nl.Contains("опыта"))
                return true;
            if (nl.Contains("преддипломная") && nl.Contains("практик"))
                return true;
            if (nl.StartsWith("выпускной квалификационной работы", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        private static bool IsCardOnlyGarbageExportPlaceholder(string? disciplineName)
        {
            if (string.IsNullOrWhiteSpace(disciplineName))
                return true;
            var t = disciplineName.Trim();
            if (t.StartsWith("НАЗВАНИЕ ДИСЦИПЛИНЫ", StringComparison.OrdinalIgnoreCase))
                return true;
            if (t.Contains("Тема курсовой работы", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        private void SetCellValue(string cellAddr, object value)
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
                _logger.LogWarning($"{cellAddr} has received a null value.");
            }
            else
            {
                cellValue = value;
            }

            activeWorksheet.Cells[cellAddr].Value = cellValue;
            AddComment(cellAddr);

            _logger.LogTrace($"Для ячейки {cellAddr} установлено значение \"{cellValue}\"");
        }

        private string FormatDate(DateOnly date)
        {
            // Массив названий месяцев в родительном падеже
            string[] monthsGenitive =
            {
                "января", "февраля", "марта", "апреля", "мая", "июня",
                "июля", "августа", "сентября", "октября", "ноября", "декабря"
            };

            // date.Month возвращает от 1 до 12, поэтому вычитаем 1 для индекса
            string monthName = monthsGenitive[date.Month - 1];

            return $"{date.Day} {monthName} {date.Year} года";
        }
        public async Task<(MemoryStream, string)> ExportDiplomaSupplement(DiplomaSupplementExportRequest request)
        {
            ArgumentNullException.ThrowIfNull(request);
            activeWorksheet = null;

            var (templatePath, e1, m1) = ChooseTemplateFile(request.manufacturer, request.educationLevel);
            _logger.LogInformation($"Найден файл шаблона: {m1}, {e1}");

            var diplomaFile = await FillDiplomaSupplementAsync(templatePath, request.data, request.educationLevel);
            _logger.LogInformation($"Приложение диплома сформировано.");

            var fileNameParts = new List<string?> { "Приложение диплома", request.data.Surname, request.data.Name };
            var patronymicTrim = request.data.Patronymic?.Trim();
            if (!string.IsNullOrEmpty(patronymicTrim) && patronymicTrim != "-" && patronymicTrim != "\u2013")
                fileNameParts.Add(patronymicTrim);
            fileNameParts.Add(m1);
            fileNameParts.Add(e1);
            string fileName = string.Join(" ", fileNameParts.Where(static s => !string.IsNullOrWhiteSpace(s)));

            return (diplomaFile, fileName);
        }

        private string? GetDiplomaExportConfigValue(string key)
        {
            var fromSection = _configuration.GetSection("ExportDiploma").GetValue<string>(key);
            if (!string.IsNullOrEmpty(fromSection))
                return fromSection;
            return _configuration.GetValue<string>(key);
        }

        private (string, string, string) ChooseTemplateFile(string manufacturer, string educationLevel)
        {
            var m = GetDiplomaExportConfigValue($"Manufacturers:{manufacturer}:path");
            var m1 = GetDiplomaExportConfigValue($"Manufacturers:{manufacturer}:name");
            var e = GetDiplomaExportConfigValue($"EducationLevels:{educationLevel}:path");
            var e1 = GetDiplomaExportConfigValue($"EducationLevels:{educationLevel}:name");

            var templatePath = Path.Combine(Directory.GetCurrentDirectory(), m, e); //Путь шаблона

            if (!File.Exists(templatePath))
            {
                _logger.LogError($"File {templatePath} doesn't exist.");
                throw new FileNotFoundException("Файл шаблона не найден. Обратитесь к администратору");
            }


            return (templatePath, e1, m1);

        }

        private async Task<MemoryStream> FillDiplomaSupplementAsync(string templatePath, DiplomaSupplementData supplementData, string? educationLevel)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var diplomaSupplement = new MemoryStream(); //Используем временное хранилище

            using (var package = new ExcelPackage(new FileInfo(templatePath)))
            {
                FillExcel(supplementData, package, educationLevel);
                await package.SaveAsAsync(diplomaSupplement); //Сохраняем файл
            }
            diplomaSupplement.Position = 0; //Сбрасываем позицию

            return diplomaSupplement;
        }

        private void FillExcel(DiplomaSupplementData data, ExcelPackage package, string? educationLevel)
        {
            if (package.Workbook.Worksheets.Count == 0)
                throw new InvalidOperationException("Файл шаблона не содержит листов");
            activeWorksheet = null;

            foreach (ExcelWorksheet s in package.Workbook.Worksheets)
                s.Protection.IsProtected = false;

            var ws = package.Workbook.Worksheets;
            var diplomaSheet = ws.Where(w => w.Name.Trim()
                .Contains("диплом", StringComparison.OrdinalIgnoreCase)) // ищем "Диплом"
                .MinBy(w => w.Name.Length);
            var ownerSheet = ws.Where(w => w.Name.Trim().Contains("обладат", StringComparison.OrdinalIgnoreCase))
                .First(); // ищем "1 Обладатель диплома"
            var programMasteringSheet = ws.Where(w => w.Name.Trim().Contains("програм", StringComparison.OrdinalIgnoreCase))
                .First(); // ищем "3 Освоение программы"
            var otherDataSheet = ws.FirstOrDefault(w =>
            {
                var n = w.Name.Trim();
                return n.Contains("доп", StringComparison.OrdinalIgnoreCase)
                    && n.Contains("свед", StringComparison.OrdinalIgnoreCase);
            });

            _logger.LogInformation($"Начинаем заполнение страницы диплома (лист \"{diplomaSheet.Name}\")...");
            FillDiplomaSheet(diplomaSheet, data);
            _logger.LogInformation("Заполнение страницы диплома завершено.");

            _logger.LogInformation($"Начинаем заполнение страницы обладателя (лист \"{ownerSheet.Name}\")...");
            FillOwnerSheet(ownerSheet, data);
            _logger.LogInformation("Заполнение страницы обладателя завершено.");

            _logger.LogInformation($"Начинаем заполнение страницы освоения программы (лист \"{programMasteringSheet.Name}\")...");
            FillProgramMasteringSheet(programMasteringSheet, data);
            _logger.LogInformation("Заполнение страницы освоения программы завершено.");

            if (otherDataSheet != null)
            {
                _logger.LogInformation($"Начинаем заполнение страницы доп. сведений (лист \"{otherDataSheet.Name}\")...");
                FillOtherDataSheet(otherDataSheet, data, educationLevel);
                _logger.LogInformation("Заполнение страницы доп. сведений завершено.");
            }
            else
                _logger.LogWarning("Лист «4 доп.сведения» не найден по имени, пропуск.");
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
            ApplyAttentionFillToHonorStatusRow("B6");

            activeWorksheet = null;
        }

        private void FillOwnerSheet(ExcelWorksheet sheet, DiplomaSupplementData data)
        {
            activeWorksheet = sheet;

            SetCellValue("B3", data.Surname);
            SetCellValue("B4", data.Name);
            SetCellValue("B5", OwnerSheetPatronymicCellValue(data.Patronymic));
            SetCellValue("B6", data.BirthdayDate);
            SetCellValue("B8", MapEducationDocumentForSupplementOwnerB8(data.EducationReceived));
            SetCellValue("B9", data.EducationReceivedDate is not null
                ? $"{data.EducationReceivedDate.Value.Year} год"
                : null);
            if (!string.IsNullOrWhiteSpace(data.SupplementOwnerQualification))
            {
                SetCellValue("B11", data.SupplementOwnerQualification);
                ApplyAttentionFillToHonorStatusRow("B11");
            }
            SetCellValue("B12", data.DiplomaWithHonors == true ? "с отличием" : null);
            ApplyAttentionFillToHonorStatusRow("B12");
            SetCellValue("B14", data.CourseOfTraining);

            activeWorksheet = null;
        }

        private void FillOtherDataSheet(ExcelWorksheet sheet, DiplomaSupplementData data, string? educationLevel)
        {
            activeWorksheet = sheet;

            var b5Probe = sheet.Cells["B5"].Text?.Trim() ?? "";
            var isSbmLayout = !string.IsNullOrWhiteSpace(b5Probe);
            var isSpecialist = string.Equals(educationLevel?.Trim(), "specialist", StringComparison.OrdinalIgnoreCase);

            string? nameLine = null;
            var opopRaw = data.SupplementAdditionalSheetOpopName;
            if (!string.IsNullOrWhiteSpace(opopRaw))
            {
                var t = opopRaw.Trim();
                nameLine = isSpecialist ? "Специализация: " + t : t;
            }

            var nameWrapped = string.IsNullOrWhiteSpace(nameLine)
                ? null
                : string.Join(Environment.NewLine, SplitText(nameLine.Trim(), 95));

            if (isSbmLayout)
            {
                if (isSpecialist)
                {
                    if (!string.IsNullOrWhiteSpace(nameWrapped))
                        SetCellValue("B5", nameWrapped);
                    if (!string.IsNullOrWhiteSpace(data.SupplementAdditionalSheetStudyFormLine))
                        SetCellValue("B6", data.SupplementAdditionalSheetStudyFormLine);
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(nameWrapped))
                        SetCellValue("B6", nameWrapped);
                    if (!string.IsNullOrWhiteSpace(data.SupplementAdditionalSheetStudyFormLine))
                        SetCellValue("B7", data.SupplementAdditionalSheetStudyFormLine);
                }
            }
            else
            {
                if (isSpecialist)
                {
                    if (!string.IsNullOrWhiteSpace(nameWrapped))
                        SetCellValue("B2", nameWrapped);
                    if (!string.IsNullOrWhiteSpace(data.SupplementAdditionalSheetStudyFormLine))
                        SetCellValue("B3", data.SupplementAdditionalSheetStudyFormLine);
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(nameWrapped))
                        SetCellValue("B3", nameWrapped);
                    if (!string.IsNullOrWhiteSpace(data.SupplementAdditionalSheetStudyFormLine))
                        SetCellValue("B4", data.SupplementAdditionalSheetStudyFormLine);
                }
            }

            _logger.LogInformation(
                "Лист доп. сведений: сценарий={Layout}, уровень={Level}, специалитет={Spec}",
                isSbmLayout ? "СБМ(B5 не пусто)" : "верхние строки(B5 пусто)",
                educationLevel ?? "",
                isSpecialist);

            activeWorksheet = null;
        }

        private static string? MapEducationDocumentForSupplementOwnerB8(string? educationReceivedRaw)
        {
            var s = NormalizeOwnerEducationField(educationReceivedRaw);
            if (string.IsNullOrWhiteSpace(s))
                return null;
            var low = s.ToLowerInvariant();
            if ((low.Contains("среднем") && low.Contains("профессиональн")) ||
                (low.Contains("среднее") && low.Contains("профессиональн")))
                return "Диплом о среднем профессиональном образовании";
            if ((low.Contains("среднем") && low.Contains("общем")) ||
                (low.Contains("среднее") && low.Contains("общее")))
                return "Аттестат о среднем общем образовании";
            if (low.Contains("высшем") && (low.Contains("квалификац") || low.Contains("образовани")))
                return "Документ о высшем образовании";
            return s;
        }

        private static string NormalizeOwnerEducationField(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return "";
            var s = raw.Replace('\u00A0', ' ').Trim();
            s = Regex.Replace(s, @"\s+", " ");
            if (s.Length >= 2 && s.StartsWith('"') && s.EndsWith('"'))
                s = s.Substring(1, s.Length - 2).Trim();
            return s;
        }

        private static object OwnerSheetPatronymicCellValue(string? patronymic)
        {
            var t = patronymic?.Trim() ?? "";
            if (t.Length == 0 || t == "-" || t == "\u2013")
                return " ";
            return t;
        }

        private const int ProgramMasteringColumnFPixels = 60;

        private static double ColumnWidthFromPixels(int pixels) => (pixels - 5) / 7.0;

        private void FillProgramMasteringSheet(ExcelWorksheet sheet, DiplomaSupplementData data)
        {
            activeWorksheet = sheet;
            _logger.LogInformation("Заполнение листа освоения программы...");

            sheet.Column(6).Width = ColumnWidthFromPixels(ProgramMasteringColumnFPixels);
            sheet.Cells[1, 6].Value = "За семестр";

            // 1. Очистка данных; E — как в шаблоне РИД, формула LEN(B), F — служебные надписи
            sheet.Cells["A2:F140"].Value = null;

            // 2. Подготовка и нормализация данных
            var allResults = data.DisciplineResults ?? new List<StudentDisciplineResult>();

            var rawPractices = new List<StudentDisciplineResult>();
            var rawGia = new List<StudentDisciplineResult>();
            var rawCourseWorks = new List<StudentDisciplineResult>();
            var rawElectives = new List<StudentDisciplineResult>();
            var rawDisciplines = new List<StudentDisciplineResult>();
            var rawCardOnlyUnmatched = new List<StudentDisciplineResult>();

            foreach (var item in allResults)
            {
                if (IsProgramMasteringNoiseName(item.DisciplineName))
                    continue;

                if (item.IsCardOnlyUnmatchedPlan)
                {
                    rawCardOnlyUnmatched.Add(item);
                    continue;
                }

                string nameLower = item.DisciplineName?.ToLower()?.Trim() ?? "";
                string controlLower = item.ControlType?.ToLower()?.Trim() ?? "";

                bool isProjectPractice = nameLower.Contains("проектный практикум");

                // курсовые — как раньше, перекрывает привязку плана при конфликте
                bool isCourseWork =
                    controlLower.Contains("курсовая") ||
                    nameLower.Contains("курсовая") ||
                    nameLower.StartsWith("название дисциплины");

                if (isCourseWork)
                {
                    rawCourseWorks.Add(item);
                    continue;
                }

                var pb = item.PlanBucket;

                if (pb == SupplementPlanBucket.Practice)
                {
                    if (!isProjectPractice)
                        rawPractices.Add(item);
                    else
                        rawDisciplines.Add(item);
                    continue;
                }

                if (pb == SupplementPlanBucket.Elective)
                {
                    rawElectives.Add(item);
                    continue;
                }

                if (pb == SupplementPlanBucket.Gia)
                {
                    rawGia.Add(item);
                    continue;
                }

                bool treatAsPractice =
                    ((pb == SupplementPlanBucket.Discipline ||
                       pb == SupplementPlanBucket.Unknown ||
                       !pb.HasValue) &&
                     IsLikelyPracticeBlockDisciplineName(item.DisciplineName) &&
                     !isProjectPractice &&
                     !IsProbableGiaDisciplineName(nameLower));

                if (treatAsPractice)
                {
                    rawPractices.Add(item);
                    continue;
                }

                bool treatAsGia =
                    (pb == SupplementPlanBucket.Discipline ||
                     pb == SupplementPlanBucket.Unknown ||
                     !pb.HasValue) &&
                    IsProbableGiaDisciplineName(nameLower) &&
                    !IsLikelyPracticeBlockDisciplineName(item.DisciplineName) &&
                    !isProjectPractice;

                if (treatAsGia)
                {
                    rawGia.Add(item);
                    continue;
                }

                if (pb == SupplementPlanBucket.Discipline)
                {
                    rawDisciplines.Add(item);
                    continue;
                }

                if (IsProbableGiaDisciplineName(nameLower))
                    rawGia.Add(item);
                else
                    rawDisciplines.Add(item);
            }

            // 3. Агрегация и сортировка
            var disciplines = ProcessDisciplines(rawDisciplines);
            var practices = ProcessDisciplines(rawPractices);
            var electives = ProcessDisciplines(rawElectives);
            var giaResults = ProcessDisciplines(rawGia);
            var courseWorks = ProcessDisciplines(rawCourseWorks);
            var cardOnlyTail = ProcessDisciplines(rawCardOnlyUnmatched)
                .Where(x => !IsCardOnlyGarbageExportPlaceholder(x.DisciplineName))
                .ToList();

            var practicesOutsidePlanCanon = practices
                .Where(p => !PracticeTitleLooksCanonicalFromPlanSheet(p.DisciplineName))
                .ToList();
            var practicesMain = practices
                .Where(p => PracticeTitleLooksCanonicalFromPlanSheet(p.DisciplineName))
                .ToList();

            var manualReconciliationRows = new List<StudentDisciplineResult>();
            manualReconciliationRows.AddRange(practicesOutsidePlanCanon);
            foreach (var lone in cardOnlyTail)
            {
                if (manualReconciliationRows.Any(m =>
                        string.Equals(m.DisciplineName?.Trim(), lone.DisciplineName?.Trim(),
                            StringComparison.OrdinalIgnoreCase)))
                    continue;
                manualReconciliationRows.Add(lone);
            }

            int attentionRowCount =
                disciplines.Count(x => x.RequiresManualValidation)
                + practicesMain.Count(x => x.RequiresManualValidation)
                + giaResults.Count(x => x.RequiresManualValidation)
                + courseWorks.Count(x => x.RequiresManualValidation)
                + electives.Count(x => x.RequiresManualValidation)
                + manualReconciliationRows.Count;

            // для логики по умолчанию (до формул)
            double summedCredits = disciplines.Sum(x => ConvertToDouble(x.CreditUnits))
                                + practicesMain.Sum(x => ConvertToDouble(x.CreditUnits))
                                + giaResults.Sum(x => ConvertToDouble(x.CreditUnits));
            double fallbackTotalCredits = data.TargetProgramCredits ?? summedCredits;

            double totalAudHours = disciplines.Sum(x => ConvertToDouble(x.AudHours))
                                 + practicesMain.Sum(x => ConvertToDouble(x.AudHours))
                                 + giaResults.Sum(x => ConvertToDouble(x.AudHours));

            _currentRow = 2;

            foreach (var item in disciplines)
            {
                WriteDisciplineRow(
                    item.DisciplineName,
                    FormatCreditsCell(item),
                    FormatGradeCell(item),
                    requiresManualAttention: item.RequiresManualValidation,
                    yellowAuxBlock: false,
                    columnFValueOverride: FormatSemesterColumnF(item));
            }

            if (practicesMain.Any())
            {
                double practiceCreditsSummed = practicesMain.Sum(p => ConvertToDouble(p.CreditUnits));
                double practiceCreditsForHeaderRow = practiceCreditsSummed;
                if (data.TargetPracticeCreditsFromPlan is double tPrac && tPrac > 0)
                    practiceCreditsForHeaderRow = tPrac;

                int bundleRows = MeasureNameRowLines("Практики") + MeasureNameRowLines("в том числе:")
                    + MeasureNameRowLines(practicesMain[0].DisciplineName);
                AdvancePastFirstSheetPrintBandIfNeeded(bundleRows);

                WriteDisciplineRow(
                    "Практики",
                    practiceCreditsForHeaderRow > 0
                        ? string.Format(CultureInfo.InvariantCulture, "{0} з.е.", practiceCreditsForHeaderRow)
                        : FormatCredits(practiceCreditsSummed > 0 ? practiceCreditsSummed : null),
                    Xmark,
                    requiresManualAttention: false,
                    yellowAuxBlock: true,
                    skipPageBandAdjustment: true);

                WriteDisciplineRow("в том числе:", null, null,
                    requiresManualAttention: false,
                    yellowAuxBlock: true,
                    skipPageBandAdjustment: true);

                foreach (var item in practicesMain)
                {
                    WriteDisciplineRow(
                        item.DisciplineName,
                        FormatCreditsCell(item),
                        FormatGradeCell(item),
                        requiresManualAttention: item.RequiresManualValidation,
                        yellowAuxBlock: true,
                        columnFValueOverride: FormatSemesterColumnF(item));
                }
            }

            if (giaResults.Any())
            {
                double giaCreditsSummed = giaResults.Sum(g => ConvertToDouble(g.CreditUnits));
                double giaCreditsForHeaderRow = giaCreditsSummed;
                if (giaCreditsForHeaderRow <= 0 &&
                    data.TargetGiaCreditsFromPlan is double tGia &&
                    tGia > 0)
                {
                    giaCreditsForHeaderRow = tGia;
                }

                int giaBundle = MeasureNameRowLines("Государственная итоговая аттестация")
                    + MeasureNameRowLines("в том числе:")
                    + MeasureNameRowLines(giaResults[0].DisciplineName);
                AdvancePastFirstSheetPrintBandIfNeeded(giaBundle);

                WriteDisciplineRow(
                    "Государственная итоговая аттестация",
                    giaCreditsForHeaderRow > 0
                        ? string.Format(CultureInfo.InvariantCulture, "{0} з.е.", giaCreditsForHeaderRow)
                        : FormatCredits(giaCreditsSummed > 0 ? giaCreditsSummed : null),
                    Xmark,
                    requiresManualAttention: false,
                    yellowAuxBlock: true,
                    skipPageBandAdjustment: true);

                WriteDisciplineRow("в том числе:", null, null,
                    requiresManualAttention: false,
                    yellowAuxBlock: true,
                    skipPageBandAdjustment: true);

                foreach (var item in giaResults)
                {
                    WriteDisciplineRow(
                        item.DisciplineName,
                        FormatCreditsCell(item),
                        FormatGradeCell(item),
                        requiresManualAttention: item.RequiresManualValidation,
                        yellowAuxBlock: true,
                        columnFValueOverride: FormatSemesterColumnF(item));
                }
            }

            int volHeadingRows = MeasureNameRowLines("Объем образовательной программы");
            AdvancePastFirstSheetPrintBandIfNeeded(volHeadingRows);

            string? sumCreditsDisplay = fallbackTotalCredits > 0
                ? string.Format(CultureInfo.InvariantCulture, "{0} з.е.", fallbackTotalCredits)
                : FormatCredits(null);

            int rowVol = _currentRow;
            WriteDisciplineRow("Объем образовательной программы", sumCreditsDisplay ?? Xmark, Xmark,
                requiresManualAttention: false,
                yellowAuxBlock: false,
                skipPageBandAdjustment: true);

            if (activeWorksheet != null)
            {
                var cVol = activeWorksheet.Cells[rowVol, 3];
                cVol.Formula = null;
                cVol.Style.Numberformat.Format = "General";

                object volShown = sumCreditsDisplay ?? Xmark;
                if (data.TargetProgramCredits.HasValue)
                    volShown = string.Format(CultureInfo.InvariantCulture, "{0} з.е.", data.TargetProgramCredits.Value);
                cVol.Value = volShown;

                activeWorksheet.Cells[rowVol, 4].Value = Xmark;

                if (cVol.Comment != null)
                {
                    cVol.Comment.Text = cellComment;
                    cVol.Comment.Author = commentAuthor;
                }
                else
                    cVol.AddComment(cellComment, commentAuthor);
            }

            int contactHeadBundle = MeasureNameRowLines("в том числе объем контактной работы обучающихся")
                + MeasureNameRowLines("во взаимодействии с преподавателем в академических часах:");
            AdvancePastFirstSheetPrintBandIfNeeded(contactHeadBundle);

            WriteDisciplineRow("в том числе объем контактной работы обучающихся", null, null,
                requiresManualAttention: false,
                yellowAuxBlock: false,
                skipPageBandAdjustment: true);

            int rowAud = _currentRow;
            string? audShown = FormatAudHours(totalAudHours);
            string contactCellText = audShown ?? Xmark;
            if (data.TargetContactHoursFromPlan.HasValue)
                contactCellText = string.Format(CultureInfo.InvariantCulture, "{0} ак. час.", data.TargetContactHoursFromPlan.Value);

            WriteDisciplineRow("во взаимодействии с преподавателем в академических часах:",
                contactCellText, Xmark,
                requiresManualAttention: false,
                yellowAuxBlock: false,
                skipPageBandAdjustment: true);

            if (activeWorksheet != null)
            {
                var cAud = activeWorksheet.Cells[rowAud, 3];
                cAud.Formula = null;
                cAud.Value = contactCellText;
                cAud.Style.Numberformat.Format = "General";

                activeWorksheet.Cells[rowAud, 4].Value = Xmark;

                if (cAud.Comment != null)
                {
                    cAud.Comment.Text = cellComment;
                    cAud.Comment.Author = commentAuthor;
                }
                else
                    cAud.AddComment(cellComment, commentAuthor);
            }

            foreach (var item in courseWorks)
            {
                WriteDisciplineRow(item.DisciplineName, FormatCreditsCell(item), FormatGradeCell(item),
                    requiresManualAttention: item.RequiresManualValidation, yellowAuxBlock: true,
                    columnFValueOverride: FormatSemesterColumnF(item));
            }

            if (electives.Any())
            {
                WriteDisciplineRow("Факультативные дисциплины (модули)", null, null,
                    requiresManualAttention: false, yellowAuxBlock: true);

                WriteDisciplineRow("в том числе:", null, null,
                    requiresManualAttention: false, yellowAuxBlock: true);

                foreach (var item in electives)
                {
                    WriteDisciplineRow(item.DisciplineName, FormatCreditsCell(item), FormatGradeCell(item),
                        requiresManualAttention: item.RequiresManualValidation, yellowAuxBlock: true,
                        columnFValueOverride: FormatSemesterColumnF(item));
                }
            }

            if (manualReconciliationRows.Any())
            {
                _currentRow += 3;
                WriteManualReconciliationBanner(
                    "Неопределенные записи, требующие ручной проверки:");
                foreach (var item in manualReconciliationRows.OrderBy(x => x.DisciplineName, StringComparer.OrdinalIgnoreCase))
                {
                    WriteDisciplineRow(
                        item.DisciplineName,
                        FormatCreditsCell(item),
                        FormatGradeCell(item),
                        requiresManualAttention: true,
                        yellowAuxBlock: true,
                        columnATextOverride: "!",
                        columnFValueOverride: FormatSemesterColumnF(item));
                }
            }

            if (activeWorksheet != null)
                activeWorksheet.Cells[134, 6].Value = "последняя строка правой таблицы";

            _logger.LogInformation("Лист освоения программы: строк дисциплин с пометкой для проверки оператора: {Count}", attentionRowCount);


            activeWorksheet = null;
        }

        /// <summary>
        /// Объединяет дублирующиеся дисциплины (суммирует часы/з.е., берет оценку за последний семестр)
        /// и сортирует список по возрастанию семестра.
        /// </summary>
        private List<StudentDisciplineResult> ProcessDisciplines(List<StudentDisciplineResult> input)
        {
            if (input == null || !input.Any()) return new List<StudentDisciplineResult>();

            return input
                .GroupBy(d => d.DisciplineName?.Trim(), StringComparer.OrdinalIgnoreCase)
                .Select(g =>
                {
                    // Берем запись с максимальным семестром (или годом)
                    var lastEntry = g.OrderByDescending(x => x.Semester ?? 0)
                                     .ThenByDescending(x => x.Year ?? 0)
                                     .First();

                    SupplementPlanBucket? pickBucket = g
                        .Select(x => x.PlanBucket)
                        .FirstOrDefault(b => b.HasValue && b.Value != SupplementPlanBucket.Unknown);
                    if (!pickBucket.HasValue)
                        pickBucket = g.Select(x => x.PlanBucket).FirstOrDefault(x => x.HasValue);

                    return new StudentDisciplineResult
                    {
                        DisciplineName = g.Key,
                        PlanBucket = pickBucket ?? lastEntry.PlanBucket,
                        CreditUnits = g.Sum(x => ConvertToDouble(x.CreditUnits)),
                        AudHours = g.Sum(x => ConvertToDouble(x.AudHours)),
                        Score = lastEntry.Score,
                        ControlType = lastEntry.ControlType,
                        Semester = lastEntry.Semester,
                        Year = lastEntry.Year,
                        PlanOrder = g.Min(x => x.PlanOrder),
                        RequiresManualValidation = g.Any(x => x.RequiresManualValidation),
                        IsCardOnlyUnmatchedPlan = g.Any(x => x.IsCardOnlyUnmatchedPlan)
                    };
                })
                .OrderBy(x => x.PlanOrder ?? int.MaxValue)
                .ThenBy(x => x.Semester ?? 0)
                .ThenBy(x => x.DisciplineName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static object? FormatSemesterColumnF(StudentDisciplineResult item)
        {
            if (!item.Semester.HasValue)
                return null;
            return item.Semester.Value.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Запись строки данных (или заголовка) в таблицу.
        /// Колонка E — формула LEN(B…) как в шаблоне РИД; F может быть пустая до служебных строк.
        /// </summary>
        private void WriteDisciplineRow(
            string? name,
            string? creditsValue,
            string? gradeValue,
            bool requiresManualAttention,
            bool yellowAuxBlock,
            bool skipPageBandAdjustment = false,
            string? columnATextOverride = null,
            object? columnFValueOverride = null)
        {
            if (string.IsNullOrWhiteSpace(name)) return;

            var nameLines = SplitText(name, 75);
            int rowsNeeded = nameLines.Count;

            if (!skipPageBandAdjustment)
                AdvancePastFirstSheetPrintBandIfNeeded(rowsNeeded);

            int rowStart = _currentRow;

            for (int i = 0; i < rowsNeeded; i++)
            {
                int currentRowToWrite = _currentRow + i;

                if (activeWorksheet != null)
                {
                    if (columnATextOverride != null)
                        activeWorksheet.Cells[currentRowToWrite, 1].Value = columnATextOverride;
                    else
                        activeWorksheet.Cells[currentRowToWrite, 1].Value = currentRowToWrite - 1;
                }

                SetCellValue(currentRowToWrite, 2, nameLines[i]);

                if (activeWorksheet != null)
                {
                    var eCell = activeWorksheet.Cells[currentRowToWrite, 5];
                    eCell.Formula = $"LEN(B{currentRowToWrite})";
                    eCell.Style.Numberformat.Format = "General";
                }

                if (i == rowsNeeded - 1)
                {
                    SetCellValue(currentRowToWrite, 3, creditsValue);
                    SetCellValue(currentRowToWrite, 4, gradeValue);
                    if (columnFValueOverride != null)
                        SetCellValue(currentRowToWrite, 6, columnFValueOverride);
                }
            }

            int rowEnd = rowStart + rowsNeeded - 1;

            if ((requiresManualAttention || yellowAuxBlock) && activeWorksheet != null)
                ApplyAttentionRowFill(rowStart, rowEnd);

            _currentRow += rowsNeeded;
        }

        private void WriteManualReconciliationBanner(string headingText)
        {
            if (string.IsNullOrWhiteSpace(headingText) || activeWorksheet == null)
                return;

            var lines = SplitText(headingText.Trim(), 75);
            int rowsNeeded = lines.Count;
            AdvancePastFirstSheetPrintBandIfNeeded(rowsNeeded);
            int rowStart = _currentRow;

            for (int i = 0; i < rowsNeeded; i++)
            {
                int rw = _currentRow + i;
                activeWorksheet.Cells[rw, 1].Value = "!";

                SetCellValue(rw, 2, lines[i]);
                activeWorksheet.Cells[rw, 2].Style.Font.Bold = true;

                var eCell = activeWorksheet.Cells[rw, 5];
                eCell.Formula = $"LEN(B{rw})";
                eCell.Style.Numberformat.Format = "General";
            }

            int rowEnd = rowStart + rowsNeeded - 1;
            ApplyAttentionRowFill(rowStart, rowEnd);
            _currentRow += rowsNeeded;
        }

        private void ApplyAttentionFillToHonorStatusRow(string cellAddr)
        {
            if (activeWorksheet == null) return;
            int row = activeWorksheet.Cells[cellAddr].Start.Row;
            ApplyAttentionRowFill(row, row);
        }

        private void ApplyAttentionRowFill(int rowFrom, int rowTo)
        {
            if (activeWorksheet == null) return;
            var fillColor = Color.FromArgb(255, 248, 210);
            for (int r = rowFrom; r <= rowTo; r++)
            {
                for (int col = 2; col <= 6; col++)
                {
                    var cell = activeWorksheet.Cells[r, col];
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(fillColor);
                }
            }
        }

        

        private static bool PracticeTitleLooksCanonicalFromPlanSheet(string? disciplineName)
        {
            if (string.IsNullOrWhiteSpace(disciplineName))
                return false;
            var t = disciplineName.Trim();
            return t.StartsWith("Учебная практика,", StringComparison.OrdinalIgnoreCase)
                || t.StartsWith("Производственная практика,", StringComparison.OrdinalIgnoreCase)
                || t.StartsWith("Преддипломная практика,", StringComparison.OrdinalIgnoreCase);
        }

        private int MeasureNameRowLines(string? name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return 0;
            return SplitText(name.Trim(), 75).Count;
        }

        /// <summary>Граница печати: строка Excel 67 — последняя «верхней» области; не резать текст посередине.</summary>
        private void AdvancePastFirstSheetPrintBandIfNeeded(int rowsNeeded)
        {
            if (rowsNeeded <= 0)
                return;
            int endRow = _currentRow + rowsNeeded - 1;
            if (_currentRow <= 67 && endRow > 67)
                _currentRow = 68;
            else if (_currentRow <= 134 && endRow > 134)
                _currentRow = 135;
        }

        private List<string> SplitText(string text, int limit)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(text)) return result;

            var words = text.Split(' ');
            string currentLine = "";

            foreach (var word in words)
            {
                if (currentLine.Length + word.Length + 1 <= limit)
                {
                    currentLine += (currentLine.Length > 0 ? " " : "") + word;
                }
                else
                {
                    if (!string.IsNullOrEmpty(currentLine))
                        result.Add(currentLine);
                    currentLine = word;
                }
            }
            if (!string.IsNullOrEmpty(currentLine))
                result.Add(currentLine);

            return result;
        }

        private string? FormatCreditsCell(StudentDisciplineResult item)
        {
            return FormatCredits(item.CreditUnits);
        }

        private string? FormatGradeCell(StudentDisciplineResult item)
        {
            var baseText = GetGradeTextCore(item);
            if (!item.RequiresManualValidation)
                return baseText;
            if (string.Equals(baseText, "ТРЕБУЕТ ПРОВЕРКИ", StringComparison.Ordinal))
                return baseText;
            return (baseText ?? Xmark) + ManualReviewTail;
        }

        private static string NormalizeScoreForGradeParsing(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return "";
            var s = raw.Replace('\u00A0', ' ').Trim();
            if (s.Contains(',') && !s.Contains('.'))
                s = s.Replace(",", ".");
            return s.Trim();
        }

        private string? GetGradeTextCore(StudentDisciplineResult result)
        {
            string scoreStr = NormalizeScoreForGradeParsing(result.Score);
            string controlRaw = result.ControlType?.Replace('\u00A0', ' ')?.Trim() ?? "";
            string controlType = controlRaw.ToLower();

            bool isBarePassFailZach =
                controlType.Equals("зачёт", StringComparison.OrdinalIgnoreCase) ||
                controlType.Equals("зачет", StringComparison.OrdinalIgnoreCase);

            if (isBarePassFailZach)
            {
                if (scoreStr.Equals("зачтено", StringComparison.OrdinalIgnoreCase))
                    return "зачтено";
                if (scoreStr.Equals("не зачтено", StringComparison.OrdinalIgnoreCase))
                    return "не зачтено";
                if (double.TryParse(scoreStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double zachScore))
                {
                    if (Math.Abs(zachScore - 15d) < 0.0001d) return "зачтено";
                    if (Math.Abs(zachScore) < 0.0001d) return "не зачтено";
                    if (zachScore > 1d && zachScore <= 15d) return MapNumericScore(zachScore);
                }

                if (string.IsNullOrEmpty(scoreStr))
                    return Xmark;
                _logger.LogWarning($"Нераспознанное значение оценки (зачёт): \"{scoreStr}\". Требуется проверка оператора.");
                return "ТРЕБУЕТ ПРОВЕРКИ";
            }

            if (double.TryParse(scoreStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double scoreNum))
            {
                if (scoreNum < 0d || scoreNum > 15d)
                {
                    _logger.LogWarning($"Нераспознанное значение оценки (вне шкалы 0..15): \"{scoreStr}\". Требуется проверка оператора.");
                    return "ТРЕБУЕТ ПРОВЕРКИ";
                }
                return MapNumericScore(scoreNum);
            }

            if (!string.IsNullOrEmpty(scoreStr))
            {
                if (KnownTextGradeMap.TryGetValue(scoreStr, out var mapped)) return mapped;
                _logger.LogWarning($"Нераспознанное значение оценки: \"{scoreStr}\". Требуется проверка оператора.");
                return "ТРЕБУЕТ ПРОВЕРКИ";
            }

            return Xmark;
        }

        /// <summary>Шкала 0..15 для экзамена/диф.зачёта: &lt;7 … 13–15 отл</summary>
        private static string MapNumericScore(double score) => score switch
        {
            < 7d => "неудовлетворительно",
            < 10d => "удовлетворительно",
            < 13d => "хорошо",
            _ => "отлично",
        };

        private string? FormatCredits(object? credits)
        {
            double d = ConvertToDouble(credits);
            if (d == 0) return Xmark;
            return string.Format(CultureInfo.InvariantCulture, "{0} з.е.", d);
        }

        private string? FormatAudHours(double audHours)
        {
            if (audHours == 0) return Xmark;
            return string.Format(CultureInfo.InvariantCulture, "{0} ак. час.", audHours);
        }

        private double ConvertToDouble(object? val)
        {
            if (val == null) return 0;
            if (val is double d) return d;
            if (val is int i) return i;
            if (val is string s)
            {
                if (double.TryParse(s.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double res))
                    return res;
            }
            return 0;
        }
    }
}