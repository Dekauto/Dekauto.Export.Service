using Dekauto.Export.Service.Domain.Entities;
using Dekauto.Export.Service.Domain.Entities.DTO;
using Dekauto.Export.Service.Domain.Interfaces;
using OfficeOpenXml;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
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
            activeWorksheet.Cells[cell].AddComment(cellComment, commentAuthor);
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
                    cell.AddComment(cellComment, commentAuthor);
                }
            }
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
            _logger.LogInformation("Очистка содержимого (B-D) листа освоения программы...");

            // 1. Очистка диапазона данных (строки 2-140, только колонки B, C, D)
            // Столбец A (№ п/п) и E (скрытый) не трогаем, Шапку (строка 1) не трогаем.
            // Используем Value = null, чтобы сохранить границы и шрифт.
            sheet.Cells["B2:D140"].Value = null;

            // 2. Подготовка и нормализация данных
            var allResults = data.DisciplineResults ?? new List<StudentDisciplineResult>();

            // Списки для распределения
            var rawPractices = new List<StudentDisciplineResult>();
            var rawGia = new List<StudentDisciplineResult>();
            var rawCourseWorks = new List<StudentDisciplineResult>();
            var rawElectives = new List<StudentDisciplineResult>();
            var rawDisciplines = new List<StudentDisciplineResult>();

            foreach (var item in allResults)
            {
                string nameLower = item.DisciplineName?.ToLower()?.Trim() ?? "";
                string controlLower = item.ControlType?.ToLower()?.Trim() ?? "";

                // Логика определения типа
                // Исключение: "Проектный практикум" считается дисциплиной, а не практикой
                bool isProjectPractice = nameLower.Contains("проектный практикум");

                if (nameLower.Contains("практик") && !isProjectPractice)
                {
                    rawPractices.Add(item);
                }
                else if (nameLower.Contains("государственная") ||
                         nameLower.Contains("выпускная") ||
                         nameLower.Contains("квалификационная") ||
                         nameLower.Contains("защита вкр") ||
                         nameLower.Contains("итоговый"))
                {
                    rawGia.Add(item);
                }
                else if (nameLower.Contains("курсовая") || controlLower.Contains("курсовая"))
                {
                    rawCourseWorks.Add(item);
                }
                else if (nameLower.Contains("факультатив") || nameLower.Contains("спортивного мастерства"))
                {
                    rawElectives.Add(item);
                }
                else
                {
                    rawDisciplines.Add(item);
                }
            }

            // 3. Агрегация и сортировка
            // Дисциплины и Практики нужно схлопнуть и отсортировать
            var disciplines = ProcessDisciplines(rawDisciplines);
            var practices = ProcessDisciplines(rawPractices);
            var electives = ProcessDisciplines(rawElectives);

            // ГИА и Курсовые обычно не схлопываются по семестрам, но для порядка прогоним через сортировку
            var giaResults = ProcessDisciplines(rawGia);
            var courseWorks = ProcessDisciplines(rawCourseWorks);


            // 4. Подсчет итогов (считаем по обработанным спискам или исходным - математически сумма равна)
            // Исключаем факультативы из общего объема
            double totalCredits = disciplines.Sum(x => ConvertToDouble(x.CreditUnits))
                                + practices.Sum(x => ConvertToDouble(x.CreditUnits))
                                + giaResults.Sum(x => ConvertToDouble(x.CreditUnits));

            // Часы считаем все (включая факультативы? Обычно нет, но по ТЗ "Объем образовательной программы". 
            // Будем считать аналогично кредитам - без факультативов).
            double totalAudHours = disciplines.Sum(x => ConvertToDouble(x.AudHours))
                                 + practices.Sum(x => ConvertToDouble(x.AudHours))
                                 + giaResults.Sum(x => ConvertToDouble(x.AudHours));


            // 5. Последовательная запись блоков
            _currentRow = 2; // Данные начинаются со 2-й строки

            // Блок 1: Дисциплины
            foreach (var item in disciplines)
            {
                WriteDisciplineRow(item.DisciplineName, FormatCredits(item.CreditUnits), GetGradeText(item));
            }

            // Блок 2: Практики
            if (practices.Any())
            {
                double practiceCredits = practices.Sum(p => ConvertToDouble(p.CreditUnits));

                // Заголовок
                WriteDisciplineRow("Практики", $"{practiceCredits} з.е.", null);
                WriteDisciplineRow("в том числе:", null, null);

                foreach (var item in practices)
                {
                    WriteDisciplineRow(item.DisciplineName, FormatCredits(item.CreditUnits), GetGradeText(item));
                }
            }

            // Блок 3: ГИА
            if (giaResults.Any())
            {
                double giaCredits = giaResults.Sum(g => ConvertToDouble(g.CreditUnits));

                WriteDisciplineRow("Государственная итоговая аттестация", $"{giaCredits} з.е.", null);
                WriteDisciplineRow("в том числе:", null, null);

                foreach (var item in giaResults)
                {
                    // Для ВКР кредиты обычно не дублируются в строке названия (идут в шапке ГИА)
                    // Оценка пишется
                    WriteDisciplineRow(item.DisciplineName, null, GetGradeText(item));
                }
            }

            // Блок 4: Объем образовательной программы (Итого)
            // Строго 3 строки
            WriteDisciplineRow("Объем образовательной программы", $"{totalCredits} з.е.", null);
            WriteDisciplineRow("в том числе объем контактной работы обучающихся", null, null);
            WriteDisciplineRow("во взаимодействии с преподавателем в академических часах:", $"{totalAudHours} ак. час.", null);

            // Блок 5: Курсовые работы
            // По ТЗ: "Наименование", Оценка. З.Е. нет (null).
            foreach (var item in courseWorks)
            {
                WriteDisciplineRow(item.DisciplineName, null, GetGradeText(item));
            }

            // Блок 6: Факультативы
            if (electives.Any())
            {
                WriteDisciplineRow("Факультативные дисциплины (модули)", null, null);
                WriteDisciplineRow("в том числе:", null, null);

                foreach (var item in electives)
                {
                    WriteDisciplineRow(item.DisciplineName, FormatCredits(item.CreditUnits), GetGradeText(item));
                }
            }

            activeWorksheet = null;
        }

        /// <summary>
        /// Объединяет дублирующиеся дисциплины (суммирует часы/з.е., берет оценку за последний семестр)
        /// и сортирует список по возрастанию семестра.
        /// </summary>
        private List<StudentDisciplineResult> ProcessDisciplines(List<StudentDisciplineResult> input)
        {
            if (input == null || !input.Any()) return new List<StudentDisciplineResult>();

            var grouped = input
                .GroupBy(d => d.DisciplineName?.Trim(), StringComparer.OrdinalIgnoreCase)
                .Select(g =>
                {
                    // Находим запись с максимальным семестром (или годом) для определения итоговой оценки
                    var lastEntry = g.OrderByDescending(x => x.Semester ?? 0)
                                     .ThenByDescending(x => x.Year ?? 0)
                                     .First();

                    return new StudentDisciplineResult
                    {
                        DisciplineName = g.Key, // Используем нормализованное имя
                        // Суммируем трудоемкость
                        CreditUnits = g.Sum(x => ConvertToDouble(x.CreditUnits)),
                        AudHours = g.Sum(x => ConvertToDouble(x.AudHours)),
                        // Данные оценки берем из последнего периода
                        Score = lastEntry.Score,
                        ControlType = lastEntry.ControlType,
                        Semester = lastEntry.Semester,
                        Year = lastEntry.Year
                    };
                })
                .OrderBy(x => x.Semester ?? 0) // Сортировка по возрастанию семестра
                .ThenBy(x => x.DisciplineName) // Вторичная сортировка по алфавиту
                .ToList();

            return grouped;
        }

        /// <summary>
        /// Запись строки данных (или заголовка) в таблицу
        /// </summary>
        private void WriteDisciplineRow(string? name, string? creditsValue, string? gradeValue)
        {
            if (string.IsNullOrWhiteSpace(name)) return;

            // 1. Разбиваем название на строки по 75 символов
            var nameLines = SplitText(name, 75);
            int rowsNeeded = nameLines.Count;

            // 2. Проверка пагинации
            int endRow = _currentRow + rowsNeeded - 1;

            // 67 - последняя строка 1-го листа
            // 134 - последняя строка 2-го листа
            if (_currentRow <= 67 && endRow > 67)
            {
                _currentRow = 68;
            }
            else if (_currentRow <= 134 && endRow > 134)
            {
                _currentRow = 135;
            }

            // 3. Запись данных
            for (int i = 0; i < rowsNeeded; i++)
            {
                int currentRowToWrite = _currentRow + i;

                // Столбец B (2): Название
                // Используем SetCellValue без авто-комментариев для массовой вставки
                SetCellValue(currentRowToWrite, 2, nameLines[i]);

                // Столбцы C (3) и D (4): Кредиты и Оценка
                // Пишутся ТОЛЬКО в последней строке блока названия
                if (i == rowsNeeded - 1)
                {
                    SetCellValue(currentRowToWrite, 3, creditsValue);
                    SetCellValue(currentRowToWrite, 4, gradeValue);
                }
            }

            _currentRow += rowsNeeded;
        }

        // --- ОБНОВЛЕННЫЕ ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ---

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

        private string? GetGradeText(StudentDisciplineResult result)
        {
            string scoreStr = result.Score?.ToString()?.Trim() ?? "";
            string controlType = result.ControlType?.ToLower()?.Trim() ?? "";

            // 1. Зачет
            if (controlType == "зачёт" || controlType == "зачет")
            {
                if (scoreStr == "15" || scoreStr.Equals("зачтено", StringComparison.OrdinalIgnoreCase))
                    return "зачтено";
                if (scoreStr == "0" || scoreStr.Equals("не зачтено", StringComparison.OrdinalIgnoreCase))
                    return "не зачтено";

                // Для всех прочих случаев зачета (пусто или странные цифры) по ТЗ "зачтено" 
                // если это успешно пройденная дисциплина, но тут безопаснее вернуть null или что есть
                return string.IsNullOrEmpty(scoreStr) ? null : "зачтено";
            }

            // 2. Оценка (Экзамен, Диф.зачет, Курсовая)
            if (double.TryParse(scoreStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double scoreNum))
            {
                if (scoreNum < 7) return "неудовлетворительно";
                if (scoreNum >= 7 && scoreNum <= 9) return "удовлетворительно";
                if (scoreNum >= 10 && scoreNum <= 12) return "хорошо";
                if (scoreNum >= 13) return "отлично";
            }

            // Если уже текст
            if (!string.IsNullOrEmpty(scoreStr)) return scoreStr;

            return null; // Пусто, никаких "х"
        }

        private string? FormatCredits(object? credits)
        {
            double d = ConvertToDouble(credits);
            if (d == 0) return null; // 0 не отображаем
            return $"{d} з.е.";
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