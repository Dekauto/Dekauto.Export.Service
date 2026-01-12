using Dekauto.Export.Service.Domain.Entities;
using Dekauto.Export.Service.Domain.Interfaces;
using OfficeOpenXml;
using System.IO.Compression;

namespace Dekauto.Export.Service.Domain.Services
{
    public class StudentsCardService : IStudentsCardService
    {
        private IConfiguration _configuration;
        private string exportCardName;
        public StudentsCardService(IConfiguration configuration)
        {
            _configuration = configuration;
            exportCardName = _configuration.GetValue<string>("ExportCardName") ?? throw new ArgumentNullException(nameof(exportCardName));

        }
        public async Task<MemoryStream> ConvertStudentsToExcel(List<Student> students)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            if (students == null || students.Count == 0) throw new ArgumentNullException(nameof(students));


            var stream = new MemoryStream(); //Используем временное хранилище

            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
            {
                foreach (var student in students)
                {
                    var templatePath = Path.Combine(Directory.GetCurrentDirectory(), exportCardName); //Путь шаблона 

                    if (!File.Exists(templatePath))
                    {
                        throw new FileNotFoundException("Файл шаблона не найден. Обратитесь к администратору");
                    }

                    using (var package = new ExcelPackage(new FileInfo(templatePath)))
                    {
                        FillExcel(student, package);

                        var entry = archive.CreateEntry($"{student.Surname} {student.Name} {student.Patronymic}.xlsx");
                        using (var entryStream = entry.Open())
                        {
                            await package.SaveAsAsync(entryStream); //Сохраняем файл
                        }

                    }

                }
            }
            stream.Position = 0;
            return stream;
        }

        public async Task<MemoryStream> ConvertStudentToExcel(Student student)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            if (student == null) throw new ArgumentNullException(nameof(student));

            var templatePath = Path.Combine(Directory.GetCurrentDirectory(), exportCardName); //Путь шаблона

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException("Файл шаблона не найден. Обратитесь к администратору");
            }

            var stream = new MemoryStream(); //Используем временное хранилище

            using (var package = new ExcelPackage(new FileInfo(templatePath)))
            {
                FillExcel(student, package);
                await package.SaveAsAsync(stream); //Сохраняем файл
            }
            stream.Position = 0; //Сбрасываем позицию
            return stream;
        }
        public void FillExcel(Student student, ExcelPackage package)
        {
            if (package.Workbook.Worksheets.Count == 0)
            {
                throw new InvalidOperationException("Файл шаблона не содержит листов");
            }

            var worksheet = package.Workbook.Worksheets[0]; //Выбираем первый лист

            //Персональные данные
            worksheet.Cells["B4"].Value = student.Name;
            worksheet.Cells["B3"].Value = student.Surname;
            worksheet.Cells["B5"].Value = student.Patronymic;
            worksheet.Cells["G3"].Value = student.GradeBook;
            worksheet.Cells["G4"].Value = student.GradeBook;
            if (student.Gender == true) worksheet.Cells["B6"].Value = "М";
            else worksheet.Cells["B6"].Value = "Ж";
            worksheet.Cells["E6"].Value = student.BirthdayDate;
            worksheet.Cells["A7"].Value = student.BirthdayPlace;

            //Контактные данные
            worksheet.Cells["C9"].Value = student.PhoneNumber;
            worksheet.Cells["G9"].Value = student.Email;

            //Паспортные данные
            worksheet.Cells["D10"].Value = student.PassportSerial;
            worksheet.Cells["F10"].Value = student.PassportNumber;
            worksheet.Cells["A11"].Value = student.PassportIssuancePlace;
            worksheet.Cells["B12"].Value = student.PassportIssuanceDate;
            worksheet.Cells["F12"].Value = student.PassportIssuanceCode;

            //Адрес регистрации
            worksheet.Cells["C13"].Value = student.Citizenship;
            worksheet.Cells["B15"].Value = student.AddressRegistrationIndex;
            worksheet.Cells["G15"].Value = student.AddressRegistrationOblKrayAvtobl;
            worksheet.Cells["B16"].Value = student.AddressRegistrationDistrict;
            if ((student.AddressRegistrationType != "") && (student.AddressRegistrationType != null)) worksheet.Cells["E16"].Value = $"{student.AddressRegistrationType}:";
            worksheet.Cells["G16"].Value = student.AddressRegistrationCity;
            worksheet.Cells["B17"].Value = student.AddressRegistrationStreet;
            worksheet.Cells["F17"].Value = student.AddressRegistrationHouse;
            if ((student.AddressRegistrationHousingType != "") && (student.AddressRegistrationHousingType != null)) worksheet.Cells["G17"].Value = $"{student.AddressRegistrationHousingType}:";
            worksheet.Cells["H17"].Value = student.AddressRegistrationHousing;
            worksheet.Cells["J17"].Value = student.AddressRegistrationApartment;

            //Адрес проживания
            worksheet.Cells["B19"].Value = student.AddressResidentialIndex;
            worksheet.Cells["G19"].Value = student.AddressResidentialOblKrayAvtobl;
            worksheet.Cells["B20"].Value = student.AddressResidentialDistrict;
            if ((student.AddressResidentialType != "") && (student.AddressResidentialType != null)) worksheet.Cells["E20"].Value = $"{student.AddressResidentialType}:";
            worksheet.Cells["G20"].Value = student.AddressResidentialCity;
            worksheet.Cells["B21"].Value = student.AddressResidentialStreet;
            worksheet.Cells["F21"].Value = student.AddressResidentialHouse;
            if ((student.AddressResidentialHousingType != "") && (student.AddressResidentialHousingType != null)) worksheet.Cells["G21"].Value = $"{student.AddressResidentialHousingType}:";
            worksheet.Cells["H21"].Value = student.AddressResidentialHousing;
            worksheet.Cells["J21"].Value = student.AddressResidentialApartment;
            if (student.LivingInDormitory == true) worksheet.Cells["D22"].Value = "да";
            else worksheet.Cells["D22"].Value = "нет";

            //Основания зачисления в МГППУ
            worksheet.Cells["A34"].Value = student.EnrollementOrderDate;
            worksheet.Cells["C34"].Value = student.EnrollementOrderNum;

            worksheet.Cells["B37"].Value = student.GiaExam1Name;
            worksheet.Cells["E37"].Value = student.GiaExam1Score;
            if (student.GiaExam1Note != "") worksheet.Cells["F37"].Value = student.GiaExam1Note ?? "вступительные испытания в форме ЕГЭ";
            worksheet.Cells["B38"].Value = student.GiaExam2Name;
            worksheet.Cells["E38"].Value = student.GiaExam2Score;
            if (student.GiaExam2Note != "") worksheet.Cells["F38"].Value = student.GiaExam2Note ?? "вступительные испытания в форме ЕГЭ";
            worksheet.Cells["B39"].Value = student.GiaExam3Name;
            worksheet.Cells["E39"].Value = student.GiaExam3Score;
            if (student.GiaExam3Note != "") worksheet.Cells["F39"].Value = student.GiaExam3Note ?? "вступительные испытания в форме ЕГЭ";

            if (student.EducationReceived == "Среднее общее") worksheet.Cells["D40"].Value = "среднее общее образование";
            else if (student.EducationReceived == "Высшее образование") worksheet.Cells["D40"].Value = "высшее образование";
            else worksheet.Cells["D40"].Value = "среднее профессиональное образование";

            worksheet.Cells["B42"].Value = student.EducationReceivedNum;
            worksheet.Cells["D42"].Value = student.EducationReceivedSerial;
            worksheet.Cells["H42"].Value = student.EducationReceivedDate;
            worksheet.Cells["C44"].Value = student.OOName;
            worksheet.Cells["C45"].Value = student.OOAddress;
            worksheet.Cells["C46"].Value = student.EducationReceivedEndYear;
            if (student.BonusScores != 0) worksheet.Cells["C47"].Value = "да";
            else worksheet.Cells["C47"].Value = "нет";

            //Доп данные
            if (student.MaritalStatus == true) worksheet.Cells["D56"].Value = "да";
            else worksheet.Cells["D56"].Value = "нет";
            if (student.MilitaryService == true) worksheet.Cells["J56"].Value = "да";
            else worksheet.Cells["J56"].Value = "нет";

            //Текущие координаты обучения и статус			
            worksheet.Cells["C76"].Value = student.Education;
            worksheet.Cells["H76"].Value = student.EducationForm;
            worksheet.Cells["B77"].Value = student.Faculty;
            worksheet.Cells["C78"].Value = student.CourseOfTraining;
            worksheet.Cells["C79"].Value = student.Course;
            worksheet.Cells["C80"].Value = "адаптированная для лиц с ОВЗ"; // Данные статичны
            worksheet.Cells["B81"].Value = student.GroupName;
            worksheet.Cells["F81"].Value = student.EducationStartYear;
            worksheet.Cells["J81"].Value = student.EducationFinishYear;
            worksheet.Cells["I82"].Value = student.EducationTime;
            worksheet.Cells["C84"].Value = student.EducationBase;
            worksheet.Cells["C85"].Value = student.EducationRelationForm;
            worksheet.Cells["G85"].Value = student.EducationRelationNum;
            worksheet.Cells["I85"].Value = student.EducationRelationDate;

            //Заполнение данных дисциплин
            FillDisciplineData(student, package);

        }

        private void FillDisciplineData(Student student, ExcelPackage package)
        {
            if (student.DisciplineResults == null || student.DisciplineResults.Count == 0)
                return;

            if (!student.EducationStartYear.HasValue)
                return;

            // Листы курсов: Ро_1 курс, Ро_2 курс, Ро_3 курс, Ро_4 курс
            var courseSheetNames = new[] { "Ро_1 курс", "Ро_2 курс", "Ро_3 курс", "Ро_4 курс" };

            for (int courseIndex = 0; courseIndex < courseSheetNames.Length; courseIndex++)
            {
                var sheetName = courseSheetNames[courseIndex];

                // Сначала пытаемся найти лист по точному имени
                var worksheet = package.Workbook.Worksheets[sheetName];

                // Если не нашли, пытаемся найти по частичному совпадению
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.FirstOrDefault(ws =>
                        ws.Name.Contains(sheetName, StringComparison.OrdinalIgnoreCase) ||
                        sheetName.Contains(ws.Name, StringComparison.OrdinalIgnoreCase));
                }

                if (worksheet == null)
                    continue;

                var courseNumber = courseIndex + 1;

                // Вычисляем номер курса на основе Year дисциплины, Semester и EducationStartYear
                // Учебный год начинается осенью:
                // - Нечетный семестр (1, 3, 5...) в году Y → учебный год Y-Y+1
                // - Четный семестр (2, 4, 6...) в году Y → учебный год Y-1-Y
                // Курс определяется по году начала учебного года
                var disciplinesForCourse = student.DisciplineResults
                    .Where(d =>
                    {
                        if (!d.Year.HasValue || !d.Semester.HasValue)
                            return false;

                        short disciplineYear = d.Year.Value;
                        short semester = d.Semester.Value;
                        short startYear = student.EducationStartYear.Value;

                        // Определяем год начала учебного года для этой дисциплины
                        short academicYearStart;
                        if (semester % 2 == 1) // Нечетный семестр (осень)
                        {
                            // Учебный год начинается в этом же году
                            academicYearStart = disciplineYear;
                        }
                        else // Четный семестр (весна)
                        {
                            // Учебный год начался в предыдущем году
                            academicYearStart = (short)(disciplineYear - 1);
                        }

                        // Вычисляем курс для дисциплины
                        short disciplineCourse;
                        if (academicYearStart < startYear)
                        {
                            // Если год начала учебного года меньше года начала обучения, относим к 1 курсу
                            disciplineCourse = 1;
                        }
                        else
                        {
                            // Курс = разница в годах начала учебного года + 1
                            disciplineCourse = (short)(academicYearStart - startYear + 1);
                        }

                        return disciplineCourse == courseNumber;
                    })
                    .ToList();

                if (disciplinesForCourse.Count == 0)
                    continue;

                // Группируем по семестрам
                var oddSemesterDisciplines = disciplinesForCourse
                    .Where(d => d.Semester.HasValue && d.Semester.Value % 2 == 1)
                    .OrderBy(d => d.Semester)
                    .ToList();

                var evenSemesterDisciplines = disciplinesForCourse
                    .Where(d => d.Semester.HasValue && d.Semester.Value % 2 == 0)
                    .OrderBy(d => d.Semester)
                    .ToList();

                // Заполняем нечетные семестры (строки 9-21)
                FillSemesterDisciplines(worksheet, oddSemesterDisciplines, startRow: 9, endRow: 21);

                // Заполняем четные семестры (строки 39-51)
                FillSemesterDisciplines(worksheet, evenSemesterDisciplines, startRow: 39, endRow: 51);

                // Заполняем курсовые работы
                FillCourseWork(worksheet, oddSemesterDisciplines, nameRow: 25, scoreRow: 24);
                FillCourseWork(worksheet, evenSemesterDisciplines, nameRow: 55, scoreRow: 54);

                // Заполняем даты семестров
                FillSemesterDates(worksheet, oddSemesterDisciplines, evenSemesterDisciplines);
            }
        }

        private void FillSemesterDisciplines(OfficeOpenXml.ExcelWorksheet worksheet, List<StudentDisciplineResult> disciplines, int startRow, int endRow)
        {
            int currentRow = startRow;
            var regularDisciplines = disciplines
                .Where(d => d.ControlType?.ToLower().Trim() != "курсовая")
                .ToList();

            foreach (var discipline in regularDisciplines.Take(endRow - startRow + 1))
            {
                // Столбец 2: название дисциплины
                worksheet.Cells[currentRow, 2].Value = discipline.DisciplineName;

                // Столбец 3: зачетные единицы
                if (discipline.CreditUnits.HasValue)
                    worksheet.Cells[currentRow, 3].Value = discipline.CreditUnits.Value;

                // Столбец 5: аудиторные часы
                if (discipline.AudHours.HasValue)
                    worksheet.Cells[currentRow, 5].Value = discipline.AudHours.Value;

                // Столбец 7: форма аттестации (практика заменяется на "зачёт с оценкой")
                var controlTypeForDisplay = discipline.ControlType?.ToLower().Trim() == "практика"
                    ? "зачёт с оценкой"
                    : discipline.ControlType;
                worksheet.Cells[currentRow, 7].Value = controlTypeForDisplay;

                // Столбец 8: оценка
                if (discipline.Score.HasValue)
                    worksheet.Cells[currentRow, 8].Value = discipline.Score.Value;

                // Столбец 9: интерпретация
                worksheet.Cells[currentRow, 9].Value = GetInterpretation(discipline.Score, discipline.ControlType);

                // Столбец 10: документ ("В" если есть данные)
                if (HasDisciplineData(discipline))
                    worksheet.Cells[currentRow, 10].Value = "В";

                currentRow++;
            }
        }

        private string GetInterpretation(double? score, string? controlType)
        {
            if (!score.HasValue || string.IsNullOrEmpty(controlType))
                return string.Empty;

            var scoreValue = score.Value;
            var controlTypeLower = controlType.ToLower().Trim();

            // Формы зачета/с оценкой (учитываем варианты написания: зачет, зачёт, зачет с оценкой, зачёт с оценкой)
            // Практика обрабатывается как зачет с оценкой
            if (controlTypeLower == "зачет" || controlTypeLower == "зачёт" ||
                controlTypeLower == "зачет с оценкой" || controlTypeLower == "зачёт с оценкой" ||
                controlTypeLower == "практика")
            {
                if (scoreValue >= 13 && scoreValue <= 15)
                    return "зачтено (5, отлично)";
                else if (scoreValue >= 10 && scoreValue <= 12)
                    return "зачтено (4, хорошо)";
                else if (scoreValue >= 7 && scoreValue <= 9)
                    return "зачтено (3, удовлетворительно)";
                else if (scoreValue < 7)
                    return "не зачтено (2, не удовлетворительно)";
            }
            // Формы не зачета (экзамен, контрольная)
            else if (controlTypeLower == "экзамен" || controlTypeLower == "контрольная" || controlTypeLower == "контрольная работа")
            {
                if (scoreValue >= 13 && scoreValue <= 15)
                    return "5, отлично";
                else if (scoreValue >= 10 && scoreValue <= 12)
                    return "4, хорошо";
                else if (scoreValue >= 7 && scoreValue <= 9)
                    return "3, удовлетворительно";
                else if (scoreValue < 7)
                    return "2, не удовлетворительно";
            }

            return string.Empty;
        }

        private bool HasDisciplineData(StudentDisciplineResult discipline)
        {
            return !string.IsNullOrEmpty(discipline.DisciplineName) ||
                   discipline.Score.HasValue ||
                   discipline.CreditUnits.HasValue ||
                   !string.IsNullOrEmpty(discipline.ControlType);
        }

        private void FillCourseWork(OfficeOpenXml.ExcelWorksheet worksheet, List<StudentDisciplineResult> disciplines, int nameRow, int scoreRow)
        {
            var courseWork = disciplines
                .FirstOrDefault(d => d.ControlType?.ToLower().Trim() == "курсовая");

            if (courseWork != null && !string.IsNullOrEmpty(courseWork.DisciplineName))
            {
                // Название курсовой работы в объединенных ячейках столбцов 3-12
                var cellRange = worksheet.Cells[nameRow, 3, nameRow, 12];
                if (!cellRange.Merge)
                {
                    cellRange.Merge = true;
                }
                worksheet.Cells[nameRow, 3].Value = courseWork.DisciplineName;

                // Оценка в строке scoreRow на столбце 8
                if (courseWork.Score.HasValue)
                {
                    worksheet.Cells[scoreRow, 8].Value = courseWork.Score.Value;
                }

                // Интерпретация в строке scoreRow на столбце 9 (как для не-зачета: экзамен, контрольная)
                var interpretation = GetInterpretationForNonCredit(courseWork.Score);
                if (!string.IsNullOrEmpty(interpretation))
                {
                    worksheet.Cells[scoreRow, 9].Value = interpretation;
                }
            }
        }

        private string GetInterpretationForNonCredit(double? score)
        {
            if (!score.HasValue)
                return string.Empty;

            var scoreValue = score.Value;

            // Интерпретация как для экзамена/контрольной (не-зачета)
            if (scoreValue >= 13 && scoreValue <= 15)
                return "5, отлично";
            else if (scoreValue >= 10 && scoreValue <= 12)
                return "4, хорошо";
            else if (scoreValue >= 7 && scoreValue <= 9)
                return "3, удовлетворительно";
            else if (scoreValue < 7)
                return "2, не удовлетворительно";

            return string.Empty;
        }

        private void FillSemesterDates(OfficeOpenXml.ExcelWorksheet worksheet, List<StudentDisciplineResult> oddSemesterDisciplines, List<StudentDisciplineResult> evenSemesterDisciplines)
        {
            // Формат даты: "YYYY-YYYY" где первая часть - год начала учебного года (осень), вторая - конец (весна)
            short? oddYearStart = null;
            short? evenYearStart = null;

            // Определяем год начала учебного года для нечетного семестра
            var oddDiscipline = oddSemesterDisciplines.FirstOrDefault();
            if (oddDiscipline != null && oddDiscipline.Year.HasValue && oddDiscipline.Semester.HasValue)
            {
                // Нечетный семестр (осень) - учебный год начинается в этом же году
                oddYearStart = oddDiscipline.Year.Value;
            }

            // Определяем год начала учебного года для четного семестра
            var evenDiscipline = evenSemesterDisciplines.FirstOrDefault();
            if (evenDiscipline != null && evenDiscipline.Year.HasValue && evenDiscipline.Semester.HasValue)
            {
                // Четный семестр (весна) - учебный год начался в предыдущем году
                evenYearStart = (short)(evenDiscipline.Year.Value - 1);
            }

            // Формируем дату для нечетного семестра (строка 3)
            if (oddYearStart.HasValue)
            {
                string oddDateFormat = $"{oddYearStart.Value}-{oddYearStart.Value + 1}";
                var oddDateRange = worksheet.Cells[3, 8, 3, 9];
                if (!oddDateRange.Merge)
                {
                    oddDateRange.Merge = true;
                }
                worksheet.Cells[3, 8].Value = oddDateFormat;
            }

            // Формируем дату для четного семестра (строка 33)
            if (evenYearStart.HasValue)
            {
                string evenDateFormat = $"{evenYearStart.Value}-{evenYearStart.Value + 1}";
                var evenDateRange = worksheet.Cells[33, 8, 33, 9];
                if (!evenDateRange.Merge)
                {
                    evenDateRange.Merge = true;
                }
                worksheet.Cells[33, 8].Value = evenDateFormat;
            }
        }
    }
}
