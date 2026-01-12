namespace Dekauto.Export.Service.Domain.Entities.DTO
{
    public class DiplomaSupplementExportRequest
    {
        public DiplomaSupplementData supplementData { get; set; } // Данные для парсинга
        public string manufacturer { get; set; } // выбранный производитель шаблона
        public string educationLevel { get; set; } // выбранный уровень обучения
    }
}
