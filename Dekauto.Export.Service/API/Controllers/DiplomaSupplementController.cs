using Dekauto.Export.Service.Domain.Entities.DTO;
using Dekauto.Export.Service.Domain.Interfaces;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Dekauto.Export.Service.API.Controllers
{
    [Route("api/diploma-supplement")]
    [ApiController]
    [Authorize]
    public class DiplomaSupplementController : Controller
    {
        private readonly IDiplomaSupplementExportService exportService;
        private readonly IExportApiHelper apiHelper;
        private string defaultLatFileName = "exported_diploma_supplement";
        private readonly ILogger<DiplomaSupplementController> logger;

        public DiplomaSupplementController(IDiplomaSupplementExportService exportService,
            ILogger<DiplomaSupplementController> logger,
            IExportApiHelper apiHelper)
        {
            this.exportService = exportService;
            this.logger = logger;
            this.apiHelper = apiHelper;
        }

        [HttpPost]
        public async Task<IActionResult> ExportStudentsAsync([FromBody] DiplomaSupplementExportRequest exportRequest)
        {
            try
            {
                logger.LogInformation($"Начало экспорта приложения диплома...");
                var (stream, fileName) = await exportService.ExportDiplomaSupplement(exportRequest);
                // INFO: данные в имени файла не должны содержать спецсимволы!
                apiHelper.SetHeaderFileNames(Response, defaultLatFileName, fileName);

                // Возвращаем файл БЕЗ указания имени в третьем параметре
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Ошибка при экспорте приложения диплома: {ex.Message}");
                return HandleException(ex);
            }
        }


        private IActionResult HandleException(Exception ex)
        {
            switch (ex)
            {
                case ArgumentNullException argumentNullException:
                    return BadRequest(argumentNullException.Message);
                case FileNotFoundException fileNotFoundException:
                    return NotFound($"{fileNotFoundException.Message} {fileNotFoundException.FileName}");
                case InvalidOperationException invalidOperationException:
                    return BadRequest(invalidOperationException.Message);
                default:
                    return StatusCode(500, "Неизвестная ошибка, обратитесь к администратору");
            }
        }
    }
}
