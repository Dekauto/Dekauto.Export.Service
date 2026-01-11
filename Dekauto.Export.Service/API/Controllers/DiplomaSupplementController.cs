using Dekauto.Export.Service.Domain.Interfaces;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Dekauto.Export.Service.API.Controllers
{
    [Route("api/diploma/supplement")]
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

    }
}
