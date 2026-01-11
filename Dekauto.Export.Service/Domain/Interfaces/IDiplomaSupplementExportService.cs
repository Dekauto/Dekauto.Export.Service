using Dekauto.Export.Service.Domain.Entities.DTO;

namespace Dekauto.Export.Service.Domain.Interfaces
{
    public interface IDiplomaSupplementExportService
    {
        public Task<(MemoryStream, string)> ExportDiplomaSupplement(DiplomaSupplementExportRequest request);
    }
}
