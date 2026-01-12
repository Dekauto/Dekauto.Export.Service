using Dekauto.Export.Service.Domain.Interfaces;

namespace Dekauto.Export.Service.Domain.Utils
{
    public class ExportApiHelper : IExportApiHelper
    {
        public void SetHeaderFileNames(HttpResponse response, string fileName, string fileNameStar)
        {
            if (response == null)
            {
                throw new InvalidOperationException("Response не инициализирован.");
            }
            var encodedFileName = Uri.EscapeDataString(fileNameStar);
            response.Headers.Append(
                "Content-Disposition",
                $"attachment; filename=\"{fileName}.xlsx\"; filename*=UTF-8''{encodedFileName}"
            );
        }
    }
}
