namespace Dekauto.Export.Service.Domain.Interfaces
{
    public interface IExportApiHelper
    {
        /// <summary>
        /// Формируем http-заголовок с поддержкой UTF-8 (для поддержки кириллицы в http-заголовках), потому что
        /// без этого передается только сам файл, а его название автомат. вписывается в заголовки, но без поддержки кириллицы.
        /// </summary>
        /// <param name="fileName">Только латиница, для правильной передачи файла в запросе.</param>
        /// <param name="fileNameStar">Отображаемое имя файла, которое будет видно пользователю.</param>
        void SetHeaderFileNames(HttpResponse response, string fileName, string fileNameStar);
    }
}
