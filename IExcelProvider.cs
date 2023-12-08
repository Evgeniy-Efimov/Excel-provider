using System.Collections.Generic;
using System.IO;

namespace ExcelProvider
{
    public interface IExcelProvider
    {
        IEnumerable<RowImportResult<TModel>> ReadFile<TModel>(Stream file, string worksheetName = "") where TModel : new();
        MemoryStream ExportToFile<TModel>(IEnumerable<TModel> rowsData, string worksheetName) where TModel : new();
    }

    public class RowImportResult<TModel>
    {
        public TModel Model { get; set; }
        public int RowNumber { get; set; }
        public bool IsSuccessfullyProcessed { get; set; }
        public string Message { get; set; }
    }
}
