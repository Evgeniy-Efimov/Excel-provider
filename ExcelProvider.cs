using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;
using ExcelProvider.Helpers;

namespace ExcelProvider
{
    public class ExcelProvider : IExcelProvider
    {
        public IEnumerable<RowImportResult<TModel>> ReadFile<TModel>(Stream file, string worksheetName = "") where TModel : new()
        {
            using (var workbook = new ExcelWorkbook(file))
            {
                var worksheet = workbook.GetWorksheet(worksheetName);

                foreach (var row in worksheet.DataRows)
                {
                    var readResult = new RowImportResult<TModel>() { RowNumber = row.RowNum + 1 };

                    try
                    {
                        readResult.Model = worksheet.ReadRow<TModel>(row);
                        readResult.IsSuccessfullyProcessed = true;
                    }
                    catch (Exception ex)
                    {
                        readResult.Message = ex.Message;
                    }

                    yield return readResult;
                }
            }
        }

        public MemoryStream ExportToFile<TModel>(IEnumerable<TModel> rowsData, string worksheetName) where TModel : new()
        {
            var workbook = new ExcelWorkbook();
            workbook.CreateWorksheet(rowsData, worksheetName);

            return workbook.WriteToStream();
        }
    }

    public class ExcelWorkbook : IDisposable
    {
        private IWorkbook Workbook { get; set; }

        public ExcelWorkbook(Stream file)
        {
            Workbook = WorkbookFactory.Create(file);

            if (Workbook is XSSFWorkbook)
            {
                XSSFFormulaEvaluator.EvaluateAllFormulaCells(Workbook);
            }
            else
            {
                HSSFFormulaEvaluator.EvaluateAllFormulaCells(Workbook);
            }
        }

        public ExcelWorkbook()
        {
            Workbook = new XSSFWorkbook();
        }

        public ExcelWorksheet GetWorksheet(string worksheetName)
        {
            var worksheet = string.IsNullOrWhiteSpace(worksheetName) || Workbook.GetSheetIndex(worksheetName) < 0
                ? Workbook.GetSheetAt(0) : Workbook.GetSheet(worksheetName);

            return new ExcelWorksheet(worksheet);
        }

        public ExcelWorksheet CreateWorksheet<TModel>(IEnumerable<TModel> rowsData, string worksheetName)
        {
            var worksheet = Workbook.CreateSheet(worksheetName);
            var rowIndex = 0;
            var cellIndex = 0;
            var columns = typeof(TModel).GetPropertiesNames();
            var headerRow = worksheet.CreateRow(rowIndex);

            foreach (var columnName in columns)
            {
                var cell = headerRow.CreateCell(cellIndex);
                cell.SetCellValue(columnName);
                cellIndex++;
            }

            foreach (var rowData in rowsData)
            {
                rowIndex++;
                cellIndex = 0;
                var row = worksheet.CreateRow(rowIndex);

                foreach (var columnName in columns)
                {
                    var cell = row.CreateCell(cellIndex);
                    cell.SetCellValue(rowData.GetPropertyValue(columnName)?.ToString() ?? string.Empty);
                    cellIndex++;
                }
            }

            return new ExcelWorksheet(worksheet);
        }

        public MemoryStream WriteToStream()
        {
            var stream = new MemoryStream();
            Workbook.Write(stream);
            stream.Close();
            return stream;
        }

        public void Dispose()
        {
            Workbook.Dispose();
        }
    }

    public class ExcelWorksheet
    {
        private ISheet Worksheet { get; set; }

        public ExcelWorksheet(ISheet worksheet)
        {
            Worksheet = worksheet;
        }

        public IRow HeaderRow
        {
            get
            {
                IRow row = null;

                for (int rowNumber = Worksheet.FirstRowNum; rowNumber <= Worksheet.LastRowNum; rowNumber++)
                {
                    row = Worksheet.GetRow(rowNumber);

                    if (row?.Cells != null && row.Cells.Any(c => !string.IsNullOrEmpty(GetCellValue(c))))
                    {
                        return row;
                    }
                }

                return row;
            }
        }

        public IEnumerable<string> Columns
        {
            get
            {
                return HeaderRow.Cells.Select(c => GetCellValue(c).NormalizeForComparing());
            }
        }

        public IEnumerable<IRow> DataRows
        {
            get
            {
                var startRowNumber = (HeaderRow?.RowNum ?? Worksheet.FirstRowNum) + 1;
                var lastRowNumber = Worksheet.LastRowNum > startRowNumber ? Worksheet.LastRowNum : startRowNumber;

                for (int rowNumber = startRowNumber; rowNumber <= lastRowNumber; rowNumber++)
                {
                    var row = Worksheet.GetRow(rowNumber);

                    if (row != null)
                    {
                        yield return row;
                    }
                }
            }
        }

        public TModel ReadRow<TModel>(IRow row) where TModel : new()
        {
            var model = new TModel();
            var columns = Columns.ToList();
            var cells = row.Cells.ToList();
            var propertyColumnName = string.Empty;
            var cellIndex = -1;
            string value = null;

            foreach (var property in typeof(TModel).GetProperties())
            {
                propertyColumnName = property.GetPropertyName();

                if (!string.IsNullOrWhiteSpace(propertyColumnName))
                {
                    cellIndex = columns.IndexOf(propertyColumnName.NormalizeForComparing());

                    if (cellIndex >= 0)
                    {
                        value = GetCellValue(cells.ElementAtOrDefault(cellIndex));

                        if (value != null)
                        {
                            property.SetPropertyValue(model, value);
                        }
                    }
                }
            }

            model.ValidateModel();

            return model;
        }

        public void FormatCells(ICellStyle cellStyle, IEnumerable<ICell> cells)
        {
            foreach (var cell in cells)
            {
                cell.CellStyle = cellStyle;
            }
        }

        public void AutoSizeColumns()
        {
            for (int columnNumber = 0; columnNumber < Columns.Count(); columnNumber++)
            {
                Worksheet.AutoSizeColumn(columnNumber);
            }
        }

        private string GetCellValue(ICell cell)
        {
            switch (cell?.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    return DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue.ToLongDateString() : cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    cell.SetCellType(CellType.String);
                    return cell.StringCellValue;
                default: return null;
            }
        }
    }
}
