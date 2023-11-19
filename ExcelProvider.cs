using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelProvider
{
    public class ExcelProvider : IExcelProvider
    {
        private readonly List<Type> NumericTypes = new List<Type>
        {
            typeof(byte),
            typeof(short),
            typeof(int),
            typeof(decimal),
            typeof(float),
            typeof(double)
        };

        private bool IsNumericType(Type type)
        {
            return NumericTypes.Contains(type) || NumericTypes.Contains(Nullable.GetUnderlyingType(type));
        }

        public IEnumerable<RowImportResult<TModel>> ReadFile<TModel>(Stream file, string worksheetName = "") where TModel : new()
        {
            using (var workbook = WorkbookFactory.Create(file))
            {
                var worksheet = GetWorksheet(workbook, worksheetName);
                var columns = GetColumns(worksheet).ToList();

                foreach (var row in GetRows(worksheet))
                {
                    var readResult = new RowImportResult<TModel>() { RowNumber = row.RowNum + 1 };

                    try
                    {
                        readResult.Model = GetModel<TModel>(GetCells(row).ToList(), columns);
                        ValidateModel(readResult.Model);
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

        private void ValidateModel<TModel>(TModel model)
        {
            var validationContext = new ValidationContext(model, serviceProvider: null, items: null);
            var validationResults = new List<ValidationResult>();
            var isValid = Validator.TryValidateObject(model, validationContext, validationResults, true);

            if (!isValid)
            {
                throw new ArgumentException("Model is not valid: " + string.Join(", ", validationResults.Select(s => s.ErrorMessage).ToArray()));
            }
        }

        private TModel GetModel<TModel>(List<ICell> cells, List<string> columns) where TModel : new()
        {
            var model = new TModel();
            var propertyColumnName = string.Empty;
            var cellIndex = -1;
            string value = null;

            foreach (var property in typeof(TModel).GetProperties())
            {
                propertyColumnName = (property.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() as DisplayNameAttribute)?.DisplayName ?? property.Name;

                if (!string.IsNullOrWhiteSpace(propertyColumnName))
                {
                    cellIndex = columns.IndexOf(NormalizeColumnName(propertyColumnName));

                    if (cellIndex >= 0)
                    {
                        value = GetCellValue(cells.ElementAtOrDefault(cellIndex));

                        if (value != null)
                        {
                            SetPropertyValue(property, model, value);
                        }
                    }
                }
            }

            return model;
        }

        private ISheet GetWorksheet(IWorkbook workbook, string worksheetName)
        {
            if (workbook is XSSFWorkbook)
            {
                XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
            }
            else
            {
                HSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
            }

            return string.IsNullOrWhiteSpace(worksheetName) || !HasWorksheet(workbook, worksheetName)
                ? GetFirstWorksheet(workbook) : workbook.GetSheet(worksheetName);
        }

        private ISheet GetFirstWorksheet(IWorkbook workbook)
        {
            return workbook.GetSheetAt(0);
        }

        private bool HasWorksheet(IWorkbook workbook, string worksheetName)
        {
            try
            {
                return workbook.GetSheetIndex(worksheetName) >= 0;
            }
            catch { return false; }
        }

        private IEnumerable<string> GetColumns(ISheet worksheet)
        {
            return GetCells(worksheet.GetRow(worksheet.FirstRowNum)).Select(c => NormalizeColumnName(GetCellValue(c)));
        }

        private IEnumerable<IRow> GetRows(ISheet worksheet)
        {
            var startRowNumber = worksheet.FirstRowNum + 1;
            var lastRowNumber = worksheet.LastRowNum > startRowNumber ? worksheet.LastRowNum : startRowNumber;

            for (int rowNumber = startRowNumber; rowNumber <= lastRowNumber; rowNumber++)
            {
                var row = worksheet.GetRow(rowNumber);

                if (row != null)
                {
                    yield return row;
                }
            }
        }

        private IEnumerable<ICell> GetCells(IRow row)
        {
            var lastCellNumber = row?.LastCellNum ?? -1;

            for (int cellNumber = 0; cellNumber <= lastCellNumber; cellNumber++)
            {
                yield return row.GetCell(cellNumber);
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

        private void SetPropertyValue<TModel>(PropertyInfo property, TModel model, string value)
        {
            if (IsNumericType(property.PropertyType))
            {
                value = new string((value as string).ToCharArray().Where(c => !char.IsWhiteSpace(c)).ToArray()).Replace(",", ".");
            }

            property.SetValue(model, Convert.ChangeType(value, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType, new CultureInfo("en-US")));
        }

        private string NormalizeColumnName(string str)
        {
            return (str ?? string.Empty).Trim().ToUpper();
        }
    }
}
