using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelProvider.Helpers
{
    public static class ValidationHelper
    {
        public static void ValidateModel<TModel>(this TModel model)
        {
            var validationContext = new ValidationContext(model, serviceProvider: null, items: null);
            var validationResults = new List<ValidationResult>();
            var isValid = Validator.TryValidateObject(model, validationContext, validationResults, true);

            if (!isValid)
            {
                throw new ArgumentException("Model is not valid: " + string.Join(", ", validationResults.Select(s => s.ErrorMessage).ToArray()));
            }
        }
    }
}
