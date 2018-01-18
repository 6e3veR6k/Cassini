using System;
using System.Globalization;
using System.Windows.Controls;

namespace Cassini.UI.ViewModel
{
    public class ValidationDate: ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            DateTime selectedDate;
            if (!DateTime.TryParse((value ?? "").ToString(),
                CultureInfo.CurrentCulture,
                DateTimeStyles.AssumeLocal | DateTimeStyles.AllowWhiteSpaces,
                out selectedDate)) return new ValidationResult(false, "Не вірна дата");

            return selectedDate.Date.Month > DateTime.Now.Date.Month
                ? new ValidationResult(false, "Не вірний період")
                : ValidationResult.ValidResult;
        }
    }
}   