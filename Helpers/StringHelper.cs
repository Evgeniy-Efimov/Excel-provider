using System.IO;
using System.Linq;

namespace ExcelProvider.Helpers
{
    public static class StringHelper
    {
        public static string NormalizeForComparing(this string str)
        {
            return string.Concat((str ?? string.Empty).Where(ch => !char.IsWhiteSpace(ch))).ToUpper();
        }

        public static string FormatFileName(this string str)
        {
            return string.Concat(string.Concat((str ?? string.Empty)
                .Select(ch => char.IsWhiteSpace(ch) ? "_" : ch.ToString()))
                .Split(Path.GetInvalidFileNameChars()));
        }
    }
}
