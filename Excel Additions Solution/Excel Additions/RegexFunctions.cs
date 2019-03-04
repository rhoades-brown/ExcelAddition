using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace Excel_Additions
{
    public static class RegexFunctions
    {
        //The following strings are used by multiple function descriptions.
        const string Category = "Lookup & Reference";

        const string RegexDescription = "is the regular expression for the match. This does not require start/end slashes and uses the .NET Framework regular expression format.";
        const string Within_textDescription = "is the text containing the text to compare against the expression";
        const string IgnoreCaseDescription = "if TRUE ignores the case of within_text, if FALSE within_text is treated case sensitive. Default is false";
        const string MultiLineDescription = "if TRUE use multiline mode, where ^ and $ match the beginning and end of each line (instead of the beginning and end of the input string). Default is false";
        const string SingleLineDescription = "if true use single-line mode, where the period (.) matches every character (instead of every character except \n). Default is false";

        private static RegexOptions SetRegexOptions(bool ignore_case, bool Multiline, bool Singleline)
        {
            var regexOptions = new RegexOptions();
            if (ignore_case)
                regexOptions |= RegexOptions.IgnoreCase;
            if (Multiline)
                regexOptions |= RegexOptions.Multiline;
            if (Singleline)
                regexOptions |= RegexOptions.Singleline;
            return regexOptions;
        }

        [ExcelFunction(
            Description = "Looks if the regular expression provided matches a string and returns a boolean",
            Category = Category
        )]
        public static bool REGEXISMATCH(
            [ExcelArgument(Name = "regular_expression", Description = RegexDescription)]string regular_expression,
            [ExcelArgument(Name = "within_text", Description = Within_textDescription)]string within_text,
            [ExcelArgument(Name = "[ignore_case]", Description = IgnoreCaseDescription)]bool ignore_case,
            [ExcelArgument(Name = "[multi_line]", Description = MultiLineDescription)]bool Multiline,
            [ExcelArgument(Name = "[single_line]", Description = SingleLineDescription)]bool Singleline
            )
        {
            return Regex.IsMatch(
                within_text,
                regular_expression,
                SetRegexOptions(ignore_case, Multiline, Singleline));
        }

        [ExcelFunction(
            Description = "Looks if the regular expression provided matches a string and returns the match",
            Category = Category
        )]
        public static object REGEXMATCH(
            [ExcelArgument(Name = "regular_expression", Description = RegexDescription)] string regular_expression,
            [ExcelArgument(Name = "within_text", Description = Within_textDescription)]string within_text,
            [ExcelArgument(Name = "[ignore_case]", Description = IgnoreCaseDescription)]bool ignore_case,
            [ExcelArgument(Name = "[multi_line]", Description = MultiLineDescription)]bool Multiline,
            [ExcelArgument(Name = "[single_line]", Description = SingleLineDescription)]bool Singleline
            )
        {
            Match Result = Regex.Match(
                     within_text,
                     regular_expression,
                     SetRegexOptions(ignore_case, Multiline, Singleline)
                     );

            return Result.Success ? Result.Value : (object)ExcelError.ExcelErrorNA;
        }

        [ExcelFunction(
            Description = "Looks if the regular expression provided matches a string and returns the match",
            Category = "Lookup & Reference"
        )]
        public static object REGEXREPLACE(
            [ExcelArgument(Name = "regular_expression", Description = RegexDescription)] string regular_expression,
            [ExcelArgument(Name = "within_text", Description = Within_textDescription)]string within_text,
            [ExcelArgument(Name = "replace_text", Description = "is the text to replace with")]string replace_text,
            [ExcelArgument(Name = "[ignore_case]", Description = IgnoreCaseDescription)]bool ignore_case,
            [ExcelArgument(Name = "[multi_line]", Description = MultiLineDescription)]bool Multiline,
            [ExcelArgument(Name = "[single_line]", Description = SingleLineDescription)]bool Singleline
            )
        {
            var Result = Regex.Replace(
                within_text,
                regular_expression,
                replace_text,
                SetRegexOptions(ignore_case, Multiline, Singleline)
                );

            return Result.Length > 0 ? Result : (object)ExcelError.ExcelErrorNA;
        }

    }
}
