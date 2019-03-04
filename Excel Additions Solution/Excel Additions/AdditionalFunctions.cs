using ExcelDna.Integration;
using System;
using System.Collections;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace Excel_Additions
{
    public static class AdditionalFunctions
    {
        #region Helper Functions

        private static void ProcessItem(ArrayList resultsArray, object item)
        {
            switch (item.GetType().ToString())
            {
                case "System.Object[,]":
                    ProcessArray(resultsArray, (Object[,])item);
                    break;

                case "System.Object[]":
                case "System.Array":
                    ProcessItem(resultsArray, item);
                    break;

                case "System.String":
                    resultsArray.Add(item);
                    break;

                default:
                    resultsArray.Add(item.ToString());
                    break;
            }
        }

        private static void ProcessArray(ArrayList resultsArray, object[,] arrayObject)
        {
            for (int i = 0; i < arrayObject.GetLength(0); i++)
            {
                for (int j = 0; j < arrayObject.GetLength(1); j++)
                {
                    ProcessItem(resultsArray, arrayObject[i, j]);
                }
            }
        }

        private static bool LoopandReturnError(object[,] arrayObject)
        {
            for (int i = 0; i < arrayObject.GetLength(0); i++)
            {
                for (int j = 0; j < arrayObject.GetLength(1); j++)
                {
                    var x = arrayObject[i, j];
                    if (x.GetType() == typeof(ExcelError))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        #endregion Helper Functions

        [ExcelFunction(
            Description = "Joins a list of strings together, similar to concatenate, but with a seperator",
            Category = "Text",
            ExplicitRegistration = true
        )]
        public static string JOINSTRINGS(
            [ExcelArgument(
                Name = "seperator",
                Description = "is the seperator to join with"
            )]string separator,

            [ExcelArgument(
                Name = "values",
                Description = "are values or selections to be joined"
            )]params object[] values
        )
        {
            var resultsArray = new System.Collections.ArrayList();
            foreach (var value in values)
            {
                ProcessItem(resultsArray, value);
            }

            return string.Join(separator, resultsArray.ToArray());
        }

        [ExcelFunction(
            ExplicitRegistration = true
        )]
        public static bool TESTFORERRORS(
            [ExcelArgument(
                Name = "values",
                Description = "are values or selections to be joined"
            )]params object[] values
            )
        {
            Type[] arrayTypes = { typeof(object[,]), typeof(object[]), typeof(Array) };
            foreach (var value in values)
            {
                if (arrayTypes.Contains(value.GetType()))
                    if (!(LoopandReturnError(value as object[,])))
                        return true;
                    else if (value.GetType() == typeof(ExcelError))
                        return true;
            }

            return false;
        }

        [ExcelFunction(
            Description = "Returns the vesrion number of the module as a string",
            Category = "Text",
            IsVolatile = true
        )]
        public static bool TESTVERSIONNUMBER(string VersionNumber)
        {
            try
            {
                return Assembly.GetExecutingAssembly().GetName().Version >= Version.Parse(VersionNumber);
            }
            catch
            {
                return false;
            }
        }

        [ExcelFunction(Description = "Returns the vesrion number of the module as a string", Category = "Text", IsVolatile = true)]
        public static string GETVERSIONNUMBER() => Assembly.GetExecutingAssembly().GetName().Version.ToString();

        [ExcelFunction(Description = "A test function, used to return the alphabet in an array formula", Category = "Text", ExplicitRegistration = true)]
        public static object[] RETURNALPHABET() => Enumerable.Range('A', 26).Select(x => ((char)x).ToString()).ToArray();
    }
}