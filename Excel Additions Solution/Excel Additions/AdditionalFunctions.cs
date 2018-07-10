using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Additions
{
    public static class AdditionalFunctions
    {
        #region Helper Functions
        private static void ProcessItem(System.Collections.ArrayList resultsArray, object item)
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

        static void ProcessArray(System.Collections.ArrayList resultsArray, object[,] arrayObject)
        {
            for (int i = 0; i < arrayObject.GetLength(0); i++)
            {
                for (int j = 0; j < arrayObject.GetLength(1); j++)
                {
                    ProcessItem(resultsArray, arrayObject[i, j]);
                }
            }
        }
        #endregion

        [ExcelFunction(
            Description = "Joins a list of strings together, similar to concatenate, but with a seperator",
            Category = "Text"
        )]
        public static string JoinStrings(
            [ExcelArgument(Name = "seperator", Description = "is the seperator to join with")]string separator,
            [ExcelArgument(Name = "values", Description = "are values or selections to be joined")]params object[] values)
        {
            var resultsArray = new System.Collections.ArrayList();
            foreach (var value in values)
            {
                ProcessItem(resultsArray, value);
            }

            return String.Join(separator, resultsArray.ToArray());
        }
    }
    }
}
