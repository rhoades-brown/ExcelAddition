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
        [ExcelFunction(
            ExplicitRegistration = true,
            Description = "My first .NET function"
        )]
        public static string JOIN(
            [ExcelArgument(Name = "Delimiter", Description = "The delimiter between the objects")]string delimiter,            
            [ExcelArgument(Name = "Value", Description = "gives the Rest")]params object[] args
            )
        {
            var arrayResult = new System.Collections.ArrayList();
            foreach (var arg in args)
            {
                if (!(arg is ExcelEmpty)) {
                    arrayResult.Add(arg.ToString());
                }
            }
            return String.Join(delimiter, arrayResult.ToArray());
        }
    }
}
