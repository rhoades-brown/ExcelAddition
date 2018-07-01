using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace Excel_Additions
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            RegisterFunctions();
        }

        public void AutoClose()
        {
        }

        public void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions()
                             .Select(UpdateHelpTopic)
                             .RegisterFunctions();
        }

        public ExcelFunctionRegistration UpdateHelpTopic(ExcelFunctionRegistration funcReg)
        {
            funcReg.FunctionAttribute.HelpTopic = "http://www.bing.com";
            return funcReg;
        }
    }

    public class Functions
    {
        [ExcelFunction(HelpTopic = "http://www.google.com")]
        public static object SayHello()
        {
            return "Hello!!!";
        }
    }
}
