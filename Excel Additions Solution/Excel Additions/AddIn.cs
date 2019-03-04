using ExcelDna.Integration;
using ExcelDna.Registration;
using ExcelDna.IntelliSense;
using System.Linq;

namespace Excel_Additions
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
            RegisterFunctions();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        public void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions()
                             .ProcessParamsRegistrations()
                             .Select(UpdateHelpTopic)
                             .RegisterFunctions();
        }

        public ExcelFunctionRegistration UpdateHelpTopic(ExcelFunctionRegistration funcReg)
        {
            funcReg.FunctionAttribute.HelpTopic = "http://www.bing.com";
            return funcReg;
        }
    }
}