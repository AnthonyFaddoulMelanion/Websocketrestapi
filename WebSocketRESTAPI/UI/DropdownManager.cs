using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebSocketRESTAPI.UI
{
    public static class DropdownManager
    {
        public static void CreateSymbolDropdown(Range cell, List<string> symbols)
        {
            cell.Validation.Delete();

            List<string> limitedSymbols = symbols.GetRange(0, System.Math.Min(100, symbols.Count));
            string formulaList = string.Join(",", limitedSymbols);

            cell.Validation.Add(
                XlDVType.xlValidateList,
                XlDVAlertStyle.xlValidAlertStop,
                XlFormatConditionOperator.xlBetween,
                formulaList);

            cell.Validation.IgnoreBlank = true;
            cell.Validation.InCellDropdown = true;
        }
    }

}
