using NPOIWrapper.Excel;
using NPOIWrapper.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIWrapper
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                SpreadSheetWrapper xlsWrapper, xlsxWrapper;
                // create excel document
                xlsWrapper = new SpreadSheetWrapper(SpreadSheetType.XLS);
                SpreadSheetISheet xlsSheet = xlsWrapper.CreateSheet();
                xlsSheet.Cell("A1", true).SetValue("Text ini ditulis dari program");
                xlsWrapper.SaveAs("Tes Excel 2003.xls");

                xlsxWrapper = new SpreadSheetWrapper(SpreadSheetType.XLSX);
                SpreadSheetISheet xlsxSheet = xlsxWrapper.CreateSheet();
                xlsxSheet.Cell("A1", true).SetValue("Text ini ditulis dari program");
                xlsxWrapper.SaveAs("Tes Excel 2007.xlsx");

                // read excel document
                xlsWrapper = new SpreadSheetWrapper("Tes Excel 2003.xls");
                Console.WriteLine("A1 (2003) value => " + xlsWrapper.GetSheet().Cell("A1", false).GetValue());
                xlsxWrapper = new SpreadSheetWrapper("Tes Excel 2007.xlsx");
                Console.WriteLine("A1 (2007) value => " + xlsxWrapper.GetSheet().Cell("A1", false).GetValue());
                // write other text
                xlsWrapper.GetSheet().Cell("A2", true).SetValue("Ini text keduanya");
                xlsWrapper.SaveAs("Tes Excel 2003 Edited.xls");
                xlsxWrapper.GetSheet().Cell("A2", true).SetValue("Ini text keduanya");
                xlsxWrapper.SaveAs("Tes Excel 2007 Edited.xlsx");
            }
            catch(Exception e)
            {
                Console.WriteLine("Something wrong with the code");
            }


        }
    }
}
