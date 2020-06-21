using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOIWrapper.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace NPOIWrapper.Excel
{
    public class SpreadSheetISheet
    {
        private Logger Logger = new Logger("SpreadSheetISheet");


        private ISheet sheet;
        // TODO: it called indexer

        public SpreadSheetISheet (ISheet sheet)
        {
            this.sheet = sheet;
        }

        public ISheet GetISheet()
        {
            return sheet;
        }

        public SpreadSheetICell Cell(string address, bool createIfNotExists)
        {
            CellReference reference = new CellReference(address);
            try
            {
                IRow row = sheet.GetRow(reference.Row);
                if (row == null)
                    if (createIfNotExists)
                        row = sheet.CreateRow(reference.Row);
                    else
                        throw new Exception("Row doesn't exists");
                ICell cell = row.GetCell(reference.Col, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (cell == null)
                    if (createIfNotExists)
                        cell = row.CreateCell(reference.Col);
                    else
                        throw new Exception("Cell doesn't exists");
                return new SpreadSheetICell(cell);
            }catch(Exception e)
            {
                Logger.Error(e, "Cell");
                return new SpreadSheetICell(null);
            }
        }

        
    }
}
