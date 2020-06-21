using NPOI.SS.UserModel;
using NPOIWrapper.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIWrapper.Excel
{
    public class SpreadSheetICell
    {
        private ICell cell;
        private Logger Logger = new Logger("SpreadSheetICell");

        public SpreadSheetICell(ICell cell)
        {
            this.cell = cell;
        }

        /// <summary>
        /// Return cell value to string
        /// </summary>
        /// <returns>
        /// if value exists return cell value in string otherwise return null
        /// </returns>
        public string GetStringValue()
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                if (cell.CellType == CellType.Blank)
                    return null;
                else
                {
                    return cell.StringCellValue;
                }
            }catch(NullReferenceException e)
            {
                Logger.Error(e, "GetValue : NullReference");
                return null;
            }
            catch (Exception e)
            {
                Logger.Error(e, "GetValue +> row: " + cell.Row + " col: " + cell.ColumnIndex);
                return null;
            }
        }

        /// <summary>
        /// Return cell value to richtextstring
        /// </summary>
        /// <returns>
        /// if value exists return cell value in IRichTextString otherwise return null
        /// </returns>
        public IRichTextString GetRichSTextValue()
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                if (cell.CellType == CellType.Blank)
                    return null;
                else
                {
                    return cell.RichStringCellValue;
                }
            }
            catch (NullReferenceException e)
            {
                Logger.Error(e, "GetValue : NullReference");
                return null;
            }
            catch (Exception e)
            {
                Logger.Error(e, "GetValue +> row: " + cell.Row + " col: " + cell.ColumnIndex);
                return null;
            }
        }

        /// <summary>
        /// Return cell value to double
        /// </summary>
        /// <returns>
        /// if value exists return cell value in double otherwise return 0
        /// </returns>
        public double GetDoubleValue()
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                if (cell.CellType == CellType.Blank)
                    return 0d;
                else
                {
                    return cell.NumericCellValue;
                }
            }
            catch (NullReferenceException e)
            {
                Logger.Error(e, "GetValue : NullReference");
                return 0d;
            }
            catch (Exception e)
            {
                Logger.Error(e, "GetValue +> row: " + cell.Row + " col: " + cell.ColumnIndex);
                return 0d;
            }
        }

        /// <summary>
        /// Return cell value to bool
        /// </summary>
        /// <returns>
        /// return true if cell in boolean type otherwise return false
        /// </returns>
        public bool GetBoolValue()
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                if (cell.CellType == CellType.Blank)
                    return false;
                else
                {
                    return cell.BooleanCellValue;
                }
            }
            catch (NullReferenceException e)
            {
                Logger.Error(e, "GetValue : NullReference");
                return false;
            }
            catch (Exception e)
            {
                Logger.Error(e, "GetValue +> row: " + cell.Row + " col: " + cell.ColumnIndex);
                return false;
            }
        }

        /// <summary>
        /// Return cell value in datetime object
        /// </summary>
        /// <returns>
        /// If value exiest in datetime type return value otherwise return new DateTime()
        /// Check if DateTime has exists with [dateTime.HasValue]
        /// </returns>
        public DateTime GetDateTimeValue()
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                if (cell.CellType == CellType.Blank)
                    return new DateTime();
                else
                {
                    return cell.DateCellValue;
                }
            }
            catch (NullReferenceException e)
            {
                Logger.Error(e, "GetValue : NullReference");
                return new DateTime();
            }
            catch (Exception e)
            {
                Logger.Error(e, "GetValue +> row: " + cell.Row + " col: " + cell.ColumnIndex);
                return new DateTime();
            }
        }

        /// <summary>
        /// Return non null or non blank value to string
        /// </summary>
        /// <returns>
        /// return empty string ("") if cell type is blank otherwise return null
        /// </returns>
        public string GetValue()
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Numeric:
                        return Convert.ToString(cell.NumericCellValue);
                    case CellType.Boolean:
                        return Convert.ToString(cell.BooleanCellValue);
                    case CellType.Blank:
                        return "";
                    case CellType.Error:
                        return Convert.ToString(cell.ErrorCellValue);
                    default:
                        return null;
                }
            }
            catch (NullReferenceException e)
            {
                Logger.Error(e, "GetValue : NullReference");
                return null;
            }
            catch (Exception e)
            {
                Logger.Error(e, "GetValue +> row: " + cell.Row + " col: " + cell.ColumnIndex);
                return null;
            }
        }

        public void SetValue(string value)
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                cell.SetCellValue(value);
            }catch (Exception e)
            {
                Logger.Error(e, "SetValue with value: " + value);
            }
        }

        public void SetValue(bool value)
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                cell.SetCellValue(value);
            }
            catch (Exception e)
            {
                Logger.Error(e, "SetValue with value: " + Convert.ToString(value));
            }
        }

        public void SetValue(DateTime value)
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                cell.SetCellValue(value);
            }
            catch (Exception e)
            {
                Logger.Error(e, "SetValue with value: " + value.ToString());
            }
        }

        public void SetValue(IRichTextString value)
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                cell.SetCellValue(value);
            }
            catch (Exception e)
            {
                Logger.Error(e, "SetValue with value: " + value.String);
            }
        }

        public void SetValue(double value)
        {
            try
            {
                if (cell == null)
                    throw new Exception("null cell");
                cell.SetCellValue(value);
            }
            catch (Exception e)
            {
                Logger.Error(e, "SetValue with value: " + Convert.ToString(value));
            }
        }

        public ICell GetICell()
        {
            return cell;
        }
    }
}
