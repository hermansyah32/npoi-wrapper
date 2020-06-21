using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using NPOIWrapper.Util;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIWrapper.Excel
{

    public class SpreadSheetWrapper
    {
        
        private SpreadSheetType spreadSheetType;
        private IWorkbook workbook;
        private string sourceFile;
        private Logger Logger = new Logger("SpreadSheetWrapper");

        private FileMode fileMode = FileMode.Open;
        private FileAccess fileAccess = FileAccess.ReadWrite;

        public SpreadSheetWrapper(SpreadSheetType spreadSheetType)
        {
            this.spreadSheetType = spreadSheetType;
            switch (this.spreadSheetType)
            {
                case SpreadSheetType.XLS:
                    workbook = new HSSFWorkbook();
                    break;
                case SpreadSheetType.XLSX:
                    workbook = new XSSFWorkbook();
                    break;

                default:
                    Logger.Debug("Unknown File Format");
                    break;
            }
        }

        /// <summary>
        /// Open spreadsheet document with devaul setting
        /// need to close spreadsheet document before read file
        /// </summary>
        /// <param name="sourceFile">Source file to be read</param>
        public SpreadSheetWrapper(string sourceFile)
        {
            this.sourceFile = sourceFile;
            try
            {
                switch (Path.GetExtension(this.sourceFile))
                {
                    case ".xls":
                        workbook = new HSSFWorkbook(new FileStream(this.sourceFile, fileMode, fileAccess));
                        break;
                    case ".xlsx":
                        workbook = new XSSFWorkbook(this.sourceFile);
                        break;
                    default:
                        throw new Exception("Unknown file type");
                }
            }catch(Exception e)
            {
                Logger.Error(e, "SpreadSheetWrapper");   
            }
        }

        /// <summary>
        /// Open spreadsheet document with file stream
        /// </summary>
        /// <param name="sourceFile">Source file to be read</param>
        /// <param name="fileMode">File mode (Open, Append, etc...)</param>
        /// <param name="fileAccess">File Accread (Read, ReadWrite, etc..)</param>
        public SpreadSheetWrapper(string sourceFile, FileMode fileMode, FileAccess fileAccess)
        {
            this.sourceFile = sourceFile;
            this.fileMode = fileMode;
            this.fileAccess = fileAccess;
            try
            {
                switch (Path.GetExtension(this.sourceFile))
                {
                    case ".xls":
                        workbook = new HSSFWorkbook(new FileStream(this.sourceFile, this.fileMode, this.fileAccess));
                        break;
                    case ".xlsx":
                        workbook = new XSSFWorkbook(new FileStream(this.sourceFile, this.fileMode, this.fileAccess));
                        break;
                    default:
                        throw new Exception("Unknown file type");
                }
            }
            catch (Exception e)
            {
                Logger.Error(e, "SpreadSheetWrapper");
            }
        }

        public SpreadSheetISheet CreateSheet(string sheetName = "Sheet1")
        {
            try
            {
                return new SpreadSheetISheet(workbook.CreateSheet(sheetName));
            }catch(Exception e)
            {
                Logger.Error(e, "CreateSheet");
                return null;
            }
        }

        public bool IsSheetExist (string sheetName)
        {
            try
            {
                if (workbook.GetSheet(sheetName) != null)
                    return true;
                else
                    return false;
            }catch(Exception e)
            {
                Logger.Error(e, "IsSheetExist");
                return false;
            }
        }

        public SpreadSheetISheet GetSheet()
        {
            try
            {
                if (workbook.NumberOfSheets > 0)
                {
                    ISheet sheet = workbook.GetSheetAt(workbook.ActiveSheetIndex);
                    MaxRow = sheet.LastRowNum;
                    return new SpreadSheetISheet(sheet);
                }
                else
                    return new SpreadSheetISheet(workbook.CreateSheet("Sheet" + Convert.ToString(workbook.NumberOfSheets + 1)));
            }catch(Exception e)
            {
                Logger.Error(e, "GetSheet");
                return null;
            }
        }

        public SpreadSheetISheet GetSheet(int index)
        {
            try
            {
                return new SpreadSheetISheet(workbook.GetSheetAt(index));
            }
            catch (Exception e)
            {
                Logger.Error(e, "GetSheet");
                return null;
            }
        }

        public SpreadSheetISheet GetSheet(string sheetName)
        {
            try
            {
                return new SpreadSheetISheet(workbook.GetSheet(sheetName));
            }
            catch (Exception e)
            {
                Logger.Error(e, "GetSheet");
                return null;
            }
        }

        public IWorkbook GetIWorkbook()
        {
            return workbook;
        }

        public void Save()
        {
            try
            {
                if (sourceFile == null)
                    throw new Exception("Please specify destination path");
                if (fileMode == FileMode.Open && fileAccess == FileAccess.Read)
                    throw new Exception("Can't save to read-only file access");
                FileStream fileStream = new FileStream(sourceFile, FileMode.Create);
                workbook.Write(fileStream);
                fileStream.Close();
            }catch(Exception e)
            {
                Logger.Error(e, "Save");
            }
        }

        public void SaveAs(string destination)
        {
            try
            {
                FileStream fileStream = new FileStream(destination, FileMode.Create);
                workbook.Write(fileStream);
                fileStream.Close();
            }
            catch (Exception e)
            {
                Logger.Error(e, "SaveAs");
            }
        }

        /**
         * Property section
         */
        public int MaxRow { get; set; } = 0;
    }
}
