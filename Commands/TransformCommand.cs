using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;

namespace XmlSerializationResearch
{
    public class TransformCommand
    {
        protected string Warning;
        protected string Error;
        protected string StackTrace;

        protected uint? startColumnBound = null;
        protected uint? endColumnBound = null;
        protected uint? startRowBound = null;
        protected uint? endRowBound = null;

        protected Microsoft.Office.Interop.Excel.Application excel = null;
        protected string filePath = null;
        protected Workbook workbook = null;
        protected int sheetNumber = 1;

        public void SetExcelApp(Microsoft.Office.Interop.Excel.Application excel)
        {
            this.excel = excel;
            this.excel.Visible = true;
        }

        public void SetFilePath(string filePath)
        {
            this.filePath = filePath;
        }

        protected Worksheet worksheet
        {
            get
            {
                if (workbook == null)
                    workbook = excel.Workbooks.Open(filePath);

                Sheets sheets = workbook.Worksheets;
                return (Worksheet)sheets.get_Item(sheetNumber);
            }
        }

        protected void Init(TransformCommand another)
        {
            if (another != null)
            {
                this.endColumnBound = another.endColumnBound;
                this.endRowBound = another.endRowBound;
                this.excel = another.excel;
                this.sheetNumber = another.sheetNumber;
                this.startColumnBound = another.startColumnBound;
                this.startRowBound = another.startRowBound;
                this.workbook = another.workbook;
            }
        }

        protected virtual bool CheckDeserialization()
        {
            throw new NotImplementedException();
        }

        protected object GetNextColumn(object column)
        {
            if (column.GetType() == typeof(uint))
                return (uint)column + 1;
            else
                return GetNextBase26(column.ToString().ToLowerInvariant());
        }

        #region http://stackoverflow.com/questions/1011732/iterating-through-the-alphabet-c-a-caz/1011912#1011912

        private static string GetNextBase26(string a)
        {
            return Base26Sequence().SkipWhile(x => x != a).Skip(1).First();
        }

        private static IEnumerable<string> Base26Sequence()
        {
            long i = 0L;
            while (true)
                yield return Base26Encode(i++);
        }

        private static char[] base26Chars = "abcdefghijklmnopqrstuvwxyz".ToCharArray();
        private static string Base26Encode(Int64 value)
        {
            string returnValue = null;
            do
            {
                returnValue = base26Chars[value % 26] + returnValue;
                value /= 26;
            } while (value-- != 0);
            return returnValue;
        }

        /// <summary>
        /// Возвращает число, закодированное строкой английского алфавита.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        protected static uint Base26Decode(string value)
        {
            uint index = 0;
            int digit = 0;
            string lowerValue = value.ToLowerInvariant();
            for (int i = lowerValue.Length - 1; i >= 0; i--)
            {
                char c = lowerValue[i];
                index += (uint)((Array.IndexOf<char>(base26Chars, c) + 1) * Math.Pow(26, digit++));
            }

            return index;
        }
        #endregion

        protected ushort GetFirstNonEmptyColumn()
        {
            // определяем первый непустой столбец
            ushort column = 1;

            while (excel.WorksheetFunction.CountA(worksheet.Columns[column].EntireColumn) == 0)
            {
                column++;
            }

            startColumnBound = column;
            return column;
        }

        protected ushort GetFirstNonEmptyRowHeader()
        {
            // определяем первую непустую строку
            ushort row = 1;

            while (excel.WorksheetFunction.CountA(worksheet.get_Range("A" + row.ToString(), "A" + row.ToString()).EntireRow) == 0)
            {
                row++;
            }

            startRowBound = row;
            return row;
        }

        protected int GetColumnIndex(Array values, string oldName)
        {
            for (int i = 1; i < values.Length; i++)
            {
                if (values.GetValue(1, i) != null)
                    if (values.GetValue(1, i).ToString() == oldName)
                        return i;
            }

            // не нашли
            return 0;
        }

        protected uint GetLastNonEmptyRowForward(uint startIndex)
        {
            uint row = startIndex;

            while (excel.WorksheetFunction.CountA(worksheet.get_Range("A" + row.ToString(), "A" + row.ToString()).EntireRow) != 0)
            {
                row++;
            }

            return (uint)(row - 1);
        }

        protected uint GetIndexForNewColumn(uint headerRowIndex, uint lastRowIndex)
        {
            uint columnIndex = 256;

            if (this.endColumnBound == null)
            {
                while (excel.WorksheetFunction.CountA(GetColumnRange(worksheet, headerRowIndex, lastRowIndex, columnIndex)) == 0)
                {
                    columnIndex--;
                }

                this.endColumnBound = columnIndex + 1;
            }
            else
                this.endColumnBound++;

            return this.endColumnBound.Value;
        }

        protected uint GetLastColumnIndex(uint headerRowIndex, uint lastRowIndex)
        {
            uint columnIndex = 256;

            if (this.endColumnBound == null)
            {
                while (excel.WorksheetFunction.CountA(GetColumnRange(worksheet, headerRowIndex, lastRowIndex, columnIndex)) == 0)
                {
                    columnIndex--;
                }

                this.endColumnBound = columnIndex;
            }

            return this.endColumnBound.Value;
        }

        protected Range GetColumnRange(Worksheet worksheet, uint headerRowIndex, uint lastRowIndex, uint columnIndex)
        {
            Range range1 = worksheet.Cells[headerRowIndex, columnIndex];
            Range range2 = worksheet.Cells[lastRowIndex, columnIndex];
            Range cells = worksheet.Range[range1, range2];

            return cells;
        }

        protected uint GetStartRowBound()
        {
            if (this.startRowBound == null)
            {
                this.startRowBound = GetFirstNonEmptyRowHeader();
            }

            return this.startRowBound.Value;
        }

        protected uint GetEndRowBound(uint headerRowIndex)
        {
            if (this.endRowBound == null)
            {
                this.endRowBound = GetLastNonEmptyRowForward(headerRowIndex);
            }

            return this.endRowBound.Value;
        }
    }
}
