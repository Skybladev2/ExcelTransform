using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace XmlSerializationResearch
{
    /// <summary>
    /// Обрабатывает все или указанный список столбцов
    /// </summary>
    public class TransformRangeCommand : TransformCommand
    {
        // если колонки пусты, то заменяем везде, кроме заголовков
        [XmlArrayItem("Column")]
        public List<string> Source;

        [XmlIgnoreAttribute]
        public bool SourceSpecified;

        protected List<object> sourceColumnIdentifiers = new List<object>();

        protected override bool CheckDeserialization()
        {
            uint number = 0;

            foreach (string column in Source)
            {
                if (uint.TryParse(column, out number))
                {
                    sourceColumnIdentifiers.Add(number);
                }
                else
                {
                    Regex regex = new Regex("[a-zA-Z]+");
                    Match match = regex.Match(column);

                    if (match.Success)
                    {
                        sourceColumnIdentifiers.Add(column);
                    }
                    else
                        return false;
                }
            }

            return true;
        }

        protected virtual void FindBoundsAndProcess()
        {
            uint headerRowIndex = GetStartRowBound();
            uint lastRowIndex = GetEndRowBound(headerRowIndex);

            // если указаны колонки - обрабатываем их, иначе все ячейки с данными
            if (SourceSpecified)
                foreach (object columnIdentifier in sourceColumnIdentifiers)
                {
                    ProcessRows(columnIdentifier);
                }
            else
            {
                // надеемся, что какие-нибудь данные всё же есть :)
                GetLastColumnIndex(headerRowIndex, lastRowIndex);
                GetFirstNonEmptyColumn();

                for (uint i = this.startColumnBound.Value; i < this.endColumnBound; i++)
                {
                    ProcessRows(i);
                }
            }
        }

        protected virtual void ProcessRows(object columnIdentifier)
        {
            for (uint i = this.startRowBound.Value + 1; i <= endRowBound; i++)
            {
                Range cell = worksheet.Cells[i, columnIdentifier];

                string cellValue = null;
                //cell.NumberFormat = "@";
                if (cell.NumberFormatLocal == "ДД.ММ.ГГГГ")
                {
                    cellValue = cell.Text;
                    cell.NumberFormatLocal = "@";
                }
                else
                {
                    cellValue = cell.Value == null ? "" : cell.Text.ToString();
                    cell.NumberFormatLocal = "@";
                }

                //string cellValue = cell.Value == null ? "" : cell.Value.ToString();
                //string cellValue = cell.Value == null ? "" : cell.Text.ToString();
                cell.Value = ProcessCell(cellValue);
            }
        }

        protected virtual string ProcessCell(string cellValue)
        {
            throw new NotImplementedException();
        }
    }
}
