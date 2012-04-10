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
    /// Обрабатывает один столбец
    /// </summary>
    public class TransformColumnCommand : TransformCommand
    {
        public string Column;

        protected object columnIdentifier;

        protected override bool CheckDeserialization()
        {
            uint number = 0;

            if (uint.TryParse(Column, out number))
            {
                columnIdentifier = number;
                return true;
            }
            else
            {
                Regex regex = new Regex("[a-zA-Z]+");
                Match match = regex.Match(Column);

                if (match.Success)
                {
                    columnIdentifier = Column;
                    return true;
                }

                return false;
            }
        }

        protected virtual void FindBoundsAndProcess()
        {
            uint headerRowIndex = GetStartRowBound();
            uint lastRowIndex = GetEndRowBound(headerRowIndex);

            ProcessRows();
        }

        protected virtual void ProcessRows()
        {
            for (uint i = this.startRowBound.Value + 1; i <= endRowBound; i++)
            {
                Range cell = worksheet.Cells[i, columnIdentifier];
                string cellValue = cell.Value == null ? "" : cell.Text.ToString();
                cell.Value = ProcessCell(i, cellValue);
            }
        }

        protected virtual string ProcessCell(uint row, string cellValue)
        {
            throw new NotImplementedException();
        }
    }
}
