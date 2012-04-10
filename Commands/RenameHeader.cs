using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Excel;

namespace XmlSerializationResearch
{
    public class RenameHeader : TransformCommand, ITransformCommand
    {
        [XmlArrayItem ("Item")]
        public List<string> OldName;

        public string NewName;

        #region ITransformCommand Members

        public TransformCommand Execute(TransformCommand previousCommand)
        {
            Init(previousCommand);

            uint row = GetFirstNonEmptyRowHeader();

            Array values = (Array)worksheet.get_Range("A" + row.ToString()).EntireRow.Cells.Value;

            foreach (string oldName in OldName)
            {
                int column = GetColumnIndex(values, oldName);

                if (column != 0)
                {
                    Range cell = worksheet.Cells[row, column];
                    cell.Value = NewName;
                }
            }

            return this;
        }

        #endregion
    }
}
