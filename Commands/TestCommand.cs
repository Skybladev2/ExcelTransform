using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace XmlSerializationResearch
{
    public class TestCommand : TransformCommand, ITransformCommand
    {
        public string Column = null;

        private object columnIdentifier;

        #region ITransformCommand Members

        public void CastColumnIdentifier()
        {
            uint number = 0;

            if (uint.TryParse(Column, out number))
                columnIdentifier = number;
            else
                columnIdentifier = Column;
        }

        public TransformCommand Execute(TransformCommand previousCommand)
        {
            Range cell = worksheet.Cells[3, columnIdentifier];
            cell.Value = "ABCD";

            return null;
        }

        #endregion
    }
}
